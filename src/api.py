"""
Google GenAI API Handler Module
===============================

This module provides a high-level wrapper around the Google GenAI SDK (v1beta),
designed to facilitate robust interaction with Gemini models for the OS Agent.

It handles:
- Chat session initialization with custom tools and safety settings.
- Automatic retry logic for transient API errors (503, 429).
- Parsing of complex responses (Thought/Action pattern, Grounding metadata).
- Streaming capabilities for real-time agent feedback.

**IMPORTANT:**
Please refer to the `README.md` for installation instructions and required
environment variables (specifically `GOOGLE_API_KEY`) before running this code.

Author: Pio Borgelt
"""

import asyncio
import base64
import json
import time
import os
from typing import List, Optional, Dict, Any, AsyncGenerator

from google import genai
from google.genai import types

from src.utils import logger

class ApiHandler:
    def __init__(self, api_key: str, model_name: str = "gemini-2.5-flash-lite"):
        """
        Initializes the ApiHandler with the Google GenAI client.

        Args:
            api_key (str): The Google Cloud API Key.
            model_name (str): The default model identifier (e.g., 'gemini-2.5-flash-lite').
        
        Raises:
            ValueError: If the api_key is missing or empty.
        """
        if not api_key:
            logger.error("ApiHandler initialized without API Key.")
            raise ValueError("API Key is required.")
            
        self.api_key = api_key
        self.model_name = model_name
        

        self.client = genai.Client(api_key=self.api_key, http_options={"api_version": "v1beta"})
        logger.info(f"ApiHandler initialized with model: {self.model_name}")

    def create_chat_session(self, 
                            system_instruction: str = None, 
                            tool_definitions: List[Dict] = None, 
                            enable_google_search: bool = False,
                            model_override: str = None) -> Any:
        """
        Configures and creates a new chat session with specific capabilities.

        Disables standard safety settings to allow for code generation and administrative 
        tasks, and configures tool usage (Function Calling or Google Search).

        Args:
            system_instruction (str, optional): The system prompt defining the agent's behavior.
            tool_definitions (List[Dict], optional): A list of tool schemas for function calling.
            enable_google_search (bool): If True, enables the native Google Search tool.
            model_override (str, optional): Overrides the default model for this specific session.

        Returns:
            Any: A configured Google GenAI ChatSession object.
        """
        model = model_override or self.model_name
        

        safety_settings = [
            types.SafetySetting(category="HARM_CATEGORY_HATE_SPEECH", threshold="OFF"),
            types.SafetySetting(category="HARM_CATEGORY_DANGEROUS_CONTENT", threshold="OFF"),
            types.SafetySetting(category="HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold="OFF"),
            types.SafetySetting(category="HARM_CATEGORY_HARASSMENT", threshold="OFF"),
        ]

        tools_list = []
        tool_config = None

   
        if tool_definitions:
            funcs = [self._convert_tool_definition(t) for t in tool_definitions]
            if funcs:
                tools_list.append(types.Tool(function_declarations=funcs))
             
                tool_config = types.ToolConfig(
                    function_calling_config=types.FunctionCallingConfig(mode="AUTO")
                )

        if enable_google_search:
            tools_list.append(types.Tool(google_search=types.GoogleSearch()))

        config = types.GenerateContentConfig(
            temperature=0.1,
            max_output_tokens=2048,
            tools=tools_list if tools_list else None,
            tool_config=tool_config,
            system_instruction=system_instruction,
            safety_settings=safety_settings
        )

        return self.client.chats.create(model=model, config=config)

    def send_chat_message(self, chat_session, content_parts: List[Any], max_retries: int = 3) -> Dict[str, Any]:
        """
        Sends a message to the active chat session with automatic retry logic.

        Handles mixed content types (text and base64-encoded images) and catches 
        transient API errors (like overloads) with exponential backoff.

        Args:
            chat_session: The active GenAI ChatSession object.
            content_parts (List[Any]): A list containing text strings or image dictionaries.
            max_retries (int): Number of retry attempts for transient errors.

        Returns:
            Dict[str, Any]: A structured dictionary containing 'thought', 'actions', 
                            and 'grounding_info'.
        """
        retry_delay = 2
        
        for attempt in range(max_retries + 1):
            try:
             
                final_parts = []
                for part in content_parts:
                    if isinstance(part, dict) and "mime_type" in part:
                        final_parts.append(types.Part(
                            inline_data=types.Blob(mime_type=part['mime_type'], data=base64.b64decode(part['data']))
                        ))
                    elif isinstance(part, str):
                        final_parts.append(types.Part(text=part))
                    else:
                        final_parts.append(part)

                response = chat_session.send_message(final_parts)
                return self._parse_response(response)

            except Exception as e:
                error_msg = str(e).lower()
                is_transient = "503" in error_msg or "429" in error_msg or "overloaded" in error_msg
                
                if is_transient and attempt < max_retries:
                    wait_time = retry_delay * (2 ** attempt)
                    logger.warning(f"API Overload ({error_msg}). Retrying in {wait_time}s...")
                    time.sleep(wait_time)
                    continue
                
                logger.error(f"Fatal API Error: {e}")
                return {
                    "thought": json.dumps({
                        "status": "FAILED", 
                        "milestone_name": "API Error", 
                        "sprint_steps": [],
                        "grounding_notes": f"API Error: {e}"
                    }), 
                    "actions": []
                }

    async def generate_content_stream(self, 
                                      contents: List[Any], 
                                      tools: List[types.Tool] = None, 
                                      system_instruction: str = None,
                                      model_override: str = None) -> AsyncGenerator[Any, None]:

        model = model_override or self.model_name
        
        gen_config = types.GenerateContentConfig(
            temperature=0.1,
            max_output_tokens=1024,
            tools=tools,
            tool_config=types.ToolConfig(function_calling_config=types.FunctionCallingConfig(mode="AUTO")),
            system_instruction=system_instruction
        )

     
        response_stream = await self.client.aio.models.generate_content_stream(
            model=model,
            contents=contents,
            config=gen_config
        )

        async for chunk in response_stream:
            yield chunk

    def _convert_tool_definition(self, tool_def: Dict) -> types.FunctionDeclaration:

        return types.FunctionDeclaration(
            name=tool_def["name"],
            description=tool_def["description"],
            parameters=types.Schema(**tool_def["parameters"])
        )

    def _parse_response(self, response) -> Dict[str, Any]:

        result = {
            "thought": "",
            "actions": [],
            "grounding_info": []
        }

        try:
            if not response.candidates: return result
            candidate = response.candidates[0]

            if candidate.grounding_metadata:
                gm = candidate.grounding_metadata
                chunks = getattr(gm, 'grounding_chunks', None)
                if chunks:
                    for chunk in chunks:
                        web = getattr(chunk, 'web', None)
                        if web:
                            result["grounding_info"].append({
                                "title": getattr(web, 'title', 'Source'),
                                "url": getattr(web, 'uri', '#'),
                            })

            if candidate.content and candidate.content.parts:
                for part in candidate.content.parts:
                    if part.text:
                        result["thought"] += part.text + "\n"
                    
                    if part.function_call:
                        args = {}
                        if part.function_call.args:
                            try:
                                args = dict(part.function_call.args)
                            except (TypeError, ValueError):
                                args = {}
                        
                        result["actions"].append({
                            "command": part.function_call.name,
                            "parameters": args
                        })
        except Exception as e:
            logger.error(f"Response parsing error: {e}")
            if not result["thought"]:
                result["thought"] = f"Error parsing: {e}"
        
        return result