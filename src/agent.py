"""
OS Agent Core Logic
-------------------
This module implements the `OSAgent` class, acting as the autonomous "brain" 
for Windows OS interaction. It utilizes a multi-model architecture (Planner/Executor)
to perceive the screen, reason about tasks, and execute low-level inputs via IPC.

Capabilities:
- Visual Perception (Screenshots & Accessibility Tree)
- Strategic Planning (Breaking tasks into "Sprints")
- Execution (Mouse/Keyboard control, Fuzzy App Launching)
- Coding (Specialized sub-agent for file/script operations)

⚠️ IMPORTANT: 
Please read the README.md file before running this agent. 
Correct environment setup (Redis, IPC Server, Google API Key) is required.

Author: Pio Borgelt
"""

import asyncio
import base64
import json
import re
import time
import os
import traceback
import io
from multiprocessing import Queue
from typing import Any, Dict, List, Optional

import redis
from PIL import Image, ImageDraw, ImageFont
from google.genai import types


from src.api import ApiHandler
from src.utils import logger, CONFIG
from src.core import (
    get_all_windows_from_server,
    observe_os_state,
    execute_os_action,
    check_ipc_handle_exists,
    get_ipc_active_window,
    get_handle_rect,
    prepare_images_for_model,
    get_system_app_index,
    bring_window_to_front,
    read_local_file,
    write_to_local_file
)

class OSAgent:
    def __init__(self, redis_client: redis.Redis):
  
        api_key = CONFIG.get("GOOGLE_API_KEY")
        if not api_key:
            raise ValueError("GOOGLE_API_KEY missing in CONFIG")
        self.api = ApiHandler(api_key=api_key)

        self.redis_client = redis_client
        self.ipc_addr = CONFIG.get("IPC_SERVER_ADDRESS", ('localhost', 6000))
  
        auth = CONFIG.get("IPC_AUTHKEY", b'secret')
        self.ipc_key = auth.encode('utf-8') if isinstance(auth, str) else auth
        self.ipc_host = CONFIG.get("IPC_HOST")
        self.ipc_port = CONFIG.get("IPC_PORT")
        self.planner_model = "gemini-2.5-pro"
        self.executor_model = "gemini-2.5-flash-lite"
        self.coder_model = "gemini-2.5-pro"


        self.focus_stack = []     
        self.known_windows = {} 
        self.active_app_name = None 
        self.focus_handle = None
        self.focus_rect = None
        self.is_running = False
        self.step_count = 0
        self.current_elements = []
        

        self.current_sprint_plan = []
        self.last_sprint_result = "None. Starting fresh."
        self.grounding_context = ""
        self.app_index = get_system_app_index()
        

        self.planner_instruction = (
                    "### ⚠️ STRATEGY RULES\n"
                    "STOP MICROMANAGING!\n"
                    "1. **GROUP ACTIONS INTO LOGICAL FLOWS:** The Executor is smart. It can handle chains of 10-20 steps.\n"
                    "   - **BAD SPRINT:** [Launch 'Settings'] -> STOP. (Reason: Pointless interruption)\n"
                    "   - **BAD SPRINT:** [Click 'Personalization'] -> STOP. (Reason: Still not done)\n"
                    "   - **GOOD SPRINT:** [Launch 'Settings', Click 'Personalization', Click 'Colors', Change to 'Dark Mode', Verify result].\n"
                    "1. **DELEGATE TASKS, NOT CLICKS:** The Executor is smart. Don't micro-manage. \n"
                    "   - BAD: [Click Start, Type 'Calc', Click Enter]\n"
                    "   - GOOD: [Open Calculator and calculate 5*5]\n"
                    "2. **AVOID LOOPS:** If a previous plan failed or the error persists, DO NOT try the exact same steps again. Escalate or try a radical workaround (e.g. Force Quit, Reboot VM).\n"
                    "3. **VM MANAGEMENT:** If a VM throws errors, the solution is often to Power Off the VM via the main VirtualBox window, not clicking 'OK' inside the VM.\n"
                    "4. **MEMORY IS TRUTH (CRITICAL):**\n"
                    "   - Review the 'GROUNDING CONTEXT' (Ref Info) first.\n"
                    "   - If the error message is ALREADY recorded there (e.g., 'Windows cannot read setting...'), **DO NOT REPRODUCE IT AGAIN**.\n"
                    "   - **NEVER** start a sprint just to 'confirm' what is already in the notes. Move immediately to the FIX or RESEARCH phase.\n\n"
                    
                    "5. **GROUNDING NOTES USAGE:**\n"
                    "   - `grounding_notes` are for **PAST FACTS** and **DIAGNOSIS** only.\n"
                    "   -  BAD NOTE: 'The VM is off. The next step is to start it.' (This creates infinite loops!)\n"
                    "   - GOOD NOTE: 'Error [X] confirmed. VM is powered off to allow settings changes.'\n"
                    "   - **NEVER write 'The next step is...' into grounding notes.** Leave the planning to the 'thought' process.\n"

                    "**1. YOU (The Planner/Researcher):**\n"
                    "   - **Capabilities:** High-level reasoning, Memory (`grounding_notes`), **INTERNET ACCESS** (`google_search`).\n"
                    "   - **RULE:** If the user asks for information (Stock prices, Weather, 'Who is X', Docs), YOU must search for it. \n"
                    "   - **CRITICAL:** NEVER, UNDER ANY CIRCUMSTANCES, tell the Executor to open a browser and search. The Executor is blind to the web.\n"
                    "\n"
                    "**2. THE CODER (The Engineer):**\n"
                    "   - **Trigger:** Return JSON status `'CODING_REQUEST'`.\n"
                    "   - **Tools Available:** `save_file_content` (Write Code/Text), `execute_cmd` (Run Python/PIP/Shell), `read_file`.\n"
                    "   - **Use Case:** Creating files (Text, CSV, Python scripts), Debugging code, Installing dependencies, Generating Binary files (Images/PDFs via Python script).\n"
                    "\n"
                    "**3. THE EXECUTOR (The Blind Hand):**\n"
                    "   - **Trigger:** Return JSON status `'CONTINUE'` with `sprint_steps`.\n"
                    "   - **Tools Available:** `click_element`, `type_text`, `launch_app`, `execute_cmd` (Simple shell queries only), `wait`, `scroll`.\n"
                    "   - **Use Case:** GUI Interaction, Settings, Moving Windows, Clicking Buttons.\n"
                    "   - **LIMITATION:** NO Internet. BUT Brain. It can reason and adapt. You can give it complex instructions.\n"
                    "\n"
                    "--- \n\n"
                    "YOU CAN NOT DIRECTLY ASK THE USER ANYTHING. IF YOU NEED MORE INFO, USE `google_search`, OR TRY TO REPLICATE THE USE CASE.\n\n"
                    "###  SPRINT SCOPING & LOGICAL PARTITIONING\n"
                    "The Executor has a context limit and a 'stamina' of about 25 steps per sprint.\n"
                    "**DO NOT** shove unrelated tasks into one single Sprint.\n\n"
                    
                    "**EXAMPLE 1: \"Change Windows to Light Mode and then calculate 5 * 4 in Calculator.\"**\n"
                    " **BAD PLAN:** One huge list: [Open Settings, Click Personalization, Click Light Mode, Open Calculator, Type 5*4...]\n"
                    "   *Why? The Executor might fail midway, and retrying the whole chain is inefficient.*\n\n"
                    " **GOOD PLAN (Iterative):**\n"
                    "   - **Turn 1:** Output `sprint_steps` ONLY for changing the Theme. Goal: \"Windows is in Light Mode\".\n"
                    "   - **Turn 2:** (After system returns success) Output `sprint_steps` for opening Calculator and doing math.\n\n"
                    "**RULE:** If the request involves distinct apps or contexts, BREAK IT DOWN into separate Sprints (Loops).\n"
                    "**EXAMPLE 2: \"Fix an Error in a Windows 10 installation process.\"**\n"
                    " **BAD PLAN:** Very small steps. (e.g. Open Virtualbox, Open Windows 10 VM, Reproduce the error, Power off VM, Open Settings, etc.)\n"
                    "   *Why? The Executor can execute complex tasks at once. Doing one sprint for each is inefficient.*\n\n"
                    " **GOOD PLAN (Iterative):**\n"
                    "   - **Turn 1:** Output `sprint_steps` for starting the VM and reproducing the error. Goal: \"Error message is displayed\".\n"
                    "   - **Turn 2:** Research the error and start a sprint to fix it.\n\n"
                    "\n"
                    "--- \n\n"
                    "###  OS ENVIRONMENT & MULTI-TASKING (CRITICAL)\n"
                    "The Executor interacts with a real Windows PC, but it has **TUNNEL VISION**.\n"
                    "1. **VISIBILITY:** It only sees/analyzes the **CURRENTLY FOCUSED WINDOW**. Everything else is invisible to it.\n"
                    "2. **OPEN APPS REGISTRY:** You will be provided with a list of `OPEN APPS`. These are running in the background.\n"
                    "3. **SWITCHING CONTEXT:**\n"
                    "   - If you need to use an app listed in `OPEN APPS`, instructing `launch_app` again is wasteful.\n"
                    "   - Instead, instruct the Executor to use `focus_window(app_name='...')`.\n"
                    "   - **Scenario:** Copy text from Browser to Notepad.\n"
                    "     * Wrong Plan: [Read Browser, Open Notepad, Write]\n"
                    "     * Correct Plan: [Focus Browser, Copy Text] -> [Focus Notepad, Paste Text]\n"
                    "\n"
                    "### DECISION FLOW (Follow this Order)\n"
                    "1. **MISSING INFO?** -> Call `google_search`. (Do NOT generate sprint_steps).\n"
                    "2. **NEED A FILE/SCRIPT?** -> Return `'CODING_REQUEST'`. (Do NOT ask Executor to open Notepad and type).\n"
                    "3. **READY TO CLICK?** -> Return `'CONTINUE'` with `sprint_steps`.\n"
                    "   - Ensure the steps are atomic: \"Launch 'settings'\", \"Click 'Personalization'\", \"Click 'Colors'\"...\n"
                    "4. **JOB DONE?** -> Return `'COMPLETED'`.\n"
                    "\n"
                    "--- \n\n"

                    "### VISUAL DASHBOARD (SVG) - THE USER INTERFACE\n"
                        "The User sees ONLY this SVG to judge your intelligence. Make it look like a Sci-Fi HUD.\n"
                        "**SVG Canvas:** `viewBox='0 0 1000 300'`.\n"
                        "**Colors:** Background: Transparent. Text: #e5e5e5. Accent: #00ADB5 (Cyan) or #FF2E63 (Red for alerts).\n\n"
                        "**REQUIRED ELEMENTS:**\n"
                        "1. **TOP RIGHT - METRICS:**\n"
                        "   - Display **'EST. TIME: <X> min'** (Your best guess).\n"
                        "   - Display **'CONFIDENCE: <Y>%'** (How sure are you about this plan?).\n"
                        "2. **CENTER - PROGRESS NODES:**\n"
                        "   - Show a visual map of the milestones. DO NOT use generic names like 'Execution'.\n"
                        "   - Use CONCRETE names: `[Start] -> [Change Theme] -> [Open Calc] -> [Math] -> [Done]`.\n"
                        "   - Highlight the current node.\n"
                        "3. **BOTTOM - TERMINAL LOG (The 'Hacker' Thought):**\n"
                        "   - Font: 'Courier New', Monospace. Color: **#39ff14 (Neon Green)**.\n"
                        "   - Text: A direct, human-readable sentence explaining your strategy.\n"
                        "   - Example: `> SYSTEM_LOG: Initiating UI sequence. Instructing Executor to navigate Settings > Personalization to switch Theme.`\n"
                        "   - Example: `> SYSTEM_LOG: Researching stock data via Google before generating report.`\n\n"
                        "###  CRITICAL OUTPUT FORMAT RULES \n"
                    "You must output VALID JSON. To ensure parsing succeeds:\n"
                    "1. **NO STRING CONCATENATION**: Never use `\"text\" + \"text\"`. Write the full string.\n"
                    "2. **NO COMMENTS**: Do not use `//` inside the JSON.\n"
                    "3. **ESCAPE CHARACTERS**: Use `\\\\` for backslashes in paths (e.g. `C:\\\\Users`).\n"
                    "4. Output **ONLY** the JSON block wrapped in ```json ... ``` code fences.\n"
                    "JSON OUTPUT FORMAT:\n"
                    "{\n"
                    "  \"status\": \"CONTINUE\" | \"COMPLETED\" | \"FAILED\" | \"CODING_REQUEST\",\n"
                    "  \"milestone_name\": \"Name of the CURRENT Sprint (e.g. 'Switching Theme')\",\n"
                    "  \"success_condition\": \"What must be visible to consider this sprint done? (e.g. 'Calculator window open')\",\n"
                    "  \"sprint_steps\": [\"launch_app 'settings'\", \"click_element 42\"], \n"
                    "  \"svg_code\": \"<svg...>...</svg>\",\n"
                    "  \"grounding_notes\": \"(Optional) Info for the next turn\",\n"
                    "  \"coding_params\": { \"path\": \"...\", \"instruction\": \"...\" } \n"
                    "}"
                )

        self.executor_instruction = (
            "You are the EXECUTIVE OS AGENT. You are intelligent, fast, but also PATIENT. YOU ARE CONTROLLING A GERMAN WINDOWS 11 PC.\n"
            "WARNING: You DO NOT have internet access. You cannot browse the web. If you need Web Information, finish the sprint with a note to search the web for your problem.\n"
            "Your Manager has given you a SPRINT PLAN and a SUCCESS CONDITION.\n\n"
            "### CORE DIRECTIVE: ACT, THEN OBSERVE, THEN WAIT.\n"
            "You must execute the steps, but DO NOT stop immediately after clicking. You must WATCH the screen until the Success Condition is met.\n\n"
            "### THE LOOP RULES:\n"
            "1. **Check Condition**: Does the screen match the Manager's `success_condition`?\n"
            "   - YES -> IMMEDIATELY Call `finish_sprint(success)`. DO NOT GO ANY FURTHER. EXAMPLE: If the success condition is 'The Windows 10 VM starts and displays an error message', and you see the error, call `finish_sprint(success)` IMMEDIATELY without clicking the error or interacting with it!!!\n"
            "   - NO -> Continue to rule 2.\n"
            "2. **Detect Transience (The Waiting Game)**:\n"
            "   - Do you see a Spinner, Loading Bar, Windows Logo, Black Screen, or 'Please Wait'?\n"
            "   - Did you just click something that takes time (like 'Start', 'Install', 'Download')?\n"
            "   - **IF YES:** Call the `wait` tool (e.g. 5, 10, or 20 seconds). Do not click anything else.\n"
            "   - The system will pause and then show you the new screen state.\n"
            "3. **Action**:\n"
            " ALWAYS FOCUS A WINDOW BEFORE INTERACTING WITH IT!!!\n"
            "   - If not waiting and steps remain -> Execute next step tool.\n"
            "   - You can execute more tools at once if needed, like `click_element` followed by `type_text`, or click_element followed by `click_element`.\n\n"
            "    IMPORTANT: YOU CAN NOT CLICK THE DESKTOP. ALWAYS OPEN A APPLICATION TO INTERACT WITH.\n\n"
            "TEXT EDITING: NEVER EDIT TEXT/CODE YOURSELF. THE PLANNER HAS A SPECIAL CODER MODEL. REQUEST IT IF NEEDED, AND SPECIFY THE PATH FOR THE FILE. HOWEVER, YOU CAN TEST THESE FILES WITH THE CMD COMMAND LINE TOOL (e.g. python with py filename.py)\n\n"
            "### COMMAND LINE CAPABILITY:\n"
            "- You have a tool `execute_cmd` to run shell commands (dir, ipconfig, echo, type, etc.).\n"
            "- **RESTRICTION**: You are running as a STANDARD USER. Do not attempt Admin/Root commands.\n"
            "- Use this for: Checking files, network status, listing processes, or creating simple logs.\n\n"
            "Note: To open the Explorer, use the 'launch_app' tool with 'explorer', For the calculator use 'Rechner'. The base path to the current user is C:\\Users\\borge.\n"
            "### LAUNCHING APPS\n"
            "- You do not have a full list of installed apps. \n"
            "- Just guess the name (e.g. 'Calculator', 'Chrome', 'VirtualBox').\n"
            "- The system has a fuzzy search. If you are wrong, it will give you suggestions.\n"
            "\n"
            "###  CRITICAL EXECUTION RULES \n"
            "1. **NO TEXT ACTIONS:** NEVER write 'ACTION: click...' in your text response. It does nothing.\n"
            "2. **USE TOOLS:** To perform an action, you MUST use the provided Native Function/Tool.\n"
            "3. **ONE STEP AT A TIME:** Do not simulate the result. Do not write '-> RESULT:'. Just call the tool and stop.\n"
            "4. **STATELESS VISION:** You only see the CURRENT screen. Trust what you see NOW.\n"
            "5. **If a click on an element fails or causes no change on the screen, do NOT try it again. Instead, consider whether you need to click a different element or if you should use a different tool.**\n"
            "6. **Analyze screen** The goal is to set the theme to light mode but light mode is already on? Finish the sprint. Often times elements are disabled/grayed out. If so, do NOT click them. Also DONT click random text fields, focus on buttons.**\n"
            "ONLY DO WHAT THE SPRINT TELLS YOU TO DO. DO NOT DO ANY EXTRA ACTIONS. IF THE success_condition is met, stop IMMEDIATELY. Other behavior will result in instant termination.\n"
            " If you want to close/kill a window, use the 'close_window' tool to close the currently focused window.\n\n"
            "TEXT EDITING RULES:\n"
            "1. **CHECK STATUS**: Look at 'KEYBOARD STATUS' in the input. Is it 'TYPE_READY'?\n"
            "2. **FOCUS FIRST**: If status is NOT 'TYPE_READY', you MUST use `click_element` on the text field BEFORE calling `type_text`.\n"
            "3. **CHAINING**: You can output `[\"click_element ID\", \"wait 2\", \"type_text 'hello'\"]` in one turn to ensure focus.\n"
            "4. NEVER blindly type unless you just clicked the field.\n"
        )

    def _get_executor_tools(self) -> List[types.Tool]:
    
        funcs = [
            {"name": "execute_cmd", "description": "Executes CMD command.", "parameters": {"type": "OBJECT", "properties": {"command": {"type": "STRING"}}, "required": ["command"]}},
            {"name": "click_element", "description": "Clicks UI element by ID.", "parameters": {"type": "OBJECT", "properties": {"element_id": {"type": "INTEGER"}}, "required": ["element_id"]}},
            {"name": "double_click_element", "description": "Double-clicks ID.", "parameters": {"type": "OBJECT", "properties": {"element_id": {"type": "INTEGER"}}, "required": ["element_id"]}},
            {"name": "right_click_element", "description": "Right-clicks ID.", "parameters": {"type": "OBJECT", "properties": {"element_id": {"type": "INTEGER"}}, "required": ["element_id"]}},
            {"name": "type_text", "description": "Types text.", "parameters": {"type": "OBJECT", "properties": {"text": {"type": "STRING"}, "element_id": {"type": "INTEGER"}}, "required": ["text"]}},
            {"name": "scroll", "description": "Scrolls.", "parameters": {"type": "OBJECT", "properties": {"direction": {"type": "STRING", "enum": ["up", "down"]}}, "required": ["direction"]}},
            {"name": "launch_app", "description": "Launches app via Win+R.", "parameters": {"type": "OBJECT", "properties": {"app_name": {"type": "STRING"}}, "required": ["app_name"]}},
            {"name": "focus_window", "description": "Switches focus to open app.", "parameters": {"type": "OBJECT", "properties": {"app_name": {"type": "STRING"}}, "required": ["app_name"]}},
            {"name": "wait", "description": "Pauses execution.", "parameters": {"type": "OBJECT", "properties": {"seconds": {"type": "INTEGER"}}, "required": ["seconds"]}},
            {"name": "refresh_screen", "description": "Reloads screen.", "parameters": {"type": "OBJECT", "properties": {}, "required": []}},
            {"name": "finish_sprint", "description": "Milestone reached/Stuck.", "parameters": {"type": "OBJECT", "properties": {"result_summary": {"type": "STRING"}}, "required": ["result_summary"]}},
            {"name": "close_window", "description": "Closes active window.", "parameters": {"type": "OBJECT", "properties": {}, "required": []}},
        ]
        
        return_tools = []
        for f in funcs:
            return_tools.append(types.FunctionDeclaration(
                name=f["name"],
                description=f["description"],
                parameters=types.Schema(**f["parameters"])
            ))
        return [types.Tool(function_declarations=return_tools)]

    def _update_focus_stack(self, app_name, handle):
        self.focus_stack = [x for x in self.focus_stack if x[1] != handle]
        self.focus_stack.append((app_name, handle))
        self.active_app_name = app_name
        self.focus_handle = handle
        self.known_windows[app_name] = handle
        bring_window_to_front(handle)

    def reset(self):
        self.step_count = 0
        self.focus_rect = None
        if self.active_app_name and self.active_app_name in self.known_windows:
            potential_handle = self.known_windows[self.active_app_name]
            if check_ipc_handle_exists(self.ipc_addr[0], self.ipc_addr[1], self.ipc_key, potential_handle):
                self.focus_handle = potential_handle
            else:
                self.focus_handle = None
        else:
            self.focus_handle = None
        self.current_sprint_plan = []
        self.last_sprint_result = "Starting Task (State Preserved)."

    def _emit_event(self, event_type: str, data: Any):
        try:
            payload = json.dumps({"type": event_type, "data": data, "step": self.step_count})

            asyncio.create_task(self.redis_client.publish("agent_events", payload))
        except Exception as e:
            print(f"Redis Error: {e}")

    async def _capture_state(self):
        """
        Captures the visual and structural state of the Operating System.

        This method performs a hybrid analysis:
        1. Takes a screenshot of the target monitor.
        2. Retrieves the Accessibility Tree (UI Elements).
        3. Applies 'Focus Heuristics': If a specific window is tracked as focused,
           it crops the image to that window to improve model resolution and 
           filters UI elements to exclude background noise.

        Returns:
            tuple: (PIL.Image, List[UI_Elements], (x_offset, y_offset))
        """
        pil_img = None
        all_elements = []
        final_offset = (0, 0)
        target_handle = None
        

        if self.active_app_name:
            if self.active_app_name in self.known_windows:
                potential_handle = self.known_windows[self.active_app_name]
                if check_ipc_handle_exists(self.ipc_host, self.ipc_port, self.ipc_key, potential_handle):
                    target_handle = potential_handle
                else:
                    logger.warning(f"[Focus] Window '{self.active_app_name}' died.")
                    del self.known_windows[self.active_app_name]
                    self.active_app_name = None
            else:
                self.active_app_name = None

        self.focus_handle = target_handle
        target_mon = 1 
        
        try:
           
            pil_img, all_elements, monitor_offset = await asyncio.to_thread(
                observe_os_state, target_mon, self.focus_handle
            )
            

            if self.focus_handle and not all_elements and not pil_img:
                logger.warning(f"[Focus] Targeted scan failed. Fallback to Full Scan.")
                pil_img, all_elements, monitor_offset = await asyncio.to_thread(
                    observe_os_state, target_mon, None
                )

            final_offset = monitor_offset

        except Exception as e:
            logger.error(f"Screenshot/UIA Error: {e}")
            return Image.new('RGB', (800, 600)), [], (0,0)

    
        if self.focus_handle and pil_img:
            try:
                focus_rect = None
                window_element = next((el for el in all_elements if el.get('top_level_handle') == self.focus_handle), None)
                if window_element:
                    focus_rect = window_element.get('absolute_rectangle')
                
                if not focus_rect:
                    focus_rect = get_handle_rect(self.ipc_host, self.ipc_port, self.ipc_key, self.focus_handle)

                self.focus_rect = focus_rect

                if self.focus_rect:
                    mon_x, mon_y = monitor_offset
                    f_left, f_top, f_right, f_bottom = self.focus_rect
                    
                    crop_left = f_left - mon_x
                    crop_top = f_top - mon_y
                    crop_right = f_right - mon_x
                    crop_bottom = f_bottom - mon_y
                    
                    img_w, img_h = pil_img.size
                    crop_left = max(0, min(img_w, crop_left))
                    crop_top = max(0, min(img_h, crop_top))
                    crop_right = max(0, min(img_w, crop_right))
                    crop_bottom = max(0, min(img_h, crop_bottom))
                    
                    if (crop_right - crop_left) > 20 and (crop_bottom - crop_top) > 20:
                        pil_img = pil_img.crop((crop_left, crop_top, crop_right, crop_bottom))
                        final_offset = (mon_x + crop_left, mon_y + crop_top)
                        
                        visible_elements = []
                        for el in all_elements:
                            r = el.get('absolute_rectangle')
                            if not r: continue
                            cx = (r[0] + r[2]) // 2
                            cy = (r[1] + r[3]) // 2
                            if (f_left <= cx <= f_right) and (f_top <= cy <= f_bottom):
                                visible_elements.append(el)
                        all_elements = visible_elements
                
            except Exception as e:
                logger.error(f"Crop Logic Error: {e}")

        return pil_img, all_elements, final_offset

    def _optimize_image(self, pil_img):
        target_width = 1920 
        if pil_img.width > target_width:
            ratio = target_width / pil_img.width
            new_height = int(pil_img.height * ratio)
            pil_img = pil_img.resize((target_width, new_height), Image.Resampling.LANCZOS)
        
        buf = io.BytesIO()
        pil_img.save(buf, format="JPEG", quality=70, optimize=True)
        img_bytes = buf.getvalue()
        return img_bytes, base64.b64encode(img_bytes).decode('utf-8')

    def _fast_element_format(self, elements):
        final_list = []
        for el in elements:
            e_id = el['id']
            name = el.get('name', '')
            e_type = el.get('type', 'Unknown')
            if name:
                entry = f"{e_id}:{name}<{e_type}>"
            else:
                entry = f"{e_id}:<{e_type}>"
            final_list.append(entry)

        full_text = "; ".join(final_list)
        if len(full_text) > 30000: 
            return full_text[:30000] + "...(truncated)"
        return full_text
    
    def _map_tool_to_internal_action(self, tool_name: str, params: Dict) -> Optional[Dict]:
        if tool_name == "click_element": return {"command": "click", "element_id": int(params["element_id"])}
        elif tool_name == "double_click_element": return {"command": "double_click", "element_id": int(params["element_id"])}
        elif tool_name == "right_click_element": return {"command": "right_click", "element_id": int(params["element_id"])}
        elif tool_name == "type_text": return {"command": "type", "text": params["text"], "element_id": params.get("element_id")}
        elif tool_name == "scroll": return {"command": "scroll", "direction": params["direction"]}
        elif tool_name == "launch_app": return {"command": "launch_app", "app_name": params["app_name"]}
        elif tool_name == "focus_window": return {"command": "focus_window_internal", "app_name": params.get("app_name")}
        elif tool_name == "execute_cmd": return {"command": "execute_cmd", "cmd_line": params["command"]}
        return None

    def _get_fuzzy_suggestions(self, failed_query: str) -> str:
        import difflib
        if not hasattr(self, 'app_index') or not self.app_index:
            return "No app index."
        installed = list(self.app_index.keys())
        matches = difflib.get_close_matches(failed_query.lower(), installed, n=5, cutoff=0.4)
        if not matches:
            matches = [name for name in installed if failed_query.lower() in name][:5]
        return ", ".join(matches) if matches else "Try standard names."

 

    async def run_autonomous_loop(self, user_input: str):

        """
        Starts the main autonomous agent loop.

        This method orchestrates the high-level workflow:
        1. Captures the current OS state (Screenshot & UI Tree).
        2. Consults the 'Manager' (Planner model) to determine the next 'Sprint'.
        3. Updates the Dashboard via Redis events.
        4. Delegates execution to the 'Executor' or 'Coder' based on the plan.
        5. Repeats until the objective is completed or max phases are reached.

        Args:
            user_input (str): The high-level objective provided by the user.
        """
        if self.is_running: return
        self.is_running = True
        self.reset()
        max_phases = 15
        current_phase = 0
        
        try:
            while self.is_running and current_phase < max_phases:
                current_phase += 1
                self._emit_event("status", "planning_analysis")
                
        
                pil_img, _, _ = await self._capture_state()
                _, b64_clean = self._optimize_image(pil_img)
                self._emit_event("screen_update", b64_clean)
                
                open_apps = list(self.known_windows.keys())
                focus_state = self.active_app_name if self.active_app_name else "Desktop"

             
                plan_data = await self._consult_manager(
                    user_input, b64_clean, self.last_sprint_result, 
                    self.grounding_context, open_apps, focus_state
                )
                
                status = plan_data.get("status", "CONTINUE")
                svg_code = plan_data.get("svg_code")
                if svg_code:
                    self._emit_event("dashboard_update", svg_code)
    
                if status == "CODING_REQUEST":
                    params = plan_data.get("coding_params", {})
                    path = params.get("path")
                    instr = params.get("instruction")
                    notes = plan_data.get("grounding_notes", "")
                    if path and instr:
                        self._emit_event("log", f"Manager requests Code Agent for {path}")
                        code_res = await self._run_coder_session(path, f"{instr}\nContext: {notes}")
                        self.last_sprint_result = f"CODER FINISHED:\n{code_res}"
                        continue 

                if status == "COMPLETED":
                    self._emit_event("success", "Task Completed.")
                    break
                if status == "FAILED":
                    self._emit_event("error", "Task Failed.")
                    break
                
            
                self.current_sprint_plan = plan_data.get("sprint_steps", [])
                milestone = plan_data.get("milestone_name", f"Phase {current_phase}")
                success_cond = plan_data.get("success_condition", "Visual confirmation")
                
                if "grounding_notes" in plan_data:
                     self._emit_event("log", f"Ref Info: {plan_data['grounding_notes']}")

           
                sprint_result = await self._run_executor_sprint(
                    user_input, milestone, self.current_sprint_plan, success_cond
                )
                
                self.last_sprint_result = sprint_result

        except Exception as e:
            self._emit_event("error", f"System Critical: {e}")
            traceback.print_exc()
        finally:
            self.is_running = False
            self._emit_event("status", "disconnected")
    async def _consult_manager(self, main_goal: str, b64_img: str, last_result: str, grounding_ctx: str, open_apps: List[str], current_focus: str) -> Dict:
            apps_list_str = ", ".join(open_apps) if open_apps else "None (Clean State)"
            """
            Invokes the Planner Model (Manager) to generate the next strategic Sprint.

            The Manager reviews the goal, history, and current screen state to output
            a JSON plan containing:
            - status: CONTINUE, CODING_REQUEST, or COMPLETED.
            - sprint_steps: A high-level list of actions for the Executor.
            - success_condition: What the screen must look like to finish the sprint.

            Args:
                main_goal (str): The user's original request.
                b64_img (str): Base64 encoded screenshot of the current focus.
                last_result (str): The outcome of the previous execution sprint.
                grounding_ctx (str): Accumulated facts/notes from previous turns.
                open_apps (List[str]): List of currently tracked open windows.
                current_focus (str): Name of the currently focused app.

            Returns:
                Dict: Parsed JSON response containing the plan.
            """
     
            planner_chat = self.api.create_chat_session(
                model_override=self.planner_model,
                system_instruction=self.planner_instruction,
                tool_definitions=[], 
                enable_google_search=True 
            )
            
            prompt = (
                f"MAIN GOAL: {main_goal}\n"
                f"LAST SPRINT RESULT: {last_result}\n\n"
                f"###  LONG-TERM MEMORY (GROUNDING CONTEXT)\n"
                f"WARNING: The following info contains FACTS established in previous turns. "
                f"Assume these facts are TRUE. Do NOT re-verify them unless explicitly failed.\n"
                f"-------------------------\n"
                f"{grounding_ctx}\n"
                f"-------------------------\n\n"
                f"### CURRENT SYSTEM STATE\n"
                f"- **ACTIVELY FOCUSED WINDOW**: '{current_focus}'\n"
                f"- **BACKGROUND APPS**: [{apps_list_str}]\n\n"
                "If info is missing, SEARCH GOOGLE. "
                "Then, output the JSON plan."
            )
   
            message_parts = [prompt, {"mime_type": "image/jpeg", "data": b64_img}]


            response = await asyncio.to_thread(
                self.api.send_chat_message, 
                planner_chat, 
                message_parts
            )
            
          
            if response.get("grounding_info"):
                self._emit_event("grounding", response["grounding_info"])
                
            thought = response.get("thought", "").strip()
            
        
            try:
                json_str = thought
                match = re.search(r'```json\s*(\{.*?\})\s*```', thought, re.DOTALL)
                if match:
                    json_str = match.group(1)
                else:
                    match_alt = re.search(r'(\{.*\})', thought, re.DOTALL)
                    if match_alt:
                        json_str = match_alt.group(1)

                if not json_str:
                    raise ValueError("No JSON block found in output.")

                json_str = re.sub(r'^\s*//.*$', '', json_str, flags=re.MULTILINE)
                json_str = re.sub(r'"\s*\+\s*"', "", json_str)
                json_str = re.sub(r',\s*\}', '}', json_str)
                json_str = re.sub(r',\s*\]', ']', json_str)

                return json.loads(json_str)

            except Exception as e:
                print(f"General Parsing Error: {e}")
                return {
                    "status": "FAILED", 
                    "milestone_name": "Parsing Error", 
                    "sprint_steps": []
                }

    async def _run_executor_sprint(self, main_goal: str, milestone: str, steps: List[str], success_condition: str) -> str:
            """
            Runs the Executor Loop to perform low-level OS actions.

            This loop operates in short 'Sprints' (max 25 steps) to achieve a specific milestone.
            It uses the 'Flash-Lite' model for speed. 
            
            Key Behaviors:
            - Validates the screen against the `success_condition` before every action.
            - Auto-detects popups to shift focus.
            - Executes tools: click, type, wait, launch_app, etc.
            - Stops immediately if the success condition is met.

            Returns:
                str: A summary of what was achieved or an error message.
            """
            tools_def = self._get_executor_tools() 

            sprint_active = True
            sprint_result = "Sprint timed out."
            step_safety = 0
            max_steps = 25 
            
      
            recent_history = []
            last_tool_output = "I am ready. Show me the screen."
            
            try:
         
                initial_windows = get_all_windows_from_server(self.ipc_host, self.ipc_port, self.ipc_key)
                window_snapshot = {w['handle'] for w in initial_windows}
            except:
                window_snapshot = set()
                
            print(f"\n🚀 SPRINT START: {milestone}")
            
            while sprint_active and self.is_running and step_safety < max_steps:
                step_safety += 1
                
        
                try:
                
                    current_win_list = get_all_windows_from_server(self.ipc_host, self.ipc_port, self.ipc_key)
                    current_handles_map = {w['handle']: w for w in current_win_list}
                    current_handles_set = set(current_handles_map.keys())
                    
            
                    new_handles = current_handles_set - window_snapshot
                    
                    popup_found = False
                    if new_handles:
                        for w in current_win_list:
                            if w['handle'] in new_handles:
                                rect = w.get('rect', [0,0,0,0])
                                w_w, w_h = rect[2] - rect[0], rect[3] - rect[1]
                                title = w.get('title', 'Unknown')
                                
                                ignored_titles = [
                                    "Default IME", "MSCTFIME UI", "NVIDIA GeForce Overlay", 
                                    "Windows-Widgets", "Windows Widgets", "SearchHost", 
                                    "StartMenuExperienceHost", "Cortana", "Program Manager", 
                                    "Task View", "PopupHost"
                                ]
                                
                                if title not in ignored_titles and title.strip() != "":
                                    if w_w > 20 and w_h > 20:
                                        print(f"*** AUTO-FOCUS: Popup detected '{title}' (Handle: {w['handle']}) ***")
                                        self._emit_event("log", f"Popup: {title}")
                                        self._update_focus_stack(title, w['handle'])
                                        popup_found = True
                                        await asyncio.sleep(0.3)
                                        break
                                        
                    if not popup_found and self.focus_handle:
                        if self.focus_handle not in current_handles_set:
                            print(f"*** FOCUS LOST: Window {self.focus_handle} is gone. Searching history... ***")
                            self._emit_event("log", "Window closed. Reverting focus...")
                            
                            found_fallback = False
                            while self.focus_stack:
                                last_name, last_handle = self.focus_stack.pop()
                                if self.focus_stack:
                                    prev_name, prev_handle = self.focus_stack[-1]
                                    if prev_handle in current_handles_set:
                                        print(f"*** RESTORING FOCUS to '{prev_name}' ***")
                                        self.active_app_name = prev_name
                                        self.focus_handle = prev_handle
                                        bring_window_to_front(prev_handle)
                                        found_fallback = True
                                        break
                                else:
                                    break
                            
                            if not found_fallback:
                                print("*** NO HISTORY LEFT. Switching to Full Desktop Scan. ***")
                                self.focus_handle = None
                                self.active_app_name = None

                    window_snapshot = current_handles_set

                except Exception as e:
                    print(f"Auto-focus/Recovery Error: {e}")
                    traceback.print_exc()

         
                t0_loop = time.perf_counter()
                
        
                self._emit_event("status", f"step {step_safety}/{max_steps}")
                self._emit_event("state_update", f"👁️ READING SCREEN")

                try:
                    t0_capture = time.perf_counter()
                    clean_img, elements, current_offset = await self._capture_state()
                    t1_capture = time.perf_counter()

                    t0_prep = time.perf_counter()
                    self.current_elements = elements 
                    _, annotated_bytes = prepare_images_for_model(clean_img, elements, current_offset, step_safety)
                    clean_bytes, b64_clean_for_ui = self._optimize_image(clean_img)
                    self._emit_event("screen_update", b64_clean_for_ui)
                    
                
                    active_window_info = get_ipc_active_window(self.ipc_host, self.ipc_port, self.ipc_key)
                    caret_status = "UNKNOWN"
                    if active_window_info and "caret" in active_window_info:
                        if active_window_info["caret"].get("active"):
                            caret_status = "✅ TYPE_READY"
                        else:
                            caret_status = "❌ NO TEXT FOCUS"

                    curr_focus = self.active_app_name if self.active_app_name else "Desktop/None"
                    open_apps_str = ", ".join(list(self.known_windows.keys())) if self.known_windows else "None"
                    elements_text = self._fast_element_format(elements)
                    t1_prep = time.perf_counter()
                    
                except Exception as e:
                    print(f"Capture Error: {e}")
                    traceback.print_exc()
                    break

                history_str = ""
                if recent_history:
                    formatted_entries = []
                    for entry in recent_history[-5:]:
                        clean_entry = entry.replace("THOUGHT:", "Thought:").replace("ACTION:", " | Executed:").replace("-> RESULT:", " | Outcome:")
                        formatted_entries.append(f"Step Record: {clean_entry}")
                    history_str = "\n".join(formatted_entries)
                    history_block = f"###  HISTORY (FOR CONTEXT ONLY):\n{history_str}\n"
                else:
                    history_block = "###  HISTORY: (No steps taken yet)\n"

                user_block = (
                    f"###  LIVE SYSTEM CONTEXT\n"
                    f"- **FOCUS**: '{curr_focus}'\n"
                    f"- **OPEN APPS**: [{open_apps_str}]\n"
                    f"- **KEYBOARD**: {caret_status}\n\n"
                    f"###  MISSION\n"
                    f"GOAL:{main_goal}\n"
                    f"CURRENT TASK: {milestone}\n"
                    f"EXIT CONDITION: {success_condition}\n\n"
                    f"{history_block}\n" 
                    f"###  CURRENT OBSERVATION\n"
                    f"UI ELEMENTS:\n{elements_text}\n\n"
                    f"LAST TOOL FEEDBACK: {last_tool_output}\n\n"
                    f"⚠️ INSTRUCTION:\n"
                    f"Call the next necessary Function/Tool based on the screenshot."
                )

                contents = [
                    types.Content(role="user", parts=[
                        types.Part(inline_data=types.Blob(mime_type="image/jpeg", data=clean_bytes)),
                        types.Part(inline_data=types.Blob(mime_type="image/png", data=annotated_bytes)),
                        types.Part(text=user_block)
                    ])
                ]

            
                self._emit_event("state_update", " AI THINKING")
                t0_infer = time.perf_counter()
                
                try:
          
                    full_thought = ""
                    function_calls = []

                    async for chunk in self.api.generate_content_stream(
                        contents=contents,
                        tools=tools_def,
                        system_instruction=self.executor_instruction,
                        model_override=self.executor_model
                    ):
                        if chunk.candidates and chunk.candidates[0].content.parts:
                            for part in chunk.candidates[0].content.parts:
                                if part.text:
                                    txt = part.text
                                    full_thought += txt
                                    self._emit_event("thought_partial", txt)
                                if part.function_call:
                                    function_calls.append(part.function_call)
                    
                    t1_infer = time.perf_counter()

                    clean_thought = full_thought.replace('\n', ' ').strip()
                    step_log = f"[Step {step_safety}] THOUGHT: {clean_thought[:150]}..." 
                    
                    if not function_calls:
                        last_tool_output = "No tool called. Screen updated."
                        recent_history.append(f"{step_log} -> NO ACTION.")
                        t_tools_duration = 0.0
                        await asyncio.sleep(1)
                    else:
                  
                        self._emit_event("state_update", "EXECUTING")
                        t0_tools = time.perf_counter()
                        
                        for call in function_calls:
                            cmd = call.name
                            params = dict(call.args) if call.args else {}
                            
                            self._emit_event("log", f"Exec: {cmd}")
                            step_log += f"\n   ACTION: {cmd}({params})"

                            if cmd == "finish_sprint":
                                sprint_result = params.get("result_summary", "Finished")
                                sprint_active = False
                                break 

                            if cmd == "launch_app":
                                app_req_name = params.get("app_name", "")
                                q = Queue()
                        
                                res = await execute_os_action({"command": "launch_app", "app_name": app_req_name}, self.current_elements, q)
                                
                                if res.get("status") == "success":
                                    new_h = res.get("handle")
                                    if new_h:
                                        self.known_windows[app_req_name] = new_h
                                        self.active_app_name = app_req_name
                                        tool_res = f"SUCCESS: Launched '{app_req_name}' (Handle {new_h}). Window is now FOCUSED."
                                    else:
                                        tool_res = "WARNING: App launched but window not detected immediately (Timed out)."
                                else:
                                    suggestions = self._get_fuzzy_suggestions(app_req_name)
                                    tool_res = f"ERROR: App '{app_req_name}' NOT found in index. Suggestions: [{suggestions}]"

                            elif cmd == "close_window":
                                if self.focus_handle:
                                    q = Queue()
                   
                                    close_res = await execute_os_action(
                                        {"command": "close_window", "handle": self.focus_handle}, 
                                        [], q
                                    )
                                    tool_res = close_res.get("message", "Close command sent.")
                                    
                                    if self.active_app_name:
                                        if self.active_app_name in self.known_windows:
                                            del self.known_windows[self.active_app_name]
                                        self.focus_stack = [x for x in self.focus_stack if x[1] != self.focus_handle]
                                        tool_res += f" App '{self.active_app_name}' removed from registry."
                                        self.active_app_name = None
                                        self.focus_handle = None
                                else:
                                    tool_res = "ERROR: No window is currently focused/tracked to close."                

                            elif cmd == "focus_window":
                                target_name = params.get("app_name")
                                match = next((k for k in self.known_windows if k.lower() == target_name.lower()), None)
                                if match:
                                    handle = self.known_windows[match]
                                    q = Queue()
                                    await execute_os_action({"command": "focus_window", "element_id": handle}, [], q)
                                    self.active_app_name = match
                                    await asyncio.sleep(0.5)
                                    tool_res = f"Focused '{match}'."
                                else:
                                    tool_res = f"Error: '{target_name}' is not in OPEN APPS list."

                            elif cmd == "execute_cmd":
                                c_str = params.get("command")
                                q = Queue()
                                res = await execute_os_action({"command": "execute_cmd", "cmd_line": c_str}, [], q)
                                out = res.get("output", "")
                                self.grounding_context += f"\nCMD '{c_str}':\n{out[:500]}..."
                                tool_res = f"CMD Output: {out[:300]}..."

                            elif cmd == "wait":
                                sec = int(params.get("seconds", 2))
                                await asyncio.sleep(sec)
                                tool_res = f"Waited {sec}s."
                                
                            elif cmd == "refresh_screen":
                                tool_res = "Screen refreshed."
                                
                            else:
                                internal = self._map_tool_to_internal_action(cmd, params)
                                if internal:
                                    q = Queue()
                                    await execute_os_action(internal, self.current_elements, q)
                                    tool_res = "Action executed."
                                else:
                                    tool_res = "Error: Unknown internal tool mapping."

                            step_log += f" -> RESULT: {tool_res}"
                            last_tool_output = f"Tool '{cmd}' result: {tool_res}"

                            if cmd not in ["launch_app", "focus_window", "wait"]:
                                await asyncio.sleep(0.3)
                        
                        t1_tools = time.perf_counter()
                        t_tools_duration = t1_tools - t0_tools
                        recent_history.append(step_log)


                    t_total = time.perf_counter() - t0_loop
                    t_capture_duration = t1_capture - t0_capture
                    t_prep_duration = t1_prep - t0_prep
                    t_infer_duration = t1_infer - t0_infer
                    print(f"⏱️ [STEP {step_safety} TIMING] Total: {t_total:.2f}s | Capture: {t_capture_duration:.2f}s | Prep: {t_prep_duration:.2f}s | Inference: {t_infer_duration:.2f}s | Tools: {t_tools_duration:.2f}s")

                except Exception as e:
                    print(f"Executor Loop Critical: {e}")
                    traceback.print_exc()
                    last_tool_output = f"SYSTEM ERROR: {e}"
                    recent_history.append(f"SYSTEM ERROR: {e}")
                    
            return sprint_result

    async def _run_coder_session(self, file_path: str, instruction: str) -> str:
            """
            Spawns a specialized 'Coder' session for file manipulation and scripting.

            Unlike the UI-based Executor, this agent has no vision but has access to:
            - File System (Read/Write)
            - Command Line (Execution)
            
            It is used when the Planner determines that writing code or config files 
            is more efficient than using a GUI text editor.

            Args:
                file_path (str): Target file to create or edit.
                instruction (str): What to do with the code (e.g., "Fix bug", "Write script").

            Returns:
                str: Summary of the coding session.
            """
            from src.core import read_local_file, write_to_local_file 
            import os
            

            if not os.path.isabs(file_path):
                base_dir = os.path.expanduser("~") 
                lower_path = file_path.lower()
                if lower_path.startswith("desktop"):
                    file_path = os.path.join(base_dir, file_path)
                elif lower_path.startswith("documents"):
                    file_path = os.path.join(base_dir, "Documents", file_path.split(os.sep, 1)[-1])
                else:
                    if os.sep not in file_path:
                        file_path = os.path.join(base_dir, "Desktop", file_path)
            
            file_path = os.path.abspath(file_path)
            working_dir = os.path.dirname(file_path)
            if not os.path.exists(working_dir):
                os.makedirs(working_dir, exist_ok=True)

            self._emit_event("log", f" Starting Pro Coder Session for: {os.path.basename(file_path)}")
            
   
            read_res = read_local_file(file_path)
            current_code = read_res["content"] if read_res["status"] == "success" else "(New File)"
            self._emit_event("code_view", {"code": current_code, "path": file_path, "status": "reading"})

 
            coder_tools = [
                {"name": "save_file_content", "description": "Overwrites/Creates file with new code/text.", "parameters": {"type": "OBJECT", "properties": {"new_content": {"type": "STRING"}, "target_path": {"type": "STRING"}}, "required": ["new_content"]}},
                {"name": "execute_cmd", "description": "Executes a terminal command.", "parameters": {"type": "OBJECT", "properties": {"command": {"type": "STRING"}}, "required": ["command"]}},
                {"name": "read_file", "description": "Reads a file.", "parameters": {"type": "OBJECT", "properties": {"path": {"type": "STRING"}}, "required": ["path"]}},
                {"name": "finish_coding", "description": "Call this when the task is done.", "parameters": {"type": "OBJECT", "properties": {"summary": {"type": "STRING"}}, "required": ["summary"]}}
            ]

            sys_prompt = "You are an Elite Senior DevOps Engineer. Write code, execute, debug."
            chat = self.api.create_chat_session(
                model_override=self.coder_model, 
                system_instruction=sys_prompt,
                tool_definitions=coder_tools,
                enable_google_search=False 
            )

            next_prompt = f"TARGET FILE: {file_path}\nINSTRUCTION: {instruction}\nCURRENT CONTENT:\n```\n{current_code}\n```\nStart."

            max_turns = 15
            turn = 0
            final_summary = "Coder session timed out."

            while turn < max_turns:
                turn += 1
                

                response = await asyncio.to_thread(
                    self.api.send_chat_message, chat, [next_prompt]
                )
                
                thought = response.get("thought", "")
                if thought: self._emit_event("thought_stream", f"Coder: {thought}")

                actions = response.get("actions", [])
                
                if not actions:
                    if "finish" in thought.lower(): break
                    next_prompt = "No action. If done, call finish_coding."
                    continue

                tool_outputs = []
                session_finished = False

                for action in actions:
                    cmd = action["command"]
                    params = action["parameters"]
                    
                    if cmd == "finish_coding":
                        final_summary = params.get("summary", "Done.")
                        session_finished = True
                        break
                    elif cmd == "save_file_content":
                        raw_tgt = params.get("target_path", file_path)
                        tgt = os.path.join(working_dir, raw_tgt) if not os.path.isabs(raw_tgt) else raw_tgt
                        content = params.get("new_content")
                        self._emit_event("code_view", {"code": content, "path": tgt, "status": "writing"})
                        write_to_local_file(tgt, content)
                        tool_outputs.append(f"Saved to {tgt}")
                    elif cmd == "read_file":
                        tgt = params.get("path", file_path)
                        if not os.path.isabs(tgt): tgt = os.path.join(working_dir, tgt)
                        r = read_local_file(tgt)
                        tool_outputs.append(f"Content:\n{r['content'][:500]}...")
                    elif cmd == "execute_cmd":
                        c_line = params.get("command")
                        self._emit_event("cmd_stream", {"cmd": f"(CODER) {c_line}", "output": "Running..."})
                        try:
                            proc = await asyncio.create_subprocess_shell(c_line, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE, cwd=working_dir)
                            out, err = await proc.communicate()
                            full = (out + err).decode('utf-8', errors='replace')
                            self._emit_event("cmd_stream", {"cmd": f"(CODER) {c_line}", "output": full[:1000]})
                            tool_outputs.append(f"Result:\n{full}")
                        except Exception as e:
                            tool_outputs.append(f"Error: {e}")

                if session_finished: break
                
                if tool_outputs:
                    next_prompt = "\n".join([f"TOOL_OUTPUT: {out}" for out in tool_outputs])
                else:
                    next_prompt = "Action executed."

            self._emit_event("code_view", {"close": True})
            return f"Summary: {final_summary}"