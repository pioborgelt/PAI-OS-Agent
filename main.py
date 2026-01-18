"""
PAI OS Agent - Main Entry Point
===============================

This module initializes the FastAPI server that acts as the central nervous system
for the PAI OS Agent. It coordinates the communication between the web-based
frontend, the Redis message bus, and the autonomous agent logic.

Key Responsibilities:
- Serving the User Interface (Static Files & HTML).
- Managing the application lifecycle (starting Redis clients, UI Server, and Agent).
- Exposing REST endpoints for chat interaction (`/chat`, `/stop`).
- Streaming real-time agent thoughts and logs via Server-Sent Events (SSE).

⚠️ IMPORTANT PRE-REQUISITES ⚠️
------------------------------
This system is complex and requires specific environment configuration (Redis,
IPC Server, Google API Keys, etc.).
PLEASE READ THE `README.md` FILE CAREFULLY BEFORE ATTEMPTING TO RUN THIS SCRIPT.

Author: Pio Borgelt
"""

import os
import sys
import logging
import asyncio
import subprocess
import json
from contextlib import asynccontextmanager
from typing import Dict
import uvicorn
import redis.asyncio as redis
from fastapi import FastAPI, Request, Body, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from sse_starlette.sse import EventSourceResponse
from dotenv import load_dotenv
from src.agent import OSAgent

# Konfiguration laden
load_dotenv()
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("MAIN")

ui_server_process = None
redis_client: redis.Redis = None
agent_instance: OSAgent = None

@asynccontextmanager
async def lifespan(app: FastAPI):
    """
    Lifecycle Manager:
    Startet UI Server, Redis Client und Agent.
    """
    global ui_server_process, redis_client, agent_instance

    logger.info("Starting UI Analysis Server (src/server.py)...")
    try:
        # Starte den Windows UI Server
        ui_server_process = subprocess.Popen([sys.executable, "-m", "src.server"])
    except Exception as e:
        logger.error(f"Failed to start UI server: {e}")

    # Redis Verbindung (Async)
    redis_client = redis.Redis(host='localhost', port=6379, decode_responses=True)
    
    # Agent Initialisieren
    agent_instance = OSAgent(redis_client=redis_client)

    logger.info("System initialized. Agent is ready.")
    
    yield
    
    # --- SHUTDOWN ---
    logger.info("Shutting down...")
    
    if agent_instance:
        agent_instance.is_running = False

    if ui_server_process:
        ui_server_process.terminate()
        try:
            ui_server_process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            ui_server_process.kill()
    
    if redis_client:
        await redis_client.aclose()
    
    logger.info("Shutdown complete.")

app = FastAPI(title="PAI OS Agent", lifespan=lifespan)

if not os.path.exists("static"):
    os.makedirs("static")
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
async def get_index():
    try:
        with open("static/index.html", "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return "<h1>Error: static/index.html not found.</h1>"

@app.post("/chat")
async def start_chat(payload: Dict[str, str] = Body(...)):
    if not agent_instance:
        raise HTTPException(status_code=503, detail="Agent not initialized")

    message = payload.get("message", "")
    if not message:
        raise HTTPException(status_code=400, detail="Message empty")

    if agent_instance.is_running:
        return {"status": "busy", "message": "Agent is already running a task."}

    # Task starten
    asyncio.create_task(agent_instance.run_autonomous_loop(message))

    return {"status": "started", "message": "Objective accepted."}

@app.post("/stop")
async def stop_agent():
    if agent_instance:
        agent_instance.is_running = False
    return {"status": "stopped"}

@app.get("/events")
async def event_stream(request: Request):
    """
    SSE Endpoint. Abonniert direkt den Redis Channel.
    """
    async def event_generator():
        # Initialer Status
        yield {
            "event": "state_update",
            "data": json.dumps("ACTIVE" if agent_instance and agent_instance.is_running else "IDLE")
        }

        if not redis_client:
            logger.error("No Redis client available for SSE")
            return

        pubsub = redis_client.pubsub()
        await pubsub.subscribe("agent_events")

        try:
            async for message in pubsub.listen():
                if await request.is_disconnected():
                    logger.info("Client disconnected from SSE")
                    break

                if message["type"] == "message":
                    raw_data = message["data"]
                    try:
                        # Agent sendet JSON String: {"type": "...", "data": "...", "step": ...}
                        payload = json.loads(raw_data)
                        
                        event_type = payload.get("type", "update")
                        event_data = payload.get("data", "")
                        
                        # Wir senden es als Server-Sent Event weiter
                        yield {
                            "event": event_type, 
                            "data": json.dumps(event_data) # Data muss String sein für SSE
                        }
                    except Exception as e:
                        logger.error(f"Error parsing redis message: {e}")

        except asyncio.CancelledError:
            pass
        finally:
            await pubsub.unsubscribe("agent_events")
            await pubsub.close()

    return EventSourceResponse(event_generator())

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)