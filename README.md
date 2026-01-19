
![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python)
![AI](https://img.shields.io/badge/GenAI-Gemini%20Flash%20%2F%20Pro-orange?style=for-the-badge)
![Windows](https://img.shields.io/badge/Platform-Windows%2011-0078D6?style=for-the-badge&logo=windows)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

> **"Built by a 17-year-old developer exploring the skills of Autonomous Desktop Agents using Gemini."**

## 🚀 Overview

The **PAI OS Agent** is a fully autonomous AI system capable of controlling a Windows PC to execute complex, multi-step tasks. Unlike standard automation scripts, PAI "sees" the screen and "thinks" about the next step.

It uses a **Hybrid Vision System** combining **Vision (Screenshots with OCR)** and **Accessibility Trees (UIA)** to navigate the OS, allowing it to interact with apps it has never seen before—from Calculator to VM VirtualBox.

### 🎥 Demo

https://github.com/user-attachments/assets/ec7fab50-e609-42b6-a085-f3912d2b9029

---

## 🏗️ Architecture

The system avoids the "loop-of-death" common in simple agents by splitting cognition into specialized roles. It uses a **Client-Server Architecture** to bypass Python's issues during heavy Windows UI interactions.

### The "Brain" (Agent)
*   **Planner Model (Gemini Pro):** High-level reasoning. Breaks the user's goal into logical "Sprints" (e.g., "Open Calculator", "Change Theme"). It has internet access and long-term memory.
*   **Executor Model (Gemini Flash):** Fast, low-latency execution. Handles the mouse and keyboard. 

*   **Coder Agent:** A specialized sub-agent that can write Python scripts or edit files when GUI interaction is inefficient.

### The "Body" (Perception & Action)
*   **Hybrid Vision:** Uses EasyOCR (GPU accelerated or CPU as fallback) for text on screen AND pywinauto/UIA to fetch the underlying DOM-like tree of Windows apps.
*   **IPC Server (server.py):** A standalone process that handles Win32 API calls, ensuring the UI remains responsive while the Agent "thinks".
*   **Redis Pub/Sub:** Acts as the nervous system, streaming UI Updated including real-time thoughts, logs, and screen updates to the frontend.

---

## ✨ Key Features
*   **Complex Step Planning**: The planner enables the agent to follow complex multi program tasks, like fixing an error in an Windows VM.
*   **Self-Correction:** If a click fails or a window doesn't open, the agent analyzes the new state and retries with a different strategy.
*   **Visual HUD:** A sci-fi inspired web interface (FastAPI + SSE) to visualize the Agent's "Thought Process" and future plan in real-time.
*   **Tool Use:** Can launch apps, type, click, scroll, run CMD commands, and write code.
*   **Resource Efficient:** Optimized to run on consumer hardware (tested on a standard Windows 11 machine).
*   **Dynamic Focus Logic:** To reduce hallucinations and token usage, the agent doesn't just look at the full desktop. It dynamically crops its vision to the **active window its using right now**, increasing model resolution for specific buttons while filtering out background noise.

---

## 🛠️ Installation & Setup

### Prerequisites
*   **Windows 10/11**
*   **Python 3.10+**
*   **Redis Server** (Must be running locally on port 6379)
*   **Tesseract OCR** (Install and set path in config)
*   **Google Gemini API Key**

### 1. Clone the Repository


Navigate into the project directory and install the required packages:

```bash
cd PAI-OS-Agent
pip install -r requirements.txt
```

### 2. Configuration
Create a `.env` file in the root directory. You can copy the following template:

```ini
GOOGLE_API_KEY=your_gemini_key_here
IPC_HOST=localhost
IPC_PORT=6000
IPC_AUTHKEY=supersecretkey
# Adjust path if necessary (Default location for Windows)
TESSERACT_PATH=C:\Program Files\Tesseract-OCR\tesseract.exe
LOG_LEVEL=INFO
```

### 3. Run the System
Start the main entry point. This will automatically spin up the UI Server, connect to Redis, and launch the Web Interface.

```bash
python -m src.main
```

Once the server is running, open your browser and navigate to:
`http://localhost:8000`

## 🧪 Research Findings & Reflections

Building this agent taught me that while **Agentic Workflows** are the future, current LLMs are not yet fully "OS-ready" for reliable, unsupervised use. Here are my key takeaways from testing:

### 1. The Reliability Gap
Even with my **Hybrid Architecture** (UIA + OCR), hallucinations remain a challenge, especially with faster, smaller models like *Flash Lite*. My implementation of **"Focus Logic"** (cropping the view to the active window) significantly reduced errors, but the models still occasionally misinterpret UI states or "invent" successful outcomes when an action actually failed.

### 2. Why I rejected "Grid" & "Coordinate Prediction"
I consciously decided **against** letting the LLM guess X/Y coordinates or using a "Grid Overlay".
*   **Coordinate Guessing** is too unreliable with current vision models.
*   **Grids** force a trade-off: Large cells allow the model to see context but lack precision; small cells offer precision but obscure the UI.
*   **My Solution:** The *Object-Oriented approach* (UIA Handles) is harder to implement but allows for exact interaction—when the model actually finds the element.

### 3. The Latency Bottleneck
Speed is critical for OS control. Waiting 3-5 seconds for a "click" breaks the flow. I see huge potential in the **Gemini Live API**. If Google expands "Live" capabilities from native audio to **native multimodal/screen streaming**, we could achieve near-zero latency control.

**Conclusion:** This project is a robust foundation. I will continue to test it against future model generations. The architecture is ready; we are just waiting for the models to catch up.


## 🚧 Project Status

**This repository serves as a Portfolio Showcase.**

I developed the PAI OS Agent as a "Deep Dive" into Agentic AI workflows and as a research project. It demonstrates my ability to architect complex, multi-modal systems under resource constraints.

Please note that this project is **not actively maintained**. It stands as a functional "Proof of Concept" representing my engineering skillset at this time. I am currently directing my focus toward new academic challenges and next-generation AI research.


### 🧠 Technical Deep Dive: Engineering the "Brain"

While many agents rely on simple `screenshot -> LLM -> click` loops, PAI OS employs a sophisticated architecture designed for **reliability in multi-step workflows** (e.g., "Open Calculator, perform calculation, copy result to Notepad").

#### 1. The Perception Pipeline (Hybrid Vision)
Standard Vision-Language Models (VLMs) struggle with small text and low-contrast UI elements. To solve this, I implemented a multi-stage perception stack:
*   **Layer 1: Structural Analysis (UIA):** The agent first queries the Windows Accessibility Tree via `pywinauto`. This provides exact coordinates and object types (Buttons, TextFields) without "guessing."
*   **Layer 2: Neural OCR Fallback:** If the UIA tree is empty (common in Electron apps, Games, or VMs), the system automatically triggers **EasyOCR (CUDA accelerated)**.
*   **Coordinate Transformation:** The system maps relative OCR bounding boxes back to absolute global screen coordinates, allowing the agent to click buttons inside a virtual machine window as if they were native controls.

#### 2. Focus Logics & Token Efficiency
Sending a full 4K desktop screenshot to an LLM is inefficient and vulnerable to hallucinations. PAI OS implements **Dynamic Viewport Cropping**:
*   The agent tracks the `Active Window Handle`.
*   Before inference, the visual pre-processor crops the screenshot to the active window bounds (plus a small margin).
*   **Result:** The model sees the target application in high resolution while background noise is eliminated. This significantly increases success rates for complex UIs while reducing token costs.

#### 3. Strategic State Management (The "Sprint" Protocol)
To prevent the common "Agent Loop of Death" (repeating the same failed click), the logic is split:
*   **The Planner (Gemini Pro):** Holds the "Long Term Memory" (`grounding_notes`). It creates a high-level **Sprint Plan** (a list of 5-10 atomic actions) and defines a **Success Condition** (e.g., "Calculator window is visible").
*   **The Executor (Gemini Flash):** Executes the sprint blindly but fast.
*   **Correction Logic:** If the Executor finishes the sprint but the Planner's `Success Condition` is not met via visual verification, the Planner analyzes the failure, updates the `grounding_notes` (e.g., *"Button X was disabled"*), and generates a *new* strategy instead of retrying blindly.

#### 4. Asynchronous IPC Architecture
Python's Global Interpreter Lock (GIL) often causes UI freezes when running heavy logic.
*   **Separation of Concerns:** The main Agent runs in one process, while the OS Interaction Logic runs in a completely separate process (`server.py`).
*   **Communication:** They exchange data via a custom IPC protocol (using `multiprocessing.connection`). This ensures the UI Analysis Server can poll 60Hz mouse updates while the Agent waits for API responses, preventing "Application Not Responding" states.


## 🍜 About the Project (and me)

I am a **17-year-old student** from Germany with a passion for AI Engineering.

This project started as a research experiment to see if LLMs are ready for reliable OS control. working with a limited budget (funding API credits by working part-time as a Ramen cook), I had to optimize for **efficiency and reliability** rather than brute-forcing with expensive models.

The result is a robust agentic architecture that handles latency, hallucinations, and state management effectively.

**Future Goals:**
*   Implementing better models and do further testing when I have the ressources
*   Wait for the new Gemini Flash Lite Version


---

## ⚠️ Disclaimer
This tool executes real clicks and keystrokes on your operating system. Run it with caution, ideally in a controlled environment or VM.

---
If you have ANY questions regarding this project, please feel free to reach out!

*MIT License - Created by Pio Borgelt*
