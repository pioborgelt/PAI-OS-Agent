"""
Core Agent Functions Module
===========================

This module provides the essential backend functionality for the hybrid AI agent.
It acts as the bridge between the high-level reasoning agent and the low-level
Operating System API (Windows), handling screen perception, OCR, and input simulation.

⚠️ **IMPORTANT READ** ⚠️
Before using this module, please read the **README.md** in the root directory.
This system requires a specific environment setup, including a running IPC server,
Redis, and specific Tesseract/CUDA configurations.

**DEPRECATION NOTICE (WEB / SELENIUM):**
All functions related to Selenium and direct Web-Driver interaction (e.g.,
`observe_web_state`, `execute_web_action`) are currently **DEPRECATED**.
They are present for legacy reference only and are not part of the active
autonomous loop in this version. Web capabilities may be reintroduced in future releases.

**Author:** Pio Borgelt
**Repository:** OS Agent
"""

import base64
import configparser
import io
import logging
import multiprocessing
import os
import subprocess
import time
import uuid
import socket
import json
import ctypes
from typing import Dict, Tuple, List, Optional
from multiprocessing.connection import Client

import mss
import pyautogui
import cv2
import numpy
import torch
import easyocr
from PIL import Image, ImageDraw, ImageFont, ImageOps

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, WebDriverException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.remote.remote_connection import RemoteConnection


from src.utils import CONFIG, get_logger


try:
    import pytesseract

    tess_path = CONFIG.get("TESSERACT_PATH", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    if os.path.exists(tess_path):
        pytesseract.pytesseract.tesseract_cmd = tess_path
    else:
        pytesseract = None
except ImportError:
    pytesseract = None

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


agent_logger = get_logger('Core')


ocr_reader = None
try:
    agent_logger.info("[Core] Initializing EasyOCR model...")
    if torch.cuda.is_available():
        gpu_count = torch.cuda.device_count()
        gpu_name = torch.cuda.get_device_name(0)
        agent_logger.info(f"[Core] ✅ CUDA DETECTED! Found {gpu_count} device(s).")
        agent_logger.info(f"[Core] 🚀 Using GPU: {gpu_name} for accelerated OCR.")
        use_gpu = True
    else:
        agent_logger.warning("[Core] ⚠️ CUDA NOT DETECTED. Falling back to CPU. OCR will be slow.")
        use_gpu = False

  
    ocr_reader = easyocr.Reader(['de', 'en'], gpu=use_gpu, verbose=False)
    agent_logger.info(f"[Core] EasyOCR initialized successfully.")

except Exception as e:
    agent_logger.error(f"[Core] Failed to initialize EasyOCR: {e}")
    ocr_reader = None



_APP_INDEX_CACHE = None


def perform_ocr_scan(pil_crop_img: Image.Image, global_offset_x: int, global_offset_y: int, start_id: int) -> list:
    """
    Executes an Optical Character Recognition (OCR) scan on a given image segment.

    Uses EasyOCR (CUDA-accelerated if available) to detect text elements.
    Critically, this function transforms local image coordinates back into
    global screen coordinates so the Agent can click them accurately.

    Args:
        pil_crop_img (Image.Image): The image segment to analyze.
        global_offset_x (int): X-offset of the image relative to the main screen.
        global_offset_y (int): Y-offset of the image relative to the main screen.
        start_id (int): The starting integer ID for newly detected elements to avoid collision.

    Returns:
        list: A list of dictionaries representing detected text elements with
              'OCR_TEXT' type and absolute screen coordinates.
    """
    if ocr_reader is None:
        return []

    
    cv_img = cv2.cvtColor(numpy.array(pil_crop_img), cv2.COLOR_RGB2BGR)


    try:
        results = ocr_reader.readtext(cv_img)
    except Exception as e:
        agent_logger.error(f"EasyOCR Failed: {e}")
        return []

    ocr_elements = []

    for (bbox, text, prob) in results:
    
        tl, tr, br, bl = bbox
        x1, y1 = map(int, tl)
        x2, y2 = map(int, br)
        
        abs_x1 = int(global_offset_x + x1)
        abs_y1 = int(global_offset_y + y1)
        abs_x2 = int(global_offset_x + x2)
        abs_y2 = int(global_offset_y + y2)

        abs_rect = (abs_x1, abs_y1, abs_x2, abs_y2)

        ocr_elements.append({
            "id": start_id + len(ocr_elements),
            "name": text,
            "type": "OCR_TEXT",
            "absolute_rectangle": abs_rect,
            "top_level_handle": None,
            "confidence": prob
        })

    return ocr_elements


def observe_os_state(target_monitor_num: int, focus_handle: int = None) -> tuple[Image.Image, list, tuple[int, int]]:
    """
    Captures the current state of the Operating System combining visual and structural data.

    This function performs a hybrid scan:
    1.  **Visual:** Takes a screenshot using MSS.
    2.  **Structural:** Requests a UI Automation (UIA) tree dump via the IPC server.
    3.  **Fallback:** If UIA data is sparse (e.g., inside a VM or Game), it triggers
        a GPU-accelerated OCR scan using EasyOCR.

    Args:
        target_monitor_num (int): The index of the monitor to capture.
        focus_handle (int, optional): If provided, restricts UIA analysis to this specific window handle
                                      for performance optimization.

    Returns:
        tuple:
            - PIL.Image.Image: The raw screenshot.
            - list: A list of detected UI elements (merged from UIA and OCR).
            - tuple: Global (x, y) offset of the monitor/window.
    """
 
    ipc_host = CONFIG["IPC_HOST"]
    ipc_port = CONFIG["IPC_PORT"]
    ipc_key = CONFIG["IPC_AUTHKEY"]

    t_start_total = time.perf_counter()
    agent_logger.info(f"[Core] Capturing OS state. Focus Handle: {focus_handle}")
    
    with mss.mss() as sct:
        try:
            monitor = sct.monitors[target_monitor_num]
        except IndexError:
            monitor = sct.monitors[1]
            
        sct_img = sct.grab(monitor)
        pil_img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
        monitor_offset = (monitor['left'], monitor['top'])

    t_start_uia = time.perf_counter()
    raw_elements = _trigger_ipc_analysis(ipc_host, ipc_port, ipc_key, root_handle=focus_handle)
    t_end_uia = time.perf_counter()
    
    detected_elements = []
    
    monitor_rect = (monitor['left'], monitor['top'], monitor['left'] + monitor['width'], monitor['top'] + monitor['height'])
    
    total_uia_count = 0
    uia_blockers = [] 

    blocker_types = [
        "Button", "Edit", "CheckBox", "RadioButton", "ComboBox", 
        "List", "ListItem", "MenuItem", "TabItem", "Hyperlink", 
        "TreeItem", "DataItem", "HeaderItem", "Document", "Text", "TitleBar"
    ]

    force_ocr_vm = False
    if raw_elements:
        root_name = raw_elements[0].get('name', '').lower()
        
        if "oracle" in root_name or "virtualbox" in root_name:
            force_ocr_vm = True
            agent_logger.info(f"[Core] Detected 'Oracle'/'VirtualBox' in window title ('{root_name}'). Forcing OCR.")
        
        elif "vm" in root_name:
            is_vm_word = " vm " in root_name or root_name.startswith("vm ") or root_name.endswith(" vm") or root_name == "vm"
            if is_vm_word:
                force_ocr_vm = True
                agent_logger.info(f"[Core] Detected 'VM' keyword in window title ('{root_name}'). Forcing OCR.")

    for el_data in raw_elements:
        rect = el_data['rectangle_coords'] 

        if not focus_handle:
            if not _is_rect_on_monitor(rect, monitor_rect):
                continue

        el_type = el_data.get('type', '')
        if not el_type: el_type = el_data.get('control_type', '')

        detected_elements.append({
            "id": len(detected_elements),
            "name": el_data.get('name', ''),
            "type": el_type,
            "automation_id": el_data.get('automation_id'),
            "absolute_rectangle": rect,
            "top_level_handle": el_data.get('top_level_handle')
        })

        if "Window" not in el_type and "Pane" not in el_type:
            total_uia_count += 1
        
        if any(t in el_type for t in blocker_types):
            uia_blockers.append(rect)


    run_ocr = False
    
    if force_ocr_vm:
        run_ocr = True
    elif total_uia_count < 5:
        agent_logger.info(f"[Core] Low UI density (Count: {total_uia_count}). Activating OCR fallback.")
        run_ocr = True
    else:
        agent_logger.info(f"[Core] High UI density (Count: {total_uia_count}). Native App detected. SKIPPING OCR.")

    t_ocr_duration = 0.0

    if run_ocr:
        t_start_ocr = time.perf_counter()
        ocr_start_id = 9000
        if len(detected_elements) > 0:
            ocr_start_id = max(9000, detected_elements[-1]['id'] + 100)

        final_offset_x = monitor_offset[0]
        final_offset_y = monitor_offset[1]
        
        raw_ocr = perform_ocr_scan(pil_img, final_offset_x, final_offset_y, ocr_start_id)
        
        skipped_ocr = 0
        for ocr_el in raw_ocr:
            o_rect = ocr_el['absolute_rectangle']
            cx = (o_rect[0] + o_rect[2]) // 2
            cy = (o_rect[1] + o_rect[3]) // 2
            
            is_covered = False
            

            if not force_ocr_vm:
                for u_rect in uia_blockers:
                    if (u_rect[0] <= cx <= u_rect[2]) and (u_rect[1] <= cy <= u_rect[3]):
                        is_covered = True
                        break
            
            if not is_covered:
                detected_elements.append(ocr_el)
            else:
                skipped_ocr += 1
        
        t_end_ocr = time.perf_counter()
        t_ocr_duration = t_end_ocr - t_start_ocr
        if skipped_ocr > 0:
            agent_logger.info(f"[Core] Filtered {skipped_ocr} OCR elements overlapped by UIA.")
    
    t_total = time.perf_counter() - t_start_total
    agent_logger.info(f"[Core-Timing] Capture Total: {t_total:.2f}s | UIA: {(t_end_uia - t_start_uia):.2f}s | OCR: {t_ocr_duration:.2f}s")
    
    return pil_img, detected_elements, monitor_offset


def get_system_app_index() -> Dict[str, str]:
    """
    Builds a searchable index of installed applications using PowerShell.

    On the first run, it uses `Get-StartApps` to map human-readable names
    (e.g., "Google Chrome") to execution commands (e.g., "chrome.exe" or AppIDs).
    Results are cached in memory (`_APP_INDEX_CACHE`) to prevent delay in subsequent calls.

    Also includes a hardcoded manual override list for common Windows tools
    (explorer, calc, notepad, settings) to ensure reliability.

    Returns:
        Dict[str, str]: A dictionary mapping lowercase app names to launch commands.
    """
    global _APP_INDEX_CACHE
    
    if _APP_INDEX_CACHE is not None:
        return _APP_INDEX_CACHE

    agent_logger.info("[AppIndex] Indexing installed applications via PowerShell (First Run)...")
    app_map = {}


    manual_entries = {
        "explorer": "explorer.exe",
        "datei-explorer": "explorer.exe",
        "ordner": "explorer.exe",
        "cmd": "cmd.exe",
        "terminal": "wt.exe",
        "powershell": "powershell.exe",
        "notepad": "notepad.exe",
        "editor": "notepad.exe",
        "calc": "calc.exe",
        "calculator": "calc.exe",
        "rechner": "calc.exe",
        "chrome": "chrome.exe",
        "firefox": "firefox.exe",
        "edge": "msedge.exe",
        "browser": "msedge.exe",
        "task manager": "taskmgr.exe",
        "control panel": "control.exe",
        "systemsteuerung": "control.exe",
        "settings": "ms-settings:",          
        "einstellungen": "ms-settings:",
        "ausführen": "explorer.exe Shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}",
        "snipping tool": "snippingtool.exe",
        "vscode": "code",
        "code": "code",
        "spotify": "spotify"
    }
    app_map.update(manual_entries)

    ps_command = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Get-StartApps | Select-Object Name, AppID | ConvertTo-Json"
    
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_command],
            capture_output=True, 
            text=True, 
            encoding='utf-8', 
            errors='replace', 
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        
        if result.returncode == 0 and result.stdout and result.stdout.strip():
            try:
                data = json.loads(result.stdout)
                if isinstance(data, dict): data = [data]
                
                for entry in data:
                    name = entry.get("Name", "")
                    app_id = entry.get("AppID", "")
                    
                    if name and app_id:
                        clean_name = name.strip().lower()
                        if clean_name not in app_map:
                            app_map[clean_name] = f"shell:AppsFolder\\{app_id}"
                            
            except json.JSONDecodeError as je:
                agent_logger.error(f"[AppIndex] JSON Parse Error: {je}")
        else:
            if result.stderr:
                agent_logger.warning(f"[AppIndex] PowerShell Error: {result.stderr}")

    except Exception as e:
        agent_logger.error(f"[AppIndex] Critical fail: {e}")

    agent_logger.info(f"[AppIndex] Scan complete. Found {len(app_map)} applications.")
    
    _APP_INDEX_CACHE = app_map
    return app_map


async def observe_web_state(driver: webdriver.Firefox) -> tuple[Image.Image, list]:
    agent_logger.info("[Core] Capturing Web state via atomic JS call.")
    try:
        screenshot_bytes = driver.get_screenshot_as_png()
        pil_img = Image.open(io.BytesIO(screenshot_bytes))
    except Exception as e:
        agent_logger.error(f"[Core] Screenshot failed: {e}")
        return None, []

    js_script = """
    const getPathTo = (element) => {
        if (element.id) return `id("${element.id}")`;
        if (element === document.body) return element.tagName.toLowerCase();
        let ix = 0;
        const siblings = element.parentNode.childNodes;
        for (let i = 0; i < siblings.length; i++) {
            const sibling = siblings[i];
            if (sibling === element)
                return getPathTo(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';
            if (sibling.nodeType === 1 && sibling.tagName === element.tagName)
                ix++;
        }
    };

    const elements = document.querySelectorAll("a, button, input, textarea, select, [role='button'], [role='link'], [onclick], [contenteditable='true']");
    const elementData = [];
    
    elements.forEach(el => {
        const rect = el.getBoundingClientRect();
        if ((rect.width === 0 || rect.height === 0) || el.disabled) {
            return;
        }

        const tag = el.tagName.toLowerCase();
        let supported_actions = ['klicken'];
        if (tag === 'input' || tag === 'textarea' || el.isContentEditable) {
            supported_actions.push('schreiben');
        }

        const style = window.getComputedStyle(el);
        const isVisible = style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;

        elementData.push({
            tag: tag,
            text: (el.textContent || el.ariaLabel || el.title || "").trim().substring(0, 100),
            xpath: getPathTo(el),
            rect: { x: rect.left, y: rect.top, width: rect.width, height: rect.height },
            supported_actions: supported_actions,
            is_visible: isVisible
        });
    });
    return elementData;
    """

    try:
        raw_elements = driver.execute_script(js_script)
        elements = []
        for i, el_data in enumerate(raw_elements):
            rect_data = el_data['rect']
            el_data['id'] = i
            el_data['rect'] = (rect_data['x'], rect_data['y'], rect_data['x'] + rect_data['width'], rect_data['y'] + rect_data['height'])
            elements.append(el_data)
        
        agent_logger.info(f"[Core] {len(elements)} interactive Web elements found.")
        return pil_img, elements
    except Exception as e:
        agent_logger.error(f"[Core] Error executing JS script: {e}", exc_info=True)
        return pil_img, []


def _filter_nested_elements(elements: list) -> list:

    if not elements:
        return []

    def get_rect(el):
        return el.get('rect') or el.get('absolute_rectangle')

    def get_area(r):
        if not r: return 0
        return (r[2] - r[0]) * (r[3] - r[1])

    with_area = []
    for el in elements:
        r = get_rect(el)
        if r:
            with_area.append({'el': el, 'area': get_area(r), 'rect': r})


    with_area.sort(key=lambda x: x['area'])

    kept_elements = []

    for item in with_area:
        current_el = item['el']
        current_rect = item['rect']
        current_area = item['area']
        
        if current_area <= 0: continue

        is_redundant = False
        
        for kept_item in kept_elements:
            kept_rect = kept_item['rect']
            kept_area = kept_item['area']


            if (kept_rect[0] >= current_rect[0] - 5 and  
                kept_rect[1] >= current_rect[1] - 5 and
                kept_rect[2] <= current_rect[2] + 5 and
                kept_rect[3] <= current_rect[3] + 5):
                
                coverage_ratio = kept_area / current_area
                if coverage_ratio > 0.80:
                    is_redundant = True
                    break
        
        if not is_redundant:
            kept_elements.append(item)

    return [x['el'] for x in kept_elements]


def prepare_images_for_model(img: Image.Image, elements: list, offset: tuple, step: int = 0) -> tuple[bytes, bytes]:
    """
    Prepares visual data for the Vision Language Model (VLM).

    This function creates two versions of the screen:
    1.  Clean: The raw screenshot (for the AI to see content).
    2.  Annotated: The screenshot with bounding boxes and numeric IDs drawn
        over interactive elements.

    It also filters overlapping/redundant elements to reduce visual clutter (token usage)
    for the AI model.

    Args:
        img (Image.Image): The source screenshot.
        elements (list): List of detected UI elements (with coordinates).
        offset (tuple): (x, y) screen offset to align coordinates.
        step (int): Current step counter (used for debug filename generation).

    Returns:
        tuple: (clean_image_bytes, annotated_image_bytes) in PNG format.
    """
    clean_buffer = io.BytesIO()
    img.save(clean_buffer, format="PNG")
    clean_image_bytes = clean_buffer.getvalue()


    visible_elements = _filter_nested_elements(elements)

    annotated_img = img.copy()                
    draw = ImageDraw.Draw(annotated_img)
    
    try:
        font = ImageFont.truetype("arialbd.ttf", 16) 
    except IOError:
        try:
            font = ImageFont.truetype("arial.ttf", 16)
        except:
            font = ImageFont.load_default()

    monitor_offset_x, monitor_offset_y = offset
    img_width, img_height = img.size

    debug_dir = os.path.join(CONFIG.get("LOG_DIR", "logs"), "debug_annotated")
    os.makedirs(debug_dir, exist_ok=True)

    for el in visible_elements:
        el_id = str(el.get('id', '?'))
        
        rect_key = 'rect' if 'rect' in el else 'absolute_rectangle'
        abs_rect = el.get(rect_key)
        
        if not abs_rect: 
            continue

        x1 = abs_rect[0] - monitor_offset_x
        y1 = abs_rect[1] - monitor_offset_y
        x2 = abs_rect[2] - monitor_offset_x
        y2 = abs_rect[3] - monitor_offset_y


        if x2 < 0 or y2 < 0 or x1 > img_width or y1 > img_height:
            continue

        is_ocr = (el.get("type") == "OCR_TEXT")
        color = "#FF00FF" if is_ocr else "#FF2222"
        
        draw.rectangle((x1, y1, x2, y2), outline=color, width=2)
        
        label_text = str(el_id)
        tb = draw.textbbox((0, 0), label_text, font=font)
        text_w = tb[2] - tb[0]
        text_h = tb[3] - tb[1]
        
        pad = 2

        if is_ocr:
            text_x = x1
            text_y = y1 - text_h - (pad * 2)
            if text_y < 0:
                text_y = y2 + 2
        else:
            text_x = x1 + pad
            text_y = y1 + pad
            
            box_w = x2 - x1
            box_h = y2 - y1
            
            if box_w < text_w + 4:
                text_x = x1 + (box_w - text_w) // 2
            if box_h < text_h + 4:
                text_y = y1 + (box_h - text_h) // 2

        if text_x < 0: text_x = 0
        if text_x + text_w > img_width: text_x = img_width - text_w
        
        if not is_ocr or (is_ocr and text_y >= 0 and text_y + text_h <= img_height):
             if text_y < 0: text_y = 0
             if text_y + text_h > img_height: text_y = img_height - text_h

        draw.rectangle(
            (text_x - pad, text_y - pad, text_x + text_w + pad, text_y + text_h + pad), 
            fill=color, outline=None
        )
        draw.text((text_x, text_y), label_text, fill="black", font=font)

    try:
        filename = f"step_{step}_annotated_{int(time.time())}.png"
        filepath = os.path.join(debug_dir, filename)
        annotated_img.save(filepath)
    except Exception as e:
        print(f"Failed to save debug image: {e}")

    annotated_buffer = io.BytesIO()
    annotated_img.save(annotated_buffer, format="PNG") 
    annotated_image_bytes = annotated_buffer.getvalue()
    
    return clean_image_bytes, annotated_image_bytes


async def execute_os_action(action: dict, all_elements: list, cursor_queue: multiprocessing.Queue) -> Dict | None:
    """
    Executes a physical action on the Operating System based on the Agent's decision.

    Handles mapping abstract commands (e.g., "click element 42") to low-level
    inputs (PyAutoGUI or IPC interaction).
    
    Capabilities include:
    - Mouse interaction (Click, Double Click, Right Click) via Coordinates or UIA Handles.
    - Keyboard input (Type text).
    - Application management (Launch, Close, Focus).
    - Command Line execution (User level).

    Args:
        action (dict): The command dictionary (e.g., {'command': 'click', 'element_id': 10}).
        all_elements (list): Current list of UI elements to resolve IDs to coordinates/handles.
        cursor_queue (multiprocessing.Queue): Queue for UI feedback (e.g., moving a fake cursor on the dashboard).

    Returns:
        Dict | None: Result dictionary (e.g., {'status': 'success'}) or None.
    """
    ipc_host = CONFIG["IPC_HOST"]
    ipc_port = CONFIG["IPC_PORT"]
    ipc_key = CONFIG["IPC_AUTHKEY"]

    command = action.get("command")
    description = action.get("description", "OS action...")
    agent_logger.info(f"[Core] >>> START ACTION: '{command}'")

    if command in ["click", "double_click", "right_click"]:
        element_id = action.get("element_id")
        target = next((el for el in all_elements if el['id'] == element_id), None)
        
        if not target:
            agent_logger.error(f"[Core] FAIL: Element ID {element_id} not found in current UI state.")
            return

        agent_logger.info(f"[Core] Target identified: ID={element_id}, Name='{target.get('name')}', Type='{target.get('type')}'")
        
        if target.get("type") == "OCR_TEXT":
            rect = target.get('absolute_rectangle')
            if rect:
                center_x = (rect[0] + rect[2]) // 2
                center_y = (rect[1] + rect[3]) // 2
                
                agent_logger.info(f"[Core] MODE: OCR Direct Click at ({center_x}, {center_y})")
                
                if cursor_queue:
                    cursor_queue.put({'action': 'move', 'x': center_x, 'y': center_y, 'text': description})
                    time.sleep(0.3) 

                try:
                    pyautogui.moveTo(center_x, center_y)
                    if command == "click":
                        pyautogui.click()
                    elif command == "double_click":
                        pyautogui.doubleClick()
                    elif command == "right_click":
                        pyautogui.rightClick()
                except Exception as e:
                    agent_logger.error(f"[Core] PyAutoGUI failed: {e}")
            return

        if not target.get('top_level_handle'): 
            if target.get('absolute_rectangle'):
                 agent_logger.warning("[Core] FALLBACK: Using blind coordinates click because handle is missing.")
                 rect = target.get('absolute_rectangle')
                 cx, cy = (rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2
                 pyautogui.click(cx, cy)
            return

        target_rect = target.get('absolute_rectangle')
        if target_rect and cursor_queue:
            center_x, center_y = (target_rect[0] + target_rect[2]) // 2, (target_rect[1] + target_rect[3]) // 2
            cursor_queue.put({'action': 'move', 'x': center_x, 'y': center_y, 'text': description})
            time.sleep(0.2) 

        interaction_map = {
            "click": "click",
            "double_click": "double_click",
            "right_click": "right_click"
        }
        
        _trigger_ipc_interaction(
            ipc_host, ipc_port, ipc_key,
            target['top_level_handle'], 
            target.get('automation_id'), 
            target.get('name'), 
            target.get('type'), 
            interaction_map[command],
            target_rect=target_rect 
        )

    elif command == "execute_cmd":
        cmd_str = action.get("cmd_line")
        agent_logger.info(f"[Core] Running CMD (User Level): {cmd_str}")
        try:
            result = subprocess.run(
                cmd_str, shell=True, capture_output=True, text=True, timeout=15,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            output = (result.stdout + "\n" + result.stderr).strip()
            if not output: output = "Command executed successfully (No Output)."
            return {"status": "success", "output": output}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    elif command == "scroll":
        direction = action.get("direction", "down")
        if direction == "down": pyautogui.press('pagedown')
        else: pyautogui.press('pageup')
        time.sleep(0.5)

    elif command == "type":
        text_to_type = action.get("text", "")
        element_id = action.get("element_id")

        if element_id is not None:
            target = next((el for el in all_elements if el['id'] == element_id), None)
            
            if target and target.get("type") == "OCR_TEXT":
                 rect = target.get('absolute_rectangle')
                 if rect:
                     center_x, center_y = (rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2
                     if cursor_queue:
                         cursor_queue.put({'action': 'move', 'x': center_x, 'y': center_y, 'text': 'Clicking to focus text...'})
                     pyautogui.click(center_x, center_y)
                     time.sleep(0.3) 
                     pyautogui.write(text_to_type, interval=0.02)
                 return

            if not target or not target.get('top_level_handle'): 
                agent_logger.warning("[Core] Type requested but target invalid. Typing blindly.")
                pyautogui.write(text_to_type, interval=0.01)
                return

            target_rect = target.get('absolute_rectangle')
            

            _trigger_ipc_interaction(
                ipc_host, ipc_port, ipc_key,
                target['top_level_handle'], target.get('automation_id'), target.get('name'), target.get('type'), 
                'click', text_to_type=None, target_rect=target_rect
            )
            time.sleep(0.2)
            

            _trigger_ipc_interaction(
                ipc_host, ipc_port, ipc_key,
                target['top_level_handle'], target.get('automation_id'), target.get('name'), target.get('type'), 
                'type', text_to_type=text_to_type, target_rect=target_rect
            )
        else:
            pyautogui.write(text_to_type, interval=0.01)

    elif command == "launch_app":
            app_name_input = action.get("app_name", "").lower().strip()
            if not app_name_input: raise ValueError("'launch_app' requires 'app_name'.")

            windows_before = get_all_windows_from_server(ipc_host, ipc_port, ipc_key)
            handles_before = {w['handle'] for w in windows_before}

            direct_aliases = {
                "settings": "ms-settings:", "einstellungen": "ms-settings:",
                "calc": "calc.exe", "rechner": "calc.exe", "calculator": "calc.exe",
                "terminal": "wt.exe", "cmd": "cmd.exe",
                "explorer": "explorer.exe", "datei": "explorer.exe",
                "code": "code", "vscode": "code",
                "notepad": "notepad.exe", "editor": "notepad.exe",
                "browser": "msedge.exe", "edge": "msedge.exe",
                "chrome": "chrome.exe", "firefox": "firefox.exe"
            }

            target_cmd = direct_aliases.get(app_name_input)
            
            if not target_cmd:
                app_map = get_system_app_index()
                target_cmd = app_map.get(app_name_input)
                
                if not target_cmd:
                    candidates = []
                    for name, cmd in app_map.items():
                        score = 0
                        if app_name_input in name: score += 100 - (len(name) - len(app_name_input))
                        elif name in app_name_input: score += 80
                        if score > 0: candidates.append((score, cmd, name))
                    
                    if candidates:
                        candidates.sort(key=lambda x: x[0], reverse=True)
                        target_cmd = candidates[0][1]
            
            if not target_cmd: 
                if " " in app_name_input and not app_name_input.startswith('"'): target_cmd = f'"{app_name_input}"'
                else: target_cmd = app_name_input

            agent_logger.info(f"[Core] Launching '{app_name_input}' via command: '{target_cmd}'... Waiting for window.")
            
            try:
                if target_cmd.lower().startswith("shell:") or "://" in target_cmd or target_cmd.lower().startswith("ms-"):
                    final_cmd = f'explorer.exe "{target_cmd}"'
                    subprocess.Popen(final_cmd, shell=True)
                else:
                    subprocess.Popen(target_cmd, shell=True)
            except Exception as e:
                agent_logger.error(f"[Core] Failed to launch process: {e}")
                return {"status": "error", "message": f"Failed to launch: {e}"}

            new_window_handle = None
            new_window_title = ""
            max_attempts = 20
            
            for i in range(max_attempts): 
                time.sleep(1.0)
                windows_after = get_all_windows_from_server(ipc_host, ipc_port, ipc_key)
                handles_after = {w['handle'] for w in windows_after}
                
                new_handles = handles_after - handles_before
                
                if new_handles:
                    agent_logger.info(f"[Core] Window Diff Scan {i+1}/{max_attempts}: Found new handles: {new_handles}")
                    
                    candidates = [w for w in windows_after if w['handle'] in new_handles]
                    
                    valid_candidates = []
                    for cand in candidates:
                        w_w = cand['rect'][2] - cand['rect'][0]
                        w_h = cand['rect'][3] - cand['rect'][1]
                        title_lower = cand['title'].lower()
                        
                
                        if w_w > 50 and w_h > 50:
                            score = 0
                            if app_name_input in title_lower: score += 10
                            if "oracle" in title_lower or "virtualbox" in title_lower: score += 5
                            if "vm" in title_lower: score += 2
                            valid_candidates.append((score, cand))
                    
                    if valid_candidates:
                        valid_candidates.sort(key=lambda x: x[0], reverse=True)
                        best_match = valid_candidates[0][1]
                        new_window_handle = best_match['handle']
                        new_window_title = best_match['title']
                        break
            
            if new_window_handle:
                agent_logger.info(f"[Core] SUCCESS: Detected new window: '{new_window_title}' (Handle: {new_window_handle})")
                bring_window_to_front(new_window_handle)
                return {
                    "status": "success", 
                    "action_type": "launched",
                    "handle": new_window_handle, 
                    "detected_title": new_window_title
                }
            else:
                agent_logger.warning(f"[Core] Time out ({max_attempts}s). App launched, but no NEW window detected via Diff.")
                return {"status": "success", "message": "Command executed, but window detection timed out."}

    elif command == "focus_window":
        handle = action.get("element_id") 
        if handle:
            bring_window_to_front(handle)
            return {"status": "success", "message": "Window focused."}
    elif command == "close_window":
        handle_to_close = action.get("handle")
        if not handle_to_close:
            return {"status": "error", "message": "No handle provided for close_window."}
            
        agent_logger.info(f"[Core] Closing window with handle: {handle_to_close}")
        try:
            with Client((ipc_host, ipc_port), authkey=ipc_key) as conn:
                conn.send({'command': 'close_window', 'payload': {'handle': handle_to_close}})
                if conn.poll(timeout=2):
                    res = conn.recv()
                    return res
        except Exception as e:
            return {"status": "error", "message": str(e)}
    elif command == "press_enter":
        pyautogui.press('enter')
        time.sleep(0.5)
        
    elif command == "wait":
        duration = max(1, min(int(action.get("duration", 3)), 30))
        time.sleep(duration)
        
    else:
        agent_logger.warning(f"[Core] Unknown command: {command}")
        
    time.sleep(0.5)


async def execute_web_action(action: dict, driver: webdriver.Firefox, cursor_queue: multiprocessing.Queue) -> Dict | None:
    command = action.get("command")
    description = action.get("description", "Web action is being executed...")
    agent_logger.info(f"[Core] Executing Web command '{command}': {description}")

    try: 
        if command == "web_navigate":
            url = action.get("url")
            if not url: raise ValueError("Command 'web_navigate' requires a 'url'.")
            driver.get(url)

        elif command in ["web_click", "web_type"]:
            xpath = action.get("xpath")
            if not xpath: raise ValueError(f"Command '{command}' requires an 'xpath'.")
            
            try:
                element = driver.find_element(By.XPATH, xpath)
            except NoSuchElementException:
                agent_logger.warning(f"[Core] Element with XPath '{xpath}' not found.")
                return None 
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", element)
            time.sleep(0.3)

            try:
                if cursor_queue:
                    window_pos = driver.get_window_position()
                    elem_loc = element.location
                    elem_size = element.size
                    center_x = window_pos.get('x', 0) + elem_loc.get('x', 0) + (elem_size.get('width', 0) / 2)
                    center_y = window_pos.get('y', 0) + elem_loc.get('y', 0) + (elem_size.get('height', 0) / 2)
                    cursor_queue.put({'action': 'move', 'x': int(center_x), 'y': int(center_y), 'text': description})
                    time.sleep(0.2)
            except Exception as cursor_error:
                 agent_logger.warning(f"[Core] Could not move cursor for Web action: {cursor_error}")

            if command == "web_click":
                driver.execute_script("arguments[0].click();", element)
            elif command == "web_type":
                text_to_type = action.get("text", "")
                element.click(); time.sleep(0.1)
                element.send_keys(Keys.CONTROL, 'a')
                element.send_keys(Keys.BACKSPACE)
                element.send_keys(text_to_type)
        
        elif command == "wait":
            duration = max(1, min(int(action.get("duration", 3)), 30))
            time.sleep(duration)

        elif command == "ask_user":
            question = action.get('question')
            if not question: raise ValueError("'ask_user' requires a 'question' parameter.")
            clarification_id = str(uuid.uuid4())
            agent_logger.info(f"Agent asks for clarification (ID: {clarification_id}): '{question}'")
            return {
                'type': 'clarification_needed', 
                'question': question,
                'clarification_id': clarification_id
            }
        
        time.sleep(1.0) 
        return None

    except (WebDriverException, StaleElementReferenceException) as e:
        error_msg = str(e).split('\n')[0]
        agent_logger.error(f"[Core] WEBDRIVER ERROR during '{command}': {error_msg}")
        raise e


def connect_to_web_agent(service_name: str, config_ports: dict) -> webdriver.Firefox:
    """
    Connects to a running Firefox instance via Marionette/GeckoDriver.
    """
    port = config_ports.get(service_name)
    if not port:
        raise ValueError(f"No port found for service '{service_name}' in configuration.")

    agent_logger.info(f"[Core] Attempting to connect to '{service_name}' agent on Marionette port {port}...")

    try:
        options = Options()
        driver = webdriver.Remote(
            command_executor=f'http://127.0.0.1:{port}',
            options=options
        )
        _ = driver.title 
        agent_logger.info(f"Successfully connected to running '{service_name}' instance via Marionette!")
        return driver
    except Exception as e:
        agent_logger.error(f"FATAL: Could not connect to '{service_name}' agent. Error: {e}", exc_info=True)
        raise ConnectionError(f"Connection to '{service_name}' agent failed.") from e


def _trigger_ipc_analysis(ipc_host: str, ipc_port: int, ipc_key: bytes, root_handle: int = None) -> list:
    try:
        with Client((ipc_host, ipc_port), authkey=ipc_key) as conn:
            conn.send({'command': 'analyze', 'payload': {'root_handle': root_handle}})
            
            timeout = 2.0 if root_handle else 120.0
            
            if conn.poll(timeout=timeout):
                result = conn.recv()
                if result.get("status") == "success":
                    return result.get("data", [])
                elif result.get("message") == "HandleInvalid":
                    raise ValueError("HandleInvalid")
                else:
                    return []
            else:
                if root_handle: raise ValueError("HandleInvalid")
                raise TimeoutError("No response from Analysis Server.")
    except Exception as e:
        if "HandleInvalid" in str(e): raise
        agent_logger.error(f"[Core] IPC Analysis Error: {e}")
        return []


def _trigger_ipc_interaction(ipc_host: str, ipc_port: int, ipc_key: bytes, top_handle: int, auto_id: str, name: str, control_type: str, interaction_type: str, text_to_type: str = None, target_rect: tuple = None) -> bool:
    try:  
        with Client((ipc_host, ipc_port), authkey=ipc_key) as conn:
            payload = {
                'top_level_handle': top_handle, 
                'automation_id': auto_id,
                'name': name, 
                'control_type': control_type,
                'interaction_type': interaction_type, 
                'text_to_type': text_to_type,
                'target_rect': target_rect
            }
            conn.send({'command': 'interact', 'payload': payload})
            if conn.poll(timeout=15):
                result = conn.recv()
                return result.get("status") == "success"
            else:
                agent_logger.error("[Core] IPC Interaction Timeout.")
                return False
    except Exception as e:
        agent_logger.error(f"[Core] IPC Interaction Error: {e}", exc_info=True)
        return False


def get_all_windows_from_server(ipc_host: str, ipc_port: int, ipc_key: bytes):
    try:
        with Client((ipc_host, ipc_port), authkey=ipc_key) as conn:
            conn.send({'command': 'get_window_list'})
            if conn.poll(timeout=2):
                res = conn.recv()
                if res.get("status") == "success":
                    return res.get("windows", [])
    except Exception as e:
        agent_logger.error(f"Failed to get window list: {e}")
    return []


def bring_window_to_front(handle: int):
    try:
        user32 = ctypes.windll.user32
        if user32.IsIconic(handle):
            user32.ShowWindow(handle, 9)
        else:
            user32.ShowWindow(handle, 5) 
        user32.SetForegroundWindow(handle)
        time.sleep(0.2)
    except Exception as e:
        agent_logger.error(f"Failed to focus window {handle}: {e}")


def get_ipc_active_window(ipc_host: str, ipc_port: int, ipc_key: bytes) -> dict:
    try:
        with Client((ipc_host, ipc_port), authkey=ipc_key) as conn:
            conn.send({'command': 'get_active_window'})
            if conn.poll(timeout=2):
                result = conn.recv()
                if result.get("status") == "success":
                    return result
    except Exception as e:
        agent_logger.error(f"[Core] Failed to get active window: {e}")
    return {}


def check_ipc_handle_exists(ipc_host: str, ipc_port: int, ipc_key: bytes, handle: int) -> bool:
    try:
        with Client((ipc_host, ipc_port), authkey=ipc_key) as conn:
            conn.send({'command': 'check_handle', 'payload': {'handle': handle}})
            if conn.poll(timeout=1):
                return conn.recv().get("exists", False)
    except:
        return False
    return False


def read_local_file(file_path: str) -> dict:
    agent_logger.info(f"[Core] Reading file: {file_path}")
    if not os.path.exists(file_path):
        return {"status": "error", "content": "File not found."}
    
    try:
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        return {"status": "success", "content": content}
    except Exception as e:
        return {"status": "error", "content": str(e)}


def write_to_local_file(file_path: str, content: str) -> dict:
    agent_logger.info(f"[Core] Writing to file: {file_path}")
    try:
        backup_path = file_path + ".bak"
        if os.path.exists(file_path) and not os.path.exists(backup_path):
            import shutil
            shutil.copy2(file_path, backup_path)
            
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return {"status": "success", "message": "File saved successfully."}
    except Exception as e:
        return {"status": "error", "message": str(e)}


def _is_rect_on_monitor(window_rect: tuple, monitor_rect: tuple) -> bool:
    win_center_x = (window_rect[0] + window_rect[2]) / 2
    win_center_y = (window_rect[1] + window_rect[3]) / 2
    return (monitor_rect[0] <= win_center_x < monitor_rect[2]) and \
           (monitor_rect[1] <= win_center_y < monitor_rect[3])

def get_handle_rect(ipc_addr: str, ipc_key: bytes, handle: int) -> tuple:
    try:
        with Client(ipc_addr, authkey=ipc_key) as conn:
            conn.send({'command': 'check_handle', 'payload': {'handle': handle}})
            if conn.poll(timeout=1):
                res = conn.recv()
                if res.get("exists") and res.get("rect"):
                    return res.get("rect")
    except Exception as e:
        agent_logger.error(f"[Core] Failed to validate handle rect: {e}")
    return None