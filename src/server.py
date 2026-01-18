"""
Analysis Server
===============

A standalone Windows UI Automation Server acting as the bridge between the 
OS Agent and the Operating System. It handles direct Win32 interactions, 
UIA tree scanning, and input simulation via a secure IPC listener.

This server executes strictly local operations and exposes an API for:
- Scanning the UI Element Tree (pywinauto/UIA).
- Performing interactions (Click, Type, Focus).
- retrieving low-level window states (Caret position, Handles).

WARNING:
    Before running this server, ensure you have read the README.md file
    in the root directory for configuration and dependency requirements.
    Running this allows remote control of the mouse/keyboard via the configured port.

Author: Pio Borgelt
"""

import sys
import os
import time
import json
import traceback
import re
import datetime
import ctypes
from ctypes import wintypes
from multiprocessing.connection import Listener
import pythoncom
import pywinauto
import win32con
import win32gui
import win32process
from pywinauto.uia_defines import NoPatternInterfaceError


try:
    sys.coinit_flags = 2
except AttributeError:
    pass

from src.utils import CONFIG, get_logger

logger = get_logger("AnalysisServer")

class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("hwndActive", wintypes.HWND),
        ("hwndFocus", wintypes.HWND),
        ("hwndCapture", wintypes.HWND),
        ("hwndMenuOwner", wintypes.HWND),
        ("hwndMoveSize", wintypes.HWND),
        ("hwndCaret", wintypes.HWND),
        ("rcCaret", wintypes.RECT)
    ]

def clean_text(text: str) -> str:
    if not isinstance(text, str):
        return ""
    return re.sub(r'[\u202a-\u202f\u2066-\u2069]', '', text)

def _dump_debug_info(elements, root_handle):

    try:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        handle_str = f"_handle_{root_handle}" if root_handle else "_full_desktop"
        filename = f"ui_dump_{timestamp}{handle_str}.json"
        filepath = os.path.join(CONFIG["DEBUG_DIR"], filename)

        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(elements, f, indent=2, ensure_ascii=False)
        
 
        files = sorted([os.path.join(CONFIG["DEBUG_DIR"], f) for f in os.listdir(CONFIG["DEBUG_DIR"])], key=os.path.getmtime)
        while len(files) > 20:
            os.remove(files.pop(0))
    except Exception as e:
        logger.error(f"Failed to dump debug info: {e}")

def fetch_raw_elements(root_handle=None):
    mode = f"STRICT TARGET (Handle: {root_handle})" if root_handle else "FULL DESKTOP SCAN"
    logger.info(f"Starting forensic analysis... Mode: {mode}")
    
    all_elements = []
    debug_elements = [] 
    
    try:
        desktop = pywinauto.Desktop(backend="uia")
        windows_to_scan = []

        if root_handle:
            try:
                target_win = desktop.window(handle=root_handle)
                if target_win.exists():
                    windows_to_scan = [target_win]
                else:
                    return {"status": "error", "message": "Target window handle invalid or closed"}
            except Exception as e:
                return {"status": "error", "message": f"Focus handle access failed: {e}"}
        else:
            windows_to_scan = desktop.windows()

        for win in windows_to_scan:
            try:
                if not root_handle and not (win.is_visible() and not win.is_minimized()):
                    continue
         
                try:
                    p_rect = win.rectangle()
                    win_text = clean_text(win.window_text())
                    win_data = {
                        'id': 0, 'name': win_text, 'type': "Window",
                        'top_level_handle': int(win.handle), 'automation_id': "",
                        'rectangle_coords': (p_rect.left, p_rect.top, p_rect.right, p_rect.bottom)
                    }
                    all_elements.append(win_data)
                    debug_elements.append({**win_data, 'class_name': win.class_name(), 'raw_rect': str(p_rect)})
                except: pass

 
                descendants = win.descendants()
                for i, elem in enumerate(descendants):
                    try:
                        if not elem.is_visible(): continue
                        rect = elem.rectangle()
                        if rect.width() <= 0 or rect.height() <= 0: continue

                        auto_id = str(elem.automation_id())
                        e_type = elem.friendly_class_name()
                        name = clean_text(elem.window_text())
                        
                        all_elements.append({
                            'name': name, 'type': e_type, 'automation_id': auto_id,
                            'top_level_handle': int(win.handle),
                            'rectangle_coords': (rect.left, rect.top, rect.right, rect.bottom)
                        })
                        
        
                        clickable_point = None
                        try:
                            pt = elem.element_info.clickable_point
                            if pt: clickable_point = (pt.x, pt.y)
                        except: clickable_point = "N/A"

                        debug_elements.append({
                            'index': i, 'name': name, 'type': e_type, 'automation_id': auto_id,
                            'rect': (rect.left, rect.top, rect.right, rect.bottom),
                            'clickable_point': clickable_point
                        })
                    except: continue
            except Exception: continue

        _dump_debug_info(debug_elements, root_handle)
        logger.info(f"Analysis completed. Found {len(all_elements)} elements.")
        return {"status": "success", "data": all_elements}

    except Exception as e:
        logger.error(f"Critical error in fetch_raw_elements: {repr(e)}", exc_info=True)
        return {"status": "error", "message": str(e)}

def get_active_window_info():
    try:
        hwnd = win32gui.GetForegroundWindow()
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
            rect = win32gui.GetWindowRect(hwnd)
            return {"status": "success", "handle": hwnd, "rect": rect}
    except Exception: pass
    return {"status": "error", "message": "No active window found"}

def get_caret_status():
    try:
        hwnd_foreground = win32gui.GetForegroundWindow()
        if not hwnd_foreground: return {"active": False, "reason": "No Foreground Window"}
        remote_thread_id, _ = win32process.GetWindowThreadProcessId(hwnd_foreground)
        
        gui_info = GUITHREADINFO()
        gui_info.cbSize = ctypes.sizeof(GUITHREADINFO)
        
        if ctypes.windll.user32.GetGUIThreadInfo(remote_thread_id, ctypes.byref(gui_info)):
            has_caret = bool(gui_info.hwndCaret) and (gui_info.rcCaret.right - gui_info.rcCaret.left > 0)
            return {
                "active": has_caret, 
                "hwnd_focus": gui_info.hwndFocus,
                "caret_rect": (gui_info.rcCaret.left, gui_info.rcCaret.top, gui_info.rcCaret.right, gui_info.rcCaret.bottom)
            }
    except Exception as e:
        logger.error(f"Caret Check Error: {e}")
    return {"active": False, "reason": "Exception"}

def _perform_interaction(top_level_handle, automation_id, name, control_type, interaction_type, text_to_type=None, target_rect=None):
    import pywinauto.mouse
    import pywinauto.keyboard
    
    logger.info(f"[Interact] START. Target: Name='{name}', ID='{automation_id}'")
    try:
        desktop = pywinauto.Desktop(backend="uia")
        app_win = desktop.window(handle=top_level_handle)
        
        if not app_win.exists():
            logger.error(f"[Interact] Target window {top_level_handle} gone.")
            return False

        if app_win.is_minimized(): app_win.restore()
        if win32gui.GetForegroundWindow() != top_level_handle: app_win.set_focus()

        element_to_interact = None
        
 
        if automation_id:
            try:
                candidates = [e for e in app_win.descendants(auto_id=automation_id) if e.is_visible()]
                if candidates: element_to_interact = candidates[0]
            except: pass

        if not element_to_interact and (name or control_type):
            criteria = {}
            if name: criteria['title'] = clean_text(name)
            if control_type: criteria['control_type'] = control_type
            try:
                candidates = [e for e in app_win.descendants(**criteria) if e.is_visible()]
                if candidates: element_to_interact = candidates[0]
            except: pass

        final_click_point = None
        if element_to_interact:
            try:
                rect = element_to_interact.rectangle()
                final_click_point = ((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
            except: pass
        elif target_rect:
       
             logger.info(f"[Interact] Using coordinate fallback: {target_rect}")
             final_click_point = ((target_rect[0] + target_rect[2]) // 2, (target_rect[1] + target_rect[3]) // 2)
        
        if not final_click_point:
             logger.error("[Interact] FAILED: No element or coords found.")
             return False

        logger.info(f"[Interact] ACTION: {interaction_type.upper()} at {final_click_point}")
        pywinauto.mouse.move(coords=final_click_point)
        time.sleep(0.05)
        
        if interaction_type == 'click': pywinauto.mouse.click(coords=final_click_point)
        elif interaction_type == 'double_click': pywinauto.mouse.double_click(coords=final_click_point)
        elif interaction_type == 'right_click': pywinauto.mouse.right_click(coords=final_click_point)
        elif interaction_type == 'type':
            pywinauto.mouse.click(coords=final_click_point)
            time.sleep(0.1)
            if text_to_type:
                if element_to_interact: element_to_interact.type_keys(text_to_type, with_spaces=True)
                else: pywinauto.keyboard.send_keys(text_to_type, with_spaces=True)
        
        return True
    except Exception as e:
        logger.error(f"[Interact] Error: {e}")
        traceback.print_exc()
        return False

def get_all_visible_windows():
    window_list = []
    def enum_cb(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetWindowTextLength(hwnd) > 0:
                title = win32gui.GetWindowText(hwnd)
                if title not in ["Program Manager", "Settings", "NVIDIA GeForce Overlay"]:
                    results.append({'handle': hwnd, 'title': title, 'rect': win32gui.GetWindowRect(hwnd)})
    win32gui.EnumWindows(enum_cb, window_list)
    return window_list

def main():
    address = (CONFIG["IPC_HOST"], CONFIG["IPC_PORT"])
    authkey = CONFIG["IPC_AUTHKEY"]
    
    logger.info(f"Analysis Server started on {address}")
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except pythoncom.com_error:
        pass
    
    try:
        with Listener(address, authkey=authkey) as listener:
            while True:
                try:
                    with listener.accept() as conn:
                        msg = conn.recv()
                        cmd = msg.get("command")
                        
                        if cmd == 'analyze':
                            root_handle = msg.get("payload", {}).get("root_handle")
                            conn.send(fetch_raw_elements(root_handle))
                        elif cmd == 'check_handle':
                            h = msg.get("payload", {}).get("handle")
                            exists = win32gui.IsWindow(h) if h else False
                            rect = win32gui.GetWindowRect(h) if exists else None
                            conn.send({"status": "success", "exists": exists, "rect": rect})
                        elif cmd == 'close_window':
                            h = msg.get("payload", {}).get("handle")
                            try:
                                win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                                conn.send({"status": "success"})
                            except Exception as e:
                                conn.send({"status": "error", "message": str(e)})
                        elif cmd == 'get_active_window':
                            win_info = get_active_window_info()
                            win_info['caret'] = get_caret_status()
                            conn.send(win_info)
                        elif cmd == 'interact':
                            success = _perform_interaction(**msg.get("payload", {}))
                            conn.send({"status": "success" if success else "error"})
                        elif cmd == 'get_window_list':
                            conn.send({"status": "success", "windows": get_all_visible_windows()})
                        elif cmd == 'ping':
                            conn.send({"status": "pong"})
                        elif cmd == 'shutdown':
                            conn.send({"status": "ok"})
                            return
                        else:
                            conn.send({"status": "error", "message": "Unknown command"})
                except Exception as e:
                    logger.error(f"Connection loop error: {e}")
                    time.sleep(1)
    finally:
        pythoncom.CoUninitialize()
        logger.info("Server terminated.")

if __name__ == '__main__':
    main()