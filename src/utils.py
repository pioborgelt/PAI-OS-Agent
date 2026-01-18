"""
OS Agent Utilities & Configuration
----------------------------------
This module serves as the central configuration hub for the OS Agent system.
It handles:
- Environment variable loading via .env
- Centralized configuration dictionary (CONFIG)
- Logging setup and directory management

⚠️ IMPORTANT:
Before running this project, ensure you have read the README.md file.
Proper environment setup (including the .env file and API keys) is critical
for the agent to function correctly.

Author: Pio Borgelt
"""

import os
import logging
import sys
from pathlib import Path
from dotenv import load_dotenv

# Load .env file
load_dotenv()

# Central Configuration
CONFIG = {
    "GOOGLE_API_KEY": os.getenv("GOOGLE_API_KEY", ""),
    "IPC_HOST": os.getenv("IPC_HOST", "localhost"),
    "IPC_PORT": int(os.getenv("IPC_PORT", 6000)),
    "IPC_AUTHKEY": os.getenv("IPC_AUTHKEY", "changeme").encode("utf-8"),
    "TESSERACT_PATH": os.getenv("TESSERACT_PATH", r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    "LOG_LEVEL": os.getenv("LOG_LEVEL", "INFO").upper(),
    "LOG_DIR": os.getenv("LOG_DIR", "logs"),
    "DEBUG_DIR": os.path.join(os.getenv("LOG_DIR", "logs"), "ui_snapshots")
}

def get_logger(name: str) -> logging.Logger:
    """
    Returns a configured logger instance.
    Ensures log directories exist.
    """
    log_dir = Path(CONFIG["LOG_DIR"])
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Create Debug Dir for UI Snapshots if needed
    Path(CONFIG["DEBUG_DIR"]).mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger(name)
    
    # Avoid adding multiple handlers if get_logger is called repeatedly
    if not logger.handlers:
        logger.setLevel(CONFIG["LOG_LEVEL"])
        
        formatter = logging.Formatter(
            '[%(asctime)s] [%(name)s] [%(levelname)s] - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        # File Handler
        log_file = log_dir / "agent.log"
        file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        # Stream Handler (Console)
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(formatter)
        logger.addHandler(stream_handler)

    return logger

logger = get_logger("OSAgent")