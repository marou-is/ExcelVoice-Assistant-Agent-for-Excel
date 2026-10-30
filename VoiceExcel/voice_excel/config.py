import os
from pathlib import Path

from dotenv import load_dotenv

ROOT_DIR = Path(__file__).resolve().parent.parent
HISTORY_DIR = ROOT_DIR / "history"

load_dotenv(ROOT_DIR / ".env")

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")

WHISPER_MODEL = os.getenv("WHISPER_MODEL", "small")
WHISPER_DEVICE = os.getenv("WHISPER_DEVICE", "cpu")
WHISPER_COMPUTE_TYPE = os.getenv("WHISPER_COMPUTE_TYPE", "int8")
SPEECH_LANGUAGE = os.getenv("SPEECH_LANGUAGE", "fr")

AUTO_FIX_ATTEMPTS = int(os.getenv("AUTO_FIX_ATTEMPTS", "1"))

MIC_DEVICE = os.getenv("MIC_DEVICE", "").strip()

SAMPLE_RATE = 16000
