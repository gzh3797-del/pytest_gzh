import time
import json
from pathlib import Path
from datetime import datetime


def timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def load_json(file_path: str | Path) -> dict:
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(data: dict, file_path: str | Path):
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def wait_seconds(seconds: float):
    time.sleep(seconds)
