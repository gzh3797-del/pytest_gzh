import time
from datetime import datetime


def timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def wait(seconds: float = 1.0):
    time.sleep(seconds)
