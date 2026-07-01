import os
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent

# ── 浏览器配置 ──────────────────────────────────────────────────────────────
BROWSER   = os.getenv("BROWSER", "chromium")
HEADLESS  = os.getenv("HEADLESS", "false").lower() == "true"
SLOW_MO   = int(os.getenv("SLOW_MO", "0"))
TIMEOUT   = int(os.getenv("TIMEOUT", "30000"))

# ── 目标设备 ────────────────────────────────────────────────────────────────
BASE_URL  = os.getenv("BASE_URL", "https://192.168.2.8")

# ── 登录凭据 ────────────────────────────────────────────────────────────────
DEFAULT_USERNAME = os.getenv("WEB_USERNAME", "admin")
DEFAULT_PASSWORD = os.getenv("WEB_PASSWORD", "Admin@110002")

# ── 输出目录（运行时自动创建） ─────────────────────────────────────────────
SCREENSHOT_DIR = BASE_DIR / "screenshots"
REPORT_DIR     = BASE_DIR / "reports"
SCREENSHOT_DIR.mkdir(exist_ok=True)
REPORT_DIR.mkdir(exist_ok=True)
