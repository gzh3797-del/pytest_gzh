import os
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent

#########################Browser settings浏览器配置
# 1. 选择浏览器：默认 chromium，支持 chromium / firefox / webkit
BROWSER = os.getenv("BROWSER", "chromium")

# 2. 是否无头模式（服务器无界面必须开）：默认 false（有界面）
HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"

# 3. 操作延迟（毫秒）：防止太快被网站拦截，默认 1ms
SLOW_MO = int(os.getenv("SLOW_MO", "300"))

# 4. 全局超时时间：默认 30 秒
TIMEOUT = int(os.getenv("TIMEOUT", "30000"))  # ms


############################## Target URL
# 专门用来存放目标网站的基础网址，后面脚本里所有页面都基于这个地址拼接
BASE_URL = os.getenv("BASE_URL", "https://192.168.2.8")

############################### Screenshot / video
# 项目根路径
SCREENSHOT_DIR = BASE_DIR / "screenshots"
REPORT_DIR = BASE_DIR / "reports"

# 自动创建文件夹screenshots、reports，exist_ok=True，如果文件夹已经存在，就不报错、不覆盖。
SCREENSHOT_DIR.mkdir(exist_ok=True)
REPORT_DIR.mkdir(exist_ok=True)

# Default credentials (override via .env)
# 默认凭据（可以通过 .env 文件覆盖）
# 规则：优先读取环境变量，没有就用默认值 "admin"
DEFAULT_USERNAME = os.getenv("WEB_USERNAME", "admin")
DEFAULT_PASSWORD = os.getenv("WEB_PASSWORD", "Admin@110002")
