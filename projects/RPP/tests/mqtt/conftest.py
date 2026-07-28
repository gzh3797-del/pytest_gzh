# -*- coding: utf-8 -*-
"""
conftest.py — MQTT 测试套件共享 fixtures

Session 级（整个 session 只执行一次）：
  driver      Selenium Chrome，登录一次后全程复用
  mqtt_page   MQTTPage 页面对象，导航至 MQTT 配置页后全程复用

运行方式（仓库根目录）：
  pytest Protocols/MQTT/test_mqtt.py -v
  pytest Protocols/MQTT/test_mqtt.py -v -m lv0          # 只跑冒烟
  pytest Protocols/MQTT/test_mqtt.py -v -m "lv0 or lv1" # 只跑高优先级
  pytest Protocols/MQTT/test_mqtt.py -v -m integration   # 只跑集成测试
"""
from __future__ import annotations

import logging
import re
import sys
import time
from pathlib import Path
from typing import Any

import pytest
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# ── 路径注册 ──────────────────────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # 仓库根
sys.path.insert(0, str(Path(__file__).parent.parent))         # projects/RPP/
sys.path.insert(0, str(Path(__file__).parent))                # projects/RPP/mqtt/

import settings as config                                      # projects/RPP/settings.py
from mqtt_page import (
    MQTTPage,
    MQTTConfig,
    MQTTGeneralConfig,
    MQTTCredentialConfig,
    MQTTSSLConfig,
    MQTTLWTConfig,
    MQTTTopicConfig,
)

log = logging.getLogger(__name__)


# ── 覆盖项目级 screenshot_on_failure ─────────────────────────────────────────
# 项目级 conftest 的 screenshot_on_failure 依赖 pytest-playwright 的 `page`
# fixture，会在 MQTT（Selenium）测试中意外启动一个多余的 Playwright 浏览器。
# 此处以无 page 依赖的空实现覆盖，仅保留 Selenium driver 的截图逻辑。
@pytest.fixture(autouse=True)
def screenshot_on_failure(request):
    yield
    node = request.node
    if (hasattr(node, "rep_call") and node.rep_call.failed
            and "driver" in request.fixturenames):
        try:
            drv = request.getfixturevalue("driver")
            from pathlib import Path as _P
            import datetime as _dt
            ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            out = _P(__file__).parent / "screenshots"
            out.mkdir(exist_ok=True)
            drv.save_screenshot(str(out / f"FAIL_{node.name}_{ts}.png"))
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════════════════════
# Fixtures
# ══════════════════════════════════════════════════════════════════════════════

@pytest.fixture(scope="session")
def ensure_certs():
    """mTLS 证书不存在时自动调用 gen_certs.generate_certs() 生成，
    有证书则直接复用，整个 session 只执行一次。
    依赖：cryptography 库（pip install cryptography）
    """
    import socket
    ca = Path(config.MQTT_SSL_CA_CERT)
    srv = Path(config.MQTT_SSL_SERVER_CERT)
    # 服务端 SAN 必须包含 www.accu.com（设备 Broker Address 使用的域名）和本机 IP
    _extra = [config.WEB_MQTT_BROKER_ADDRESS]
    try:
        _extra.append(socket.gethostbyname(socket.gethostname()))
    except OSError:
        pass
    if not ca.exists():
        log.info("[ensure_certs] 证书不存在（%s），自动生成…", ca)
        from gen_certs import generate_certs          # Protocols/MQTT/gen_certs.py
        generate_certs(out_dir=ca.parent, extra_hosts=_extra, days=3650)
        log.info("[ensure_certs] 证书生成完成：%s", ca.parent)
    elif not srv.exists():
        log.info("[ensure_certs] server.crt 不存在，重新生成服务端证书…")
        from gen_certs import generate_server_cert
        generate_server_cert(out_dir=ca.parent, extra_hosts=_extra, days=3650)
    else:
        log.info("[ensure_certs] 证书已存在，跳过生成")


@pytest.fixture(scope="session")
def driver():
    """会话级 Chrome WebDriver，整个 test session 只启动一次浏览器。

    TODO(需现场确认)：下方登录表单的 XPath 选择器（Enter User Name / Enter Password /
    Sign In 按钮）照抄自 AcuHMI-1-7 登录页，尚未对着 RPP 真机
    （http://192.168.2.94:3030）验证过，登录表单结构可能不同。
    """
    options = Options()
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors=yes")
    options.add_argument("--allow-running-insecure-content")
    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    drv = webdriver.Chrome(options=options)

    drv.get(config.GATEWAY_WEB_URL)
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    wait = WebDriverWait(drv, 15)
    wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Enter User Name' or @type='text']")))
    drv.find_element(By.XPATH, "//input[@placeholder='Enter User Name' or @type='text']").send_keys(config.GATEWAY_WEB_USER)
    drv.find_element(By.XPATH, "//input[@placeholder='Enter Password' or @type='password']").send_keys(config.GATEWAY_WEB_PASS)
    drv.find_element(By.XPATH, "//button[contains(normalize-space(.),'Sign In')]").click()
    time.sleep(3)
    log.info("已登录 %s", config.GATEWAY_WEB_URL)
    time.sleep(3)

    yield drv

    drv.quit()
    log.info("浏览器已关闭")


@pytest.fixture(scope="session")
def mqtt_page(driver):
    """会话级 MQTTPage，导航至 MQTT 配置页后整个 session 复用。

    url_root 覆盖为 RPP 真实路由（见 knowledge/meters/RPP/requirements/context/mqtt_context.md）：
    Settings → Protocols（左侧栏）→ MQTT（横向 Tab，hover 展开）→ General。
    """
    page = MQTTPage(driver, url_root="protocols/mqtt/general")
    page.navigate()
    return page


# ── pytest-html 报告：用例编号列 ──────────────────────────────────────────────
# 与 tests/BacnetIP/conftest.py 的同名逻辑保持一致：MQTT 用例同样不是参数化用例，
# nodeid 默认只含 `文件::类::函数名`，报告里看不到官方用例编号（TestCase_AcuHMI17_MQTT_子模块号_序号）。
# 编号写在每个用例 docstring 首段，收集阶段读出后以 `[编号]` 形式追加到 nodeid 末尾。

_CASE_ID_RE = re.compile(r"(TestCase_AcuHMI[\w\-]+)")


def pytest_collection_modifyitems(items: list[pytest.Item]) -> None:
    """把用例编号注入 MQTT 用例的 nodeid，让 pytest-html 报告的 Test 列直接显示编号。

    子目录 conftest 的 pytest_collection_modifyitems 会收到**全量** items，故先按本
    conftest 所在目录过滤，只改 MQTT 自己的用例；并对结尾做幂等判断，避免重复追加。
    """
    here = Path(__file__).parent.resolve()
    for item in items:
        try:
            item_path = Path(str(item.fspath)).resolve()
        except Exception:
            continue
        if here != item_path.parent and here not in item_path.parents:
            continue
        doc = (getattr(getattr(item, "function", None), "__doc__", "") or "").strip()
        m = _CASE_ID_RE.match(doc)
        if not m:
            continue
        case_id = m.group(1)
        suffix = f"[{case_id}]"
        if not item._nodeid.endswith(suffix):
            item._nodeid = f"{item._nodeid}{suffix}"


@pytest.fixture(autouse=True)
def _record_case_id(request: pytest.FixtureRequest, record_property: Any) -> None:
    """将用例编号写入 JUnit XML property（兼容 CI/CD 系统）。"""
    doc = (getattr(request.node.function, "__doc__", "") or "").strip()
    m = _CASE_ID_RE.match(doc)
    if m:
        record_property("用例编号", m.group(1))
