"""
pytest fixtures — Azure IoT 测试套件（AcuHMI-1-7 Playwright 版）

用法:
    pytest tests/azure_iot/ -m azure_iot -v
    pytest tests/azure_iot/ -m "azure_iot and not slow" -v
"""

import logging
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pytest
import yaml

_THIS_DIR     = Path(__file__).resolve().parent          # tests/azure_iot/
_PROJECT_ROOT = _THIS_DIR.parent.parent                  # AcuHMI_1_7/

sys.path.insert(0, str(_PROJECT_ROOT))

CONFIG_PATH = _THIS_DIR / "config.yaml"

_log = logging.getLogger(__name__)


# ─── Allure 报告自动生成（每次 pytest 结束后生成带时间戳的报告目录）────────────────

def pytest_sessionfinish(session, exitstatus):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    results_dir = _THIS_DIR / "reports" / "allure-results"
    report_dir  = _THIS_DIR / "reports" / f"allure-report-{ts}"
    allure_bat  = Path(r"C:\work\tools\allure\allure-2.32.0\bin\allure.bat")
    if results_dir.exists() and allure_bat.exists():
        subprocess.run(
            [str(allure_bat), "generate", str(results_dir),
             "-o", str(report_dir), "--clean"],
            check=False,
        )
        _log.info("Allure 报告已生成：%s", report_dir)


# ─── 共享配置 ──────────────────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def azure_cfg():
    return yaml.safe_load(CONFIG_PATH.read_text(encoding="utf-8"))


# ─── 函数级 fixture（每条用例独立，带 navigate + teardown disable）───────────────

@pytest.fixture(scope="function")
def azure_page(app_page, azure_cfg):
    """
    函数级 AzureIoTPage fixture：
    - Setup：导航到 Azure IoT 配置页面
    - 用例结束后 teardown：disable Azure IoT（防止残留数据污染）
    """
    from pages.protocols.azure_iot_page import AzureIoTPage
    page = AzureIoTPage(app_page)
    page.navigate_to_azure_iot()
    yield page
    try:
        page.navigate_to_azure_iot()
        if page.is_enabled():
            page.disable()
    except Exception:
        pass


# ─── Session 级 fixture（整个 session 共用一个浏览器，用于 Interval 测试）─────────

@pytest.fixture(scope="session")
def azure_session_page(app_page, azure_cfg):
    """
    Session 级 AzureIoTPage fixture：
    - 整个 session 共享同一 Playwright 页面实例
    - Setup：导航 + 完整初始化配置（连接串、Interval、设备选择）
    - Session 结束后 teardown：disable Azure IoT
    """
    from pages.protocols.azure_iot_page import AzureIoTPage
    azure = azure_cfg["azure_iot"]

    page = AzureIoTPage(app_page)
    page.navigate_to_azure_iot()

    page.ensure_enabled()
    page.set_primary_conn_str(azure.get("primary_conn_str", ""))
    if azure.get("secondary_conn_str"):
        page.set_secondary_conn_str(azure["secondary_conn_str"])
    page.set_interval(azure.get("interval", "30 seconds"))

    device_name = azure.get("interval_test_device", "")
    page.select_only_device(device_name or "")
    page.configure_all_devices_parameters(checked_only=True)
    page.save()

    yield page

    try:
        page.navigate_to_azure_iot()
        page.disable()
    except Exception:
        pass
