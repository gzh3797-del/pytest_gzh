"""Alarm 模块（Alarm Config 用例组）fixtures。

复用项目根 conftest 的 session 级 browser；本模块自建 context（目标机与 RPP
config.yaml 的 base_url 不同：当前按 AcuHMI-1-7 真机执行，见 config_alarm.py）。
"""
from __future__ import annotations

import pytest
from playwright.sync_api import Browser, BrowserContext, Page

from projects.RPP.tests.Alarm import config_alarm as cfg
from projects.RPP.tests.Alarm import helpers_alarm as ha


@pytest.fixture(scope="session")
def alarm_ctx(browser: Browser) -> BrowserContext:
    ctx = browser.new_context(
        ignore_https_errors=True,
        viewport={"width": 1280, "height": 720},
    )
    ctx.set_default_timeout(15_000)
    yield ctx
    ctx.close()


@pytest.fixture(scope="session")
def _alarm_session_page(alarm_ctx: BrowserContext) -> Page:
    page = alarm_ctx.new_page()
    ha.login(page)
    yield page
    # session 兜底清理：删除本模块创建的告警规则、恢复确认开关默认 Enable。
    # 各用例自身 finally 已清理，这里只兜异常中断的残留，失败不影响测试结论。
    try:
        ha.ensure_logged_in(page)
        ha.cleanup_test_rules(page)
        ha.set_ack_enable(page, True)
    except Exception:
        pass


@pytest.fixture
def app_page(_alarm_session_page: Page) -> Page:
    """各用例共用的已登录页面（名字必须叫 app_page：根 conftest 失败截图按此名取）。"""
    page = _alarm_session_page
    ha.ensure_logged_in(page)
    return page
