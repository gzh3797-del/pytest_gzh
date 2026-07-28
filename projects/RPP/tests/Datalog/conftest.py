"""Datalog 模块（接入设备 Data Log 下载用例组）fixtures。

与 tests/Alarm 同款结构：复用项目根 browser，自建 context + 登录；
目标机与账号配置沿用 tests/Alarm/config_alarm.py（同一台 AcuHMI-1-7）。
"""
from __future__ import annotations

import pytest
from playwright.sync_api import Browser, BrowserContext, Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


@pytest.fixture(scope="session")
def datalog_ctx(browser: Browser) -> BrowserContext:
    ctx = browser.new_context(
        ignore_https_errors=True,
        viewport={"width": 1280, "height": 720},
        accept_downloads=True,
    )
    ctx.set_default_timeout(15_000)
    yield ctx
    ctx.close()


@pytest.fixture(scope="session")
def _datalog_session_page(datalog_ctx: BrowserContext) -> Page:
    page = datalog_ctx.new_page()
    ha.login(page)
    yield page


@pytest.fixture
def app_page(_datalog_session_page: Page) -> Page:
    """已登录页面（名字必须叫 app_page：根 conftest 失败截图按此名取）。"""
    page = _datalog_session_page
    ha.ensure_logged_in(page)
    return page
