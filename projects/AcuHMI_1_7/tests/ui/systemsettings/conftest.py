# -*- coding: utf-8 -*-
"""systemsettings 子模块共享 fixture。

`system_settings_page`：登录 AcuHMI 后返回已就绪的 Playwright Page，供本目录下
各用例（005_06_*、005_08_* 等）复用。用例内再通过 hash-URL 跳转到具体子页，
例如 ``page.goto(page.url.split("#")[0] + "#/systemSettings/certificate")``，
因此此处只保证「已登录」这一前置条件，不替用例预设具体子页（避免菜单名差异引入脆弱性）。
"""
import pytest
from playwright.sync_api import Page

from projects.AcuHMI_1_7.pages.login_page import LoginPage


@pytest.fixture
def system_settings_page(login_page: LoginPage) -> Page:
    """登录 AcuHMI 并返回 Page；具体系统设置子页由各用例自行导航。"""
    login_page.open()
    login_page.login()
    return login_page.page
