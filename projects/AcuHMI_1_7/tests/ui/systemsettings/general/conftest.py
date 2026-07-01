# -*- coding: utf-8 -*-
"""
systemsettings/general 组级 conftest

提供 datetime_guard fixture：会修改系统时间 / 时区 / NTP 的用例专用。
用例在修改任何设置之前调用 datetime_guard.snapshot(page) 记录原始状态，
用例结束后由本 fixture 自动恢复（还原时区 / NTP 配置，并用有效 NTP 重新同步时钟），
确保设备时钟回到正确当前时间，不污染后续用例。
"""
import pytest

from projects.AcuHMI_1_7.tests.ui.systemsettings.helpers.datetime_helpers import (
    snapshot_datetime_settings,
    restore_datetime_settings,
)


class _DateTimeGuard:
    """承载执行前快照并在用例结束后触发恢复。"""

    def __init__(self) -> None:
        self.page = None
        self.snap = None

    def snapshot(self, page) -> None:
        """在修改任何设置之前调用，记录时区 / NTP 开关 / NTP Server。"""
        self.page = page
        self.snap = snapshot_datetime_settings(page)


@pytest.fixture
def datetime_guard():
    guard = _DateTimeGuard()
    yield guard
    # teardown：仅当用例确实做过 snapshot 才恢复；恢复失败不影响用例结果
    if guard.page is not None and guard.snap is not None:
        try:
            restore_datetime_settings(guard.page, guard.snap)
        except Exception:
            pass
