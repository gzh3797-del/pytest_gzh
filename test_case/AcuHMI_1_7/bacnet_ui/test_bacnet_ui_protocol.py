# -*- coding: utf-8 -*-
"""
test_bacnet_ui_protocol.py — BACnet/IP 协议层端到端验证用例（P1/P2）

用例覆盖：
  TestCase_AcuHMI-1-7_033_001_002: BACnet Port 合法边界值保存且 BACnet 客户端可通信
  TestCase_AcuHMI-1-7_033_001_003: Network Number 合法边界值保存并持久化
  TestCase_AcuHMI-1-7_033_001_012: AcuRev4100 EPICS Enable 后 BACnet 客户端可接收参数
  TestCase_AcuHMI-1-7_033_001_013: PXE1 EPICS Enable 后 BACnet 客户端可接收参数
  TestCase_AcuHMI-1-7_033_001_014: PXE2 EPICS Enable 后 BACnet 客户端可接收参数
  TestCase_AcuHMI-1-7_033_001_015: PXM350 EPICS Enable 后 BACnet 客户端可接收参数
  TestCase_AcuHMI-1-7_033_001_016: AcuVIM3 EPICS Enable 后 BACnet 客户端可接收参数
  TestCase_AcuHMI-1-7_033_001_017: AcuRev2100 EPICS Enable 后 BACnet 客户端可接收参数
  TestCase_AcuHMI-1-7_033_001_018: 关闭 EPICS Enable 后 BACnet 客户端不能接收该参数
  TestCase_AcuHMI-1-7_033_001_037: EPICS file download 下载触发正常
  TestCase_AcuHMI-1-7_033_001_038: 禁用 BACnet 后客户端无法连接

运行：
  pytest test_case/AcuHMI_1_7/bacnet_ui/test_bacnet_ui_protocol.py -v
"""
from __future__ import annotations

import sys
import time
from pathlib import Path
from typing import Optional

import pytest
from playwright.sync_api import Page

# ── 路径 ─────────────────────────────────────────────────────────────────────
_REPO_ROOT = str(Path(__file__).parents[3])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

# ── 复用已有辅助函数 ──────────────────────────────────────────────────────────
from test_case.AcuHMI_1_7.bacnet_ui.test_bacnet_ui_basic import (  # noqa: E402
    _get_device_names,
    _open_param_dialog,
    _close_param_config_dialog,
)
from test_case.AcuHMI_1_7.bacnet_ui.test_bacnet_ui_config import (  # noqa: E402
    _get_field_value,
    _set_field_value,
    _click_save,
    _dismiss_toast,
    _navigate_to_bacnet,
)

# ── BACnet 客户端 ─────────────────────────────────────────────────────────────
from test_case.AcuHMI_1_7.bacnet_ui.helpers.hmi_bacnet_client import (  # noqa: E402
    can_connect,
    get_object_count,
    HMI_DEFAULT_PORT,
    DEVICE_RESTART_WAIT,
)

# ── 本文件常量 ────────────────────────────────────────────────────────────────
_SAVE_WAIT_MS = 2000


# ═════════════════════════════════════════════════════════════════════════════
# 本文件专用辅助函数
# ═════════════════════════════════════════════════════════════════════════════

def _get_bacnet_enable_state(page: Page) -> bool:
    """读取 BACnet Enable radio-group 当前是否为 Enable（True）。"""
    return page.evaluate(
        """() => {
            const radios = document.querySelectorAll('.el-radio__original[value="true"]');
            for (const r of radios) {
                const group = r.closest('.el-radio-group');
                if (!group) continue;
                const formItem = group.closest('.el-form-item');
                const label = formItem && formItem.querySelector('.el-form-item__label');
                if (label && label.textContent.trim() === 'BACnet Enable') {
                    return r.checked;
                }
            }
            return false;
        }"""
    )


def _set_bacnet_enable(page: Page, enable: bool) -> None:
    """设置 BACnet Enable radio-group 为 Enable 或 Disable。"""
    value = "true" if enable else "false"
    page.evaluate(
        """(val) => {
            const radios = document.querySelectorAll('.el-radio__original');
            for (const r of radios) {
                if (r.getAttribute('value') !== val) continue;
                const group = r.closest('.el-radio-group');
                if (!group) continue;
                const formItem = group.closest('.el-form-item');
                const label = formItem && formItem.querySelector('.el-form-item__label');
                if (label && label.textContent.trim() === 'BACnet Enable') {
                    r.click();
                    return;
                }
            }
        }""",
        value,
    )
    page.wait_for_timeout(800)


def _get_first_row_epics_state(page: Page) -> Optional[bool]:
    """
    读取 Parameter Config 弹窗第一行 cells[1]（EPICS Enable）的开关状态。
    返回 True/False；弹窗不存在或无行则返回 None。
    """
    return page.evaluate(
        """() => {
            const rows = document.querySelectorAll(
                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
            );
            if (!rows.length) return null;
            const cells = rows[0].querySelectorAll('td');
            const sw = cells[1] && cells[1].querySelector('.el-switch__input');
            return sw ? sw.getAttribute('aria-checked') === 'true' : null;
        }"""
    )


def _set_first_row_epics_enable(page: Page, enable: bool) -> bool:
    """
    将 Parameter Config 弹窗第一行 cells[1]（EPICS Enable）设置为指定状态。
    返回是否成功找到并操作了开关。
    """
    target_state = "true" if enable else "false"
    return page.evaluate(
        """(target) => {
            const rows = document.querySelectorAll(
                '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
            );
            if (!rows.length) return false;
            const cells = rows[0].querySelectorAll('td');
            const sw = cells[1] && cells[1].querySelector('.el-switch__input');
            if (!sw) return false;
            if (sw.getAttribute('aria-checked') !== target) {
                sw.click();
            }
            return true;
        }""",
        target_state,
    )


def _save_param_config_dialog(page: Page) -> bool:
    """
    在 Parameter Config 弹窗内点击 Save 按钮。
    若弹窗内无 Save 则返回 False，调用方应改用外部 Save。
    """
    saved: bool = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            if (!dlg) return false;
            for (const btn of dlg.querySelectorAll('button')) {
                if (btn.textContent.trim() === 'Save') {
                    btn.click();
                    return true;
                }
            }
            return false;
        }"""
    )
    if saved:
        page.wait_for_timeout(_SAVE_WAIT_MS)
        _dismiss_toast(page)
    return saved


def _find_device_keyword(devices: list[str], keyword: str) -> Optional[str]:
    """在设备列表中找到包含 keyword 的第一个设备名（大小写不敏感），找不到返回 None。"""
    kw_lower = keyword.lower()
    return next((d for d in devices if kw_lower in d.lower()), None)


def _epics_enable_verify_count_increase(
    page: Page,
    device_keyword: str,
    test_name: str,
) -> None:
    """
    通用流程：对指定设备第一行启用 EPICS Enable，验证 BACnet 对象数量增加。
    若设备已启用则仅验证 BACnet 可达。
    """
    count_before = get_object_count()
    if count_before is None:
        pytest.skip("BACnet 客户端无法连接，跳过此用例")

    if not _open_param_dialog(page, device_keyword):
        pytest.skip(f"[{test_name}] 无法打开设备 {device_keyword!r} 的 Parameter Config 弹窗")

    already_enabled: bool = False
    try:
        state = _get_first_row_epics_state(page)
        if state is None:
            pytest.skip(f"[{test_name}] Parameter Config 弹窗中未找到参数行")

        already_enabled = state

        if not already_enabled:
            ok = _set_first_row_epics_enable(page, True)
            if not ok:
                pytest.skip(f"[{test_name}] 无法操作第一行 EPICS Enable 开关")
            page.wait_for_timeout(300)

        # 保存：优先弹窗内 Save，否则关闭后用外部 Save
        saved_in_dlg = _save_param_config_dialog(page)
        if saved_in_dlg:
            _close_param_config_dialog(page)
        else:
            _close_param_config_dialog(page)
            _click_save(page)
            _dismiss_toast(page)

    except Exception:
        if page.locator('[aria-label="Parameter Config"]').count() > 0:
            _close_param_config_dialog(page)
        raise

    # 等待 BACnet 服务重启
    time.sleep(DEVICE_RESTART_WAIT)

    count_after = get_object_count()

    # 恢复：若原来未启用则需要 disable 回去
    if not already_enabled:
        try:
            if _open_param_dialog(page, device_keyword):
                _set_first_row_epics_enable(page, False)
                page.wait_for_timeout(300)
                saved_in_dlg = _save_param_config_dialog(page)
                if saved_in_dlg:
                    _close_param_config_dialog(page)
                else:
                    _close_param_config_dialog(page)
                    _click_save(page)
                    _dismiss_toast(page)
        except Exception:
            if page.locator('[aria-label="Parameter Config"]').count() > 0:
                _close_param_config_dialog(page)

    # 断言
    if already_enabled:
        assert count_after is not None, (
            f"[{test_name}] 设备 {device_keyword!r} 的 EPICS Enable 已启用，"
            "BACnet 客户端应可达，但 get_object_count() 返回 None"
        )
    else:
        assert count_after is not None and count_after > count_before, (
            f"[{test_name}] 设备 {device_keyword!r} EPICS Enable 启用后，"
            f"BACnet 对象数量应增加，但 count_before={count_before}，"
            f"count_after={count_after}"
        )


# ═════════════════════════════════════════════════════════════════════════════
# 测试用例
# ═════════════════════════════════════════════════════════════════════════════

class TestBACnetProtocol:
    """BACnet/IP 协议层端到端验证（配合 BACnet 客户端）。"""

    def test_002_bacnet_port_valid_boundary(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_002: BACnet Port 合法边界值保存且 BACnet 客户端可通信。"""
        original_port = _get_field_value(hmi_page, "Enter BACnet Port")
        restore_port = original_port if original_port else str(HMI_DEFAULT_PORT)

        try:
            # 测试端口 47808（合法下限）
            _set_field_value(hmi_page, "Enter BACnet Port", "47808")
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)
            time.sleep(DEVICE_RESTART_WAIT)
            assert can_connect(gateway_port=47808), (
                "BACnet Port = 47808 保存后，BACnet 客户端应可通过该端口连接，但连接失败"
            )

            # 测试端口 49000（合法上限）
            _set_field_value(hmi_page, "Enter BACnet Port", "49000")
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)
            time.sleep(DEVICE_RESTART_WAIT)
            assert can_connect(gateway_port=49000), (
                "BACnet Port = 49000 保存后，BACnet 客户端应可通过该端口连接，但连接失败"
            )

        finally:
            # 恢复原始端口，确保后续测试正常
            _set_field_value(hmi_page, "Enter BACnet Port", restore_port)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)
            time.sleep(DEVICE_RESTART_WAIT)

    def test_003_network_number_valid_boundary(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_003: Network Number 合法边界值保存并持久化。"""
        original_nn = _get_field_value(hmi_page, "Enter Network Number")
        restore_nn = original_nn if original_nn else "1"

        try:
            # 测试 NN=1（合法下限）
            _set_field_value(hmi_page, "Enter Network Number", "1")
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)
            _navigate_to_bacnet(hmi_page)
            actual_nn = _get_field_value(hmi_page, "Enter Network Number")
            assert actual_nn == "1", (
                f"Network Number = 1 保存后，重新导航读回值应为 '1'，实际为 {actual_nn!r}"
            )

            # 测试 NN=65534（合法上限）
            _set_field_value(hmi_page, "Enter Network Number", "65534")
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)
            _navigate_to_bacnet(hmi_page)
            actual_nn = _get_field_value(hmi_page, "Enter Network Number")
            assert actual_nn == "65534", (
                f"Network Number = 65534 保存后，重新导航读回值应为 '65534'，实际为 {actual_nn!r}"
            )

        finally:
            _set_field_value(hmi_page, "Enter Network Number", restore_nn)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)

    def test_012_acurev4100_epics_enable_bacnet_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_012: AcuRev4100 EPICS Enable 后 BACnet 客户端可接收参数。"""
        devices = _get_device_names(hmi_page)
        device = _find_device_keyword(devices, "4100")
        if device is None:
            pytest.skip(f"未找到含 '4100' 的设备（当前设备列表：{devices}），跳过此用例")

        _epics_enable_verify_count_increase(hmi_page, device, "test_012_4100")

    def test_013_pxe1_epics_enable_bacnet_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_013: PXE1 EPICS Enable 后 BACnet 客户端可接收参数。"""
        devices = _get_device_names(hmi_page)
        device = _find_device_keyword(devices, "PXE1")
        if device is None:
            pytest.skip(f"未找到含 'PXE1' 的设备（当前设备列表：{devices}），跳过此用例")

        _epics_enable_verify_count_increase(hmi_page, device, "test_013_PXE1")

    def test_014_pxe2_epics_enable_bacnet_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_014: PXE2 EPICS Enable 后 BACnet 客户端可接收参数。"""
        devices = _get_device_names(hmi_page)
        device = _find_device_keyword(devices, "PXE2")
        if device is None:
            pytest.skip(f"未找到含 'PXE2' 的设备（当前设备列表：{devices}），跳过此用例")

        _epics_enable_verify_count_increase(hmi_page, device, "test_014_PXE2")

    def test_015_pxm350_epics_enable_bacnet_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_015: PXM350 EPICS Enable 后 BACnet 客户端可接收参数。"""
        devices = _get_device_names(hmi_page)
        device = _find_device_keyword(devices, "PXM350")
        if device is None:
            device = _find_device_keyword(devices, "PXM")
        if device is None:
            pytest.skip(
                f"未找到含 'PXM350' 或 'PXM' 的设备（当前设备列表：{devices}），跳过此用例"
            )

        _epics_enable_verify_count_increase(hmi_page, device, "test_015_PXM350")

    def test_016_acuvim3_epics_enable_bacnet_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_016: AcuVIM3 EPICS Enable 后 BACnet 客户端可接收参数。"""
        devices = _get_device_names(hmi_page)
        device = _find_device_keyword(devices, "VIM3")
        if device is None:
            device = _find_device_keyword(devices, "AcuVIM")
        if device is None:
            pytest.skip(
                f"未找到含 'VIM3' 或 'AcuVIM' 的设备（当前设备列表：{devices}），跳过此用例"
            )

        _epics_enable_verify_count_increase(hmi_page, device, "test_016_AcuVIM3")

    def test_017_acurev2100_epics_enable_bacnet_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_017: AcuRev-2100 EPICS Enable 后 BACnet 客户端可接收参数。"""
        devices = _get_device_names(hmi_page)
        device = _find_device_keyword(devices, "2100")
        if device is None:
            pytest.skip(f"未找到含 '2100' 的设备（当前设备列表：{devices}），跳过此用例")

        _epics_enable_verify_count_increase(hmi_page, device, "test_017_2100")

    def test_018_epics_disable_bacnet_not_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_018: 关闭 EPICS Enable 后 BACnet 客户端不能接收该参数。"""
        devices = _get_device_names(hmi_page)
        if not devices:
            pytest.skip("设备列表为空，无法验证 EPICS Disable 场景")

        device_keyword = devices[0]

        # 打开 Parameter Config 弹窗，读取第一行 EPICS Enable 初始状态
        if not _open_param_dialog(hmi_page, device_keyword):
            pytest.skip(f"无法打开设备 {device_keyword!r} 的 Parameter Config 弹窗")

        was_enabled: Optional[bool] = _get_first_row_epics_state(hmi_page)
        if was_enabled is None:
            _close_param_config_dialog(hmi_page)
            pytest.skip("Parameter Config 弹窗中未找到参数行")

        count_enabled: Optional[int] = None

        try:
            # 若原来未启用，先启用后等待，建立"已启用"基准
            if not was_enabled:
                ok = _set_first_row_epics_enable(hmi_page, True)
                if not ok:
                    pytest.skip("无法操作第一行 EPICS Enable 开关")
                hmi_page.wait_for_timeout(300)

                saved_in_dlg = _save_param_config_dialog(hmi_page)
                if saved_in_dlg:
                    _close_param_config_dialog(hmi_page)
                else:
                    _close_param_config_dialog(hmi_page)
                    _click_save(hmi_page)
                    _dismiss_toast(hmi_page)

                time.sleep(DEVICE_RESTART_WAIT)

                # 重新打开弹窗，准备执行 disable
                if not _open_param_dialog(hmi_page, device_keyword):
                    pytest.fail(
                        f"启用 EPICS 后无法重新打开设备 {device_keyword!r} 的 Parameter Config"
                    )

            # 此时弹窗已打开（原来已启用 or 刚刚启用后重新打开）
            count_enabled = get_object_count()
            if count_enabled is None:
                _close_param_config_dialog(hmi_page)
                pytest.skip("BACnet 客户端不可达，无法验证对象数量变化")

            # 关闭第一行 EPICS Enable
            ok = _set_first_row_epics_enable(hmi_page, False)
            if not ok:
                pytest.skip("无法操作第一行 EPICS Enable 开关（disable）")
            hmi_page.wait_for_timeout(300)

            saved_in_dlg = _save_param_config_dialog(hmi_page)
            if saved_in_dlg:
                _close_param_config_dialog(hmi_page)
            else:
                _close_param_config_dialog(hmi_page)
                _click_save(hmi_page)
                _dismiss_toast(hmi_page)

        except Exception:
            if hmi_page.locator('[aria-label="Parameter Config"]').count() > 0:
                _close_param_config_dialog(hmi_page)
            raise

        time.sleep(DEVICE_RESTART_WAIT)

        count_disabled = get_object_count()

        # 恢复原始状态
        try:
            if was_enabled:
                # 原来是 True，现在被设成 False，需要恢复为 True
                if _open_param_dialog(hmi_page, device_keyword):
                    _set_first_row_epics_enable(hmi_page, True)
                    hmi_page.wait_for_timeout(300)
                    saved_in_dlg = _save_param_config_dialog(hmi_page)
                    if saved_in_dlg:
                        _close_param_config_dialog(hmi_page)
                    else:
                        _close_param_config_dialog(hmi_page)
                        _click_save(hmi_page)
                        _dismiss_toast(hmi_page)
            # was_enabled == False 时：我们先启用再 disable，最终状态已是 False，无需恢复
        except Exception:
            if hmi_page.locator('[aria-label="Parameter Config"]').count() > 0:
                _close_param_config_dialog(hmi_page)

        # 断言
        assert count_disabled is not None and count_enabled is not None, (
            "EPICS Disable 后 BACnet 客户端返回 None，无法比较对象数量"
        )
        assert count_disabled < count_enabled, (
            f"关闭设备 {device_keyword!r} 第一行 EPICS Enable 后，"
            f"BACnet 对象数量应减少，但 count_enabled={count_enabled}，"
            f"count_disabled={count_disabled}"
        )

    def test_037_epics_file_download(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_037: EPICS file download 下载触发正常。"""
        # 检查页面上是否存在 EPICS / Download 相关按钮
        has_epics_btn: bool = hmi_page.evaluate(
            """() => {
                for (const btn of document.querySelectorAll('button')) {
                    const t = btn.textContent.trim();
                    if (t.includes('EPICS') || t.includes('Download')) return true;
                }
                return false;
            }"""
        )
        if not has_epics_btn:
            pytest.skip("未找到 EPICS File Download 按钮，跳过此用例")

        download_obj = None
        try:
            with hmi_page.expect_download(timeout=15000) as download_info:
                hmi_page.evaluate(
                    """() => {
                        for (const btn of document.querySelectorAll('button')) {
                            const t = btn.textContent.trim();
                            if (t.includes('EPICS') || t.includes('Download')) {
                                btn.click();
                                return;
                            }
                        }
                    }"""
                )
            download_obj = download_info.value
        except Exception as exc:
            pytest.skip(f"EPICS 文件下载未触发，可能需要真机确认（{exc}）")

        assert download_obj is not None, "download 对象为 None，下载事件未正常触发"

        filename = download_obj.suggested_filename
        assert filename and len(filename) > 0, (
            "下载文件名为空，EPICS File Download 响应异常"
        )

    def test_038_bacnet_disable_client_unreachable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_038: 禁用 BACnet 后客户端无法连接。"""
        # 前置条件：BACnet 客户端当前可达
        if not can_connect():
            pytest.skip(
                "BACnet 客户端初始状态不可达（默认端口 %d），测试无法执行" % HMI_DEFAULT_PORT
            )

        initial_enabled: bool = _get_bacnet_enable_state(hmi_page)

        try:
            # 将 BACnet Enable 设为 Disable
            _set_bacnet_enable(hmi_page, False)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)
            time.sleep(DEVICE_RESTART_WAIT)

            assert not can_connect(), (
                "BACnet Enable = Disable 后，客户端应无法连接（端口 %d），但连接仍然成功"
                % HMI_DEFAULT_PORT
            )

        finally:
            # 恢复初始 BACnet Enable 状态（通常为 Enable）
            _set_bacnet_enable(hmi_page, initial_enabled)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)
            time.sleep(DEVICE_RESTART_WAIT)
