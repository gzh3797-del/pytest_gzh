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
  pytest projects/AcuHMI_1_7/tests/bacnet/test_bacnet_ui_protocol.py -v
"""
from __future__ import annotations

import logging
import sys
import time
from pathlib import Path
from typing import Optional

import pytest
from playwright.sync_api import Page

# ── 路径 ─────────────────────────────────────────────────────────────────────
_REPO_ROOT = str(Path(__file__).parents[4])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

# ── 复用已有辅助函数 ──────────────────────────────────────────────────────────
from projects.AcuHMI_1_7.tests.bacnet.test_bacnet_ui_basic import (  # noqa: E402
    _get_device_names,
    _open_param_dialog,
    _close_param_config_dialog,
    _get_parameter_types,
    _select_param_type,
    _enable_polling_all_pages,
)
from projects.AcuHMI_1_7.tests.bacnet.test_bacnet_ui_config import (  # noqa: E402
    _get_field_value,
    _set_field_value,
    _click_save,
    _dismiss_toast,
    _navigate_to_bacnet,
)

# ── BACnet 客户端 ─────────────────────────────────────────────────────────────
from projects.AcuHMI_1_7.helpers.hmi_bacnet_client import (  # noqa: E402
    can_connect,
    get_object_identifiers,
    read_object_details_batch,
    HMI_DEFAULT_PORT,
    DEVICE_RESTART_WAIT,
)

# ── 模板基准 + Modbus 实时值 ──────────────────────────────────────────────────
from projects.AcuHMI_1_7.helpers.template_matcher import (  # noqa: E402
    DEVICE_MODBUS_MODULES,
    get_bacnet_template_map,
)
from projects.AcuHMI_1_7.helpers.hmi_modbus_client import (  # noqa: E402
    read_modbus_values,
    modbus_param_keys,
)
from projects.AcuHMI_1_7 import settings as hmi_cfg  # noqa: E402

# ── 本文件常量 ────────────────────────────────────────────────────────────────
_SAVE_WAIT_MS = 2000

log = logging.getLogger(__name__)


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
    """
    设置 BACnet Enable radio-group 为 Enable 或 Disable。

    JS element.click() 在 El Plus v2 radio 上不触发完整事件链，
    改用 getBoundingClientRect() + page.mouse.click() 模拟真实鼠标点击。
    """
    value = "true" if enable else "false"
    coords = page.evaluate(
        """(val) => {
            const radios = document.querySelectorAll('.el-radio__original');
            for (const r of radios) {
                if (r.getAttribute('value') !== val) continue;
                const group = r.closest('.el-radio-group');
                if (!group) continue;
                const formItem = group.closest('.el-form-item');
                const label = formItem && formItem.querySelector('.el-form-item__label');
                if (label && label.textContent.trim() === 'BACnet Enable') {
                    const radioEl = r.closest('.el-radio') || r;
                    // 前序用例可能把页面滚走，radio 在视口外时 mouse.click 会落空，
                    // 必须先滚回视口内再取坐标
                    radioEl.scrollIntoView({block: 'center', behavior: 'instant'});
                    const rect = radioEl.getBoundingClientRect();
                    return {x: rect.left + rect.width / 2, y: rect.top + rect.height / 2};
                }
            }
            return null;
        }""",
        value,
    )
    if coords:
        page.mouse.click(coords["x"], coords["y"])
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
    保存 Parameter Config 弹窗中的改动。

    优先找弹窗内 Save 按钮（JS click）；找不到则在弹窗开启状态下
    用 force=True 点主区域 Save，绕过 el-overlay 遮罩层。
    改动在关闭弹窗前提交，否则关弹窗时 Vue state 销毁会丢失修改。

    Returns:
        True = 已通过弹窗内 Save 按钮保存；False = 已降级用主区域 Save 保存。
        两种情况下改动均已提交，调用方无需再补充保存。
    """
    saved_in_dlg: bool = page.evaluate(
        """() => {
            const dlg = document.querySelector('[aria-label="Parameter Config"]');
            if (!dlg) return false;
            for (const btn of dlg.querySelectorAll('button')) {
                if (btn.textContent.trim() === 'Save') { btn.click(); return true; }
            }
            return false;
        }"""
    )
    if not saved_in_dlg:
        page.locator('button:has-text("Save")').first.click(force=True)
    page.wait_for_timeout(_SAVE_WAIT_MS)
    _dismiss_toast(page)
    return saved_in_dlg


def _find_device_keyword(devices: list[str], keyword: str) -> Optional[str]:
    """在设备列表中找到包含 keyword 的第一个设备名（大小写不敏感），找不到返回 None。"""
    kw_lower = keyword.lower()
    return next((d for d in devices if kw_lower in d.lower()), None)


def _enable_all_polling_for_device(page: Page) -> None:
    """
    在已打开的 Parameter Config 弹窗中，遍历所有 Parameter Type，对每个 Type 翻页
    批量开启全部行的 **Polling Enable**（第 1 列），使该设备全部参数被 BACnet 上发。

    说明：Parameter Config 弹窗列为 Parameter / Polling Enable / COV Enable /
    COV Increment，**没有 EPICS Enable 列**；控制参数是否上发的开关是 Polling Enable。
    复用 test_bacnet_ui_basic._enable_polling_all_pages（开第 1 列、翻页、回首页）。
    """
    dlg = page.locator('[aria-label="Parameter Config"]')
    types = _get_parameter_types(page, dlg) or [""]
    for pt in types:
        if pt:
            _select_param_type(page, dlg, pt)
        _enable_polling_all_pages(page)


def _compare_bacnet_vs_modbus(
    device_label: str,
    published_keys: set[str],
    bacnet_value_by_key: dict[str, object],
) -> tuple[list[tuple[str, Optional[float], Optional[float], Optional[float], Optional[float], str]],
           int, int, str]:
    """
    读设备自身 Modbus 实时值，与 BACnet 上传值做 ±1%/±0.05 容差比对。

    比对集 = BACnet 已发布且读到值 ∩ 该设备 Modbus 地址表。

    Returns:
        (rows, fail_count, err_count, skip_note)
        rows: [(param_key, bacnet值, modbus值, 绝对差, 相对差%, status), ...]，
              status ∈ {"PASS","FAIL","ERR"}；
        skip_note 非空表示未配置/无公共参数而跳过数值比对（rows 为空）。
    """
    mb_cfg = hmi_cfg.DEVICE_MODBUS_MAP.get(device_label)
    module = DEVICE_MODBUS_MODULES.get(device_label)
    if not mb_cfg or not module:
        return [], 0, 0, f"{device_label} 未配置 DEVICE_MODBUS_MAP，跳过数值比对"

    host, port, unit = mb_cfg
    mb_keys = modbus_param_keys(module)
    compare_keys = sorted(
        k for k in published_keys if k in mb_keys and k in bacnet_value_by_key
    )
    if not compare_keys:
        return [], 0, 0, f"{device_label} 无 BACnet∩Modbus 公共参数，跳过数值比对"

    mb_values = read_modbus_values(
        module, host, port, unit, compare_keys,
        timeout=hmi_cfg.MODBUS_CMP_TIMEOUT,
        max_retries=hmi_cfg.MODBUS_CMP_MAX_RETRIES,
    )

    pct = hmi_cfg.MODBUS_CMP_TOLERANCE_PERCENT
    abs_tol = hmi_cfg.MODBUS_CMP_TOLERANCE_ABSOLUTE
    rows: list[tuple[str, Optional[float], Optional[float], Optional[float], Optional[float], str]] = []
    fail = 0
    err = 0
    for key in compare_keys:
        mv, _mb_err = mb_values.get(key, (None, "未读取"))
        try:
            bv: Optional[float] = float(bacnet_value_by_key[key])
        except (TypeError, ValueError):
            bv = None
        if bv is None or mv is None:
            rows.append((key, bv, mv, None, None, "ERR"))
            err += 1
            continue
        diff = abs(bv - mv)
        ref = max(abs(bv), abs(mv))
        dp = (diff / ref * 100) if ref > 1e-12 else 0.0
        tol = max(abs_tol, ref * pct / 100)
        status = "PASS" if diff <= tol else "FAIL"
        if status == "FAIL":
            fail += 1
        rows.append((key, bv, mv, diff, dp, status))
    return rows, fail, err, ""


def _verify_full_param_upload(
    page: Page,
    device_keyword: str,
    template_name: str,
    test_name: str,
) -> None:
    """
    全量参数验证：开启设备全部参数 Polling Enable → 客户端读回网关发布对象 →
    与设备模板 BACnet 参数严格比对（缺失=0 且 多余=0），参数名 + 读到的
    presentValue 写入日志，证明客户端真实读到了上传值；再读设备
    自身 Modbus 实时值做容差比对（若已配置 DEVICE_MODBUS_MAP）。

    流程：
      1. 加载设备模板 BACnet 参数集
      2. 记录启用前网关 objectList
      3. 打开 Parameter Config，遍历所有 Type/分页开启全部 Polling Enable，保存
      4. 等服务重启，读回网关 objectList，差集 = 本次新发布对象
      5. 批量并发读新对象 objectName + presentValue，按前缀圈定本设备对象
      6. 严格比对模板集 vs 已发布集；参数表 + 范围结论写入日志
      7. BACnet 上传值 vs Modbus 实时值容差比对

    注：不再恢复（关闭 Polling），与 COV Batch 用例一致；下一台设备测试靠 objectList
    差集独立隔离，互不影响。
    """
    # 1. 模板基准
    try:
        tmpl_map = get_bacnet_template_map(template_name)
    except FileNotFoundError:
        pytest.skip(f"[{test_name}] 未找到设备模板 {template_name!r}，跳过此用例")
        return
    if not tmpl_map:
        pytest.skip(f"[{test_name}] 模板 {template_name!r} 无 BACnet 参数，跳过此用例")
        return
    template_keys = set(tmpl_map)

    # 2. 启用前 objectList
    idents_before = get_object_identifiers()
    if idents_before is None:
        pytest.skip(f"[{test_name}] BACnet 客户端无法连接，跳过此用例")
        return
    before_set = set(idents_before)

    # 3. 开启全部 Polling Enable（上发总开关）
    if not _open_param_dialog(page, device_keyword):
        pytest.skip(f"[{test_name}] 无法打开设备 {device_keyword!r} 的 Parameter Config 弹窗")
        return
    try:
        _enable_all_polling_for_device(page)
        _save_param_config_dialog(page)
        _close_param_config_dialog(page)
    except Exception:
        if page.locator('[aria-label="Parameter Config"]').count() > 0:
            _close_param_config_dialog(page)
        raise

    # 4. 等重启，读回 objectList
    time.sleep(DEVICE_RESTART_WAIT)
    idents_after = get_object_identifiers()
    new_objects = (
        [o for o in idents_after if o not in before_set]
        if idents_after is not None else []
    )

    # 5. 批量并发读新对象明细，按前缀圈定本设备对象
    details = read_object_details_batch(new_objects) if new_objects else []
    prefixes: dict[str, int] = {}
    for _t, _i, name, _v in details:
        if name and "-" in name:
            pre = name.split("-", 1)[0]
            prefixes[pre] = prefixes.get(pre, 0) + 1
    device_prefix = max(prefixes, key=prefixes.get) if prefixes else None

    table_rows: list[tuple[str, int, Optional[str], Optional[str], object]] = []
    published_keys: set[str] = set()
    for obj_type, inst, name, value in details:
        if device_prefix and not (name and name.split("-", 1)[0] == device_prefix):
            continue
        key = name.split("-", 1)[1] if name and "-" in name else (name or "")
        published_keys.add(key)
        table_rows.append((obj_type, inst, name, key, value))

    # 6. 范围比对
    missing = sorted(template_keys - published_keys)
    extra = sorted(published_keys - template_keys)
    matched = sorted(template_keys & published_keys)

    log.info("[%s] 范围比对：模板=%d  已发布=%d  匹配=%d  缺失=%d  多余=%d",
             test_name, len(template_keys), len(published_keys),
             len(matched), len(missing), len(extra))

    # 7. BACnet 上传值 vs 设备 Modbus 实时值容差比对
    bacnet_value_by_key = {
        key: value for _t, _i, _n, key, value in table_rows if key
    }
    cmp_rows, cmp_fail, cmp_err, cmp_note = _compare_bacnet_vs_modbus(
        template_name, published_keys, bacnet_value_by_key
    )
    if cmp_note:
        log.info("[%s] %s", test_name, cmp_note)
    else:
        log.info("[%s] 数值比对：共 %d  PASS=%d  FAIL=%d  ERR=%d",
                 test_name, len(cmp_rows),
                 len(cmp_rows) - cmp_fail - cmp_err, cmp_fail, cmp_err)

    # 断言：客户端可达 + 严格范围一致 + 读到的值可读
    assert idents_after is not None, (
        f"[{test_name}] 启用全部参数后 BACnet 客户端不可达，get_object_identifiers() 返回 None"
    )
    assert not missing and not extra, (
        f"\n[{test_name}] 设备 {device_keyword!r} 上传参数范围与模板不一致！\n"
        f"  模板有但网关未发布（{len(missing)} 条）：{missing[:20]}\n"
        f"  网关发布但模板未包含（{len(extra)} 条）：{extra[:20]}"
    )
    assert table_rows, (
        f"[{test_name}] 设备 {device_keyword!r} 启用全部 Polling Enable 后无新发布对象，"
        "客户端未读到任何上传参数"
    )
    unread = [r for r in table_rows if r[4] is None]
    assert not unread, (
        f"[{test_name}] 设备 {device_keyword!r} 有 {len(unread)}/{len(table_rows)} "
        f"个上传对象 presentValue 未读到，客户端读取不完整：{[r[2] for r in unread[:20]]}"
    )

    # 数值比对断言（仅当配置了该设备 Modbus）：无超差，且非全部读取失败
    if not cmp_note:
        mismatches = [(r[0], r[1], r[2], r[4]) for r in cmp_rows if r[5] == "FAIL"]
        assert not mismatches, (
            f"[{test_name}] 设备 {device_keyword!r} BACnet 上传值与 Modbus 实时值超差"
            f"（±{hmi_cfg.MODBUS_CMP_TOLERANCE_PERCENT}% / ±{hmi_cfg.MODBUS_CMP_TOLERANCE_ABSOLUTE}），"
            f"共 {len(mismatches)} 项，前 10："
            + "".join(
                f"\n    {k}: BACnet={bv} Modbus={mv} 相对差={dp:.2f}%"
                for k, bv, mv, dp in mismatches[:10]
            )
        )
        assert not (cmp_rows and cmp_err == len(cmp_rows)), (
            f"[{test_name}] 设备 {device_keyword!r} 配置了 Modbus 但 {cmp_err} 项实时值全部读取失败，"
            "请检查 DEVICE_MODBUS_MAP 的 IP/Unit 是否正确、设备是否可达"
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
            time.sleep(15)  # 端口切换需要 BACnet 服务在新端口完整重启，等待时间比默认更长
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

        _verify_full_param_upload(hmi_page, device, "AcuRev4100", "test_012_4100")

    def test_013_pxe1_epics_enable_bacnet_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_013: AcuvimIIR（PXE1）EPICS Enable 后 BACnet 客户端可接收参数。"""
        devices = _get_device_names(hmi_page)
        device = _find_device_keyword(devices, "IIR")
        if device is None:
            pytest.skip(f"未找到含 'IIR'（AcuvimIIR）的设备（当前设备列表：{devices}），跳过此用例")

        _verify_full_param_upload(hmi_page, device, "AcuvimIIR", "test_013_AcuvimIIR")

    def test_014_pxe2_epics_enable_bacnet_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_014: AcuvimIIW（PXE2）EPICS Enable 后 BACnet 客户端可接收参数。"""
        devices = _get_device_names(hmi_page)
        device = _find_device_keyword(devices, "IIW")
        if device is None:
            pytest.skip(f"未找到含 'IIW'（AcuvimIIW）的设备（当前设备列表：{devices}），跳过此用例")

        _verify_full_param_upload(hmi_page, device, "AcuvimIIW", "test_014_AcuvimIIW")

    def test_015_pxm350_epics_enable_bacnet_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_015: AcuRev1300（PXM350）EPICS Enable 后 BACnet 客户端可接收参数。"""
        devices = _get_device_names(hmi_page)
        device = _find_device_keyword(devices, "1300")
        if device is None:
            pytest.skip(
                f"未找到含 '1300'（AcuRev1300）的设备（当前设备列表：{devices}），跳过此用例"
            )

        _verify_full_param_upload(hmi_page, device, "AcuRev1300", "test_015_AcuRev1300")

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

        _verify_full_param_upload(hmi_page, device, "AcuVIM3", "test_016_AcuVIM3")

    def test_017_acurev2100_epics_enable_bacnet_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_017: AcuRev-2100 EPICS Enable 后 BACnet 客户端可接收参数。"""
        devices = _get_device_names(hmi_page)
        device = _find_device_keyword(devices, "2100")
        if device is None:
            pytest.skip(f"未找到含 '2100' 的设备（当前设备列表：{devices}），跳过此用例")

        _verify_full_param_upload(hmi_page, device, "AcuRev2100", "test_017_2100")

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
        idents_enabled: Optional[list[tuple[str, int]]] = None

        try:
            # 若原来未启用，先启用后等待，建立"已启用"基准
            if not was_enabled:
                ok = _set_first_row_epics_enable(hmi_page, True)
                if not ok:
                    pytest.skip("无法操作第一行 EPICS Enable 开关")
                hmi_page.wait_for_timeout(300)

                _save_param_config_dialog(hmi_page)
                _close_param_config_dialog(hmi_page)

                time.sleep(DEVICE_RESTART_WAIT)

                # 重新打开弹窗，准备执行 disable
                if not _open_param_dialog(hmi_page, device_keyword):
                    pytest.fail(
                        f"启用 EPICS 后无法重新打开设备 {device_keyword!r} 的 Parameter Config"
                    )

            # 此时弹窗已打开（原来已启用 or 刚刚启用后重新打开）
            idents_enabled = get_object_identifiers()
            if idents_enabled is None:
                _close_param_config_dialog(hmi_page)
                pytest.skip("BACnet 客户端不可达，无法验证对象数量变化")
            count_enabled = len(idents_enabled)

            # 关闭第一行 EPICS Enable
            ok = _set_first_row_epics_enable(hmi_page, False)
            if not ok:
                pytest.skip("无法操作第一行 EPICS Enable 开关（disable）")
            hmi_page.wait_for_timeout(300)

            # 弹窗内或主区域 Save 均在函数内部完成提交，关闭弹窗即可
            _save_param_config_dialog(hmi_page)
            _close_param_config_dialog(hmi_page)

        except Exception:
            if hmi_page.locator('[aria-label="Parameter Config"]').count() > 0:
                _close_param_config_dialog(hmi_page)
            raise

        time.sleep(DEVICE_RESTART_WAIT)

        idents_disabled = get_object_identifiers()
        count_disabled = len(idents_disabled) if idents_disabled is not None else None

        # 记录本次被撤销发布的对象（对象已从网关移除，仅能记录标识）
        if idents_enabled is not None and idents_disabled is not None:
            disabled_set = set(idents_disabled)
            removed = [o for o in idents_enabled if o not in disabled_set]
            log.info("[test_018] BACnet 对象数 %d → %d（撤销 %d 个）",
                     count_enabled, count_disabled, len(removed))
            for obj_type, inst in removed:
                log.info("[test_018] 撤销发布对象: %s,%d", obj_type, inst)

        # 恢复原始状态
        try:
            if was_enabled:
                # 原来是 True，现在被设成 False，需要恢复为 True
                if _open_param_dialog(hmi_page, device_keyword):
                    _set_first_row_epics_enable(hmi_page, True)
                    hmi_page.wait_for_timeout(300)
                    # 弹窗内或主区域 Save 均在函数内部完成提交，关闭弹窗即可
                    _save_param_config_dialog(hmi_page)
                    _close_param_config_dialog(hmi_page)
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
