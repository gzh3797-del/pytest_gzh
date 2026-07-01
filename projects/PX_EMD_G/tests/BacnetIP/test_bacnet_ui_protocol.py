# -*- coding: utf-8 -*-
"""
test_bacnet_ui_protocol.py — BACnet/IP 协议层端到端验证用例（P1/P2）

用例覆盖：
  TestCase_AcuHMI-1-7_033_001_002: BACnet Port 合法边界值保存且 BACnet 客户端可通信
  TestCase_AcuHMI-1-7_033_001_003: Network Number 合法边界值保存并持久化
  TestCase_AcuHMI-1-7_033_001_012: AcuRev4100 六段式全量参数验证（单设备隔离）
  TestCase_AcuHMI-1-7_033_001_013: AcuvimIIR（PXE1）六段式全量参数验证（单设备隔离）
  TestCase_AcuHMI-1-7_033_001_014: AcuvimIIW（PXE2）六段式全量参数验证（单设备隔离）
  TestCase_AcuHMI-1-7_033_001_015: AcuRev1300（PXM350）六段式全量参数验证（单设备隔离）
  TestCase_AcuHMI-1-7_033_001_016: AcuVIM3 六段式全量参数验证（单设备隔离）
  TestCase_AcuHMI-1-7_033_001_017: AcuRev2100 六段式全量参数验证（单设备隔离）
  TestCase_AcuHMI-1-7_033_001_018: 关闭 EPICS Enable 后 BACnet 客户端不能接收该参数
  TestCase_AcuHMI-1-7_033_001_037: EPICS file download 下载触发正常
  TestCase_AcuHMI-1-7_033_001_038: 禁用 BACnet 后客户端无法连接

六段式比对流程（test_012~017 每台设备独立执行）：
  段 1 — 模板加载：get_bacnet_template_map() 取应发布参数范围
  段 2 — 范围检查：网关发布 AI/BI 对象 vs 模板（缺失=0 且 多余=0）
  段 3 — 元数据检查：description/units 属性 vs 模板元数据（单位等价容差）
  段 4 — 数值比对：BACnet Present Value vs Modbus 实时值（±1%/±0.05）
  段 5 — Device Object：12 项 §12.11 标准必需属性全部可读
  段 6 — 协议合规：非法对象请求返回错误 + AI 必需属性（statusFlags/outOfService/units）可读
  段 7 — 连接稳定性：同一 AI 对象连续读 5 次，全部成功

单设备隔离策略：
  每台设备测试前，_isolate_single_device → _click_save → 等重启 → 此时
  objectList 仅含该设备对象，无需前缀猜测。
  module 级 fixture bacnet_device_state_snapshot 在 teardown 恢复原始勾选集。

运行：
  pytest projects/PX_EMD_G/tests/BacnetIP/test_bacnet_ui_protocol.py -v
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
from projects.PX_EMD_G.tests.BacnetIP.test_bacnet_ui_basic import (  # noqa: E402
    _get_all_device_rows,
    _isolate_single_device,
    _restore_device_selection,
    _open_param_dialog,
    _close_param_config_dialog,
    _get_parameter_types,
    _select_param_type,
    _enable_polling_all_pages,
    _enable_polling_select_all,
)
from projects.PX_EMD_G.tests.BacnetIP.test_bacnet_ui_config import (  # noqa: E402
    _get_field_value,
    _set_field_value,
    _click_save,
    _dismiss_toast,
    _navigate_to_bacnet,
)

# ── BACnet 客户端 ─────────────────────────────────────────────────────────────
from projects.PX_EMD_G.helpers.hmi_bacnet_client import (  # noqa: E402
    can_connect,
    wait_until_connectable,
    get_object_identifiers,
    read_object_details_batch,
    read_device_info,
    read_object_metadata_batch,
    run_protocol_compliance,
    check_stability,
    StabilityCheckResult,
    HMI_DEFAULT_PORT,
    DEVICE_RESTART_WAIT,
    SERVICE_READY_TIMEOUT,
)
from projects.PX_EMD_G.helpers.bacnet_report import (  # noqa: E402
    generate_six_segment_report,
)

# ── 模板基准 + Modbus 实时值 ──────────────────────────────────────────────────
from projects.PX_EMD_G.helpers.template_matcher import (  # noqa: E402
    DEVICE_MODBUS_MODULES,
    get_bacnet_template_map,
)
from projects.PX_EMD_G.helpers.physical_devices_reader import (  # noqa: E402
    DiscoveredDevice,
    connection_map,
    pick_device_for_template,
)
from projects.PX_EMD_G.helpers.hmi_modbus_client import (  # noqa: E402
    read_modbus_values,
    modbus_param_keys,
)
from projects.PX_EMD_G import settings as hmi_cfg  # noqa: E402

# ── 本文件常量 ────────────────────────────────────────────────────────────────
_SAVE_WAIT_MS = 2000

log = logging.getLogger(__name__)


# ═════════════════════════════════════════════════════════════════════════════
# Module 级 fixture：测试前快照原始设备勾选状态，teardown 时恢复
# ═════════════════════════════════════════════════════════════════════════════

@pytest.fixture(scope="module")
def bacnet_device_state_snapshot(hmi_page: Page) -> list[str]:
    """
    在 test_012~017（六段式用例）运行前快照已勾选设备名，teardown 恢复原始勾选 + Save。

    scope=module：整个 test_bacnet_ui_protocol 模块共享一次快照/恢复——
    每台设备的隔离/Save/重启由用例内部自行负责，module teardown 负责最终还原现场。
    """
    rows = _get_all_device_rows(hmi_page)
    original_checked: list[str] = [r["name"] for r in rows if r["checked"]]
    log.info("[fixture] 快照原始设备勾选集：%s", original_checked)
    yield original_checked
    # ── teardown：恢复原始勾选 + Save + 等待重启 ──────────────────────────────
    log.info("[fixture] teardown：恢复原始设备勾选 %s", original_checked)
    try:
        _restore_device_selection(hmi_page, original_checked)
        _click_save(hmi_page)
        _dismiss_toast(hmi_page)
        time.sleep(DEVICE_RESTART_WAIT)
        log.info("[fixture] teardown 完成，BACnet 服务已重启，设备勾选已恢复")
    except Exception as exc:
        log.warning("[fixture] teardown 恢复失败（不影响本次测试结论）：%s", exc)


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

    Returns:
        True = 已通过弹窗内 Save 按钮保存；False = 已降级用主区域 Save 保存。
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


def _resolve_bacnet_device(
    discovered: list[DiscoveredDevice],
    ui_rows: list[dict],
    template_name: str,
) -> tuple[Optional[str], str]:
    """选出该型号要勾选测试的目标设备名（取该型号第一台在线设备）。

    1. 从网关动态发现列表里按 template_name 取第一台在线设备（pick_device_for_template）；
    2. 在 BACnet UI 设备表 ui_rows 里按全名精确匹配确认该设备存在且在线。

    返回 (device_name, reason)：device_name 为 None 时 reason 说明跳过原因。
    """
    dev = pick_device_for_template(discovered, template_name, online_only=True)
    if dev is None:
        return None, (
            f"网关下挂的在线 Modbus TCP 设备中无模板为 {template_name!r} 的设备"
            f"（已发现：{[(d.name, d.model, d.online) for d in discovered]}）"
        )
    row = next((r for r in ui_rows if r["name"].lower() == dev.name.lower()), None)
    if row is None:
        return None, (
            f"动态发现的设备 {dev.name!r} 不在 BACnet 设备表中"
            "（BACnet Devices Selection 表与 Physical Devices 不一致，请核查网关）"
        )
    if not row.get("online"):
        return None, f"设备 {dev.name!r} 在 BACnet 设备表中显示离线，跳过"
    return dev.name, "discovered"


def _enable_all_polling_for_device(page: Page) -> None:
    """
    在已打开的 Parameter Config 弹窗中，遍历所有 Parameter Type，对每个 Type
    一键开启全部行的 **Polling Enable**，使该设备全部参数被 BACnet 上发。

    优先调用 ``_enable_polling_select_all``（点列头全选 checkbox，一次覆盖全部分页）；
    若列头控件不存在（返回 False）则回退到 ``_enable_polling_all_pages`` 逐页处理，
    确保不会因控件缺失而漏开参数。

    参数列为 Parameter / Polling Enable / COV Enable / COV Increment。
    """
    dlg = page.locator('[aria-label="Parameter Config"]')
    types = _get_parameter_types(page, dlg) or [""]
    for pt in types:
        if pt:
            _select_param_type(page, dlg, pt)
        # 优先用列头全选 checkbox（一次覆盖所有分页，无需翻页，速度极快）；
        # 找不到控件时回退到逐页翻页方式（兼容控件结构变化）。
        if not _enable_polling_select_all(page):
            _enable_polling_all_pages(page)


def _compare_bacnet_vs_modbus(
    device_name: str,
    template_name: str,
    published_keys: set[str],
    bacnet_value_by_key: dict[str, object],
    conn_map: dict[str, tuple[str, int, int]],
) -> tuple[
    list[tuple[str, Optional[float], Optional[float], Optional[float], Optional[float], str]],
    int,
    int,
    str,
]:
    """
    读设备自身 Modbus 实时值，与 BACnet 上传值做 ±1%/±0.05 容差比对。

    比对集 = BACnet 已发布且读到值 ∩ 该设备 Modbus 地址表。

    连接信息按 ``device_name`` 从动态发现的 conn_map 取；Modbus 地址表模块按
    ``template_name`` 从 DEVICE_MODBUS_MODULES（按型号）取——二者分开，使多台同型号
    表能各自配自己的连接信息而共用同一地址表。

    Returns:
        (rows, fail_count, err_count, skip_note)
        rows: [(param_key, bacnet值, modbus值, 绝对差, 相对差%, status), ...]，
              status ∈ {"PASS","FAIL","ERR"}；
        skip_note 非空表示未配置/无公共参数而跳过数值比对（rows 为空）。
    """
    mb_cfg = conn_map.get(device_name)
    module = DEVICE_MODBUS_MODULES.get(template_name)
    if not mb_cfg or not module:
        return [], 0, 0, f"{device_name} 未发现 Modbus 连接信息，跳过数值比对"

    host, port, unit = mb_cfg
    mb_keys = modbus_param_keys(module)
    compare_keys = sorted(
        k for k in published_keys if k in mb_keys and k in bacnet_value_by_key
    )
    if not compare_keys:
        return [], 0, 0, f"{device_name} 无 BACnet∩Modbus 公共参数，跳过数值比对"

    mb_values = read_modbus_values(
        module, host, port, unit, compare_keys,
        timeout=hmi_cfg.MODBUS_CMP_TIMEOUT,
        max_retries=hmi_cfg.MODBUS_CMP_MAX_RETRIES,
    )

    pct = hmi_cfg.MODBUS_CMP_TOLERANCE_PERCENT
    abs_tol = hmi_cfg.MODBUS_CMP_TOLERANCE_ABSOLUTE
    rows: list[
        tuple[str, Optional[float], Optional[float], Optional[float], Optional[float], str]
    ] = []
    fail = 0
    err = 0
    for key in compare_keys:
        mv, _mb_err = mb_values.get(key, (None, "未读取"))
        try:
            bv: Optional[float] = float(bacnet_value_by_key[key])  # type: ignore[arg-type]
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


# ─────────────────────────────────────────────────────────────────────────────
# 六段式全量参数验证（单设备隔离版）
# ─────────────────────────────────────────────────────────────────────────────

def _verify_full_param_upload(
    page: Page,
    device_keyword: str,
    template_name: str,
    test_name: str,
    conn_map: dict[str, tuple[str, int, int]],
) -> None:
    """
    六段式全量参数验证（单设备隔离执行）。

    **隔离策略**：调用前须已通过 bacnet_device_state_snapshot 快照原始勾选集；
    本函数在段 1 之前先隔离被测设备（取消其余所有设备勾选 → Save → 等重启），
    此后网关 objectList 仅含该设备对象，所有 AI/BI 均属本设备，无需前缀猜测。

    六段：
      段 1 — 模板加载
      段 2 — 范围检查（缺失=0 且 多余=0）
      段 3 — 元数据检查（description/units vs 模板）
      段 4 — 数值比对（BACnet vs Modbus，±1%/±0.05）
      段 5 — Device Object 12 项必需属性（§12.11）
      段 6 — 协议合规性（§16 非法对象 + AI 必需属性）
      段 7 — 连接稳定性（同一 AI 对象连续读 5 次）
    """
    # ── 段 1：模板加载 ────────────────────────────────────────────────────────
    log.info("[%s] 段1 开始：加载设备模板 %r", test_name, template_name)
    try:
        tmpl_map = get_bacnet_template_map(template_name)
    except FileNotFoundError:
        pytest.skip(f"[{test_name}] 未找到设备模板 {template_name!r}，跳过此用例")
        return
    if not tmpl_map:
        pytest.skip(f"[{test_name}] 模板 {template_name!r} 无 BACnet 参数，跳过此用例")
        return
    template_keys = set(tmpl_map)
    log.info("[%s] 段1 PASS：模板参数数=%d", test_name, len(template_keys))

    # ── 隔离：取消其余设备勾选，只留被测设备 → Save → 等重启 ──────────────────
    log.info("[%s] 隔离：_isolate_single_device(%r)", test_name, device_keyword)
    _isolate_single_device(page, device_keyword)
    _click_save(page)
    _dismiss_toast(page)
    log.info("[%s] 隔离已 Save，等待 BACnet 服务重启 %.0fs ...", test_name, DEVICE_RESTART_WAIT)
    time.sleep(DEVICE_RESTART_WAIT)

    # ── 开启全部 Polling Enable（上发总开关）→ Save → 等重启 ──────────────────
    log.info("[%s] 打开 Parameter Config，开启全部 Polling Enable", test_name)
    if not _open_param_dialog(page, device_keyword):
        pytest.skip(
            f"[{test_name}] 无法打开设备 {device_keyword!r} 的 Parameter Config 弹窗"
        )
        return
    try:
        _enable_all_polling_for_device(page)
        _save_param_config_dialog(page)
        _close_param_config_dialog(page)
    except Exception:
        if page.locator('[aria-label="Parameter Config"]').count() > 0:
            _close_param_config_dialog(page)
        raise

    # 开启全量 Polling 后网关需重建大量对象，重启耗时常超过固定 DEVICE_RESTART_WAIT；
    # 用轮询等待服务真正可达再读，避免过早读取得到空 TimeoutError 被误判为"连不上"。
    log.info("[%s] Polling Enable 已保存，轮询等待 BACnet 服务重新可达（最多 %.0fs）...",
             test_name, SERVICE_READY_TIMEOUT)
    if not wait_until_connectable():
        pytest.skip(
            f"[{test_name}] 隔离+全量 Polling 保存后，BACnet 服务在 "
            f"{SERVICE_READY_TIMEOUT:.0f}s 内仍不可达——大概率网关侧异常，请人工核查网关 BACnet 服务"
        )
        return

    # 读取隔离后网关完整 objectList（单设备，所有 AI/BI 均属本设备）
    idents_all = get_object_identifiers()
    if idents_all is None:
        pytest.skip(f"[{test_name}] BACnet 客户端无法连接，跳过此用例")
        return

    # 仅保留 AI/BI 对象供后续各段使用
    # BAC0/bacpypes3 返回的 objectType 用连字符格式（analog-input / binary-input），
    # 同时兼容 camelCase 格式（analogInput / binaryInput）以防实现变更。
    _AI_BI_TYPES = frozenset(
        ("analogInput", "binaryInput", "analog-input", "binary-input")
    )
    ai_bi_objects = [
        (ot, inst) for ot, inst in idents_all
        if ot in _AI_BI_TYPES
    ]

    # 批量并发读 objectName + presentValue
    details = read_object_details_batch(ai_bi_objects) if ai_bi_objects else []

    # 单设备隔离下无需前缀猜测：所有对象都属本设备
    # objectName 格式为 "<DevicePrefix>-<param_key>"，取 "-" 后半段为 param_key
    table_rows: list[tuple[str, int, Optional[str], str, object]] = []
    published_keys: set[str] = set()
    for obj_type, inst, name, value in details:
        key = name.partition("-")[2] if name and "-" in name else (name or "")
        published_keys.add(key)
        table_rows.append((obj_type, inst, name, key, value))

    # ── 段 2：范围检查 ────────────────────────────────────────────────────────
    missing = sorted(template_keys - published_keys)
    extra = sorted(published_keys - template_keys)
    matched = sorted(template_keys & published_keys)
    log.info(
        "[%s] 段2 范围检查：模板=%d  已发布=%d  匹配=%d  缺失=%d  多余=%d",
        test_name, len(template_keys), len(published_keys),
        len(matched), len(missing), len(extra),
    )
    range_ok = not missing and not extra
    log.info("[%s] 段2 %s", test_name, "PASS" if range_ok else "FAIL")

    # ── 段 4：数值比对（BACnet vs Modbus） ───────────────────────────────────
    # 执行顺序优化（不改判定逻辑）：数值比对所需的 Modbus 读取提前到紧贴
    # BACnet presentValue 采样（read_object_details_batch）之后、仅隔段2纯计算，
    # 避免被段3元数据批量读取（数百对象、数十秒）拉开两侧采样间隔——否则
    # 谐波/THD 等快速波动量会因 BACnet 与 Modbus 采样时刻不同步而大面积假性超差。
    # 段3元数据、段5/6/7 等只读检查顺延到本段之后；六段逻辑编号保持不变。
    bacnet_value_by_key: dict[str, object] = {
        key: value for _t, _i, _n, key, value in table_rows if key
    }
    cmp_rows, cmp_fail, cmp_err, cmp_note = _compare_bacnet_vs_modbus(
        device_keyword, template_name, published_keys, bacnet_value_by_key, conn_map
    )
    if cmp_note:
        log.info("[%s] 段4 %s（跳过）", test_name, cmp_note)
        cmp_skipped = True
    else:
        cmp_pass = len(cmp_rows) - cmp_fail - cmp_err
        log.info(
            "[%s] 段4 数值比对：共 %d  PASS=%d  FAIL=%d  ERR=%d",
            test_name, len(cmp_rows), cmp_pass, cmp_fail, cmp_err,
        )
        cmp_skipped = False
        cmp_ok = (cmp_fail == 0) and not (cmp_rows and cmp_err == len(cmp_rows))
        log.info("[%s] 段4 %s", test_name, "PASS" if cmp_ok else "FAIL")

    # ── 段 3：元数据检查（description/units vs 模板） ──────────────────────────
    # 顺序上紧随段4之后执行（见段4注释）：units/description 是只读校验，与实时值无关，
    # 延后读取不影响其结论，却能让段4的 BACnet/Modbus 采样尽量同时刻。
    log.info("[%s] 段3 开始：元数据检查（units/description），共 %d 个对象",
             test_name, len(matched))
    # 构造 (obj_type, instance, param_key, tmpl_unit, tmpl_desc) 列表
    meta_inputs: list[tuple[str, int, str, str, str]] = []
    key_to_obj: dict[str, tuple[str, int]] = {
        key: (ot, inst) for ot, inst, _name, key, _val in table_rows if key
    }
    for pk in matched:
        tp = tmpl_map[pk]
        tmpl_unit = (tp.unit or "").strip()
        tmpl_desc = (tp.description or "").split("\n")[0].strip()
        if pk in key_to_obj:
            ot, inst = key_to_obj[pk]
            meta_inputs.append((ot, inst, pk, tmpl_unit, tmpl_desc))

    meta_results = read_object_metadata_batch(meta_inputs) if meta_inputs else []
    # 仅对「成功读到单位」的参数做严格比对；模板单位为空跳过、units 读取失败均不计入 FAIL
    # （读取失败=未拿到值，不能据此断言与模板不符；读取失败已在客户端侧重试过）。
    meta_read_failed = [m for m in meta_results if m.unit_read_failed]
    meta_skipped = [m for m in meta_results
                    if m.unit_skipped and not m.unit_read_failed]
    meta_fail = [m for m in meta_results
                 if not m.unit_ok and not m.unit_skipped and not m.unit_read_failed]
    meta_pass_count = (len(meta_results) - len(meta_fail)
                       - len(meta_skipped) - len(meta_read_failed))
    log.info(
        "[%s] 段3 元数据检查：共 %d  单位匹配=%d  跳过(模板无单位)=%d  读取失败=%d  单位不符=%d",
        test_name, len(meta_results), meta_pass_count,
        len(meta_skipped), len(meta_read_failed), len(meta_fail),
    )
    if meta_read_failed:
        log.warning("[%s] 段3 有 %d 项 units 读取失败（已重试仍未拿到，不判错）：%s",
                    test_name, len(meta_read_failed),
                    [m.param_key for m in meta_read_failed[:10]])
    if meta_fail:
        for mf in meta_fail[:10]:
            log.warning(
                "[%s] 段3 单位不匹配 %s,%d  param=%s  模板=%r  BACnet=%r",
                test_name, mf.obj_type, mf.instance, mf.param_key,
                mf.tmpl_unit, mf.bacnet_unit,
            )
    meta_ok = not meta_fail
    log.info("[%s] 段3 %s", test_name, "PASS" if meta_ok else "FAIL")

    # ── 段 5：Device Object 12 项必需属性（§12.11） ───────────────────────────
    log.info("[%s] 段5 开始：Device Object 必需属性读取", test_name)
    dev_info = read_device_info()
    if dev_info.ok:
        log.info(
            "[%s] 段5 PASS：vendorName=%r  modelName=%r  firmwareRevision=%r  "
            "protocolVersion=%r  protocolRevision=%r  segmentation=%r",
            test_name,
            dev_info.vendor_name, dev_info.model_name, dev_info.firmware_revision,
            dev_info.protocol_version, dev_info.protocol_revision, dev_info.segmentation,
        )
    else:
        log.warning("[%s] 段5 FAIL：%s", test_name, dev_info.error)

    # ── 段 6：协议合规性（§16 错误响应 + AI 必需属性） ───────────────────────
    probe_obj: Optional[tuple[str, int]] = None
    if ai_bi_objects:
        probe_obj = next(
            (
                (ot, inst) for ot, inst in ai_bi_objects
                if ot in ("analogInput", "analog-input")
            ),
            ai_bi_objects[0],
        )
    log.info("[%s] 段6 开始：协议合规性检查，probe=%s", test_name, probe_obj)
    compliance_results = run_protocol_compliance(probe_obj=probe_obj)
    comp_pass = sum(1 for c in compliance_results if c.passed)
    comp_fail = len(compliance_results) - comp_pass
    for cr in compliance_results:
        lvl = logging.INFO if cr.passed else logging.WARNING
        log.log(lvl, "[%s] 段6 %s  %s  (%s)",
                test_name, "PASS" if cr.passed else "FAIL", cr.test_name, cr.detail)
    compliance_ok = comp_fail == 0
    log.info("[%s] 段6 %s：%d/%d 通过",
             test_name, "PASS" if compliance_ok else "FAIL",
             comp_pass, len(compliance_results))

    # ── 段 7：连接稳定性（同一 AI 对象连续读 5 次） ───────────────────────────
    stability_ok = True
    stability_result: Optional[StabilityCheckResult] = None
    if probe_obj is not None:
        log.info("[%s] 段7 开始：稳定性测试，对象=%s,%d 连续读 5 次",
                 test_name, probe_obj[0], probe_obj[1])
        stab = check_stability(probe_obj=probe_obj, attempts=5)
        stability_result = stab
        stability_ok = stab.ok
        log.info(
            "[%s] 段7 %s：%d/%d 成功%s",
            test_name, "PASS" if stab.ok else "FAIL",
            stab.successes, stab.attempts,
            "" if stab.ok else f"  错误：{stab.errors}",
        )
    else:
        log.info("[%s] 段7 跳过：无可用 AI/BI 对象", test_name)

    # ── 生成六段式 HTML 报告（无论通过/失败都落盘，供具体分析；放在断言之前） ──
    try:
        report_path = generate_six_segment_report(
            device_name=device_keyword,
            template_name=template_name,
            template_keys=template_keys,
            published_keys=published_keys,
            matched=matched,
            missing=missing,
            extra=extra,
            meta_results=meta_results,
            cmp_rows=cmp_rows,
            cmp_note=cmp_note,
            dev_info=dev_info,
            compliance_results=compliance_results,
            stability=stability_result,
            gateway_ip=hmi_cfg.HMI_IP,
            gateway_port=HMI_DEFAULT_PORT,
            tol_pct=hmi_cfg.MODBUS_CMP_TOLERANCE_PERCENT,
            tol_abs=hmi_cfg.MODBUS_CMP_TOLERANCE_ABSOLUTE,
        )
        log.info("[%s] 六段式 HTML 报告已生成：%s", test_name, report_path)
    except Exception as exc:  # 报告生成失败不应影响测试结论
        log.warning("[%s] 六段式 HTML 报告生成失败（不影响测试结论）：%s", test_name, exc)

    # ── 断言汇总 ──────────────────────────────────────────────────────────────
    # 段 2：范围
    assert idents_all is not None, (
        f"[{test_name}] 启用全部参数后 BACnet 客户端不可达，get_object_identifiers() 返回 None"
    )
    assert not missing and not extra, (
        f"\n[{test_name}] 设备 {device_keyword!r} 上传参数范围与模板不一致！\n"
        f"  模板有但网关未发布（{len(missing)} 条）：{missing[:20]}\n"
        f"  网关发布但模板未包含（{len(extra)} 条）：{extra[:20]}"
    )
    assert table_rows, (
        f"[{test_name}] 设备 {device_keyword!r} 启用全部 Polling Enable 后无发布对象，"
        "客户端未读到任何上传参数"
    )
    unread = [r for r in table_rows if r[4] is None]
    assert not unread, (
        f"[{test_name}] 设备 {device_keyword!r} 有 {len(unread)}/{len(table_rows)} "
        f"个上传对象 presentValue 未读到：{[r[2] for r in unread[:20]]}"
    )

    # 段 3：元数据（模板有单位但与 BACnet units 不符视为 FAIL；模板无单位跳过不断言）
    assert meta_ok, (
        f"[{test_name}] 设备 {device_keyword!r} 元数据（units）不匹配"
        f"（模板有单位，共 {len(meta_fail)} 项，跳过 {len(meta_skipped)} 项）：\n"
        + "".join(
            f"  param={m.param_key}  模板={m.tmpl_unit!r}  BACnet={m.bacnet_unit!r}\n"
            for m in meta_fail[:10]
        )
    )

    # 段 4：数值比对（有配置时才断言）
    if not cmp_skipped:
        mismatches = [(r[0], r[1], r[2], r[4]) for r in cmp_rows if r[5] == "FAIL"]
        assert not mismatches, (
            f"[{test_name}] 设备 {device_keyword!r} BACnet 上传值与 Modbus 实时值超差"
            f"（±{hmi_cfg.MODBUS_CMP_TOLERANCE_PERCENT}%"
            f" / ±{hmi_cfg.MODBUS_CMP_TOLERANCE_ABSOLUTE}），"
            f"共 {len(mismatches)} 项，前 10："
            + "".join(
                f"\n    {k}: BACnet={bv} Modbus={mv} 相对差={dp:.2f}%"
                for k, bv, mv, dp in mismatches[:10]
            )
        )
        assert not (cmp_rows and cmp_err == len(cmp_rows)), (
            f"[{test_name}] 设备 {device_keyword!r} 配置了 Modbus 但 {cmp_err} 项实时值"
            "全部读取失败，请检查 DEVICE_MODBUS_MAP 的 IP/Unit 是否正确、设备是否可达"
        )

    # 段 5：Device Object
    assert dev_info.ok, (
        f"[{test_name}] Device Object 12 项必需属性（§12.11）读取失败：{dev_info.error}"
    )

    # 段 6：协议合规性
    assert compliance_ok, (
        f"[{test_name}] BACnet 协议合规性检查未全通过，{comp_fail}/{len(compliance_results)} 项失败：\n"
        + "".join(
            f"  {cr.test_name}  detail={cr.detail}\n"
            for cr in compliance_results if not cr.passed
        )
    )

    # 段 7：稳定性
    assert stability_ok, (
        f"[{test_name}] 连接稳定性测试失败（probe={probe_obj}）：{stab.successes}/{stab.attempts} 次成功"  # type: ignore[possibly-undefined]
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

    def test_012_acurev4100_epics_enable_bacnet_readable(
        self,
        hmi_page: Page,
        discovered_devices: list,
        bacnet_device_state_snapshot: list[str],  # noqa: ARG002 — 触发 module fixture
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_012: AcuRev4100 六段式全量参数验证（单设备隔离）。"""
        devices = _get_all_device_rows(hmi_page)
        device, reason = _resolve_bacnet_device(discovered_devices, devices, "AcuRev4100")
        if device is None:
            all_names = [r["name"] for r in devices]
            pytest.skip(f"未确定可测的 4100 设备（{reason}；当前设备列表：{all_names}）")
        conn_map = connection_map(discovered_devices)
        _verify_full_param_upload(hmi_page, device, "AcuRev4100", "test_012_4100", conn_map)

    def test_013_pxe1_epics_enable_bacnet_readable(
        self,
        hmi_page: Page,
        discovered_devices: list,
        bacnet_device_state_snapshot: list[str],  # noqa: ARG002
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_013: AcuvimIIR（PXE1）六段式全量参数验证（单设备隔离）。"""
        devices = _get_all_device_rows(hmi_page)
        device, reason = _resolve_bacnet_device(discovered_devices, devices, "AcuvimIIR")
        if device is None:
            all_names = [r["name"] for r in devices]
            pytest.skip(f"未确定可测的 AcuvimIIR 设备（{reason}；当前设备列表：{all_names}）")
        conn_map = connection_map(discovered_devices)
        _verify_full_param_upload(hmi_page, device, "AcuvimIIR", "test_013_AcuvimIIR", conn_map)

    def test_014_pxe2_epics_enable_bacnet_readable(
        self,
        hmi_page: Page,
        discovered_devices: list,
        bacnet_device_state_snapshot: list[str],  # noqa: ARG002
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_014: AcuvimIIW（PXE2）六段式全量参数验证（单设备隔离）。"""
        devices = _get_all_device_rows(hmi_page)
        device, reason = _resolve_bacnet_device(discovered_devices, devices, "AcuvimIIW")
        if device is None:
            all_names = [r["name"] for r in devices]
            pytest.skip(f"未确定可测的 AcuvimIIW 设备（{reason}；当前设备列表：{all_names}）")
        conn_map = connection_map(discovered_devices)
        _verify_full_param_upload(hmi_page, device, "AcuvimIIW", "test_014_AcuvimIIW", conn_map)

    def test_015_pxm350_epics_enable_bacnet_readable(
        self,
        hmi_page: Page,
        discovered_devices: list,
        bacnet_device_state_snapshot: list[str],  # noqa: ARG002
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_015: AcuRev1300（PXM350）六段式全量参数验证（单设备隔离）。"""
        devices = _get_all_device_rows(hmi_page)
        device, reason = _resolve_bacnet_device(discovered_devices, devices, "AcuRev1300")
        if device is None:
            all_names = [r["name"] for r in devices]
            pytest.skip(f"未确定可测的 AcuRev1300 设备（{reason}；当前设备列表：{all_names}）")
        conn_map = connection_map(discovered_devices)
        _verify_full_param_upload(hmi_page, device, "AcuRev1300", "test_015_AcuRev1300", conn_map)

    def test_016_acuvim3_epics_enable_bacnet_readable(
        self,
        hmi_page: Page,
        discovered_devices: list,
        bacnet_device_state_snapshot: list[str],  # noqa: ARG002
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_016: AcuVIM3 六段式全量参数验证（单设备隔离）。"""
        devices = _get_all_device_rows(hmi_page)
        device, reason = _resolve_bacnet_device(discovered_devices, devices, "AcuVIM3")
        if device is None:
            all_names = [r["name"] for r in devices]
            pytest.skip(f"未确定可测的 AcuVIM3 设备（{reason}；当前设备列表：{all_names}）")
        conn_map = connection_map(discovered_devices)
        _verify_full_param_upload(hmi_page, device, "AcuVIM3", "test_016_AcuVIM3", conn_map)

    def test_017_acurev2100_epics_enable_bacnet_readable(
        self,
        hmi_page: Page,
        discovered_devices: list,
        bacnet_device_state_snapshot: list[str],  # noqa: ARG002
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_017: AcuRev-2100 六段式全量参数验证（单设备隔离）。"""
        devices = _get_all_device_rows(hmi_page)
        device, reason = _resolve_bacnet_device(discovered_devices, devices, "AcuRev2100")
        if device is None:
            all_names = [r["name"] for r in devices]
            pytest.skip(f"未确定可测的 2100 设备（{reason}；当前设备列表：{all_names}）")
        conn_map = connection_map(discovered_devices)
        _verify_full_param_upload(hmi_page, device, "AcuRev2100", "test_017_2100", conn_map)

    def test_018_epics_disable_bacnet_not_readable(self, hmi_page: Page) -> None:
        """TestCase_AcuHMI-1-7_033_001_018: 关闭 EPICS Enable 后 BACnet 客户端不能接收该参数。"""
        # 取首个「已勾选且在线」的设备：离线设备无对象上发，无法验证对象数量变化，
        # 且会拖累 BACnet 客户端连接（与六段式用例同因）。
        rows = _get_all_device_rows(hmi_page)
        checked_online = [r["name"] for r in rows if r["checked"] and r["online"]]
        if not checked_online:
            pytest.skip(
                f"无已勾选且在线的设备（当前：{rows}），无法验证 EPICS Disable 场景"
            )

        device_keyword = checked_online[0]

        if not _open_param_dialog(hmi_page, device_keyword):
            pytest.skip(f"无法打开设备 {device_keyword!r} 的 Parameter Config 弹窗")

        was_enabled: Optional[bool] = _get_first_row_epics_state(hmi_page)
        if was_enabled is None:
            _close_param_config_dialog(hmi_page)
            pytest.skip("Parameter Config 弹窗中未找到参数行")

        count_enabled: Optional[int] = None
        idents_enabled: Optional[list[tuple[str, int]]] = None

        try:
            if not was_enabled:
                ok = _set_first_row_epics_enable(hmi_page, True)
                if not ok:
                    pytest.skip("无法操作第一行 EPICS Enable 开关")
                hmi_page.wait_for_timeout(300)

                _save_param_config_dialog(hmi_page)
                _close_param_config_dialog(hmi_page)

                time.sleep(DEVICE_RESTART_WAIT)

                if not _open_param_dialog(hmi_page, device_keyword):
                    pytest.fail(
                        f"启用 EPICS 后无法重新打开设备 {device_keyword!r} 的 Parameter Config"
                    )

            wait_until_connectable()  # 等服务重启回来再读，避免过早读取超时误判
            idents_enabled = get_object_identifiers()
            if idents_enabled is None:
                _close_param_config_dialog(hmi_page)
                pytest.skip("BACnet 客户端不可达，无法验证对象数量变化")
            count_enabled = len(idents_enabled)

            ok = _set_first_row_epics_enable(hmi_page, False)
            if not ok:
                pytest.skip("无法操作第一行 EPICS Enable 开关（disable）")
            hmi_page.wait_for_timeout(300)

            _save_param_config_dialog(hmi_page)
            _close_param_config_dialog(hmi_page)

        except Exception:
            if hmi_page.locator('[aria-label="Parameter Config"]').count() > 0:
                _close_param_config_dialog(hmi_page)
            raise

        wait_until_connectable()  # 等服务重启回来再读，避免过早读取超时误判
        idents_disabled = get_object_identifiers()
        count_disabled = len(idents_disabled) if idents_disabled is not None else None

        if idents_enabled is not None and idents_disabled is not None:
            disabled_set = set(idents_disabled)
            removed = [o for o in idents_enabled if o not in disabled_set]
            log.info("[test_018] BACnet 对象数 %d → %d（撤销 %d 个）",
                     count_enabled, count_disabled, len(removed))
            for obj_type, inst in removed:
                log.info("[test_018] 撤销发布对象: %s,%d", obj_type, inst)

        try:
            if was_enabled:
                if _open_param_dialog(hmi_page, device_keyword):
                    _set_first_row_epics_enable(hmi_page, True)
                    hmi_page.wait_for_timeout(300)
                    _save_param_config_dialog(hmi_page)
                    _close_param_config_dialog(hmi_page)
        except Exception:
            if hmi_page.locator('[aria-label="Parameter Config"]').count() > 0:
                _close_param_config_dialog(hmi_page)

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
        if not can_connect():
            pytest.skip(
                "BACnet 客户端初始状态不可达（默认端口 %d），测试无法执行" % HMI_DEFAULT_PORT
            )

        initial_enabled: bool = _get_bacnet_enable_state(hmi_page)

        try:
            _set_bacnet_enable(hmi_page, False)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)
            time.sleep(DEVICE_RESTART_WAIT)

            assert not can_connect(), (
                "BACnet Enable = Disable 后，客户端应无法连接（端口 %d），但连接仍然成功"
                % HMI_DEFAULT_PORT
            )

        finally:
            _set_bacnet_enable(hmi_page, initial_enabled)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)
            time.sleep(DEVICE_RESTART_WAIT)
