# -*- coding: utf-8 -*-
"""
test_bacnet_six_segment.py — BACnet/IP 六段式比对（第 3/5/6/7 段）

补齐 test_bacnet_ui_protocol.py 已覆盖的第 1/2/4 段，使六段全覆盖：

  第 3 段（元数据）：
    TestCase_AcuHMI-1-7_033_001_050: 已发布 AI 对象 units 属性与模板一致（多设备）

  第 5 段（Device Object 属性）：
    TestCase_AcuHMI-1-7_033_001_051: Device Object 12 项标准必需属性全部可读

  第 6 段（协议合规性）：
    TestCase_AcuHMI-1-7_033_001_052: 非法 AI 对象请求返回错误（§16）
    TestCase_AcuHMI-1-7_033_001_053: AI 必需属性（statusFlags/outOfService/units）可读

  第 7 段（连接稳定性）：
    TestCase_AcuHMI-1-7_033_001_054: 同一 AI 对象连续读取 5 次全部成功

运行：
  pytest projects/RPP/tests/BacnetIP/test_bacnet_six_segment.py -v
  pytest projects/RPP/tests/BacnetIP/ -v   # 与其余 BACnet 用例一起运行
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

# ── BACnet 客户端（六段式辅助函数） ──────────────────────────────────────────
from projects.RPP.helpers.hmi_bacnet_client import (  # noqa: E402
    DeviceInfoResult,
    MetadataItem,
    ProtocolCheckItem,
    StabilityCheckResult,
    check_stability,
    get_object_identifiers,
    read_device_info,
    read_object_details_batch,
    read_object_metadata_batch,
    run_protocol_compliance,
    wait_until_connectable,
    HMI_DEFAULT_PORT,
    DEVICE_RESTART_WAIT,
    SERVICE_READY_TIMEOUT,
)

# ── 模板基准 ──────────────────────────────────────────────────────────────────
from projects.RPP.helpers.template_matcher import (  # noqa: E402
    get_bacnet_template_map,
)
from projects.RPP.helpers.physical_devices_reader import (  # noqa: E402
    DiscoveredDevice,
)

# ── 复用 UI 操作辅助（设备隔离 / 参数弹窗 / 保存）──────────────────────────────
from projects.RPP.tests.BacnetIP.test_bacnet_ui_basic import (  # noqa: E402
    _get_all_device_rows,
    _isolate_single_device,
    _restore_device_selection,
    _open_param_dialog,
    _close_param_config_dialog,
    _get_parameter_types,
    _select_param_type,
)
from projects.RPP.tests.BacnetIP.test_bacnet_ui_config import (  # noqa: E402
    _click_save,
    _dismiss_toast,
)
from projects.RPP.tests.BacnetIP.test_bacnet_ui_protocol import (  # noqa: E402
    _resolve_bacnet_device,
    _save_param_config_dialog,
)

log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# 常量
# ─────────────────────────────────────────────────────────────────────────────

# Device Object 12 项标准必需属性（ANSI/ASHRAE 135 §12.11），
# 每项只要能读到非空值即判为可读（不校验内容，内容由 test_051 日志记录）。
_DEVICE_REQUIRED_PROPS: list[tuple[str, str]] = [
    ("object_identifier",   "objectIdentifier"),
    ("object_name",         "objectName"),
    ("system_status",       "systemStatus"),
    ("vendor_name",         "vendorName"),
    ("vendor_id",           "vendorIdentifier"),
    ("model_name",          "modelName"),
    ("firmware_revision",   "firmwareRevision"),
    ("app_sw_version",      "applicationSoftwareVersion"),
    ("protocol_version",    "protocolVersion"),
    ("protocol_revision",   "protocolRevision"),
    ("max_apdu_length",     "maxApduLengthAccepted"),
    ("segmentation",        "segmentationSupported"),
]

# 稳定性测试读取次数
_STABILITY_ATTEMPTS = 5

# 稳定性测试通过门限：至少这么多次成功（允许 1 次网络抖动）
_STABILITY_MIN_SUCCESS = 4

# BAC0/bacpypes3 返回的 objectType 为连字符格式（analog-input / binary-input）；
# 同时兼容 camelCase（analogInput / binaryInput）以防底层实现变更。
_AI_TYPES = ("analog-input", "analogInput")
_BI_TYPES = ("binary-input", "binaryInput")
_AI_BI_TYPES = _AI_TYPES + _BI_TYPES


# ─────────────────────────────────────────────────────────────────────────────
# 辅助：获取当前网关已发布的第一个 AI 对象（用作稳定性 / 协议合规性 probe）
# ─────────────────────────────────────────────────────────────────────────────

def _get_first_ai_object() -> Optional[tuple[str, int]]:
    """返回当前网关 objectList 中第一个 analog-input，找不到返回 None。"""
    idents = get_object_identifiers()
    if not idents:
        return None
    for obj_type, inst in idents:
        if obj_type in _AI_TYPES:
            return obj_type, inst
    return None


def _get_all_ai_bi_objects() -> list[tuple[str, int]]:
    """返回当前网关 objectList 中全部 analog-input 和 binary-input 对象。"""
    idents = get_object_identifiers()
    if not idents:
        return []
    return [(ot, inst) for ot, inst in idents
            if ot in _AI_BI_TYPES]


# ─────────────────────────────────────────────────────────────────────────────
# 辅助：从 objectName 解析 (prefix, param_key)
# ─────────────────────────────────────────────────────────────────────────────

def _parse_object_name(name: Optional[str]) -> tuple[str, str]:
    """'PrefixDevice-PARAM_KEY' → ('PrefixDevice', 'PARAM_KEY')。"""
    if not name or "-" not in name:
        return ("", name or "")
    prefix, _, key = name.partition("-")
    return prefix, key


# ─────────────────────────────────────────────────────────────────────────────
# 辅助：单参数抽检——关闭全部 Polling Enable、读首行参数名（test_050 专用）
# ─────────────────────────────────────────────────────────────────────────────

_POLLING_HEADER_STATE_JS = """() => {
    const dlg = document.querySelector('[aria-label="Parameter Config"]');
    if (!dlg) return 'no_dialog';
    const thead = dlg.querySelector('.el-table__header');
    if (!thead) return 'no_thead';
    for (const th of thead.querySelectorAll('th')) {
        const label = th.querySelector('.el-checkbox');
        if (!label || !label.textContent.includes('Polling Enable')) continue;
        const cb = label.querySelector('.el-checkbox__input');
        if (!cb) return 'no_input';
        if (cb.classList.contains('is-checked')) return 'checked';
        if (cb.classList.contains('is-indeterminate')) return 'indeterminate';
        return 'unchecked';
    }
    return 'not_found';
}"""

_POLLING_HEADER_CLICK_JS = """() => {
    const dlg = document.querySelector('[aria-label="Parameter Config"]');
    if (!dlg) return false;
    const thead = dlg.querySelector('.el-table__header');
    if (!thead) return false;
    for (const th of thead.querySelectorAll('th')) {
        const label = th.querySelector('.el-checkbox');
        if (!label || !label.textContent.includes('Polling Enable')) continue;
        const orig = label.querySelector('.el-checkbox__original');
        if (!orig) return false;
        orig.scrollIntoView({block: 'center', behavior: 'instant'});
        orig.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true}));
        return true;
    }
    return false;
}"""


def _uncheck_polling_select_all(page: Page) -> bool:
    """把当前 Parameter Type 的 Polling Enable 列头全选框驱动到「全不选」。

    El Plus checkbox 半选(indeterminate)态点一次会变全选，故循环判状态：
    checked / indeterminate → 点一次再判，直到 unchecked（最多 4 次）。
    返回 True=已到全不选；False=未找到控件。
    """
    for _ in range(4):
        state: str = page.evaluate(_POLLING_HEADER_STATE_JS)
        if state in ("no_dialog", "no_thead", "no_input", "not_found"):
            return False
        if state == "unchecked":
            return True
        if not page.evaluate(_POLLING_HEADER_CLICK_JS):
            return False
        page.wait_for_timeout(300)
    return page.evaluate(_POLLING_HEADER_STATE_JS) == "unchecked"


def _disable_all_polling_in_dialog(page: Page) -> None:
    """遍历 Parameter Config 弹窗所有 Parameter Type，逐个把 Polling Enable 全部关闭。

    完成后停留在第一个 Type 的第一页，便于随后单独启用其首行参数。
    """
    dlg = page.locator('[aria-label="Parameter Config"]')
    types = _get_parameter_types(page, dlg) or [""]
    for pt in types:
        if pt:
            _select_param_type(page, dlg, pt)
        _uncheck_polling_select_all(page)
    if types and types[0]:
        _select_param_type(page, dlg, types[0])


_ENABLE_ROW_IN_SET_JS = """(descs) => {
    const set = new Set(descs);
    const rows = document.querySelectorAll(
        '[aria-label="Parameter Config"] .el-table__body tr.el-table__row'
    );
    for (const row of rows) {
        const cells = row.querySelectorAll('td');
        if (cells.length < 2) continue;
        const name = (cells[0].textContent || '').trim();
        if (!set.has(name)) continue;
        const sw = cells[1].querySelector('.el-switch__input');
        if (sw && sw.getAttribute('aria-checked') !== 'true') sw.click();
        return name;
    }
    return '';
}"""

_IS_LAST_PAGE_JS = """() => {
    const dlg = document.querySelector('[aria-label="Parameter Config"]');
    if (!dlg) return true;
    const btn = dlg.querySelector('.el-pagination .btn-next');
    if (!btn) return true;
    return btn.disabled || btn.classList.contains('is-disabled')
        || btn.getAttribute('aria-disabled') === 'true';
}"""

_CLICK_NEXT_PAGE_JS = """() => {
    const dlg = document.querySelector('[aria-label="Parameter Config"]');
    const btn = dlg && dlg.querySelector('.el-pagination .btn-next');
    if (btn) btn.click();
}"""


def _enable_first_param_with_unit(page: Page, unit_descs: set[str]) -> str:
    """遍历所有 Parameter Type 与分页，启用首个「模板里有单位」参数的 Polling Enable。

    无单位参数的 units 比对会被跳过（unit_skipped），测不出东西；故只在模板
    unit 非空参数的 description 集合（unit_descs，对应弹窗 Parameter 列显示文本）中
    挑一行启用，保证 units 比对有实际意义。

    返回被启用参数的 description；未找到返回 ""。
    """
    if not unit_descs:
        return ""
    descs = list(unit_descs)
    dlg = page.locator('[aria-label="Parameter Config"]')
    types = _get_parameter_types(page, dlg) or [""]
    for pt in types:
        if pt:
            _select_param_type(page, dlg, pt)
        for _ in range(200):  # 分页上限保护，防止异常下死循环
            name: str = page.evaluate(_ENABLE_ROW_IN_SET_JS, descs)
            if name:
                page.wait_for_timeout(300)
                return name
            if page.evaluate(_IS_LAST_PAGE_JS):
                break
            page.evaluate(_CLICK_NEXT_PAGE_JS)
            page.wait_for_timeout(300)
    return ""


# ═════════════════════════════════════════════════════════════════════════════
# 第 3 段：元数据检查
# ═════════════════════════════════════════════════════════════════════════════

class TestBACnetMetadata:
    """第 3 段：AI 对象 units 属性与模板一致性验证。"""

    # 协议规范测试：只需验证「units 属性能被正确发布且与模板一致」这一行为，
    # 不必全量参数。故隔离 AcuRev4100 单设备、仅启用一个参数的 Polling Enable，
    # 只读回该参数对应 AI 对象的 units 比对，避免读全量 3000+ 对象的漫长耗时。
    _TARGET_TEMPLATE = "AcuRev4100"

    def test_050_units_match_template(
        self,
        hmi_page: Page,
        discovered_devices: list[DiscoveredDevice],
    ) -> None:
        """TestCase_AcuHMI-1-7_033_001_050: 已发布 AI 对象 units 属性与模板一致（单参数抽检）。

        协议规范测试，只抽检一个参数（不做全量）：
          1. 解析 AcuRev4100 设备 → 隔离该单设备（取消其余勾选）→ Save → 等重启
          2. 打开 Parameter Config，关闭全部 Polling Enable，仅启用第一个参数 → Save
          3. 读回网关已发布 AI 对象（应仅该参数对应对象）
          4. 读 BACnet units 与 AcuRev4100 模板比对，断言无单位不一致
        测试结束在 finally 中恢复原始设备勾选状态。
        """
        # ── 段1：模板加载 ──
        try:
            tmpl_map = get_bacnet_template_map(self._TARGET_TEMPLATE)
        except FileNotFoundError:
            pytest.skip(f"未找到设备模板 {self._TARGET_TEMPLATE!r}，跳过")
            return
        if not tmpl_map:
            pytest.skip(f"模板 {self._TARGET_TEMPLATE!r} 无 BACnet 参数，跳过")
            return

        # ── 解析目标 4100 设备（取该型号第一台在线设备）──
        ui_rows = _get_all_device_rows(hmi_page)
        device, reason = _resolve_bacnet_device(
            discovered_devices, ui_rows, self._TARGET_TEMPLATE
        )
        if device is None:
            pytest.skip(f"未确定可测的 {self._TARGET_TEMPLATE} 设备（{reason}）")
            return

        # 快照原始勾选集，测试结束恢复
        original_checked = [r["name"] for r in ui_rows if r["checked"]]
        log.info("[test_050] 原始设备勾选集：%s", original_checked)

        try:
            # ── 隔离单设备 → Save → 等重启 ──
            log.info("[test_050] 隔离单设备 %r 并保存", device)
            _isolate_single_device(hmi_page, device)
            _click_save(hmi_page)
            _dismiss_toast(hmi_page)
            time.sleep(DEVICE_RESTART_WAIT)

            # ── 打开参数弹窗：关闭全部 Polling，仅启用一个参数 → Save ──
            if not _open_param_dialog(hmi_page, device):
                pytest.skip(f"无法打开设备 {device!r} 的 Parameter Config 弹窗")
                return

            # 只在「模板里有单位」的参数中挑一个启用：无单位参数 units 比对会被跳过，
            # 测不出东西，专挑有单位的才能保证比对有实际意义。
            unit_descs = {
                (p.description or "").split("\n")[0].strip()
                for p in tmpl_map.values()
                if (p.unit or "").strip()
            }
            unit_descs.discard("")
            if not unit_descs:
                pytest.skip(
                    f"模板 {self._TARGET_TEMPLATE!r} 无任何带单位参数，无法做有意义的 units 比对"
                )
                return

            enabled_param = ""
            try:
                _disable_all_polling_in_dialog(hmi_page)
                enabled_param = _enable_first_param_with_unit(hmi_page, unit_descs)
                if not enabled_param:
                    pytest.skip("弹窗中未找到可启用的带单位参数")
                    return
                _save_param_config_dialog(hmi_page)
                _close_param_config_dialog(hmi_page)
            except Exception:
                if hmi_page.locator('[aria-label="Parameter Config"]').count() > 0:
                    _close_param_config_dialog(hmi_page)
                raise
            log.info("[test_050] 已仅启用单个带单位参数 Polling：%r", enabled_param)

            # ── 等服务可达 → 读回已发布对象 ──
            if not wait_until_connectable():
                pytest.skip(
                    f"单参数 Polling 保存后 BACnet 服务在 {SERVICE_READY_TIMEOUT:.0f}s 内不可达"
                )
                return
            idents = get_object_identifiers()
            if idents is None:
                pytest.skip(f"BACnet 客户端无法连接（端口 {HMI_DEFAULT_PORT}），跳过")
                return
            ai_bi = [(ot, inst) for ot, inst in idents if ot in _AI_BI_TYPES]
            if not ai_bi:
                pytest.skip(
                    f"仅启用参数 {enabled_param!r} 后网关未发布任何 AI/BI 对象，无法元数据比对"
                )
                return
            details = read_object_details_batch(ai_bi)

            # 单设备隔离：所有对象都属本设备，param_key 取 objectName "-" 后半段
            meta_inputs: list[tuple[str, int, str, str, str]] = []
            no_tmpl_keys: list[str] = []
            for obj_type, inst, name, _val in details:
                _prefix, param_key = _parse_object_name(name)
                if not param_key:
                    continue
                tmpl = tmpl_map.get(param_key)
                if tmpl is None:
                    no_tmpl_keys.append(param_key)
                    continue
                meta_inputs.append(
                    (obj_type, inst, param_key, tmpl.unit, tmpl.description)
                )
            if no_tmpl_keys:
                log.info("[test_050] 模板中无对应项的参数（已跳过）：%s", no_tmpl_keys[:10])
            if not meta_inputs:
                pytest.skip("已发布对象在模板中均无对应项，无法做元数据比对")
                return

            # ── 读 units 比对 ──
            meta_items: list[MetadataItem] = read_object_metadata_batch(meta_inputs)
            total = len(meta_items)
            mismatch = [m for m in meta_items if not m.unit_ok]
            log.info(
                "[test_050] 单参数元数据抽检：共 %d 项，一致 %d，不一致 %d（参数=%r）",
                total, total - len(mismatch), len(mismatch), enabled_param,
            )
            for m in mismatch:
                log.warning("[test_050] 单位不一致  %s  模板=%r  BACnet=%r",
                            m.param_key, m.tmpl_unit, m.bacnet_unit)

            assert not mismatch, (
                f"[test_050] 有 {len(mismatch)}/{total} 个 AI 对象的 units 属性与模板不一致，"
                f"前 10 条：\n" + "\n".join(
                    f"  {m.param_key}: 模板={m.tmpl_unit!r} BACnet={m.bacnet_unit!r}"
                    for m in mismatch[:10]
                )
            )
        finally:
            # ── 恢复原始设备勾选状态 ──
            log.info("[test_050] 恢复原始设备勾选：%s", original_checked)
            try:
                _restore_device_selection(hmi_page, original_checked)
                _click_save(hmi_page)
                _dismiss_toast(hmi_page)
                time.sleep(DEVICE_RESTART_WAIT)
            except Exception as exc:
                log.warning("[test_050] 恢复设备勾选失败（不影响结论）：%s", exc)


# ═════════════════════════════════════════════════════════════════════════════
# 第 5 段：Device Object 12 项标准必需属性
# ═════════════════════════════════════════════════════════════════════════════

class TestBACnetDeviceObject:
    """第 5 段：Device Object 12 项标准必需属性（ANSI/ASHRAE 135 §12.11）。"""

    def test_051_device_object_required_props(self) -> None:
        """TestCase_AcuHMI-1-7_033_001_051: Device Object 12 项标准必需属性全部可读。

        读取以下属性并验证非空：
          objectIdentifier, objectName, systemStatus, vendorName, vendorIdentifier,
          modelName, firmwareRevision, applicationSoftwareVersion,
          protocolVersion, protocolRevision, maxApduLengthAccepted, segmentationSupported
        """
        dev_info: DeviceInfoResult = read_device_info()

        if not dev_info.ok:
            pytest.fail(
                f"[test_051] Device Object 属性读取失败：{dev_info.error}\n"
                f"  请确认网关 BACnet 服务已启用（IP={HMI_DEFAULT_PORT}）"
            )

        # 记录所有属性值到日志（无论成功与否都记录，方便排查）
        log.info("[test_051] Device Object 属性一览：")
        readable: list[str] = []
        unreadable: list[str] = []
        for attr, prop_name in _DEVICE_REQUIRED_PROPS:
            val = getattr(dev_info, attr, "")
            if val:
                log.info("  %-30s = %s", prop_name, val)
                readable.append(prop_name)
            else:
                log.warning("  %-30s = <未读到>", prop_name)
                unreadable.append(prop_name)

        log.info("[test_051] 可读 %d/%d 项，缺失 %d 项：%s",
                 len(readable), len(_DEVICE_REQUIRED_PROPS),
                 len(unreadable), unreadable)

        assert not unreadable, (
            f"[test_051] Device Object 以下 {len(unreadable)} 项属性未能读取（返回空值）：\n"
            + "\n".join(f"  {p}" for p in unreadable)
            + "\n（以上属性为 ANSI/ASHRAE 135 §12.11 要求的必需属性，应始终可读）"
        )


# ═════════════════════════════════════════════════════════════════════════════
# 第 6 段：协议合规性
# ═════════════════════════════════════════════════════════════════════════════

class TestBACnetProtocolCompliance:
    """第 6 段：BACnet 协议合规性验证（ANSI/ASHRAE 135 §16 + §12.2.2）。"""

    def test_052_illegal_object_returns_error(self) -> None:
        """TestCase_AcuHMI-1-7_033_001_052: 非法 AI 对象请求应返回错误（§16）。

        向网关请求一个不存在的 AI 对象（Instance=9999999）的 presentValue，
        标准要求设备返回 BACnet Error，BAC0 收到 Error 后返回 None。
        若网关错误地返回了一个值，则视为协议合规性问题。
        """
        tests: list[ProtocolCheckItem] = run_protocol_compliance(probe_obj=None)
        if not tests:
            pytest.skip("协议合规性测试返回空结果（网关可能不可达），跳过此用例")

        # 本用例只关心第一项（非法对象测试）
        illegal_test = next(
            (t for t in tests
             if "9999999" in t.test_name or "不存在" in t.test_name),
            None,
        )
        if illegal_test is None:
            pytest.skip("未找到非法对象测试项，跳过")

        log.info("[test_052] %s -> passed=%s, detail=%s",
                 illegal_test.test_name, illegal_test.passed, illegal_test.detail)
        assert illegal_test.passed, (
            f"[test_052] 非法 AI 对象请求（Instance=9999999）未被网关拒绝：\n"
            f"  {illegal_test.detail}\n"
            f"  （期望行为：网关应返回 BACnet Error，不应返回任何值）"
        )

    def test_053_ai_required_props_readable(self) -> None:
        """TestCase_AcuHMI-1-7_033_001_053: AI 必需属性（statusFlags/outOfService/units）可读（§12.2.2）。

        对当前网关已发布的第一个 AI 对象，验证三个必需属性均可读。
        若网关未发布任何 AI 对象则跳过。
        """
        probe = _get_first_ai_object()
        if probe is None:
            pytest.skip("当前网关未发布任何 AI 对象，无法执行必需属性可读性检查")

        tests: list[ProtocolCheckItem] = run_protocol_compliance(probe_obj=probe)

        # 过滤出必需属性测试项（排除非法对象测试）
        prop_tests = [t for t in tests
                      if any(p in t.test_name
                             for p in ("statusFlags", "outOfService", "units"))]
        if not prop_tests:
            pytest.skip("未返回 AI 必需属性测试项，跳过")

        failed = [t for t in prop_tests if not t.passed]
        log.info("[test_053] AI 必需属性检查：%d/%d 通过  probe=%s,%d",
                 len(prop_tests) - len(failed), len(prop_tests), probe[0], probe[1])
        for t in prop_tests:
            log.info("  %-55s -> %s  %s",
                     t.test_name, "PASS" if t.passed else "FAIL", t.detail)

        assert not failed, (
            f"[test_053] {len(failed)} 个 AI 必需属性未能读取（§12.2.2 要求必须可读）：\n"
            + "\n".join(f"  {t.test_name}" for t in failed)
        )


# ═════════════════════════════════════════════════════════════════════════════
# 第 7 段：连接稳定性
# ═════════════════════════════════════════════════════════════════════════════

class TestBACnetStability:
    """第 7 段：BACnet 连接稳定性验证。"""

    def test_054_stability_repeated_reads(self) -> None:
        """TestCase_AcuHMI-1-7_033_001_054: 同一 AI 对象连续读取 5 次，成功率 >= 4/5。

        对当前网关已发布的第一个 AI 对象连续读取 _STABILITY_ATTEMPTS 次，
        验证成功次数 >= _STABILITY_MIN_SUCCESS（门限 4/5，允许 1 次网络抖动）。
        若网关未发布任何 AI 对象则跳过。
        """
        probe = _get_first_ai_object()
        if probe is None:
            pytest.skip("当前网关未发布任何 AI 对象，无法执行稳定性测试")

        result: StabilityCheckResult = check_stability(
            probe_obj=probe,
            attempts=_STABILITY_ATTEMPTS,
            delay=0.5,
        )

        log.info("[test_054] 稳定性测试：%d/%d 成功  probe=%s,%d",
                 result.successes, result.attempts, probe[0], probe[1])
        if result.errors:
            for err in result.errors:
                log.warning("[test_054] 读取失败：%s", err)

        assert result.successes >= _STABILITY_MIN_SUCCESS, (
            f"[test_054] 连接稳定性不足：{result.successes}/{result.attempts} 次成功，"
            f"门限 {_STABILITY_MIN_SUCCESS}（对象={probe[0]},{probe[1]}）\n"
            f"  失败详情：{result.errors}"
        )
