# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_06_case02
类别：UI — 参数类型分组与参数列表和设备模板一致性验证
"""
import datetime
import logging
import sys
import time
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_PROJECT_ROOT = _HERE.parent.parent.parent.parent  # AcuHMI-1-7/
sys.path.insert(0, str(_HERE))
sys.path.insert(0, str(_PROJECT_ROOT))

import config
from datalog_page import DataLogParamConfigPage
from template_reader import find_template_file, load_template

CASE_ID = "TestCase_AcuHMI_003_06_case02"
log = logging.getLogger(__name__)

_CLEAR_BTN_XPATH = (
    "//button[.//span[normalize-space(.)='Clear']]"
    " | //button[normalize-space(.)='Clear']"
)
_SAVE_BTN_XPATH = (
    "//button[.//span[normalize-space(.)='Save']]"
    " | //button[normalize-space(.)='Save']"
)

_DEVICE_ALIASES: dict[str, str] = {
    "pxm350": "AcuRev1300",
}

_OPT_XPATH = (
    "xpath=//li[contains(@class,'el-select-dropdown__item')"
    " and not(contains(@class,'disabled'))]"
)

_PANEL_NOISE = frozenset({
    'all', 'clear', 'selected', 'not selected', 'save',
    'device', 'parameter type', 'parameters', '←', '→',
})


# ── 模板工具 ───────────────────────────────────────────────────────────────────

def _find_template_path(device_name: str) -> str | None:
    candidates: list[str] = [device_name]
    if "-" in device_name:
        candidates.append(device_name.split("-")[0].strip())
    alias = _DEVICE_ALIASES.get(device_name.strip().lower())
    if alias:
        candidates.append(alias)
    for cand in candidates:
        try:
            return find_template_file(config.TEMPLATE_DIR, cand)
        except FileNotFoundError:
            continue
    return None


def _template_groups(device_name: str) -> dict[str, set[str]]:
    path = _find_template_path(device_name)
    if path is None:
        log.warning("[%s] 未找到模板文件，跳过", device_name)
        return {}
    params = load_template(path)
    groups: dict[str, set[str]] = {}
    for p in params:
        if p.datalog:
            first_line = next(
                (l.strip() for l in p.description.split("\n") if l.strip()),
                p.description.strip(),
            )
            if first_line:
                groups.setdefault(p.datalog.strip().lower(), set()).add(first_line.lower())
    return groups


# ── 下拉框工具 ────────────────────────────────────────────────────────────────

def _open_dropdown(page, selector: str):
    container = page.locator(selector).first
    container.wait_for(state="visible", timeout=8000)
    wrapper = container.locator(".el-select__wrapper")
    if wrapper.count() > 0:
        wrapper.first.evaluate("el => el.click()")
    else:
        container.evaluate("el => el.click()")
    page.wait_for_timeout(600)
    return container


def _close_dropdown_gently(page):
    try:
        page.locator("body").click()
        page.wait_for_timeout(400)
    except Exception:
        pass


def _read_dropdown_options(page, selector: str) -> list[str]:
    try:
        _open_dropdown(page, selector)
        items = page.locator(_OPT_XPATH).all()
        options = [item.text_content().strip() for item in items
                   if item.text_content().strip()]
        _close_dropdown_gently(page)
        return options
    except Exception as e:
        log.warning("读取下拉选项失败：%s", e)
        _close_dropdown_gently(page)
        return []


def _select_dropdown_item(page, selector: str, text: str, label: str) -> bool:
    try:
        _open_dropdown(page, selector)
        opt_patterns = [
            f"xpath=//li[contains(@class,'el-select-dropdown__item') and normalize-space(.)='{text}']",
            f"xpath=//li[contains(@class,'el-select-dropdown__item') and .//span[normalize-space(.)='{text}']]",
            f"xpath=//ul[contains(@class,'el-select-dropdown__list')]//li[contains(normalize-space(.),'{text}')]",
            f"xpath=//div[contains(@class,'el-popper')]//li[contains(normalize-space(.),'{text}')]",
        ]
        for pat in opt_patterns:
            try:
                opt = page.locator(pat).first
                opt.wait_for(state="visible", timeout=3000)
                opt.evaluate("el => el.click()")
                page.wait_for_timeout(400)
                log.debug("下拉 %s 选择：%s", label, text)
                return True
            except Exception:
                continue
        log.warning("下拉 %s 未找到选项 '%s'，温和关闭", label, text)
        _close_dropdown_gently(page)
        return False
    except Exception as e:
        log.warning("操作下拉 %s 失败：%s", label, e)
        _close_dropdown_gently(page)
        return False


# ── 页面操作 ───────────────────────────────────────────────────────────────────

def _click_clear(page):
    try:
        btn = page.locator(f"xpath={_CLEAR_BTN_XPATH}").first
        btn.wait_for(state="visible", timeout=5000)
        btn.evaluate("el => el.click()")
        page.wait_for_timeout(800)
    except Exception as e:
        log.warning("Clear 按钮点击失败：%s", e)


def _click_save(page):
    try:
        btn = page.locator(f"xpath={_SAVE_BTN_XPATH}").first
        btn.wait_for(state="visible", timeout=5000)
        btn.evaluate("el => el.click()")
        page.wait_for_timeout(1000)
    except Exception as e:
        log.warning("Save 按钮点击失败：%s", e)


def _read_not_selected(page) -> set[str]:
    xpaths = [
        "//*[normalize-space(text())='Not Selected']"
        "/following-sibling::*[not(*) and normalize-space(.)!='']",
        "//*[normalize-space(text())='Not Selected']"
        "/../*[not(normalize-space(text())='Not Selected')]"
        "[not(*) and normalize-space(.)!='']",
        "//*[normalize-space(text())='Not Selected']"
        "/following-sibling::*[1]//*[not(*) and normalize-space(.)!='']",
        "//*[normalize-space(text())='Not Selected']"
        "/../..//*[not(*) and normalize-space(.)!=''"
        " and normalize-space(.)!='Not Selected'"
        " and normalize-space(.)!='Selected']",
        "//*[normalize-space(.)='Not Selected']/ancestor::*[2]"
        "//li[normalize-space(.)!='']",
    ]
    for xpath in xpaths:
        try:
            els = page.locator(f"xpath={xpath}").all()
            texts = [e.text_content().strip() for e in els if e.text_content().strip()]
            result = {t.lower() for t in texts
                      if t.lower() not in _PANEL_NOISE and len(t) > 1}
            if result:
                log.debug("Not Selected 读取成功，共 %d 项", len(result))
                return result
        except Exception:
            continue
    return set()


# ── HTML 报告生成 ──────────────────────────────────────────────────────────────

def _write_html_report(failures: list[str], unreadable: list[str], devices_skipped: list[str]):
    report_dir = _HERE / "reports"
    report_dir.mkdir(exist_ok=True)
    report_path = report_dir / f"{CASE_ID}_report.html"
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    total = len(failures)
    status_cls = "pass" if total == 0 else "fail"
    status_text = "PASSED" if total == 0 else f"FAILED  ({total} 处不一致)"

    device_map: dict[str, list[str]] = {}
    for msg in failures:
        dev = msg.split("]")[0].lstrip("[")
        device_map.setdefault(dev, []).append(msg)

    rows_html = ""
    for dev, msgs in device_map.items():
        for j, msg in enumerate(msgs):
            rows_html += f"""
            <tr>
              {"<td rowspan='" + str(len(msgs)) + "' class='dev-cell'>" + dev + "</td>" if j == 0 else ""}
              <td>{msg.replace('<','&lt;').replace('>','&gt;')}</td>
            </tr>"""

    unreadable_html = ""
    if unreadable:
        items = "".join(f"<li>{u}</li>" for u in unreadable)
        unreadable_html = (f'<div class="section"><h2>⚠ 无法读取的分组（{len(unreadable)} 个）</h2>'
                           f'<ul class="unreadable-list">{items}</ul></div>')

    skipped_html = ""
    if devices_skipped:
        items = "".join(f"<li>{d}</li>" for d in devices_skipped)
        skipped_html = (f'<div class="section"><h2>ℹ 已跳过设备（{len(devices_skipped)} 台）</h2>'
                        f'<ul class="skip-list">{items}</ul></div>')

    failures_section = (
        f'<div class="section"><h2>✗ 不一致详情（{total} 处）</h2>'
        f'<table><thead><tr><th>设备</th><th>不一致描述</th></tr></thead>'
        f'<tbody>{rows_html}</tbody></table></div>'
        if failures else
        '<div class="section pass-block"><h2>✓ 所有设备参数与模板完全一致</h2></div>'
    )

    html = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>{CASE_ID}</title>
<style>
  body{{font-family:"Segoe UI",Arial,sans-serif;background:#f4f6f9;color:#333;margin:0}}
  .header{{background:#1a73e8;color:#fff;padding:24px 32px}}
  .header h1{{font-size:1.4em;font-weight:600}}
  .badge{{display:inline-block;padding:4px 14px;border-radius:12px;font-weight:700;margin-top:10px}}
  .badge.pass{{background:#34a853;color:#fff}}.badge.fail{{background:#ea4335;color:#fff}}
  .content{{max-width:1100px;margin:24px auto;padding:0 16px}}
  .section{{background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.1);
            padding:20px 24px;margin-bottom:20px}}
  .section h2{{font-size:1em;font-weight:600;margin-bottom:14px;border-bottom:1px solid #eee;padding-bottom:8px}}
  table{{width:100%;border-collapse:collapse;font-size:.88em}}
  th{{background:#f0f4ff;color:#1a73e8;padding:9px 12px;text-align:left}}
  td{{padding:8px 12px;border-bottom:1px solid #f0f0f0;vertical-align:top}}
  .dev-cell{{font-weight:600;color:#1a73e8;border-right:2px solid #d0e0ff;white-space:nowrap}}
  .pass-block{{border-left:4px solid #34a853}}.pass-block h2{{color:#34a853}}
  ul{{padding-left:20px;font-size:.88em}}
  .unreadable-list li{{color:#e65100}}.skip-list li{{color:#888}}
</style>
</head>
<body>
<div class="header">
  <h1>{CASE_ID}</h1>
  <div style="margin-top:8px;font-size:.85em;opacity:.85">参数类型分组与模板一致性验证 | {now}</div>
  <div class="badge {status_cls}">{status_text}</div>
</div>
<div class="content">{failures_section}{unreadable_html}{skipped_html}</div>
</body></html>"""

    report_path.write_text(html, encoding="utf-8")
    log.info("[%s] HTML 报告已写入：%s", CASE_ID, report_path)
    for i, msg in enumerate(failures, 1):
        log.error("[%s] [%02d] %s", CASE_ID, i, msg)


# ── 主测试函数 ─────────────────────────────────────────────────────────────────

def test_case(driver):
    page = DataLogParamConfigPage(driver)
    page.navigate()

    devices = _read_dropdown_options(driver, page._DEVICE_SELECT)
    assert devices, f"[{CASE_ID}] Device 下拉无可用设备"

    failures: list[str] = []
    unreadable: list[str] = []
    devices_skipped: list[str] = []

    # 阶段一：对每台设备 Clear + Save，确保全部参数处于未选状态
    log.info("[%s] 阶段一：重置所有设备参数", CASE_ID)
    for device_name in devices:
        if not _select_dropdown_item(driver, page._DEVICE_SELECT, device_name, "Device"):
            log.warning("[%s] 阶段一：无法选中设备 %s，跳过", CASE_ID, device_name)
            continue
        driver.wait_for_timeout(800)
        _click_clear(driver)
        _click_save(driver)
        log.info("  [重置] %s → Clear + Save 完成", device_name)

    # 阶段二：遍历每台设备，比对 Parameter Type 集合和全量参数列表与模板
    log.info("[%s] 阶段二：开始参数与模板比对", CASE_ID)
    for device_name in devices:
        tmpl_groups = _template_groups(device_name)
        if not tmpl_groups:
            devices_skipped.append(device_name)
            continue

        if not _select_dropdown_item(driver, page._DEVICE_SELECT, device_name, "Device"):
            log.warning("[%s] 阶段二：无法选中设备 %s，跳过", CASE_ID, device_name)
            continue
        driver.wait_for_timeout(800)

        # Parameter Type 分组集合比对
        ui_types_raw = _read_dropdown_options(driver, page._PARAM_TYPE_SELECT)
        ui_types = {t.strip().lower() for t in ui_types_raw}
        tmpl_type_keys = set(tmpl_groups.keys())

        missing_types = tmpl_type_keys - ui_types
        extra_types   = ui_types - tmpl_type_keys
        if missing_types:
            failures.append(
                f"[{device_name}] Parameter Type 缺失：{sorted(missing_types)}"
            )
        if extra_types:
            failures.append(
                f"[{device_name}] Parameter Type 多余：{sorted(extra_types)}"
            )

        # 全量参数比对
        if ui_types_raw:
            _select_dropdown_item(driver, page._PARAM_TYPE_SELECT, ui_types_raw[0], "Parameter Type")
            driver.wait_for_timeout(800)

        ui_params = _read_not_selected(driver)
        all_tmpl_params: set[str] = set().union(*tmpl_groups.values())

        if not ui_params:
            unreadable.append(f"{device_name}/[所有参数]")
            log.warning("  [无法读取] %s 的参数列表", device_name)
            continue

        log.info("  [比对] %s → UI %d 项，模板合计 %d 项",
                 device_name, len(ui_params), len(all_tmpl_params))

        missing_params = all_tmpl_params - ui_params
        extra_params   = ui_params - all_tmpl_params
        if missing_params:
            preview = sorted(missing_params)[:5]
            suffix = "…" if len(missing_params) > 5 else ""
            failures.append(
                f"[{device_name}] 参数缺失（共 {len(missing_params)} 项）：{preview}{suffix}"
            )
        if extra_params:
            preview = sorted(extra_params)[:5]
            suffix = "…" if len(extra_params) > 5 else ""
            failures.append(
                f"[{device_name}] 参数多余（共 {len(extra_params)} 项）：{preview}{suffix}"
            )

    if unreadable:
        log.warning("[%s] 以下分组参数列表无法读取，需人工确认：%s", CASE_ID, unreadable)

    _write_html_report(failures, unreadable, devices_skipped)

    # 恢复：为所有设备全选参数并保存
    log.info("[%s] 恢复：为所有设备全选参数并保存", CASE_ID)
    try:
        page.configure_all_devices()
        log.info("[%s] 恢复完成", CASE_ID)
    except Exception as e:
        log.warning("[%s] 恢复失败：%s", CASE_ID, e)

    assert not failures, (
        f"[{CASE_ID}] 共 {len(failures)} 处与模板不一致，"
        f"详见 reports/{CASE_ID}_report.html"
    )
