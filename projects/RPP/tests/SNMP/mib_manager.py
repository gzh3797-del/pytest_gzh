"""
MIB 文件管理器

每次测试会话自动执行：
  1. 登录 HMI → 进入 SNMP 页面 → 下载 MIB 压缩包并解压到 mib/ 目录
  2. 遍历 Physical Devices → 进入每台设备 Settings > Connection → 读取 Template 选中值
  3. Template 值 → 匹配 mib/ 目录中对应版本的 MIB 文件
  4. 将 {device_name: {template, mib_file, entry_name}} 写入 mib_mapping.json

供 snmp_oid_map.py 在运行时动态加载对应 MIB，替代手动导入。
"""
import json
import logging
import os
import re
import tarfile
from collections import Counter
from pathlib import Path

from playwright.sync_api import sync_playwright

import sys

# 迁移到 RPP：连接配置改用 projects/RPP/settings.py，不再依赖 projects.AcuHMI_1_7。
_RPP_ROOT = Path(__file__).resolve().parents[2]  # projects/RPP/
if str(_RPP_ROOT) not in sys.path:
    sys.path.insert(0, str(_RPP_ROOT))
from settings import BASE_URL, DEFAULT_USERNAME, DEFAULT_PASSWORD

log = logging.getLogger(__name__)

_SNMP_DIR = Path(__file__).parent
MIB_DIR = _SNMP_DIR / "mib"
MIB_MAPPING_PATH = _SNMP_DIR / "mib_mapping.json"


# ── MIB 下载与解压 ────────────────────────────────────────────────────────────

def _download_mib_tarball(page) -> None:
    """从当前 SNMP 页面下载 MIB 文件，解压/保存到 mib/ 目录。"""
    MIB_DIR.mkdir(exist_ok=True)

    btn = page.locator('button:has-text("MIB"), a:has-text("MIB")').first
    if btn.count() == 0:
        log.warning("[MIB] 未找到 MIB 下载按钮")
        return

    try:
        with page.expect_download(timeout=30000) as dl_info:
            btn.click()
        dl = dl_info.value
        filename = dl.suggested_filename
        dest = MIB_DIR / filename
        dl.save_as(str(dest))
        log.info("[MIB] 下载完成: %s (%d bytes)", filename, dest.stat().st_size)
        print(f"[MIB] 下载完成: {filename}", flush=True)

        if filename.endswith((".tar.gz", ".tgz", ".tar", ".tar.bz2", ".tar.xz")):
            try:
                with tarfile.open(str(dest), "r:*") as tar:
                    tar.extractall(str(MIB_DIR))
                dest.unlink()
                log.info("[MIB] 解压完成，已删除压缩包")
                print("[MIB] 解压完成", flush=True)
            except tarfile.TarError:
                # 文件扩展名为 tar.gz 但实际可能是 zip 或裸 tar
                import zipfile
                if zipfile.is_zipfile(str(dest)):
                    with zipfile.ZipFile(str(dest)) as zf:
                        zf.extractall(str(MIB_DIR))
                    dest.unlink()
                    log.info("[MIB] 解压完成（zip），已删除压缩包")
                    print("[MIB] 解压完成（zip）", flush=True)
                else:
                    log.warning("[MIB] 无法识别压缩格式，保留原文件: %s", filename)
                    print(f"[MIB] 无法识别压缩格式，保留原文件: {filename}", flush=True)
        elif filename.endswith(".zip"):
            import zipfile
            with zipfile.ZipFile(str(dest)) as zf:
                zf.extractall(str(MIB_DIR))
            dest.unlink()
            log.info("[MIB] 解压完成（zip），已删除压缩包")
            print("[MIB] 解压完成（zip）", flush=True)
    except Exception as e:
        log.warning("[MIB] MIB 下载失败: %s", e)
        print(f"[MIB] 下载失败: {e}", flush=True)


# ── 页面导航 ──────────────────────────────────────────────────────────────────

def _nav_to_physical_devices(page) -> None:
    # 先展开 Devices 父菜单（若已展开则点击无害）
    devices_menu = page.locator(
        "a:has-text('Devices'), span:has-text('Devices'), li:has-text('Devices')"
    ).first
    devices_menu.click(timeout=5000)
    page.wait_for_timeout(300)
    # 再点 Physical Devices 子项
    page.locator(
        "a:has-text('Physical Devices'), "
        "span:has-text('Physical Devices'), "
        "li:has-text('Physical Devices')"
    ).first.click(timeout=5000)
    page.wait_for_selector("table", timeout=10000)
    page.wait_for_timeout(500)


# ── Template 值读取 ───────────────────────────────────────────────────────────

def _read_template_from_page(page) -> str:
    """
    在已打开的 Settings > Connection 页面读取 Template 下拉框的当前选中值。

    Element Plus el-select 的选中文本优先通过 JS el.value 属性读取；
    若为空，退而用页面中出现次数最多的含 'Firmware' 文本作为兜底。
    """
    # 主路：JS 读取 el-select 输入框的 value 属性
    try:
        template_item = page.locator(".el-form-item").filter(has_text="Template").first
        if template_item.count() > 0:
            inp = template_item.locator(".el-select .el-input__inner, .el-input__inner").first
            if inp.count() > 0:
                val = (inp.evaluate("el => el.value") or "").strip()
                if val:
                    return val
    except Exception:
        pass

    # 兜底：统计页面所有含 Firmware 文本，出现次数最多的即选中值
    try:
        texts: list[str] = []
        for el in page.get_by_text("Firmware", exact=False).all():
            try:
                t = el.inner_text().strip()
                if "Firmware" in t and 10 < len(t) < 200:
                    texts.append(t)
            except Exception:
                pass
        if texts:
            most_common, count = Counter(texts).most_common(1)[0]
            if count >= 2 or len(set(texts)) == 1:
                return most_common
            # 如果所有值都只出现一次，返回第一个（通常是 selected 显示值）
            return texts[0]
    except Exception:
        pass

    return ""


def _read_device_template(page, device_name: str) -> str:
    """
    点击设备名 → 展开 Settings → 点击 Connection → 读取 Template 值。
    读取后自动返回 Physical Devices 列表。
    """
    try:
        page.get_by_text(device_name, exact=True).first.click()
        page.wait_for_selector("text=Settings", timeout=10000)
        page.wait_for_timeout(500)

        conn_sel = (
            "a:has-text('Connection'), "
            "li:has-text('Connection'), "
            "span:has-text('Connection')"
        )
        settings_sel = (
            "span:has-text('Settings'), a:has-text('Settings'), li:has-text('Settings')"
        )
        # 只在 Connection 不可见时展开 Settings，最多点一次避免 toggle 关闭菜单
        if not page.locator(conn_sel).first.is_visible():
            page.locator(settings_sel).first.click()
            try:
                page.wait_for_selector(conn_sel, state="visible", timeout=3000)
            except Exception:
                page.wait_for_timeout(2000)

        page.locator(conn_sel).first.click(timeout=8000)
        page.wait_for_timeout(1000)

        return _read_template_from_page(page)

    except Exception as e:
        log.warning("[MIB] 读取 %s Template 失败: %s", device_name, e)
        return ""

    finally:
        try:
            _nav_to_physical_devices(page)
        except Exception:
            pass


# ── Template → MIB 文件匹配 ───────────────────────────────────────────────────

def match_template_to_mib(template_value: str) -> str | None:
    """
    根据 Template 值匹配 mib/ 目录中对应的 MIB 文件名（仅文件名，不含路径）。

    Template 格式: {TYPE}_{template_ver}_Firmware_v{fw_ver}
      示例: PXM350_v1.03p01_Firmware_v2.24
            PXB-M24-XMA-GEN_v1.03p01_Firmware_v1.03
            PXE1_v1.03p01_Firmware_v6.36

    MIB 文件格式: p{TYPE}Modbus_v{fw_ver_with_optional_patch}.MIB
      示例: pXM350Modbus_v2.24.MIB
            pXB-M24-XMA-GENModbus_v1.03p05.MIB
            pXE1Modbus_v6.36.MIB

    版本匹配：MIB 文件版本以 Template fw_ver 开头即视为匹配（精确匹配优先）。
    """
    if not template_value or not MIB_DIR.exists():
        return None

    m = re.match(r"^(.+?)_v[\d.p]+_Firmware_v([\d.]+)", template_value)
    if not m:
        return None
    type_prefix = m.group(1)                          # e.g. "PXM350", "PXB-M24-XMA-GEN"
    fw_ver = m.group(2)                               # e.g. "2.24", "1.03", "6.36"
    mib_type = type_prefix[0].lower() + type_prefix[1:]  # "pXM350", "pXB-M24-XMA-GEN"

    best: str | None = None
    for fpath in sorted(MIB_DIR.iterdir()):
        if fpath.suffix.upper() != ".MIB":
            continue
        stem = fpath.stem
        modbus_idx = stem.lower().find("modbus")
        if modbus_idx < 0:
            continue
        file_type = stem[:modbus_idx].rstrip("_")     # e.g. "pXM350", "AcuRev-1300"（去掉尾部 _）
        if file_type.lower() != mib_type.lower():
            continue
        ver_m = re.search(r"_v([\d.p]+)$", stem)
        if not ver_m:
            continue
        file_ver = ver_m.group(1)                     # e.g. "2.24", "1.03p05", "6.36"
        if file_ver == fw_ver:
            return fpath.name                         # 精确匹配，直接返回
        if file_ver.startswith(fw_ver):
            best = fpath.name                         # 前缀匹配，继续寻找精确

    return best


# Physical Devices 中跳过模板读取的设备名前缀（无 Settings/Connection 页的 I/O 模块等）
_SKIP_TEMPLATE_PREFIXES: frozenset[str] = frozenset({"H-IO", "H-IOM", "AcuIOM"})


# Template 类型前缀 → 内部 model_type
# PX-EMD-G 风格：前缀为内部型号代码（PXM350、PXE1 等）
# AcuHMI 风格：前缀直接是设备型号名（AcuRev-1300、AcuvimIIR 等）
_TEMPLATE_PREFIX_TO_MODEL: dict[str, str] = {
    # PX-EMD-G 风格
    "PXB":            "AcuRev4100",
    "PXE1":           "AcuvimIIR",
    "PXE2":           "AcuvimIIW",
    "PXM350":         "AcuRev1300",
    "AcuRev2100":     "AcuRev2100",
    "AcuVIM3":        "Acuvim3",
    "Acuvim3":        "Acuvim3",
    # AcuHMI 风格
    "AcuRev-1300":    "AcuRev1300",
    "AcuRev-2100":    "AcuRev2100",
    "AcuRev-4110-mA": "AcuRev4100",
    "AcuRev-4110-mV": "AcuRev4100",
    "AcuvimIIR":      "AcuvimIIR",
    "AcuvimIIW":      "AcuvimIIW",
    "AcuVim3":        "Acuvim3",
}


def _derive_model_type(template_value: str) -> str | None:
    """从 Template 值推导 model_type。"""
    if not template_value:
        return None
    m = re.match(r"^(.+?)_v", template_value)
    if not m:
        return None
    type_prefix = m.group(1)               # e.g. "PXB-M24-XMA-GEN", "PXE1", "PXM350"
    short = type_prefix.split("-")[0]      # e.g. "PXB", "PXE1", "PXM350"
    return _TEMPLATE_PREFIX_TO_MODEL.get(short) or _TEMPLATE_PREFIX_TO_MODEL.get(type_prefix)


def _derive_model_type_from_name(dev_name: str) -> str | None:
    """从 SNMP 页面设备名称推导 model_type（Physical Devices 导航失败时的回退）。

    PX-EMD-G 上设备名如 "AcuRev1300PXM350Modbus RTU" 包含 Template 前缀 "PXM350"，
    通过 _TEMPLATE_PREFIX_TO_MODEL 可正确推导出 "AcuRev1300"。
    """
    upper = dev_name.upper()
    for prefix, model_type in _TEMPLATE_PREFIX_TO_MODEL.items():
        if prefix.upper() in upper:
            return model_type
    # 直接型号名匹配（适用于设备名与型号名相同或包含型号名的情况）
    for model_type in ("AcuRev4100", "AcuRev2100", "AcuRev1300",
                       "AcuvimIIW", "AcuvimIIR", "Acuvim3"):
        if model_type.upper() in upper:
            return model_type
    # 容错匹配：PX-EMD-G 部分设备名拼写变体（如 "AcivimIIW" ≈ "AcuvimIIW"）
    if "ACIVIMIIW" in upper or "ACIVIM IIW" in upper:
        return "AcuvimIIW"
    if "ACIVIMIIR" in upper or "ACIVIM IIR" in upper:
        return "AcuvimIIR"
    return None


def _find_mib_for_model_type(model_type: str) -> str | None:
    """为 model_type 在 mib/ 目录中查找最匹配的 .MIB 文件（取版本号最大者）。
    同时兼容 PX-EMD-G 风格（pxb…）和 AcuHMI 风格（acurev-4110…）的文件名。
    """
    _STEM_PATTERNS: dict[str, list[str]] = {
        "AcuRev4100": ["pxb", "acurev-4110"],
        "AcuvimIIR":  ["pxe1", "acuvimiir"],
        "AcuvimIIW":  ["pxe2", "acuvimiiw"],
        "AcuRev1300": ["pxm350", "acurev-1300"],
        "AcuRev2100": ["pacurev2100", "acurev-2100"],
        "Acuvim3":    ["pacuvim3", "acuvim3"],
    }
    patterns = _STEM_PATTERNS.get(model_type, [])
    if not patterns or not MIB_DIR.exists():
        return None
    candidates = sorted(
        f for f in MIB_DIR.iterdir()
        if f.suffix.upper() == ".MIB"
        and any(f.stem.lower().startswith(p) for p in patterns)
    )
    return candidates[-1].name if candidates else None


def mib_filename_to_entry_name(mib_filename: str) -> str:
    """
    MIB 文件名 → DeviceReadingEntry 名称。
    规则：去掉 .MIB 后缀，将 '-' 和 '.' 替换为 '_'，拼接 'DeviceReadingEntry'。

    pXE1Modbus_v6.36.MIB        → pXE1Modbus_v6_36DeviceReadingEntry
    pXB-M24-XMA-GENModbus_v1.03p05.MIB → pXB_M24_XMA_GENModbus_v1_03p05DeviceReadingEntry
    """
    stem = Path(mib_filename).stem
    prefix = stem.replace("-", "_").replace(".", "_")
    return f"{prefix}DeviceReadingEntry"


# ── 主入口 ────────────────────────────────────────────────────────────────────

def build_and_save_mapping() -> None:
    """
    完整流程（每次测试会话自动执行）：
      1. 登录 HMI → SNMP 页面 → 下载并解压 MIB 文件到 mib/
      2. Physical Devices 列表 → 逐台读取 Template
      3. Template → 匹配 MIB 文件
      4. 结果写入 mib_mapping.json
    """
    from configure_snmp import login_and_goto_snmp

    print("\n[MIB] 开始自动下载 MIB 并读取设备 Template ...", flush=True)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        ctx = browser.new_context(
            ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080},
        )
        page = ctx.new_page()

        # Step 1: 登录 + 读取 SNMP 页面设备列表（在下载前读，此时页面完整加载）
        login_and_goto_snmp(page)
        # 点击 Enable radio（与 configure_snmp_v2c 保持一致），触发 Vue 刷新设备列表
        try:
            page.locator(".el-radio").filter(has_text="Enable").nth(0).click()
            page.wait_for_timeout(300)
            page.wait_for_selector('input[placeholder="Enter Port"]', timeout=5000)
        except Exception:
            page.wait_for_timeout(1000)
        # PX-EMD-G 表格首列为 checkbox（inner_text 为空），必须用 textContent 读整行文字
        # 与 set_device_selection 的读取方式保持一致
        snmp_page_devnames: list[str] = []
        for row in page.query_selector_all(".el-table__row"):
            if row.query_selector(".el-checkbox") is None:
                continue
            row_text = (row.evaluate("el => el.textContent") or "").strip()
            for noise in ("ON", "OFF"):     # 去掉 checkbox 状态后缀
                if row_text.endswith(noise):
                    row_text = row_text[:-len(noise)].strip()
            if row_text:
                snmp_page_devnames.append(row_text)
        print(f"[MIB] SNMP 页面设备列表（共 {len(snmp_page_devnames)} 台）: {snmp_page_devnames}",
              flush=True)

        # Step 1b: 下载 MIB（可能触发页面跳转，所以设备列表已提前读取）
        _download_mib_tarball(page)

        # Step 2b: Physical Devices 列表（部分网关型号导航标签不同，失败时回退）
        nav_ok = False
        device_names: list[str] = []
        try:
            _nav_to_physical_devices(page)
            rows = page.locator("table tbody tr")
            for i in range(rows.count()):
                cells = rows.nth(i).locator("td")
                if cells.count() >= 1:
                    name = cells.nth(0).inner_text().strip()
                    if name:
                        device_names.append(name)
            nav_ok = True
            print(f"[MIB] Physical Devices 发现 {len(device_names)} 台: {device_names}",
                  flush=True)
        except Exception as _nav_err:
            log.warning("[MIB] Physical Devices 导航失败，回退到 SNMP 页面设备名称: %s", _nav_err)
            print(f"[MIB] 回退: 使用 SNMP 页面设备名称 {snmp_page_devnames}", flush=True)
            device_names = snmp_page_devnames

        # Step 3: 逐台读取 Template
        # nav_ok=True：以 Physical Devices 短名称为 key，直接读 Template
        # nav_ok=False：回退到 SNMP 页面拼接名称为 key，从名称推导 model_type
        mapping: dict[str, dict] = {}
        for dev_name in (device_names if nav_ok else snmp_page_devnames):
            if nav_ok and any(dev_name.upper().startswith(p.upper())
                              for p in _SKIP_TEMPLATE_PREFIXES):
                print(f"[MIB]   {dev_name:<40}  (跳过：I/O 模块，无 Settings 页面)", flush=True)
                continue
            if nav_ok:
                template = _read_device_template(page, dev_name)
                if template:
                    mib_file = match_template_to_mib(template)
                    model_type = _derive_model_type(template)
                    status = f"template={template!r:<40} -> {mib_file or '[无匹配 MIB]'}"
                else:
                    model_type = _derive_model_type_from_name(dev_name)
                    mib_file = _find_mib_for_model_type(model_type) if model_type else None
                    status = f"model_type={model_type!r:<20} -> {mib_file or '[无匹配 MIB]'} (模板读取失败)"
            else:
                template = ""
                model_type = _derive_model_type_from_name(dev_name)
                mib_file = _find_mib_for_model_type(model_type) if model_type else None
                status = f"model_type={model_type!r:<20} -> {mib_file or '[无匹配 MIB]'}"
            entry_name = mib_filename_to_entry_name(mib_file) if mib_file else None
            print(f"[MIB]   {dev_name:<40}  {status}", flush=True)
            mapping[dev_name] = {
                "template":   template,
                "mib_file":   f"mib/{mib_file}" if mib_file else None,
                "entry_name": entry_name,
                "model_type": model_type,
            }

        browser.close()

    # Step 4: 写入 mib_mapping.json
    MIB_MAPPING_PATH.write_text(
        json.dumps(mapping, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"[MIB] 映射已写入 {MIB_MAPPING_PATH.name}\n", flush=True)


def load_mapping() -> dict:
    """读取 mib_mapping.json，不存在时返回空字典。"""
    if MIB_MAPPING_PATH.exists():
        try:
            return json.loads(MIB_MAPPING_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}
