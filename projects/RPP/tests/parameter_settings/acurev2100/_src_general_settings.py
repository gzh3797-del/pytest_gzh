"""
Test: AcuRev-2100 General Settings
1. Chrome 登录 https://192.168.2.9
2. Gateway Devices -> 点击 AcuRev2100
3. Settings -> General -> 设置参数
4. Modbus TCP 读取寄存器，断言与页面设置值一致
"""

import os
import sys
import time
import logging
import openpyxl
import pytest
from pymodbus.client import ModbusTcpClient
from playwright.sync_api import sync_playwright

# ── 日志 ──────────────────────────────────────────────────────────────────────
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "logs")
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f"test_{time.strftime('%Y%m%d_%H%M%S')}.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

# ── 配置 ──────────────────────────────────────────────────────────────────────
WEB_URL      = "https://192.168.2.9"
USERNAME     = "admin"
PASSWORD     = "Admin@123"

MODBUS_HOST  = "192.168.2.64"
MODBUS_PORT  = 502
SLAVE_ID     = 101

EXCEL_PATH   = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..", "..", "..", "data", "registers",
    "AcuRev2100_ Modbus Address_v1.02_20260406.xlsx"
)

# 页面设置值
# Fixed 模式下 Sub-Interval 置灰不可编辑，UI 跳过设置，Modbus 仅读取不断言
SET_VALUES = {
    "Rated Voltage":  10,
    "Demand Method":  "固定区块",  # Modbus 值=2(Fixed)；页面标签=Sliding Window；选项=Fixed Block Demand
    "Demand Window":  1,           # 页面标签=Averaging Interval Window，范围 1-30
    "Demand Sub Interval": 1,      # 页面标签=Sub-Interval，Fixed 模式置灰，仅 Modbus 读取
}

# Demand Method 页面文本 -> Modbus 寄存器值
DEMAND_METHOD_MAP = {
    "Sliding":   0,
    "Rolling":   1,
    "Fixed":     2,
    "固定区块":  2,
    "Thermal":   3,
}


# ── 从 Excel 解析 Meter Settings 寄存器地址 ───────────────────────────────────

def parse_modbus_addresses(excel_path: str) -> dict:
    """
    读取 Meter Settings sheet，逐行累加寄存器数量计算实际地址。
    返回 {描述关键词: 寄存器地址} 字典。
    """
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=False)
    ws = wb["Meter Settings"]

    rows = list(ws.iter_rows(values_only=True))
    start_addr = 2048   # 第一个数据行起始地址（十进制）
    current_addr = start_addr
    addr_map = {}

    for i, row in enumerate(rows[1:], start=2):   # 跳过表头
        desc = row[5]   # 列 F：描述
        reg  = row[8]   # 列 I：寄存器数量
        if not isinstance(reg, (int, float)):
            continue
        reg = int(reg)
        if desc:
            addr_map[str(desc).strip()] = current_addr
        current_addr += reg

    wb.close()

    # 提取目标字段地址
    targets = {
        "Rated Voltage":         _find(addr_map, "Rated Voltage"),
        "Demand Method":         _find(addr_map, "Demand Calculation Method"),
        "Demand Window":         _find(addr_map, "Demand Interval"),
        "Demand Sub Interval":   _find(addr_map, "Demand Sub-Interval"),
    }
    logging.info("[Excel] 解析寄存器地址:")
    for name, addr in targets.items():
        logging.info(f"  {name:<25} 0x{addr:04X} ({addr})")
    return targets


def _find(addr_map: dict, keyword: str) -> int:
    for desc, addr in addr_map.items():
        if keyword.lower() in desc.lower():
            return addr
    raise KeyError(f"Excel 中未找到包含 '{keyword}' 的描述行")


# ── Modbus 读取 ───────────────────────────────────────────────────────────────

def modbus_read(address: int, count: int = 1) -> int:
    """
    读取 Modbus TCP 保持寄存器，并打印完整的原始报文字节。

    Modbus TCP 请求帧结构（12 字节）：
      [0-1] Transaction ID  [2-3] Protocol ID(0x0000)
      [4-5] Length(0x0006)  [6] Unit ID  [7] FC(0x03)
      [8-9] Start Addr      [10-11] Quantity

    Modbus TCP 响应帧结构：
      [0-1] Transaction ID  [2-3] Protocol ID  [4-5] Length
      [6] Unit ID  [7] FC(0x03)  [8] Byte Count  [9+] Register Data
    """
    client = ModbusTcpClient(host=MODBUS_HOST, port=MODBUS_PORT, timeout=5)
    try:
        assert client.connect(), f"Modbus TCP 连接失败 {MODBUS_HOST}:{MODBUS_PORT}"

        # 构造并打印请求原始报文
        tx_bytes = bytes([
            0x00, 0x01,                          # Transaction ID
            0x00, 0x00,                          # Protocol ID
            0x00, 0x06,                          # Length
            SLAVE_ID & 0xFF,                     # Unit ID
            0x03,                                # Function Code
            (address >> 8) & 0xFF, address & 0xFF,   # Start Address
            (count >> 8) & 0xFF,   count & 0xFF,     # Quantity
        ])
        tx_hex = " ".join(f"{b:02X}" for b in tx_bytes)
        logging.info(
            f"[Modbus TX] {MODBUS_HOST}:{MODBUS_PORT} -> "
            f"SlaveID={SLAVE_ID}  FC=03  "
            f"Addr=0x{address:04X}({address})  Count={count}\n"
            f"           RAW: {tx_hex}"
        )

        resp = client.read_holding_registers(address, count=count, device_id=SLAVE_ID)
        assert not resp.isError(), f"Modbus 响应错误: {resp}"

        # 构造并打印响应原始报文
        data_bytes = b"".join(r.to_bytes(2, "big") for r in resp.registers)
        rx_bytes = bytes([
            0x00, 0x01,                          # Transaction ID
            0x00, 0x00,                          # Protocol ID
            0x00, len(data_bytes) + 3,           # Length
            SLAVE_ID & 0xFF,                     # Unit ID
            0x03,                                # Function Code
            len(data_bytes),                     # Byte Count
        ]) + data_bytes
        rx_hex = " ".join(f"{b:02X}" for b in rx_bytes)
        logging.info(
            f"[Modbus RX] {MODBUS_HOST}:{MODBUS_PORT} <- "
            f"FC=03  Addr=0x{address:04X}({address})  "
            f"Registers={resp.registers}  "
            f"Hex={[f'0x{r:04X}' for r in resp.registers]}\n"
            f"           RAW: {rx_hex}"
        )
        return resp.registers[0]
    finally:
        client.close()


# ── 浏览器操作 ────────────────────────────────────────────────────────────────

def run_ui():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            channel="chrome",
            headless=False,
        )
        ctx  = browser.new_context(ignore_https_errors=True)
        page = ctx.new_page()

        try:
            # 1. 登录
            logging.info(f"[UI] 打开 {WEB_URL}")
            page.goto(WEB_URL, wait_until="domcontentloaded", timeout=30000)

            page.locator("input[name='username'], input[type='text']").first.fill(USERNAME)
            page.locator("input[type='password']").first.fill(PASSWORD)
            page.get_by_role("button", name="Sign In").click()
            # 等待登录后导航菜单出现，而非等待页面 load 事件
            page.get_by_text("Gateway Devices", exact=True).first.wait_for(state="visible", timeout=15000)
            logging.info("[UI] 登录完成")

            # 2. Gateway Devices
            logging.info("[UI] 进入 Gateway Devices")
            page.get_by_text("Gateway Devices", exact=True).first.click()
            page.get_by_text("AcuRev2100").first.wait_for(state="visible", timeout=10000)

            # 3. 点击 AcuRev2100
            logging.info("[UI] 点击 AcuRev2100")
            page.get_by_text("AcuRev2100").first.click()
            page.get_by_text("Settings", exact=True).first.wait_for(state="visible", timeout=10000)

            # 4. Settings -> General
            logging.info("[UI] Settings -> General")
            page.get_by_text("Settings", exact=True).first.click()
            page.get_by_text("General", exact=True).first.wait_for(state="visible", timeout=5000)
            page.get_by_text("General", exact=True).first.click()
            page.wait_for_timeout(1500)

            # 5a. 设置 Rated Voltage，点击第一个 Save 保存上半部分
            _set_by_heading(page, "Rated Voltage", str(SET_VALUES["Rated Voltage"]))
            logging.info("[UI] 点击第一个 Save（保存 Rated Voltage 等）")
            page.get_by_role("button", name="Save").first.click()
            page.wait_for_timeout(3000)   # 等待设备写入配置
            logging.info("[UI] 第一个 Save 完成")

            # 5b. 设置 Demand 字段，点击最后一个 Save 保存 Demand 区域
            _set_demand_method(page, SET_VALUES["Demand Method"])
            _set_demand_input(page, "Averaging Interval Window", str(SET_VALUES["Demand Window"]))
            # Sub-Interval 在 Fixed 模式下置灰，跳过 UI 设置

            logging.info("[UI] 点击最后一个 Save（保存 Demand 设置）")
            page.get_by_role("button", name="Save").last.click()
            page.wait_for_timeout(3000)   # 等待设备写入配置
            logging.info("[UI] 最后一个 Save 完成")

        finally:
            ctx.close()
            browser.close()


def _set_by_heading(page, heading: str, value: str):
    """点击 section 标题下方的第一个 input（适用于 Rated Voltage 这类卡片布局）。"""
    logging.info(f"[UI] 设置 {heading} = {value}")
    inp = page.get_by_text(heading, exact=True).locator("xpath=following::input[1]")
    inp.wait_for(state="visible", timeout=8000)
    inp.click(click_count=3)
    inp.fill(value)


def _set_demand_method(page, value: str):
    """设置 Demand 区域的 Sliding Window 下拉（Demand Method）。
    兼容原生 <select> 和自定义下拉组件两种情况。
    """
    option_candidates = {
        "固定区块": ["Fixed Block Demand", "Fixed Block", "Fixed"],
        "Fixed":    ["Fixed Block Demand", "Fixed Block", "Fixed"],
        "Rolling":  ["Rolling Window Demand", "Rolling Window", "Rolling"],
        "Thermal":  ["Thermal Demand", "Thermal"],
    }
    candidates = option_candidates.get(value, [value])
    logging.info(f"[UI] 设置 Demand Method = {value}，候选: {candidates}")

    # 滚动到 Demand 区域
    page.get_by_text("Sliding Window", exact=False).first.scroll_into_view_if_needed()

    # 截图确认当前状态
    shot = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "logs", "demand_section.png")
    page.screenshot(path=shot, full_page=False)
    logging.info(f"[Screenshot] logs/demand_section.png")

    # 方式1：原生 <select>
    sel = page.get_by_text("Sliding Window", exact=False).first.locator("xpath=following::select[1]")
    try:
        sel.wait_for(state="visible", timeout=3000)
        for opt in candidates:
            try:
                sel.select_option(label=opt)
                logging.info(f"[UI] Demand Method (native select) 选中: {opt}")
                return
            except Exception:
                continue
    except Exception:
        pass

    # 方式2：自定义下拉 —— 点击触发展开，再点击选项文字
    dropdown = page.get_by_text("Sliding Window", exact=False).first.locator(
        "xpath=following::*[self::div or self::span][contains(@class,'select') "
        "or contains(@class,'dropdown') or contains(@class,'picker')][1]"
    )
    try:
        dropdown.wait_for(state="visible", timeout=3000)
        dropdown.click()
    except Exception:
        # 兜底：直接点击 Sliding Window 后面第一个可点击元素
        page.get_by_text("Sliding Window", exact=False).first.locator(
            "xpath=following::*[1]"
        ).click()
    page.wait_for_timeout(400)

    for opt in candidates:
        opt_loc = page.get_by_text(opt, exact=False).first
        try:
            opt_loc.wait_for(state="visible", timeout=2000)
            opt_loc.click()
            logging.info(f"[UI] Demand Method (custom dropdown) 选中: {opt}")
            return
        except Exception:
            continue

    raise RuntimeError(f"Demand Method 所有候选选项均未找到: {candidates}")


def _set_demand_input(page, page_label: str, value: str):
    """设置 Demand 区域指定标签旁的 input。"""
    logging.info(f"[UI] 设置 {page_label} = {value}")
    inp = page.get_by_text(page_label, exact=True).locator("xpath=following::input[1]")
    inp.wait_for(state="visible", timeout=8000)
    inp.click(click_count=3)
    inp.fill(value)



# ── 测试 ──────────────────────────────────────────────────────────────────────

def test_general_settings():
    logging.info("=" * 60)
    logging.info(f"开始时间: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info(f"Web: {WEB_URL}  Modbus: {MODBUS_HOST}:{MODBUS_PORT} SlaveID={SLAVE_ID}")

    # 读取 Excel 获取寄存器地址
    addr = parse_modbus_addresses(EXCEL_PATH)

    # 浏览器设置参数
    run_ui()

    # 等待设备写入并应用配置
    logging.info("[Wait] Save 后等待 5s，确保设备写入完成...")
    time.sleep(5)

    # Modbus 读取
    logging.info("-" * 60)
    mb_voltage = modbus_read(addr["Rated Voltage"])
    mb_method  = modbus_read(addr["Demand Method"])
    mb_window  = modbus_read(addr["Demand Window"])
    mb_sub     = modbus_read(addr["Demand Sub Interval"])  # Fixed 模式置灰，仅读取

    exp_voltage = SET_VALUES["Rated Voltage"]
    exp_method  = DEMAND_METHOD_MAP[SET_VALUES["Demand Method"]]
    exp_window  = SET_VALUES["Demand Window"]

    logging.info("=" * 60)
    logging.info("断言结果:")
    logging.info(f"  Rated Voltage       : 期望={exp_voltage}         Modbus={mb_voltage}  {'PASS' if mb_voltage == exp_voltage else 'FAIL'}")
    logging.info(f"  Demand Method       : 期望={exp_method}(Fixed)   Modbus={mb_method}   {'PASS' if mb_method  == exp_method  else 'FAIL'}")
    logging.info(f"  Demand Window       : 期望={exp_window}          Modbus={mb_window}   {'PASS' if mb_window  == exp_window  else 'FAIL'}")
    logging.info(f"  Demand Sub-Interval : Fixed模式UI置灰仅读        Modbus={mb_sub}")
    logging.info("=" * 60)

    assert mb_voltage == exp_voltage, f"Rated Voltage 不一致: 期望={exp_voltage}, Modbus={mb_voltage}"
    assert mb_method  == exp_method,  f"Demand Method 不一致: 期望={exp_method}(Fixed), Modbus={mb_method}"
    assert mb_window  == exp_window,  f"Demand Window 不一致: 期望={exp_window}, Modbus={mb_window}"
