# -*- coding: utf-8 -*-
"""
debug_server_connect.py — Post Channel 服务器连通性调试脚本

逐一用 Post Channel 1 配置 FTP / SFTP / HTTP / HTTPS，
启动对应本地服务器，点击 Test Post Channel，检测弹窗结果。

用法（从仓库根）：
    python Protocols/Datalog/debug_server_connect.py
    python Protocols/Datalog/debug_server_connect.py --protos FTP SFTP
    python Protocols/Datalog/debug_server_connect.py --protos HTTP HTTPS
"""
from __future__ import annotations

import argparse
import logging
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))        # Protocols/Datalog/
sys.path.insert(0, str(Path(__file__).parent.parent)) # Protocols/
sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # 仓库根

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

import config
from datalog_server_verifier import (
    ServerInfo,
    ServerManager,
    build_protocol_pool,
    _init_driver,
    _login,
)
from datalog_page import PostChannelPage


def _test_one_protocol(pc_page: PostChannelPage, si: ServerInfo) -> dict:
    """
    将 Post Channel 1 配置为 si 对应协议，点击 Test，返回结果字典。
    """
    cfg = si.to_post_channel_config()
    log.info("=" * 60)
    log.info("测试协议：%s  服务器：%s:%d", si.protocol, si.host, si.port)
    log.info("=" * 60)

    try:
        result_text = pc_page.configure_channel(1, cfg, enabled=True, test=True)
        result_lower = result_text.lower()
        is_success = any(kw in result_lower for kw in
                         ("success", "connected", "test success", "pass", "ok"))
        is_fail    = any(kw in result_lower for kw in
                         ("fail", "error", "failed", "test fail"))

        if result_text == "":
            status = "UNKNOWN（弹窗未检测到）"
        elif is_success and not is_fail:
            status = "PASS"
        elif is_fail:
            status = "FAIL"
        else:
            status = f"UNKNOWN（{result_text}）"

        return {"protocol": si.protocol, "status": status,
                "dialog": result_text, "error": None}

    except RuntimeError as e:
        return {"protocol": si.protocol, "status": "FAIL",
                "dialog": "", "error": str(e)}
    except Exception as e:
        return {"protocol": si.protocol, "status": "ERROR",
                "dialog": "", "error": str(e)}


def _inspect_http_form(driver):
    """导航到 Post Channel 1，选 HTTP/HTTPS，打印所有 label 和 el-select 的信息。"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.wait import WebDriverWait

    # 导航到 Post Channel 1
    base = driver.current_url.split('#')[0]
    driver.get(f"{base}#/dataLog")
    time.sleep(2)
    driver.get(f"{base}#/dataLog/postChannels/postChannel1")
    time.sleep(3)

    # 选 HTTP/HTTPS Post Method
    method_sel = (By.XPATH,
        "//label[contains(@class,'el-form-item__label') and normalize-space(.)='Post Method']"
        "/following::div[contains(@class,'el-select')][1]")
    try:
        el = WebDriverWait(driver, 8).until(EC.presence_of_element_located(method_sel))
        from selenium.webdriver.common.keys import Keys
        inner = el.find_elements(By.XPATH, ".//div[contains(@class,'el-select__wrapper')]")
        (inner[0] if inner else el).click()
        time.sleep(0.8)
        for pat in [
            "//li[contains(@class,'el-select-dropdown__item') and normalize-space(.)='HTTP/HTTPS']",
            "//div[contains(@class,'el-popper')]//li[contains(normalize-space(.),'HTTP/HTTPS')]",
        ]:
            opts = driver.find_elements(By.XPATH, pat)
            if opts:
                opts[0].click()
                break
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        time.sleep(1.5)
    except Exception as e:
        log.warning("选 HTTP/HTTPS 失败：%s", e)

    # 打印所有 el-form-item__label 的文字
    labels = driver.find_elements(By.XPATH, "//label[contains(@class,'el-form-item__label')]")
    print("\n=== 表单 Labels ===")
    for i, lbl in enumerate(labels):
        print(f"  [{i}] {repr(lbl.text.strip())}")

    # 打印所有 el-select 容器的位置
    selects = driver.find_elements(By.XPATH, "//div[contains(@class,'el-select')]")
    print(f"\n=== el-select 容器（共 {len(selects)} 个）===")
    for i, sel in enumerate(selects[:20]):
        try:
            txt = sel.text.strip()[:60]
            loc = sel.location
            print(f"  [{i}] text={repr(txt)}  y={loc.get('y','?')}")
        except Exception:
            pass

    # 打印所有 el-radio 的文字（可能是 Yes/No 类型）
    radios = driver.find_elements(By.XPATH, "//label[contains(@class,'el-radio')]")
    print(f"\n=== el-radio Labels（共 {len(radios)} 个）===")
    for i, r in enumerate(radios):
        print(f"  [{i}] {repr(r.text.strip())}")

    # 打印所有 input placeholder
    inputs = driver.find_elements(By.XPATH, "//input[@placeholder]")
    print(f"\n=== Input placeholders ===")
    for inp in inputs:
        print(f"  {repr(inp.get_attribute('placeholder'))}")
    print()


def main(protos: list[str]):
    pool = build_protocol_pool()

    # 只启动需要测试的协议对应的服务器
    active = []
    for p in protos:
        p = p.upper()
        if p not in pool:
            log.warning("未知协议 %s，跳过", p)
            continue
        si = pool[p]
        if si.protocol == "HTTPS" and not (si.ssl_certfile and si.ssl_keyfile):
            log.warning("HTTPS 未配置证书（DATALOG_SSL_CERT / DATALOG_SSL_KEY），跳过")
            continue
        active.append(si)

    if not active:
        log.error("没有可测试的协议，退出")
        return

    log.info("将测试以下协议：%s", [s.protocol for s in active])

    with ServerManager(active):
        time.sleep(1)

        driver = _init_driver(config.GATEWAY_WEB_URL)
        try:
            _login(driver, config.GATEWAY_WEB_USER, config.GATEWAY_WEB_PASS)
            pc_page = PostChannelPage(driver)

            results = []
            for si in active:
                res = _test_one_protocol(pc_page, si)
                results.append(res)
                # 协议间停顿，让页面稳定
                time.sleep(2)

        finally:
            driver.quit()

    # ── 汇总报告 ──────────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("Post Channel 连通性测试结果")
    print("=" * 60)
    all_pass = True
    for r in results:
        icon = "[PASS]" if r["status"] == "PASS" else ("[FAIL]" if "FAIL" in r["status"] else "[????]")
        print(f"  {icon} {r['protocol']:6s}  {r['status']}")
        if r["dialog"]:
            print(f"       弹窗：{r['dialog']}")
        if r["error"]:
            print(f"       错误：{r['error']}")
        if r["status"] != "PASS":
            all_pass = False
    print("=" * 60)
    if all_pass:
        print("全部通过 [OK]")
    else:
        print("存在失败项，请检查：")
        print("  1. config.py 中 DATALOG_SERVER_HOST 是否为网关可访问的本机 IP")
        print("  2. 防火墙是否放行对应端口（FTP:2121 / SFTP:2222 / HTTP:8080 / HTTPS:8443）")
        print("  3. 网关 Web UI 中 Post Channel 1 的 URL 和账号密码是否填写正确")
    print()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Post Channel 服务器连通性调试")
    parser.add_argument(
        "--protos", nargs="+",
        default=["FTP", "SFTP", "HTTP", "HTTPS"],
        help="要测试的协议列表（默认全部）",
    )
    parser.add_argument("--inspect", action="store_true",
                        help="仅打印 HTTP/HTTPS 表单的元素结构，不执行测试")
    args = parser.parse_args()

    if args.inspect:
        driver = _init_driver(config.GATEWAY_WEB_URL)
        try:
            _login(driver, config.GATEWAY_WEB_USER, config.GATEWAY_WEB_PASS)
            _inspect_http_form(driver)
        finally:
            driver.quit()
    else:
        main(args.protos)
