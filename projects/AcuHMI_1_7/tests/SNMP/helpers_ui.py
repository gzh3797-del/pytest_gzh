"""
Sprint3 SNMP 全功能测试
FTS_AcuRev4100_WEB2_008_003_case01 ~ case25

运行方式：
  pytest test_snmp_ui.py -v -s
  pytest test_snmp_ui.py -v -s -k "case01 or case02"

设计原则：
  - 全部 25 个用例共用一个浏览器 session（登录一次，结束后关闭）
  - MIB 文件只在 case01 下载一次
  - 需要硬件操作或 Trap 监听的用例标记 skip
"""


import time


import pytest


import logging


from playwright.sync_api import Page


from configure_snmp import (
    login_and_goto_snmp,
    goto_snmp,
    apply_snmp_v2c,
    apply_snmp_v3,
    download_mib_file,
)


from snmp_utils import snmp_walk_device, snmp_walk_device_v3


log = logging.getLogger(__name__)


try:
    from projects.AcuHMI_1_7.settings import BASE_URL
except ImportError:
    BASE_URL = "http://192.168.2.9"


HEADLESS  = False


SLOW_MO   = 500


VALID_COMMUNITY  = "123456789012"


V3_USERNAME      = "admin"


V3_PASSWORD      = "12345678"


V3_PRIV_PASSWORD = "12345678"


def step(msg: str):
    print(f"\n  ▶ {msg}")
    log.info("步骤: %s", msg)


def check_form_error(page: Page) -> bool:
    return len(page.query_selector_all(".el-form-item.is-error, .el-form-item__error")) > 0


def snmp_page(playwright):
    """
    全部 Sprint3 SNMP 测试共用的浏览器 session。
    登录一次，测试完成后关闭浏览器。
    """
    browser = playwright.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
    context = browser.new_context(ignore_https_errors=True)
    page    = context.new_page()
    login_and_goto_snmp(page)
    yield page
    page.wait_for_timeout(2000)
    browser.close()


class SNMPBase:
    """FTS_AcuRev4100_WEB2_008_003 case01 ~ case25"""

    _mib_downloaded = False

    # ── 内部辅助 ──────────────────────────────────────────────────────────────

    def _reload(self, page: Page):
        """每个用例前重新导航到 SNMP 页面（不重新登录）。"""
        goto_snmp(page)

    def _download_mib_once(self, page: Page):
        """第一次调用时下载 MIB，后续跳过。"""
        if not SNMPBase._mib_downloaded:
            step("点击 Download MIB File")
            filename = download_mib_file(page)
            assert filename, "MIB 文件下载失败"
            SNMPBase._mib_downloaded = True
            step(f"  ✓ MIB 已下载: {filename}")
        else:
            step("  (MIB 已在 case01 下载，跳过)")

    def _walk_v2c(self, port: int = 161, community: str = VALID_COMMUNITY,
                  timeout: int = 120, pre_wait: int = 25) -> dict:
        if pre_wait > 0:
            step(f"  等待 SNMP agent 重启 {pre_wait}s...")
            time.sleep(pre_wait)
        return snmp_walk_device(port=port, community=community, total_timeout=timeout)

    def _walk_v3(self, port: int, auth_protocol: str, password: str,
                 priv_protocol: str, priv_password: str = "",
                 timeout: int = 120, pre_wait: int = 25) -> dict:
        if pre_wait > 0:
            step(f"  等待 SNMP agent 重启 {pre_wait}s...")
            time.sleep(pre_wait)
        no_priv = priv_protocol.upper().replace("_", " ") in ("NONE PRIV",)
        return snmp_walk_device_v3(
            port=port,
            security_name=V3_USERNAME,
            auth_protocol=auth_protocol,
            auth_password=password,
            priv_protocol=priv_protocol,
            priv_password=priv_password,
            security_level="authNoPriv" if no_priv else "authPriv",
            total_timeout=timeout,
        )

    def _base_v2c(self, port="161", community=VALID_COMMUNITY,
                  trap_enable=False, trap_target="", buf="30", hold="0"):
        """生成 v2c 配置字典快捷方法。"""
        cfg = {
            "enable": True,
            "port": port,
            "ro_community": community,
            "trap_enable": trap_enable,
            "buffer_size": buf,
            "hold_time": hold,
        }
        if trap_target:
            cfg["trap_target_1"] = trap_target
        return cfg

    def _base_v3(self, port="161", auth="MD5", priv="NONE PRIV"):
        """生成 v3 配置字典快捷方法。"""
        return {
            "enable": True,
            "port": port,
            "username": V3_USERNAME,
            "password": V3_PASSWORD,
            "auth_protocol": auth,
            "priv_protocol": priv,
            "priv_password": V3_PRIV_PASSWORD,
        }

    # ── case01: v2c port=161 AcuRev4100 + 下载 MIB ───────────────────────────


    # ── case02: v2c port=16100 AcuRev2100 ────────────────────────────────────


    # ── case03: v2c port=16159 Acuvim3（设备不可用） ──────────────────────────


    # ── case04: v2c port=16199 AcuvimIIW（设备不可用） ───────────────────────


    # ── case05: 空 Community → NMS GET 超时 ──────────────────────────────────


    # ── case06: Community 不匹配 → NMS GET 超时 ──────────────────────────────


    # ── case07: 端口不匹配 → NMS GET 超时 ────────────────────────────────────


    # ── case08: 重启设备（需手动操作） ───────────────────────────────────────

    # ── case09: 网络异常恢复（需手动操作） ──────────────────────────────────

    # ── case10: 关闭 SNMP → 超时；重开 → 成功 ────────────────────────────────


    # ── case11: v3 MD5/NONE PRIV port=161 ────────────────────────────────────


    # ── case12: v3 MD5/DES port=16100 ────────────────────────────────────────


    # ── case13: v3 SHA/AES port=16100 ────────────────────────────────────────


    # ── case14: v3 MD5/AES port=16100 ────────────────────────────────────────


    # ── case15: v3 SHA/DES port=16100 ────────────────────────────────────────


    # ── case16: v3 凭据不匹配 → GET 超时 ─────────────────────────────────────


    # ── case17-21: Trap（需要 AcuIOM 设备和 Trap 监听服务，暂未实现） ──────────

    # ── case22: Report Buffer Size 边界 ──────────────────────────────────────


    # ── case23: Report Hold Time 边界 ─────────────────────────────────────────


    # ── case24: 非法端口验证 ──────────────────────────────────────────────────


    # ── case25: 非法 Trap Target IP ──────────────────────────────────────────


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s", "--tb=short"])

