# -*- coding: utf-8 -*-
"""
Metering 数据采集与比对自动化测试

测试流程（三步顺序执行）：
  1. 采集 AcuHMI 平台 Metering 页面数据 → metering_collect.json
  2. 匹配寄存器地址（AcuRev-4100 模板）→ metering_register_match.csv
  3. Modbus 直读 + 比对 → metering_compare.xlsx，断言 FAIL == 0

运行：
  cd C:\\JrJ\\auto\\autotest\\Protocols
  python -m pytest DataCollect -v -s
"""
from __future__ import annotations

import pathlib, sys

# ── 让本文件能 import DataCollect/metering.py ────────────────────────────────
_PROTO = pathlib.Path(__file__).resolve().parent.parent
_DC    = pathlib.Path(__file__).resolve().parent
for p in (_PROTO, _DC):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

import pytest
import metering as M
import config   # 网页 IP / 账号统一从上级 config.py 取（与 Pass Through / Device Mirror 一致）

# ╔══════════════════════════════════════════════════════╗
# ║       网页 IP/账号取自 config.py（HMI 网页地址）      ║
# ╚══════════════════════════════════════════════════════╝
BASE_URL = config.HMI_URL          # https://192.168.3.51
USERNAME = config.HMI_USERNAME
PASSWORD = config.HMI_PASSWORD
TOL_REL  = 0.01
TOL_ABS  = 0.05
# ══════════════════════════════════════════════════════════


@pytest.fixture(scope="session", autouse=True)
def apply_config():
    """将本文件的配置注入 metering 模块（覆盖模块顶部常量）。"""
    M.BASE_URL = BASE_URL
    M.USERNAME = USERNAME
    M.PASSWORD = PASSWORD
    M.TOL_REL  = TOL_REL
    M.TOL_ABS  = TOL_ABS
    yield


# ══════════════════════════════════════════════════════════════════════════════
# Fixtures — 按顺序执行三步，结果跨用例复用
# ══════════════════════════════════════════════════════════════════════════════

@pytest.fixture(scope="session")
def collected(apply_config):
    data = M.collect()
    assert M.JSON.exists(), f"采集结果文件未生成: {M.JSON}"
    assert data, "metering_collect.json 为空"
    return data


@pytest.fixture(scope="session")
def matched(collected):
    rows = M.match_registers(collected)
    assert M.CSV.exists(), f"匹配结果文件未生成: {M.CSV}"
    assert rows, "metering_register_match.csv 为空"
    return rows


@pytest.fixture(scope="session")
def compared(matched, collected):
    stats = M.compare(collected, matched)
    assert M.XLSX.exists(), f"比对报告未生成: {M.XLSX}"
    return stats


# ══════════════════════════════════════════════════════════════════════════════
# 测试用例
# ══════════════════════════════════════════════════════════════════════════════

class TestCollect:
    """Step 1：页面数据采集"""

    def test_json_generated(self, collected):
        """采集文件正常生成"""
        assert M.JSON.exists()

    def test_has_devices(self, collected):
        """至少采集到 1 台设备"""
        assert len(collected) >= 1, "未找到任何设备数据"

    def test_connection_info(self, collected):
        """每台设备均抓取到 Connection IP"""
        for name, info in collected.items():
            assert info.get("connection", {}).get("ip"), \
                f"设备 [{name}] 未抓取到 Connection IP"

    def test_metering_views(self, collected):
        """每台设备至少有 1 个 Metering 视图数据"""
        for name, info in collected.items():
            assert info.get("metering"), f"设备 [{name}] 无 Metering 视图数据"

    def test_param_count(self, collected):
        """总采集参数条目 > 0"""
        total = sum(
            len(rows)
            for info in collected.values()
            for vd in info["metering"].values()
            for rows in vd.values()
        )
        assert total > 0, "采集到的参数条目为 0"
        print(f"\n  采集参数条目合计: {total}")


class TestMatch:
    """Step 2：寄存器地址匹配"""

    def test_csv_generated(self, matched):
        """匹配文件正常生成"""
        assert M.CSV.exists()

    def test_match_count(self, matched):
        """匹配结果条目 > 0"""
        assert len(matched) > 0

    def test_exact_match_ratio(self, matched):
        """精确匹配率 ≥ 80%（不含 N/A）"""
        valid = [r for r in matched if not r.get("match", "").startswith("N/A")]
        if not valid:
            pytest.skip("无有效匹配行")
        exact = sum(1 for r in valid if r.get("match") == "精确")
        ratio = exact / len(valid)
        print(f"\n  精确匹配率: {exact}/{len(valid)} = {ratio:.1%}")
        assert ratio >= 0.80, f"精确匹配率 {ratio:.1%} 低于 80%"

    def test_no_unmatched_excess(self, matched):
        """未匹配条目 ≤ 20"""
        nomatch = [r for r in matched if r.get("match") == "未匹配"]
        print(f"\n  未匹配条目: {len(nomatch)}")
        assert len(nomatch) <= 20, (
            f"未匹配条目 {len(nomatch)} 超过阈值 20\n"
            + "\n".join(f"  {r['view']} / {r['dropdown']} / {r['parameter']} / {r['column']}"
                        for r in nomatch[:10])
        )


class TestCompare:
    """Step 3：网页值 ↔ Modbus 直读比对"""

    def test_xlsx_generated(self, compared):
        """比对报告正常生成"""
        assert M.XLSX.exists()

    def test_fail_count_zero(self, compared):
        """FAIL 条目 = 0（所有值在容差内）"""
        failed = compared["failed"]
        print(f"\n  FAIL 条目: {failed}")
        assert failed == 0, f"存在 {failed} 条 FAIL，请查看 {M.XLSX.name}"

    def test_pass_rate(self, compared):
        """有效比对通过率 ≥ 95%"""
        passed = compared["passed"]
        comparable = compared["total"] - compared["noread"] - compared["noregs"]
        if comparable == 0:
            pytest.skip("无可比对条目")
        rate = passed / comparable
        print(f"\n  通过率: {passed}/{comparable} = {rate:.1%}")
        assert rate >= 0.95, f"通过率 {rate:.1%} 低于 95%"

    def test_modbus_read_failure_zero(self, compared):
        """Modbus 读取失败条目 = 0"""
        noread = compared["noread"]
        print(f"\n  Modbus 读取失败: {noread}")
        assert noread == 0, f"存在 {noread} 条 Modbus 读取失败，请检查设备连通性"
