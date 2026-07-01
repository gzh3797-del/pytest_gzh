# -*- coding: utf-8 -*-
"""
Metering 数据采集与比对自动化测试

测试流程（三步顺序执行）：
  1. 采集 AcuHMI 平台 Metering 页面数据 → metering_collect.json
  2. 匹配寄存器地址（AcuRev-4100 模板）→ metering_register_match.csv
  3. Modbus 直读 + 比对 → metering_compare.xlsx，断言 FAIL == 0

运行（仓库根目录下，任意人 clone 后均可）：
  pytest "projects/AcuHMI_1_7/tests/DataCollect" -v -s
也可直接运行本文件： python "projects/AcuHMI_1_7/tests/DataCollect/test_metering_collect.py" -v -s
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
from playwright.sync_api import Browser
# metering.py 在本 DataCollect/ 目录、config.py 在上级 tests/ 目录，运行时均由上面的
# sys.path.insert 注入；PyCharm 静态分析看不到该动态路径，故标记 noinspection 抑制
# 「无法解析」误报（运行时导入正常）。
# noinspection PyUnresolvedReferences
import metering as M
# 配置来源：优先本地 tests/config.py（开发机覆盖，gitignored）；别人 clone 没有它时，
# 回退到框架配置（configs/.env + config.yaml），保证从仓库根直接 pytest 也能开箱即跑。
import importlib.util as _ilu, types as _types

def _load_local_config():
    _cfg_path = pathlib.Path(__file__).resolve().parent.parent / "config.py"
    if _cfg_path.exists():
        _spec = _ilu.spec_from_file_location("_tests_config", _cfg_path)
        _mod = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_mod)
        return _mod
    _RR = pathlib.Path(__file__).resolve().parents[4]
    if str(_RR) not in sys.path:
        sys.path.insert(0, str(_RR))
    from projects.AcuHMI_1_7 import settings as _s
    return _types.SimpleNamespace(
        HMI_URL=_s.HMI_URL, HMI_USERNAME=_s.HMI_USERNAME, HMI_PASSWORD=_s.HMI_PASSWORD,
    )

config = _load_local_config()

# ╔══════════════════════════════════════════════════════╗
# ║       网页 IP/账号取自 config.py（HMI 网页地址）      ║
# ╚══════════════════════════════════════════════════════╝
BASE_URL = config.HMI_URL          # https://192.168.3.51
USERNAME = config.HMI_USERNAME
PASSWORD = config.HMI_PASSWORD
TOL_REL  = 0.01
TOL_ABS  = 0.05
# ══════════════════════════════════════════════════════════


@pytest.fixture(scope="module", autouse=True)
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

@pytest.fixture(scope="module")
def dc_page(browser: Browser):
    """DataCollect 专用 page：复用项目级共享 browser，避免重复启动 playwright 实例。"""
    ctx = browser.new_context(ignore_https_errors=True)
    pg = ctx.new_page()
    yield pg
    ctx.close()


@pytest.fixture(scope="module")
def collected(apply_config, dc_page):
    data = M.collect(page=dc_page)
    assert M.JSON.exists(), f"采集结果文件未生成: {M.JSON}"
    assert data, "metering_collect.json 为空"
    return data


@pytest.fixture(scope="module")
def matched(collected):
    rows = M.match_registers(collected)
    assert M.CSV.exists(), f"匹配结果文件未生成: {M.CSV}"
    assert rows, "metering_register_match.csv 为空"
    return rows


@pytest.fixture(scope="module")
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
        """稳定量/其他 FAIL = 0（波动量差异单独报告，不计入失败）。

        与 Device Mirror / Pass Through 一致：稳定量（电压/频率/电能）与其他量要求网页值与
        Modbus 直读值在容差内一致；波动量（谐波/THD/相位角/电流/功率/不平衡度等）在两次取值
        之间本身会变化，差异仅作 warning 报告，不计入 FAIL。"""
        failed_stable = compared.get("failed_stable", compared["failed"])
        failed_dynamic = compared.get("failed_dynamic", 0)
        print(f"\n  FAIL(稳定量/其他): {failed_stable}  |  波动量差异(仅报告): {failed_dynamic}")
        if failed_dynamic:
            import warnings
            warnings.warn(f"波动量 网页↔Modbus 差异 {failed_dynamic} 项（不计入失败，详见 {M.XLSX.name}）")
        assert failed_stable == 0, f"存在 {failed_stable} 条稳定量/其他 FAIL，请查看 {M.XLSX.name}"

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


if __name__ == "__main__":
    # 直接运行：python "DataCollect/test_metering_collect.py" [额外 pytest 参数]
    # 用 --confcutdir 限制 pytest 只从本套件目录起加载 conftest，绕开重组后因目录大小写
    # （磁盘 acuhmi_1_7 vs 代码引用 AcuHMI_1_7）而 import 失败的上级 conftest。
    # 本套件不依赖上级 conftest 的 fixture，仅用本目录 conftest 与 config.py。
    _f = str(pathlib.Path(__file__).resolve())
    _here = str(pathlib.Path(__file__).resolve().parent)
    raise SystemExit(pytest.main([_f, "--confcutdir", _here, *sys.argv[1:]]))
