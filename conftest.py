"""仓库根 conftest：仅保留跨项目通用的 fixture（计时 / 异常兜底 / Modbus 连接）。

历史上承载的 test_case/IOM 数据驱动 fixture 与 Selenium driver fixture 已随 test_case/
移除；项目级 fixture 请放各 projects/<项目>/conftest.py。
"""
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import pytest

from comm.modbus_rtu_tcp import ModbusRtuOrTcp
from comm.source_control import close_dc_all
from framework.report.paths import detect_module, make_report_dir, update_latest

# 仓库根加入 sys.path，便于 shared.* / framework.* 顶层导入
_PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if _PROJECT_ROOT not in sys.path:
    sys.path.append(_PROJECT_ROOT)


# ── 报告路径归一（全项目通用）────────────────────────────────────────────────
# 直接 `pytest --html=reports/xxx.html` 时，相对路径会落到 reports/ 根目录，污染其他项目。
# 本钩子从命令行测试路径中识别 projects/<项目名>/ 段，把 --html 报告（无论传裸文件名还是
# 相对路径，只取文件名）统一重定向到 reports/<项目名>/<时间戳>/html/ 下，与 framework
# runner 的目录约定保持一致。仅当能唯一确定项目时才重定向；跨项目或无法识别时保持原样。
_PROJECT_PATH_RE = re.compile(r"(?:^|[\\/])projects[\\/]([^\\/]+)")


def _detect_project(config: pytest.Config) -> str | None:
    """从命令行测试路径参数中唯一识别项目名；无法唯一确定时返回 None。"""
    found = set()
    for arg in config.args:
        path_part = arg.split("::", 1)[0].replace("\\", "/")
        match = _PROJECT_PATH_RE.search(path_part)
        if match:
            found.add(match.group(1))
    return next(iter(found)) if len(found) == 1 else None


@pytest.hookimpl(tryfirst=True)
def pytest_configure(config: pytest.Config) -> None:
    """裸 pytest 版「运行目录 owner」：识别项目后建立单一 run 目录，并把 HTML 报告、
    截图、日志统一归到 reports/<项目>/<时间戳>/ 下。

    仅在传了 --html（即本次确有报告产出）时介入，避免凭空创建空 run 目录。注入的环境
    变量与 framework runner 完全一致（PROJECT/RUN_TS/REPORT_DIR/SCREENSHOT_DIR），
    用 setdefault 保证 runner 已设的值不被覆盖，两条入口对称。
    """
    htmlpath = getattr(config.option, "htmlpath", None)
    if not htmlpath:
        return
    project = _detect_project(config)
    if project is None:
        return
    # framework runner 已设置 RUN_TS 时复用其时间戳，保证单次运行报告同目录。
    run_ts = os.environ.get("RUN_TS") or datetime.now().strftime("%Y%m%d_%H%M%S")
    # 从测试路径中识别 tests/ 下的单一模块目录，与 framework runner 的目录约定保持一致；
    # 整项目/跨模块运行时为 None，报告退化为不含模块层级的旧结构。
    module = detect_module(config.args)
    dirs = make_report_dir(project, run_ts, module=module)
    os.environ.setdefault("PROJECT", project)
    os.environ.setdefault("RUN_TS", run_ts)
    os.environ.setdefault("REPORT_DIR", str(dirs.root))
    os.environ.setdefault("SCREENSHOT_DIR", str(dirs.screenshots))
    config.option.htmlpath = str(dirs.root / Path(htmlpath).name)
    update_latest(dirs.root)


# ── 报告渲染兼容补丁（全项目通用）────────────────────────────────────────────
# 现象：双击 report.html（file:// 协议）从本仓库这种含非 ASCII 字符的路径
# （C:\AI工具\工程优化\...）打开时，整张报告只剩表头，Environment / 用例表格 /
# 失败详情（traceback、用例步骤、截图）全部空白——看起来"报告没有详细信息"。
# 根因：pytest-html v4 的渲染脚本在排序/筛选时调用 history.pushState 写入
# `?sort=` 之类的 URL；file:// 文档 origin 为 'null'，叠加路径里的非 ASCII 字符，
# 浏览器抛 SecurityError，异常中断了后续把数据从内嵌 JSON 填进表格的脚本。
# 数据本身完整写在文件里，只是 JS 在画表格前就挂了。
# 修复：报告写盘后，在 <head> 顶部注入一段把 history.pushState/replaceState 包进
# try/catch 的守卫脚本（作为普通 <script> 先于 pytest-html 的 module 脚本执行），
# pushState 抛错被吞掉，渲染脚本得以继续——双击中文路径打开也能正常展开详情。
# 守卫在 pushState 成功时是无害的透传，故所有报告一律注入，不做路径判断。
_PUSHSTATE_GUARD = (
    "<script>/*phtml-pushstate-guard*/(function(){var a=['pushState','replaceState'];"
    "for(var i=0;i<a.length;i++){(function(f){var o=history[f]&&history[f].bind(history);"
    "if(o){history[f]=function(){try{return o.apply(history,arguments)}catch(e){}}}})(a[i])}})();</script>"
)


@pytest.hookimpl(trylast=True)
def pytest_unconfigure(config: pytest.Config) -> None:
    """在 pytest-html 把报告写盘后注入 pushState 守卫，修复非 ASCII 路径下 file://
    打开报告时表格/详情空白的问题。trylast 保证晚于 pytest-html 的 sessionfinish 写盘。"""
    htmlpath = getattr(config.option, "htmlpath", None)
    if not htmlpath:
        return
    report = Path(htmlpath)
    if not report.is_file():
        return
    try:
        text = report.read_text(encoding="utf-8")
    except OSError:
        return
    if "phtml-pushstate-guard" in text or "<head>" not in text:
        return  # 已注入或结构不符，跳过
    report.write_text(text.replace("<head>", "<head>\n  " + _PUSHSTATE_GUARD, 1), encoding="utf-8")


# ── Playwright 浏览器自检（仅 UI 测试会话触发）──────────────────────────────
# 现象：某个 venv 升级 playwright 包后忘了同步装浏览器，UI 用例启动时报
# "Executable doesn't exist at ...chromium-XXXX"，堆栈晦涩且反复出现——本机有多个
# venv，各自要求的 chromium 版本不同，而浏览器是全局装到 %LOCALAPPDATA%\ms-playwright，
# 升级包后没补装就对不上号。
# 本钩子在收集到 UI 用例（nodeid 含 /ui/）时，用当前解释器解析 playwright 期望的
# chromium 可执行路径并校验存在性；缺失则 fail fast，直接给出含本 venv python 的精确
# 安装命令，把晦涩堆栈变成可照抄的一行命令。非 UI 会话（纯协议/Modbus）零开销跳过；
# 未装 playwright 包的环境也跳过。
def pytest_collection_modifyitems(config: pytest.Config, items: list) -> None:
    if not any("/ui/" in item.nodeid.replace("\\", "/") for item in items):
        return
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return
    playwright = sync_playwright().start()
    try:
        chromium_exe = playwright.chromium.executable_path
    finally:
        playwright.stop()
    if not os.path.exists(chromium_exe):
        raise pytest.UsageError(
            "Playwright 浏览器缺失：当前解释器期望的 chromium 不存在。\n"
            f"    缺失路径: {chromium_exe}\n"
            f"    解释器:   {sys.executable}\n"
            "  请在该解释器对应的 venv 下补装匹配版本的浏览器：\n"
            f'    "{sys.executable}" -m playwright install chromium'
        )


# ================= 计时功能 Fixture ================= #
@pytest.fixture(scope="function")
def timer(request):
    """在函数执行前开始计时，执行完结束计时。"""
    test_name = request.node.name
    start_time = time.time()
    logging.info(f"测试函数 [{test_name}] 开始执行")
    yield
    execution_time = time.time() - start_time
    logging.info(f"测试函数 [{test_name}] 执行完成，耗时: {execution_time:.4f} 秒")


# ================= 异常处理 Fixture ================= #
@pytest.fixture(scope="function")
def test_error_handler():
    """统一处理测试函数异常，结束后关闭电源输出。"""
    try:
        yield
    except KeyboardInterrupt:
        logging.info("程序被用户中断，关闭电源输出")
    except Exception as exc:
        logging.error("程序异常终止，关闭电源输出")
        logging.error(str(exc))
    finally:
        logging.info("程序执行完毕，关闭电源输出")
        close_dc_all()


# ================= 统一连接对象 Fixture ================= #
@pytest.fixture(scope="function")
def modbus_client():
    """提供一个 Modbus 连接实例（函数级别）。"""
    client = None
    try:
        client = ModbusRtuOrTcp()
        yield client
    except Exception as exc:
        logging.error(f"Modbus 客户端初始化失败: {str(exc)}")
        raise
    finally:
        if client:
            try:
                client.close()
                time.sleep(0.5)  # 延长等待时间确保 Windows 释放串口资源
                logging.info("Modbus 客户端已关闭")
            except Exception as exc:
                logging.error(f"关闭 Modbus 客户端时出错: {str(exc)}")
