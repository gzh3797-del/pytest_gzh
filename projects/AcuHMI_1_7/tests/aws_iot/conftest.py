"""
pytest fixtures — AWS IoT 测试套件（AcuHMI-1-7 Playwright 版）

用法:
    pytest tests/protocols/aws_iot/ -m aws_iot -v
    pytest tests/protocols/aws_iot/ -m "aws_iot and not slow" -v
"""

import logging
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pytest
import yaml

_THIS_DIR    = Path(__file__).resolve().parent          # tests/aws_iot/
_PROJECT_ROOT = _THIS_DIR.parent.parent                  # AcuHMI_1_7/
_REPO_ROOT   = _PROJECT_ROOT.parent.parent               # autotest/

sys.path.insert(0, str(_PROJECT_ROOT))
# 仓库根也需在 sys.path：pages/base_page.py 等用 projects.AcuHMI_1_7.* 绝对包路径导入，
# 单独跑本目录时 rootdir 被本地 pytest.ini 截断，不加这行会 ModuleNotFoundError: projects
sys.path.insert(0, str(_REPO_ROOT))

CONFIG_PATH = _THIS_DIR / "config.yaml"

log = logging.getLogger(__name__)


# ─── Allure 报告自动生成（每次 pytest 结束后生成带时间戳的报告目录）────────────────

def pytest_collection_modifyitems(items):
    """让 004_009（无参数无上报）永远排在第一位执行，避免污染其他测试的会话状态。"""
    first = [i for i in items if "004_009" in i.nodeid]
    rest  = [i for i in items if "004_009" not in i.nodeid]
    items[:] = first + rest


def pytest_sessionfinish(session, exitstatus):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    results_dir = _THIS_DIR / "reports" / "allure-results"
    report_dir  = _THIS_DIR / "reports" / f"allure-report-{ts}"
    allure_bat  = Path(r"C:\work\tools\allure\allure-2.32.0\bin\allure.bat")
    if results_dir.exists() and allure_bat.exists():
        subprocess.run(
            [str(allure_bat), "generate", str(results_dir),
             "-o", str(report_dir), "--clean"],
            check=False,
        )
        log.info("Allure 报告已生成：%s", report_dir)


# ─── 共享配置 ──────────────────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def aws_cfg():
    return yaml.safe_load(CONFIG_PATH.read_text(encoding="utf-8"))


# ─── 动态设备跳过（读模块 DEVICE_NAME，检查 modbus.tcp.devices 和 virtual_device）─

def _device_is_configured(device_name: str, cfg: dict) -> bool:
    """检查 device_name 是否在 modbus.tcp.devices 或 virtual_device 中已配置。"""
    tcp_keys = cfg.get("modbus", {}).get("tcp", {}).get("devices", {}).keys()
    dn = device_name.lower()
    if any(dn in k.lower() or k.lower() in dn for k in tcp_keys):
        return True
    vd = (cfg.get("aws_iot") or {}).get("virtual_device", "")
    return bool(vd and (dn in vd.lower() or vd.lower() in dn))


@pytest.fixture(autouse=True)
def skip_if_device_not_configured(request, aws_cfg):
    """
    若测试模块定义了 DEVICE_NAME，且该设备在 modbus.tcp.devices / virtual_device
    中未配置，则自动跳过该用例，无需在测试文件里写 @pytest.mark.skip。
    """
    module = sys.modules.get(request.module.__name__)
    device_name = getattr(module, "DEVICE_NAME", None)
    if device_name and not _device_is_configured(device_name, aws_cfg):
        pytest.skip(f"设备 '{device_name}' 未在 config.yaml 中配置，跳过")


# ─── 浏览器 fixture（复用唯一 sync_playwright 实例，与 data_log 同模式）─────────

@pytest.fixture(scope="session")
def app_page(playwright, aws_cfg):
    """Session 级已登录 Playwright Page（复用 session 级唯一 ``playwright`` 实例）。

    不再自起 ``sync_playwright().start()``：全量套件运行时主线程已有共享实例留下的
    运行中事件循环，第二个 sync_playwright 会撞上它抛 "Playwright Sync API inside
    the asyncio loop"。改为从 session 级 ``playwright`` 实例（全量跑=项目根 conftest，
    单独跑=pytest-playwright 插件）另起独立浏览器，保留 aws_iot 自己的有头/最大化/
    忽略证书启动参数与生命周期。
    """
    gw = aws_cfg.get("gateway", {})
    url      = gw.get("url", "https://192.168.2.8")
    username = gw.get("username", "admin")
    password = gw.get("password", "Admin@110002")

    browser = playwright.chromium.launch(
        headless=False,
        args=["--ignore-certificate-errors", "--start-maximized"],
    )
    ctx = browser.new_context(ignore_https_errors=True, no_viewport=True)
    page = ctx.new_page()
    page.goto(url, wait_until="domcontentloaded")
    try:
        user_input = page.locator("input[placeholder='Enter User Name']").first
        user_input.wait_for(state="visible", timeout=10000)
        user_input.fill(username)
        page.locator("input[placeholder='Enter Password']").first.fill(password)
        page.locator("xpath=//button[span[text()='Sign In']]").first.click()
        # 等登录表单消失（Sign In 成功后表单 hidden），兼容 hash 路由（无浏览器导航事件）
        try:
            page.wait_for_selector(
                "input[placeholder='Enter User Name']",
                state="hidden",
                timeout=10000,
            )
        except Exception:
            pass  # 超时兜底，继续后续弹框处理
        # 登录成功后可能弹出 Yes / Continue / Cancel 提示框，出现就立即关闭
        _DISMISS = [
            "xpath=//button[.//span[normalize-space(.)='Yes']]",
            "xpath=//button[.//span[normalize-space(.)='Continue']]",
            "xpath=//button[.//span[normalize-space(.)='Cancel']]",
        ]
        import time as _time
        _deadline = _time.time() + 3
        while _time.time() < _deadline:
            page.wait_for_timeout(200)
            _hit = False
            for _sel in _DISMISS:
                try:
                    _btn = page.locator(_sel).first
                    if _btn.is_visible():
                        _btn.click()
                        log.info("app_page：已关闭登录后弹框")
                        _hit = True
                        # 关闭后只再等 0.5s 确认无第二个弹框，不撑满 3s
                        _deadline = min(_deadline, _time.time() + 0.5)
                        break
                except Exception:
                    pass
            if not _hit:
                break
        log.info("app_page：登录完成")
    except Exception as e:
        log.warning("app_page：登录操作异常（可能已登录）：%s", e)
    # 若停在密码/用户管理页，导航回主页
    try:
        if any(k in page.url for k in ("passwordManagement", "userManagement")):
            log.info("app_page：检测到密码管理页，导航回主页")
            page.goto(url, wait_until="domcontentloaded")
    except Exception:
        pass
    yield page
    try:
        ctx.close()
        browser.close()   # playwright 实例归 session 级 fixture 管理，此处只关自己的浏览器
    except Exception:
        pass


# ─── Session 级设备扫描与校验（在所有用例之前自动执行一次）─────────────────────

@pytest.fixture(scope="session", autouse=True)
def setup_device_config(app_page, aws_cfg):
    """
    Session 启动时自动执行：
    1. 导航到 AWS IoT 页面，扫描 Device Selection 表格中的完整设备列表
    2. 将所有设备名写入 config.yaml 的 aws_iot.expected_devices
    3. 对比 modbus.tcp.devices：若有非虚拟设备缺少配置，终止测试并打印提示
    """
    from pages.protocols.aws_iot_page import AWSIoTPage

    gw = aws_cfg.get("gateway", {})
    device_name_str = gw.get("device_name", "AcuHMI-1-7")

    scanner = AWSIoTPage(
        app_page,
        device_name=device_name_str,
        gateway_url=gw.get("url", ""),
        username=gw.get("username", "admin"),
        password=gw.get("password", "Admin@110001"),
    )
    scanner.navigate_to_aws_iot()

    # 设备表格仅在 Enable 状态下可见；扫描后恢复原状
    was_enabled = scanner.is_enabled()
    if not was_enabled:
        scanner.ensure_enabled()

    try:
        devices = scanner.get_all_devices_from_table()
    finally:
        if not was_enabled:
            try:
                scanner.disable()
            except Exception:
                pass

    if not devices:
        log.warning("setup_device_config: 未扫描到任何设备，跳过校验")
        return

    physical = [d["name"] for d in devices if not d["is_virtual"]]
    virtual  = [d["name"] for d in devices if d["is_virtual"]]
    all_names = [d["name"] for d in devices]

    log.info("Setup — 物理设备：%s", physical)
    log.info("Setup — 虚拟设备：%s", virtual)
    print(f"\n[Setup] 物理设备：{physical}", flush=True)
    print(f"[Setup] 虚拟设备：{virtual}", flush=True)

    # 写回 config.yaml expected_devices
    try:
        cfg_data = yaml.safe_load(CONFIG_PATH.read_text(encoding="utf-8"))
        cfg_data.setdefault("aws_iot", {})["expected_devices"] = all_names
        CONFIG_PATH.write_text(
            yaml.dump(cfg_data, allow_unicode=True, default_flow_style=False, sort_keys=False),
            encoding="utf-8",
        )
        aws_cfg.setdefault("aws_iot", {})["expected_devices"] = all_names
        log.info("Setup 已更新 config.yaml expected_devices → %s", all_names)
        print(f"[Setup] 已更新 config.yaml expected_devices：{all_names}", flush=True)
    except Exception as exc:
        log.warning("Setup 更新 config.yaml expected_devices 失败：%s", exc)

    # 校验 modbus.tcp.devices 覆盖（虚拟设备不要求配置）
    tcp_devices = aws_cfg.get("modbus", {}).get("tcp", {}).get("devices", {})
    missing = [dev for dev in physical if dev not in tcp_devices]

    if missing:
        lines = [
            "",
            "=" * 64,
            "[Setup 失败] 以下物理设备未在 config.yaml modbus.tcp.devices 中配置：",
        ]
        for dev in missing:
            lines.append(f"  - {dev}")
        lines += [
            "",
            "请在 config.yaml 的 modbus.tcp.devices 下补充，示例：",
            f"  {missing[0]}: {{ip: \"192.168.x.x\", port: 502, unit: 1}}",
            "=" * 64,
        ]
        pytest.exit("\n".join(lines), returncode=1)


# ─── 函数级 fixture（每条用例独立，带 navigate + teardown disable）───────────────

@pytest.fixture(scope="function")
def aws_page(app_page, aws_cfg):
    """
    函数级 AWSIoTPage fixture：
    - Setup：导航到 AWS IoT 配置页面
    - 用例结束后 teardown：disable AWS IoT（防止残留数据污染）
    """
    from pages.protocols.aws_iot_page import AWSIoTPage
    gw = aws_cfg.get("gateway", {})
    _device_name = gw.get("device_name", "AcuHMI-1-7")
    page = AWSIoTPage(
        app_page,
        device_name=_device_name,
        gateway_url=gw.get("url", ""),
        username=gw.get("username", "admin"),
        password=gw.get("password", "Admin@110001"),
    )
    page.navigate_to_aws_iot()
    # 每个函数级用例从干净的 Disable 状态开始（防止前序用例残留 Enable）
    try:
        if page.is_enabled():
            page.disable()
    except Exception:
        pass
    yield page
    try:
        page.navigate_to_aws_iot()
        if page.is_enabled():
            page.disable()
    except Exception:
        pass


# ─── Session 级 fixture（整个 session 共用一个浏览器，用于 Interval 测试）─────────

@pytest.fixture(scope="session")
def aws_session_page(app_page, aws_cfg):
    """
    Session 级 AWSIoTPage fixture：
    - 整个 session 共享同一 Playwright 页面实例
    - Setup：导航 + 完整初始化配置（连接参数、证书、设备选择）
    - Session 结束后 teardown：disable AWS IoT
    """
    from pages.protocols.aws_iot_page import AWSIoTPage
    aws = aws_cfg["aws_iot"]
    gw = aws_cfg.get("gateway", {})
    _device_name = gw.get("device_name", "AcuHMI-1-7")

    page = AWSIoTPage(
        app_page,
        device_name=_device_name,
        gateway_url=gw.get("url", ""),
        username=gw.get("username", "admin"),
        password=gw.get("password", "Admin@110001"),
    )
    page.navigate_to_aws_iot()

    page.ensure_enabled()
    page.set_client_id(aws.get("client_id", ""))
    page.set_url(aws["url"])
    page.set_topic(aws["topic"])
    page.set_interval(aws.get("interval", "60 seconds"))
    page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
    page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))

    page.select_only_device()
    page.configure_all_devices_parameters(checked_only=True)
    page.save()

    yield page

    try:
        page.navigate_to_aws_iot()
        page.disable()
    except Exception:
        pass
