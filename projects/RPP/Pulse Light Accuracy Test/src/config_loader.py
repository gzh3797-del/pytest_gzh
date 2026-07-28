"""config_loader.py — YAML config loader and device client factory.

Load config.yaml, validate required top-level sections, and build the three
device clients (CL3021Client, SDGClient, MeterClient) from it.

Usage:
    from src.config_loader import load_config, build_source, build_counter, build_meter

    cfg = load_config("config.yaml")
    source  = build_source(cfg)
    counter = build_counter(cfg)
    meter   = build_meter(cfg)
"""

import yaml

from src.cl3021_transport import TcpTransport, SerialTransport
from src.cl3021_client import CL3021Client
from src.transport import SocketTransport, VisaTransport, list_usb_resources
from src.sdg_client import SDGClient
from src.meter_client import MeterClient

_REQUIRED_SECTIONS = ("source", "counter", "meter", "test")


# ---------------------------------------------------------------------------
# load_config
# ---------------------------------------------------------------------------

def load_config(path: str) -> dict:
    """Load *path* as YAML and return the top-level dict.

    Raises
    ------
    ValueError
        If the file cannot be read, is not valid YAML, or any of the required
        top-level sections (source / counter / meter / test) is missing.
    """
    try:
        with open(path, encoding="utf-8") as fh:
            data = yaml.safe_load(fh)
    except OSError as exc:
        raise ValueError(f"config 文件无法读取: {exc}") from exc
    except yaml.YAMLError as exc:
        raise ValueError(f"config YAML 解析失败: {exc}") from exc

    if not isinstance(data, dict):
        raise ValueError("config.yaml 顶层必须是 YAML 映射（dict）")

    missing = [k for k in _REQUIRED_SECTIONS if k not in data]
    if missing:
        raise ValueError(
            f"config.yaml 缺少必需的顶层键: {', '.join(missing)}"
        )

    return data


# ---------------------------------------------------------------------------
# resolve_usb_resource
# ---------------------------------------------------------------------------

_SIGLENT_VENDOR = "0xf4ec"  # case-insensitive match target


def resolve_usb_resource(resource: str, available=None) -> str:
    """Return the VISA resource string to use for the USB counter.

    Parameters
    ----------
    resource:
        The ``counter.resource`` value from config.  Pass ``"auto"`` to
        auto-detect; anything else is returned as-is.
    available:
        Optional list of VISA resource strings to search.  When *None* and
        *resource* is ``"auto"``, ``list_usb_resources()`` is called.

    Raises
    ------
    ValueError
        When ``resource == "auto"`` and no USB devices are found.
    """
    if resource != "auto":
        return resource

    if available is None:
        available = list_usb_resources()

    if not available:
        raise ValueError(
            "未发现 USB 设备。请确认 SDG 已连接并已安装 USBTMC 驱动 / NI-VISA。"
        )

    # Prefer Siglent (vendor 0xF4EC); fall back to first entry
    for r in available:
        if _SIGLENT_VENDOR in r.lower():
            return r

    return available[0]


# ---------------------------------------------------------------------------
# Client factories
# ---------------------------------------------------------------------------

def build_source(cfg: dict) -> CL3021Client:
    """Build a CL3021Client from ``cfg['source']``.

    Supports ``mode: serial`` and ``mode: tcp``.
    """
    src = cfg["source"]
    mode = src["mode"]
    if mode == "serial":
        transport = SerialTransport(src["com"], src["baud"])
    elif mode == "tcp":
        transport = TcpTransport(src["host"], src["port"])
    else:
        raise ValueError(f"source.mode 必须是 'serial' 或 'tcp'，当前值: {mode!r}")
    return CL3021Client(transport)


def build_counter(cfg: dict, available=None) -> SDGClient:
    """Build an SDGClient from ``cfg['counter']``.

    Supports ``mode: lan`` (SocketTransport) and ``mode: usb`` (VisaTransport).
    Does NOT open/connect the transport.
    """
    ctr = cfg["counter"]
    mode = ctr["mode"]
    if mode == "lan":
        transport = SocketTransport(ctr["host"], ctr["port"])
    elif mode == "usb":
        resource = resolve_usb_resource(ctr["resource"], available)
        transport = VisaTransport(resource)
    else:
        raise ValueError(f"counter.mode 必须是 'lan' 或 'usb'，当前值: {mode!r}")
    return SDGClient(transport)


def build_meter(cfg: dict) -> MeterClient:
    """Build a MeterClient from ``cfg['meter']``.

    Passes whatever keys are present; MeterClient tolerates extra ``None``
    values.
    """
    m = cfg["meter"]
    return MeterClient(
        m["mode"],
        host=m.get("host"),
        port=m.get("port", 502),
        com=m.get("com"),
        baud=m.get("baud", 9600),
        slave=m.get("slave", 1),
    )


# ---------------------------------------------------------------------------
# Config sub-section accessors
# ---------------------------------------------------------------------------

def pulse_cfg(cfg: dict) -> dict:
    """Return ``cfg['meter']['pulse_constant']`` (register/dtype/word_order/scale)."""
    return cfg["meter"]["pulse_constant"]


def wiring_cfg(cfg: dict) -> dict:
    """Return ``cfg['meter']['wiring']`` (register/dtype/map)."""
    return cfg["meter"]["wiring"]


def get_test_params(cfg: dict) -> dict:
    """Return ``cfg['test']`` (settle_s/n_periods/freq/pf_lagging)."""
    return cfg["test"]
