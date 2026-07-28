# -*- coding: utf-8 -*-
"""
XL-9600 直流电能表检定装置 二次开发协议 (V1.0) UDP 控源驱动

依据《直流电能表检定装置二次开发协议 V1.0》(深圳市星龙科技股份有限公司) 实现。

协议要点：
  - 纯文本协议，GBK 编码
  - 报文格式：
        <命令名称>
        参数名:参数值;
        参数名:参数值;
        <End>
    每个参数名后用英文冒号 ":" 间隔，参数值后用英文分号 ";" 结束，
    一个参数值内的多个信息用英文逗号 "," 间隔。
  - 应答同样以 <应答名> 开头，以 <End> 结尾；出错时为 <错误应答> + 错误:说明。
  - UDP 默认 IP 192.168.1.105，默认监听端口 24433。

一般校准流程：
  参数配置 -> (误差上报设置) -> 源输出 -> 误差读取 -> 源停止 -> 供电关闭

注意：协议文档中“供电关闭”只在流程说明里出现，未给出独立的报文格式表，
本驱动按 <供电关闭>...<End> 的通用格式实现，如厂家命令名不同请修改
POWER_OFF_CMD 常量。
"""

from __future__ import annotations

import re
import socket
import threading
from dataclasses import dataclass, field
from typing import Dict, List, Optional


ENCODING = "gbk"           # 协议规定使用 GBK 编码
LINE_END = "\r\n"          # 报文行分隔符
# 控源类命令(参数配置/源输出/源停止/供电关闭)的等待秒数。本设备这些命令
# 常不回正常应答，只有出错才“秒回”<错误应答>；故短等一下捕获错误，超时即视为下发成功。
CONTROL_TIMEOUT = 1.5
END_TAG = "<End>"
POWER_OFF_CMD = "供电关闭"  # 见上方“注意”


class XL9600Error(Exception):
    """设备返回 <错误应答> 时抛出。"""


class XL9600Timeout(Exception):
    """在超时时间内未收到完整应答。"""


# --------------------------------------------------------------------------- #
# 报文构造 / 解析
# --------------------------------------------------------------------------- #
def build_command(name: str, params: Optional[Dict[str, object]] = None) -> bytes:
    """把命令名 + 参数字典拼成符合协议的 GBK 报文。

    例：build_command("源停止") ->  <源停止>\r\n<End>\r\n
    """
    lines: List[str] = [f"<{name}>"]
    if params:
        for key, value in params.items():
            if value is None:
                continue
            lines.append(f"{key}:{value};")
    lines.append(END_TAG)
    return (LINE_END.join(lines) + LINE_END).encode(ENCODING)


def parse_response(raw: bytes) -> Dict[str, str]:
    """解析应答报文，返回一个 dict。

    返回值约定：
      "_header"  -> 尖括号里的应答名，如 "源输出应答" / "错误应答"
      其余 key   -> 协议里的“参数名:参数值”，值为去掉末尾分号的原始字符串
    遇到 <错误应答> 时抛出 XL9600Error。
    """
    text = raw.decode(ENCODING, errors="replace")
    result: Dict[str, str] = {}

    result["_raw"] = text.strip()
    header_match = re.search(r"<([^>]+)>", text)
    result["_header"] = header_match.group(1).strip() if header_match else ""

    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("<"):
            continue
        # 形如  键:值;  键、值里都不含冒号/分号
        if ":" in line:
            key, _, value = line.partition(":")
            value = value.rstrip(";；").strip()
            result[key.strip()] = value

    if result["_header"].startswith("错误") or "错误" in result:
        # 优先用设备给的“错误:说明”，没有则把原始回复整段带出来便于排查
        msg = result.get("错误") or result["_raw"] or result["_header"]
        raise XL9600Error(f"设备返回错误: {msg}")

    return result


def to_float(value: str) -> float:
    """从带单位的数值串里提取浮点数，如 '100V' -> 100.0, '0.012' -> 0.012。"""
    m = re.search(r"[-+]?\d*\.?\d+", value.replace(",", ""))
    if not m:
        raise ValueError(f"无法从 {value!r} 解析数值")
    return float(m.group(0))


# --------------------------------------------------------------------------- #
# 数据结构
# --------------------------------------------------------------------------- #
@dataclass
class SourceParams:
    """参数配置命令的参数（单位符号可带可不带）。"""
    电流接入方式: str = "间接接入式"     # 直接接入式 / 间接接入式
    供电方式: str = "电源供电"          # 电源供电 / 线路供电
    额定电压: object = "100V"           # 单位 V
    标定电流: object = "100A"           # 单位 A
    分流器额定: object = "75mV"         # 单位 mV
    被检表阻抗: object = "1000Ω"        # 单位 Ω；阻抗小时需配置，>600 可固定 1000
    脉冲常数: object = 1000             # 单位 imp/kwh
    校验圈数: object = "自动"           # 自动 / 1~100
    校验秒数: object = 1                # 校验圈数为“自动”时有效

    def as_dict(self) -> Dict[str, object]:
        return {
            "电流接入方式": self.电流接入方式,
            "供电方式": self.供电方式,
            "额定电压": self.额定电压,
            "标定电流": self.标定电流,
            "分流器额定": self.分流器额定,
            "被检表阻抗": self.被检表阻抗,
            "脉冲常数": self.脉冲常数,
            "校验圈数": self.校验圈数,
            "校验秒数": self.校验秒数,
        }


@dataclass
class OutputPoint:
    """源输出命令的参数。检定点一般为百分比，如 '100%'。"""
    电压检定点: object = "100%"
    电流检定点: object = "100%"
    电压纹波比例: object = "0%"
    电流纹波比例: object = "0%"
    电压纹波相位: object = "0度"
    电流纹波相位: object = "0度"
    纹波频率: object = "300Hz"
    电能方向: str = "正向"              # 正向 / 反向

    def as_dict(self) -> Dict[str, object]:
        return {
            "电压检定点": self.电压检定点,
            "电流检定点": self.电流检定点,
            "电压纹波比例": self.电压纹波比例,
            "电流纹波比例": self.电流纹波比例,
            "电压纹波相位": self.电压纹波相位,
            "电流纹波相位": self.电流纹波相位,
            "纹波频率": self.纹波频率,
            "电能方向": self.电能方向,
        }


@dataclass
class ErrorResult:
    """误差读取 / 日计时误差读取的结果。"""
    均值: float
    原始值: List[float] = field(default_factory=list)
    raw: Dict[str, str] = field(default_factory=dict)


# --------------------------------------------------------------------------- #
# 设备客户端
# --------------------------------------------------------------------------- #
class XL9600:
    """XL-9600 检定装置 UDP 客户端。

    用法：
        with XL9600("192.168.1.105", 24433) as dev:
            dev.config(SourceParams(额定电压="100V", 标定电流="100A"))
            dev.source_output(OutputPoint(电压检定点="100%", 电流检定点="100%"))
            res = dev.read_error(统计次数=5)
            print(res.均值, res.原始值)
            dev.source_stop()
            dev.power_off()
    """

    def __init__(
        self,
        ip: str = "192.168.1.105",
        port: int = 24433,
        timeout: float = 5.0,
        recv_buffer: int = 8192,
    ) -> None:
        self.addr = (ip, port)
        self.timeout = timeout
        self.recv_buffer = recv_buffer
        self._sock: Optional[socket.socket] = None
        # 串行化收发：保证同一时刻只有一条命令在途，避免 UDP 请求/响应错配
        self._lock = threading.Lock()

    # -- 连接管理 ----------------------------------------------------------- #
    def open(self) -> "XL9600":
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self._sock.settimeout(self.timeout)
        return self

    def close(self) -> None:
        if self._sock is not None:
            self._sock.close()
            self._sock = None

    def __enter__(self) -> "XL9600":
        return self.open()

    def __exit__(self, *exc) -> None:
        self.close()

    # -- 底层收发 ----------------------------------------------------------- #
    def _drain(self) -> None:
        """丢弃 socket 里残留的过期数据报（上一条命令超时后才到的回复等），
        避免被下一条命令误读。"""
        assert self._sock is not None
        self._sock.setblocking(False)
        try:
            while True:
                self._sock.recvfrom(self.recv_buffer)
        except (BlockingIOError, OSError):
            pass
        finally:
            self._sock.settimeout(self.timeout)

    def _send_recv(self, payload: bytes, *, timeout_ok: bool = False,
                   recv_timeout: Optional[float] = None) -> Dict[str, str]:
        """发送一条命令并收应答。

        timeout_ok=True：等待超时不算错误，返回 {}（用于本设备常不回执的控源命令，
                         超时即“已下发”；若期间收到 <错误应答> 仍会抛 XL9600Error）。
        timeout_ok=False：超时抛 XL9600Timeout（用于必须拿到数据的读命令/探测）。
        recv_timeout：本次等待秒数，None 用 self.timeout。
        """
        if self._sock is None:
            raise RuntimeError("socket 未打开，请先调用 open() 或使用 with 语句")
        # 整个“清残包 -> 发送 -> 收应答”过程加锁，确保命令严格串行
        with self._lock:
            self._drain()
            self._sock.sendto(payload, self.addr)
            if recv_timeout is not None:
                self._sock.settimeout(recv_timeout)
            try:
                data, _ = self._sock.recvfrom(self.recv_buffer)
            except socket.timeout as exc:
                if timeout_ok:
                    return {}
                raise XL9600Timeout(
                    f"等待应答超时（{recv_timeout or self.timeout}s）："
                    f"{payload.decode(ENCODING, 'replace')!r}"
                ) from exc
            finally:
                if recv_timeout is not None:
                    self._sock.settimeout(self.timeout)
            return parse_response(data)

    def send_raw(self, name: str, params: Optional[Dict[str, object]] = None, *,
                 timeout_ok: bool = False, recv_timeout: Optional[float] = None) -> Dict[str, str]:
        """直接发送任意命令，便于扩展/调试。"""
        return self._send_recv(build_command(name, params),
                               timeout_ok=timeout_ok, recv_timeout=recv_timeout)

    def _control(self, name: str, params: Optional[Dict[str, object]] = None) -> Dict[str, str]:
        """控源类命令：短等捕获 <错误应答>，超时即视为下发成功。"""
        return self.send_raw(name, params, timeout_ok=True, recv_timeout=CONTROL_TIMEOUT)

    def ping(self) -> bool:
        """探测设备是否在线（UDP 无连接，必须真发一条命令并收到应答才算通）。

        探测用“源停止”——它无参数、必有即时应答 <源停止应答>，且连接时停掉
        输出是安全默认。只要设备回了任何应答（即使是 <错误应答>）就说明可达，
        仅超时/网络不可达才判离线。

        注意：不能用“误差上报设置”做探测——它的“应答”是误差上报流，只有在
        上报:1 且正在读误差时才会回，上报:0 时设备不回任何包，必然误判离线。
        """
        try:
            self.send_raw("源停止")     # 需要应答（timeout_ok=False）：超时/网络错=离线
            return True
        except XL9600Error:
            return True            # 设备有回复（哪怕是错误应答），说明在线
        except (XL9600Timeout, OSError):
            # 超时无应答，或 Windows 发往不可达端口时 recvfrom 抛 ConnectionResetError(10054)，
            # 均视为不可达。
            return False

    # -- 协议命令 ----------------------------------------------------------- #
    def config(self, params: SourceParams) -> Dict[str, str]:
        """参数配置（控源类命令，超时视为已下发）。"""
        return self._control("参数配置", params.as_dict())

    def set_error_report(self, enable: bool) -> Dict[str, str]:
        """误差上报设置：开启后每次误差读取到有效值会自动上报。"""
        return self._control("误差上报设置", {"上报": 1 if enable else 0})

    def source_output(self, point: OutputPoint) -> Dict[str, str]:
        """源输出（控源类命令，超时视为已下发）。设备若回执则带电压/电流/功率总值等。"""
        return self._control("源输出", point.as_dict())

    def read_error(self, 统计次数: int = 5,
                   recv_timeout: Optional[float] = None) -> ErrorResult:
        """电能误差读取（统计次数 1~1000）。

        注意：设备要测完全部 统计次数 次误差才回应答，每次至少一个校验周期，
        总时长可能远超连接超时。等待时间长的场合请传 recv_timeout（秒）。
        """
        resp = self.send_raw("误差读取", {"统计次数": 统计次数},
                             recv_timeout=recv_timeout)
        return self._to_error_result(resp)

    def read_clock_error(self, 统计次数: int = 5) -> ErrorResult:
        """日计时误差读取（统计次数 1~100，默认 5）。单位为秒。

        注意：日计时误差与电能误差不能同时读取，发其一会自动停止另一个。
        """
        resp = self.send_raw("日计时误差读取", {"统计次数": 统计次数})
        return self._to_error_result(resp)

    def source_stop(self) -> Dict[str, str]:
        """源停止（停止电压/电流/小信号输出，不关闭供电电源）。控源类，超时视为已下发。"""
        return self._control("源停止")

    def power_off(self) -> Dict[str, str]:
        """供电关闭（关闭供电电源及所有输出）。命令名见 POWER_OFF_CMD。控源类，超时视为已下发。"""
        return self._control(POWER_OFF_CMD)

    # -- 上报接收 ----------------------------------------------------------- #
    def recv_report(self, timeout: Optional[float] = None) -> ErrorResult:
        """在开启误差上报后，主动接收一次自动上报的误差。"""
        if self._sock is None:
            raise RuntimeError("socket 未打开")
        old = self._sock.gettimeout()
        try:
            if timeout is not None:
                self._sock.settimeout(timeout)
            data, _ = self._sock.recvfrom(self.recv_buffer)
        except socket.timeout as exc:
            raise XL9600Timeout("等待误差上报超时") from exc
        finally:
            self._sock.settimeout(old)
        return self._to_error_result(parse_response(data))

    # -- 工具 --------------------------------------------------------------- #
    @staticmethod
    def _to_error_result(resp: Dict[str, str]) -> ErrorResult:
        mean = to_float(resp["均值"]) if "均值" in resp else float("nan")
        raws: List[float] = []
        if "原始值" in resp:
            for piece in resp["原始值"].split(","):
                piece = piece.strip()
                if piece:
                    raws.append(to_float(piece))
        # 误差上报命令返回的是单个“误差”字段
        if "误差" in resp and not raws:
            mean = to_float(resp["误差"])
        return ErrorResult(均值=mean, 原始值=raws, raw=resp)
