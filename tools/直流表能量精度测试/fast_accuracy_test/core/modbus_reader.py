import logging
import time
from pymodbus.client import ModbusSerialClient, ModbusTcpClient

from .codec import decode_float, decode_double, decode_u32, encode_u32
from .registers import regmap, pulse_const_reg

_COUNT = {"float": 2, "double": 4}
WRITE_SETTLE_S = 2.0   # 写脉冲常数后等表应用并恢复响应的秒数(避免紧接着的读超时→Modbus失步)


def make_client(cfg):
    if cfg["conn_mode"] == "rtu":
        rtu = cfg["rtu"]
        client = ModbusSerialClient(port=rtu["port"], baudrate=rtu["baudrate"],
                                    parity=rtu["parity"], bytesize=8, stopbits=1,
                                    timeout=1)
    else:
        tcp = cfg["tcp"]
        client = ModbusTcpClient(host=tcp["ip"], port=tcp["port"], timeout=1)
    if not client.connect():
        raise IOError(f"无法建立 Modbus 连接: conn_mode={cfg['conn_mode']}")
    return client


class ModbusReader:
    def __init__(self, cfg, client):
        self.cfg = cfg
        self.client = client
        self.word_order = cfg.get("word_order", "big")
        self.map = regmap(cfg["is_dual"])
        seg = cfg["rtu"] if cfg["conn_mode"] == "rtu" else cfg["tcp"]
        self.slave_id = seg.get("slaveid", 1)
        self.retries = max(1, int(self.cfg.get("read_retries", 3)))

    def read(self, name):
        addr, kind = self.map[name]
        count = _COUNT[kind]
        last = None
        for _ in range(self.retries):
            try:
                resp = self.client.read_holding_registers(addr, count=count, slave=self.slave_id)
                if hasattr(resp, "isError") and not resp.isError() and len(resp.registers) >= count:
                    words = resp.registers[:count]
                    val = decode_float(words, self.word_order) if kind == "float" \
                        else decode_double(words, self.word_order)
                    logging.info("电表读 %s@%d(0x%04X) words=%s → %s", name, addr, addr, words, val)
                    return val
                last = f"resp={resp}"
            except Exception as e:
                last = repr(e)
            self._flush()   # 本次失败后清缓冲再重试，防止迟到应答导致后续读失步
        raise IOError(f"读取 {name}@{addr} 失败(重试{self.retries}次): {last}")

    def read_many(self, names):
        return {n: self.read(n) for n in names}

    def _flush(self):
        """清空 Modbus 串口输入缓冲，丢弃上一笔超时的迟到应答，防止后续请求收到错位应答(失步)。
        best-effort：拿不到底层串口或 TCP 模式则忽略；返回是否实际清了。"""
        try:
            sock = getattr(self.client, "socket", None)
            if sock is not None and hasattr(sock, "reset_input_buffer"):
                sock.reset_input_buffer()
                return True
        except Exception as e:
            logging.debug("清缓冲失败: %s", e)
        return False

    def write_pulse_const(self, const):
        """把脉冲常数写进电表对应寄存器(按 device_model 选寄存器+换算)，与发给源的同一个物理常数。
        写入值 = const × 系数(32位整数, 占2个寄存器, 字序按 word_order)。
        超量程则不写并告警(防写非法值)。成功返回 True，否则 False。
        注：已确认该寄存器修改后【立即生效】，无需重启/保存。"""
        model = self.cfg.get("device_model")
        info = pulse_const_reg(model)
        if info is None:
            logging.warning("write_pulse_const: 型号 %r 无脉冲常数寄存器映射，跳过", model)
            return False
        addr, scale, max_const = info
        if const is None or const <= 0:
            logging.warning("write_pulse_const: 常数无效 %r，跳过", const)
            return False
        if const > max_const:
            logging.error("write_pulse_const: 常数 %s 超过型号 %s 上限 %s，不写(防止写非法值)",
                          const, model, max_const)
            return False
        reg_val = int(round(const * scale))
        words = encode_u32(reg_val, self.word_order)
        hexs = " ".join("%04X" % w for w in words)
        logging.info("电表 -> 写脉冲常数: 型号=%s 物理常数=%s × 系数%d = %d → 寄存器%d(0x%04X) "
                     "words=%s hex=[%s] 字序=%s slave=%s",
                     model, const, scale, reg_val, addr, addr, words, hexs, self.word_order, self.slave_id)
        last = None
        for attempt in range(1, self.retries + 1):
            try:
                resp = self.client.write_registers(addr, words, slave=self.slave_id)
                if hasattr(resp, "isError") and not resp.isError():
                    logging.info("电表 <- 写脉冲常数成功 (第%d次尝试, resp=%s)", attempt, resp)
                    # 关键：写后表要花时间应用常数，期间不响应读。先等一下，再读回校验，再清串口缓冲——
                    # 否则紧接着的读会超时，迟到应答留在缓冲里把后续 V/I/P 读冲乱(Modbus 失步)。
                    time.sleep(WRITE_SETTLE_S)
                    self._verify_pulse_const(addr, reg_val)
                    flushed = self._flush()
                    logging.info("电表 写后清串口缓冲(防失步): %s", "成功" if flushed else "未执行(非串口/不支持)")
                    return True
                last = "resp=%s" % resp
                logging.warning("电表 写脉冲常数 第%d次返回错误: %s", attempt, last)
            except Exception as e:
                last = repr(e)
                logging.warning("电表 写脉冲常数 第%d次异常: %s", attempt, last)
        logging.error("write_pulse_const 失败(重试%d次): 寄存器%d, %s", self.retries, addr, last)
        return False

    def _verify_pulse_const(self, addr, expect):
        """写后读回脉冲常数寄存器，日志记录是否一致(只读校验，失败不影响主流程)。"""
        try:
            rb = self.client.read_holding_registers(addr, count=2, slave=self.slave_id)
            if hasattr(rb, "isError") and not rb.isError() and len(rb.registers) >= 2:
                words = rb.registers[:2]
                got = decode_u32(words, self.word_order)
                hexs = " ".join("%04X" % w for w in words)
                logging.info("电表 <- 脉冲常数读回: 寄存器%d words=%s hex=[%s] → 值=%d (期望 %d, %s)",
                             addr, words, hexs, got, expect, "一致" if got == expect else "不一致!")
            else:
                logging.warning("脉冲常数读回: 响应异常 resp=%s", rb)
        except Exception as e:
            logging.warning("脉冲常数读回失败: %s", e)

    def close(self):
        self.client.close()
