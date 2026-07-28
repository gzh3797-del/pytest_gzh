import re

_NUM_RE = re.compile(r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?")

def _num(s):
    """从带单位的值串（如 '10000000HZ' / '0ppm'）提取前导浮点数；无数字返回 None。"""
    if s is None:
        return None
    m = _NUM_RE.match(s.strip())
    return float(m.group(0)) if m else None

def parse_fcnt(resp):
    """解析 FCNT? 响应为字典。无效频率时 period=None。"""
    resp = resp.strip()
    if resp.upper().startswith("FCNT"):
        resp = resp[4:].strip()
    parts = [p.strip() for p in resp.split(",")]
    kv = {}
    for i in range(0, len(parts) - 1, 2):
        kv[parts[i].upper()] = parts[i + 1]
    result = {
        "state": kv.get("STATE", ""),
        "mode": kv.get("MODE", ""),
        "hfr": kv.get("HFR", ""),
        "frq": _num(kv.get("FRQ", "")),
        "duty": _num(kv.get("DUTY", "")),
        "refq": _num(kv.get("REFQ", "")),
        "trg": _num(kv.get("TRG", "")),
        "pw": _num(kv.get("PW", "")),
        "nw": _num(kv.get("NW", "")),
        "frqdev": _num(kv.get("FRQDEV", "")),
    }
    frq = result["frq"]
    result["period"] = (1.0 / frq) if frq else None
    return result


class SDGClient:
    """与 SDG 频率计交互，传输方式由注入的 transport 决定。"""

    def __init__(self, transport):
        self.transport = transport

    def connect(self):
        self.transport.open()

    def close(self):
        self.transport.close()

    def send(self, cmd):
        self.transport.send(cmd)

    def query(self, cmd):
        return self.transport.query(cmd)

    def query_fcnt(self):
        return parse_fcnt(self.query("FCNT?"))

    def set_counter(self, on):
        self.send("FCNT STATE,ON" if on else "FCNT STATE,OFF")

    def set_mode(self, mode):
        self.send("FCNT MODE," + mode)

    def set_hfr(self, on):
        self.send("FCNT HFR,ON" if on else "FCNT HFR,OFF")

    def set_refq(self, hz):
        self.send("FCNT REFQ," + str(hz))

    def set_trg(self, v):
        self.send("FCNT TRG," + str(v))

    def set_type(self, t):
        self.send("FCNT TYPE," + t)   # t = "FAST" or "SLOW" (fast/slow measurement)
