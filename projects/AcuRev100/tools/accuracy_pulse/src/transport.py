import socket


class SocketTransport:
    """网口 raw TCP socket 传输。"""

    def __init__(self, host, port=5024, timeout=5.0):
        self.host = host
        self.port = port
        self.timeout = timeout
        self.sock = None

    def open(self):
        self.sock = socket.create_connection((self.host, self.port), timeout=self.timeout)
        self.sock.settimeout(self.timeout)

    def close(self):
        if self.sock:
            try:
                self.sock.close()
            finally:
                self.sock = None

    def send(self, cmd):
        self.sock.sendall((cmd + "\n").encode("ascii"))

    def query(self, cmd):
        self.send(cmd)
        return self._recv_line()

    def _recv_line(self):
        buf = b""
        while b"\n" not in buf:
            chunk = self.sock.recv(4096)
            if not chunk:
                break
            buf += chunk
        return buf.decode("ascii", errors="replace").strip()


class VisaTransport:
    """USB (USBTMC) 传输，经 pyvisa。rm 可注入用于测试。"""

    def __init__(self, resource, timeout=5.0, rm=None):
        self.resource = resource
        self.timeout = timeout
        self._rm = rm
        self.inst = None

    def open(self):
        if self._rm is None:
            import pyvisa
            self._rm = pyvisa.ResourceManager()
        self.inst = self._rm.open_resource(self.resource)
        self.inst.timeout = int(self.timeout * 1000)
        self.inst.write_termination = "\n"
        self.inst.read_termination = "\n"

    def close(self):
        if self.inst:
            try:
                self.inst.close()
            finally:
                self.inst = None

    def send(self, cmd):
        self.inst.write(cmd)

    def query(self, cmd):
        return self.inst.query(cmd).strip()


def list_usb_resources(rm=None):
    """返回以 'USB' 开头的 VISA 资源串（USBTMC 设备）。"""
    if rm is None:
        import pyvisa
        rm = pyvisa.ResourceManager()
    return [r for r in rm.list_resources() if r.upper().startswith("USB")]
