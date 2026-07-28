"""CL3021 transport layer — TCP socket and RS-232 serial.

Frame format reminder (from cl3021_frame.py):
    Head(0x81) | RxID | TxID | Length | Cmd | Data... | CS
    Length = total byte count (Head..CS inclusive).

_read_frame(read_fn) reads exactly one complete frame:
  1. Accumulate bytes until a 0x81 start byte is found (resync on garbage).
  2. Accumulate until 4 header bytes are available; Length = header[3].
  3. Accumulate until total == Length bytes.
  4. Raise ConnectionError on EOF (read_fn returns b"").

pyserial is imported lazily inside SerialTransport.open() so the module can be
imported even when pyserial is not installed.
"""

import socket


# ---------------------------------------------------------------------------
# Shared frame-reader
# ---------------------------------------------------------------------------

def _read_frame(read_fn) -> bytes:
    """Read exactly one CL3021 frame using *read_fn(n)*.

    *read_fn* should behave like ``socket.recv`` or ``serial.read``:
    it returns up to *n* bytes and returns ``b""`` on EOF/timeout.

    Steps:
      1. Read bytes one-at-a-time (or in small chunks) until 0x81 is found.
      2. Accumulate until we have the 4-byte header.
      3. Read ``Length - 4`` more bytes.

    Raises:
        ConnectionError: if *read_fn* returns b"" before the frame is complete.
    """
    buf = bytearray()

    # --- Phase 1: find the 0x81 start-of-frame byte ---
    while True:
        chunk = read_fn(1)
        if not chunk:
            raise ConnectionError("EOF before start-of-frame byte 0x81")
        if chunk[0] == 0x81:
            buf += chunk
            break
        # else: discard garbage byte and keep looking

    # --- Phase 2: accumulate the rest of the 4-byte header ---
    while len(buf) < 4:
        chunk = read_fn(4 - len(buf))
        if not chunk:
            raise ConnectionError(
                f"EOF while reading header (have {len(buf)} of 4 bytes)"
            )
        buf += chunk

    length = buf[3]  # total frame byte count

    # --- Phase 3: accumulate the payload ---
    while len(buf) < length:
        need = length - len(buf)
        chunk = read_fn(need)
        if not chunk:
            raise ConnectionError(
                f"EOF while reading frame payload (have {len(buf)} of {length} bytes)"
            )
        buf += chunk

    return bytes(buf)


# ---------------------------------------------------------------------------
# TCP transport
# ---------------------------------------------------------------------------

class TcpTransport:
    """CL3021 transport over a TCP socket."""

    def __init__(self, host: str = "192.168.0.50", port: int = 2404, timeout: float = 3.0):
        self.host = host
        self.port = port
        self.timeout = timeout
        self.sock = None

    def open(self) -> None:
        """Open a TCP connection to (host, port)."""
        self.sock = socket.create_connection((self.host, self.port), timeout=self.timeout)
        self.sock.settimeout(self.timeout)

    def close(self) -> None:
        """Close the socket; sets sock to None regardless of errors."""
        try:
            if self.sock is not None:
                self.sock.close()
        finally:
            self.sock = None

    def send_frame(self, frame: bytes) -> None:
        """Send *frame* over the socket."""
        self.sock.sendall(frame)

    def recv_frame(self) -> bytes:
        """Read one complete CL3021 response frame from the socket."""
        return _read_frame(self.sock.recv)

    def query(self, frame: bytes) -> bytes:
        """Send *frame* and return the response frame."""
        self.send_frame(frame)
        return self.recv_frame()


# ---------------------------------------------------------------------------
# RS-232 serial transport
# ---------------------------------------------------------------------------

class SerialTransport:
    """CL3021 transport over RS-232 serial (pyserial)."""

    def __init__(
        self,
        port: str = "COM3",
        baud: int = 9600,
        bytesize: int = 8,
        parity: str = "N",
        stopbits: int = 1,
        timeout: float = 3.0,
        ser=None,
    ):
        self.port = port
        self.baud = baud
        self.bytesize = bytesize
        self.parity = parity
        self.stopbits = stopbits
        self.timeout = timeout
        self.ser = ser  # injectable for testing

    def open(self) -> None:
        """Open the serial port.  pyserial is imported lazily here."""
        if self.ser is None:
            import serial  # lazy import — keeps module importable without pyserial

            self.ser = serial.Serial(
                self.port,
                self.baud,
                bytesize=self.bytesize,
                parity=self.parity,
                stopbits=self.stopbits,
                timeout=self.timeout,
            )

    def close(self) -> None:
        """Close the serial port; sets ser to None regardless of errors."""
        try:
            if self.ser is not None:
                self.ser.close()
        finally:
            self.ser = None

    def send_frame(self, frame: bytes) -> None:
        """Write *frame* to the serial port."""
        self.ser.write(frame)

    def recv_frame(self) -> bytes:
        """Read one complete CL3021 response frame from the serial port."""
        return _read_frame(self.ser.read)

    def query(self, frame: bytes) -> bytes:
        """Send *frame* and return the response frame."""
        self.send_frame(frame)
        return self.recv_frame()
