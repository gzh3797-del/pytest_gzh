"""High-level CL3021 client for AC source (programmable power source) control.

Wraps the binary frame codec (cl3021_frame) and a transport
(cl3021_transport) to expose intent-level operations: connectivity check,
wiring/range configuration, AC output setpoints, overload clear, and a
measurement read.

Phase ordering note
-------------------
Public methods take per-phase 3-tuples in **(A, B, C)** order, matching how a
human reads phases. The CL3021 wire protocol, however, transmits each
per-phase group in **C, B, A** order. set_ac_output performs this reversal
when building the payload.

Number formats (see cl3021_frame):
    * phase angles and frequency  -> scaled_encode  (value * 10000, 4-byte LE)
    * voltage amplitude           -> int4e1_encode(value, exp=-4)  (5 bytes)
    * current amplitude           -> int4e1_encode(value, exp=-6)  (5 bytes)
"""

from src.cl3021_frame import build_frame, parse_frame, int4e1_encode, scaled_encode


class CL3021Error(Exception):
    """Raised when the CL3021 returns a failure response (cmd 0x33)."""


class CL3021Client:
    """High-level command interface to a CL3021 AC source over a transport."""

    def __init__(self, transport, rx_id=0x01, tx_id=0x25):
        self.transport = transport
        self.rx_id = rx_id
        self.tx_id = tx_id

    # -- lifecycle ---------------------------------------------------------

    def connect(self):
        """Open the underlying transport."""
        self.transport.open()

    def close(self):
        """Close the underlying transport."""
        self.transport.close()

    # -- core send/receive -------------------------------------------------

    def _send(self, cmd, data=b""):
        """Build a frame, query the transport, and return the parsed response.

        All CL3021 commands used by this client expect a response frame.
        Raises CL3021Error if the device returns a failure response (cmd 0x33).
        """
        frame = build_frame(self.rx_id, self.tx_id, cmd, data)
        resp = self.transport.query(frame)
        parsed = parse_frame(resp)
        if parsed["cmd"] == 0x33:
            raise CL3021Error(f"CL3021 returned failure (0x33) for cmd 0x{cmd:02X}")
        return parsed

    # -- commands ----------------------------------------------------------

    def get_version(self):
        """联机 / connectivity check — cmd 0xC9, no data. Returns parsed response."""
        return self._send(0xC9)

    def set_wiring(self, linemode=0x08):
        """Configure wiring mode — cmd 0xA3, data = 00 01 20 <linemode>."""
        data = bytes([0x00, 0x01, 0x20, linemode])
        return self._send(0xA3, data)

    def set_range_mode(self, auto=True):
        """Set range to auto/manual — cmd 0xA3, data = 05 40 04 <00 if auto else FF>."""
        data = bytes([0x05, 0x40, 0x04, 0x00 if auto else 0xFF])
        return self._send(0xA3, data)

    def set_ac_output(self, voltage, current, phase_v, phase_i, freq=50.0):
        """Set the full AC output — cmd 0xA3. THE CORE command.

        Each of *voltage*, *current*, *phase_v*, *phase_i* is a 3-tuple in
        **(A, B, C)** phase order. The wire protocol transmits each group in
        **C, B, A** order, so this method reverses each tuple when packing.

        Payload byte layout (all groups emitted C, B, A):
            05 46 3F
            + scaled(phase_v[C]) scaled(phase_v[B]) scaled(phase_v[A])
            + scaled(phase_i[C]) scaled(phase_i[B]) scaled(phase_i[A])
            + FF
            + int4e1(voltage[C],-4) int4e1(voltage[B],-4) int4e1(voltage[A],-4)
            + int4e1(current[C],-6) int4e1(current[B],-6) int4e1(current[A],-6)
            + scaled(freq)
            + 07 03 3F 3F        (fixed update-all flags)
        """
        va, vb, vc = voltage
        ia, ib, ic = current
        pva, pvb, pvc = phase_v
        pia, pib, pic = phase_i

        data = bytearray([0x05, 0x46, 0x3F])

        # phase voltages, C B A
        data += scaled_encode(pvc)
        data += scaled_encode(pvb)
        data += scaled_encode(pva)

        # phase currents, C B A
        data += scaled_encode(pic)
        data += scaled_encode(pib)
        data += scaled_encode(pia)

        data += bytes([0xFF])

        # voltage amplitudes, C B A (exp -4)
        data += int4e1_encode(vc, -4)
        data += int4e1_encode(vb, -4)
        data += int4e1_encode(va, -4)

        # current amplitudes, C B A (exp -6)
        data += int4e1_encode(ic, -6)
        data += int4e1_encode(ib, -6)
        data += int4e1_encode(ia, -6)

        # frequency
        data += scaled_encode(freq)

        # fixed update-all flags
        data += bytes([0x07, 0x03, 0x3F, 0x3F])

        return self._send(0xA3, bytes(data))

    def stop_output(self):
        """Stop output by driving all amplitudes to zero (freq held at 50 Hz)."""
        return self.set_ac_output(
            voltage=(0, 0, 0),
            current=(0, 0, 0),
            phase_v=(0, 0, 0),
            phase_i=(0, 0, 0),
            freq=50.0,
        )

    def clear_overload(self):
        """Clear an overload/fault latch — cmd 0xA3, data = 02 01 80 00."""
        data = bytes([0x02, 0x01, 0x80, 0x00])
        return self._send(0xA3, data)

    def read_measurements(self):
        """Request a measurement snapshot — cmd 0xA0.

        Sends data = 02 7F FF 80 3F FF FF 0F 80 and returns the parsed
        response frame.

        NOTE: The measurement payload layout is not yet documented. This
        method returns the parsed frame as-is; the caller should interpret
        ``result["data"]`` once a real-device response sample is available.
        Payload parsing is deferred pending a real-device response sample.
        """
        data = bytes([0x02, 0x7F, 0xFF, 0x80, 0x3F, 0xFF, 0xFF, 0x0F, 0x80])
        return self._send(0xA0, data)
