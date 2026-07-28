"""CL3021 binary frame codec — pure encode/decode, no networking.

Frame layout:
    Head(0x81) | RxID | TxID | Length | Cmd | Data... | CS

Length  = total byte count of the frame (Head..CS inclusive) = 6 + len(data)
CS      = XOR of bytes from RxID through last Data byte
          (i.e. RxID, TxID, Length, Cmd, Data...)

Number formats
--------------
Int4E1  (amplitudes): 5 bytes
    bytes 0-3: 4-byte little-endian signed int32 (mantissa)
    byte  4:   1 signed byte (exponent, two's-complement)
    value = mantissa * 10**exp

Scaled int (phase, frequency): 4 bytes
    round(value * 10000) as 4-byte little-endian signed int32
"""

import struct


# ---------------------------------------------------------------------------
# Checksum
# ---------------------------------------------------------------------------

def xor_checksum(data: bytes) -> int:
    """XOR of all bytes in *data*."""
    result = 0
    for b in data:
        result ^= b
    return result


# ---------------------------------------------------------------------------
# Frame builder / parser
# ---------------------------------------------------------------------------

def build_frame(rx_id: int, tx_id: int, cmd: int, data: bytes = b"") -> bytes:
    """Build a complete CL3021 frame including head and CS.

    Returns:
        bytes: Head | RxID | TxID | Length | Cmd | Data... | CS
    """
    length = 6 + len(data)           # total frame size
    cs_payload = bytes([rx_id, tx_id, length, cmd]) + data
    cs = xor_checksum(cs_payload)
    return bytes([0x81]) + cs_payload + bytes([cs])


def parse_frame(frame: bytes) -> dict:
    """Parse a CL3021 frame and validate its checksum.

    Returns:
        dict with keys: rx_id, tx_id, length, cmd, data (bytes), cs, cs_ok (bool)

    Raises:
        ValueError: if frame[0] != 0x81 or if frame is truncated
    """
    if not frame or frame[0] != 0x81:
        raise ValueError(f"Invalid CL3021 frame head: expected 0x81, got {frame[0]:#04x}")

    rx_id  = frame[1]
    tx_id  = frame[2]
    length = frame[3]
    cmd    = frame[4]

    if len(frame) < length:
        raise ValueError(f"Frame too short: expected {length} bytes, got {len(frame)}")

    data   = frame[5:length - 1]    # everything between cmd and CS
    cs     = frame[length - 1]

    cs_payload = frame[1:length - 1]   # RxID through last Data byte
    cs_ok = (xor_checksum(cs_payload) == cs)

    return {
        "rx_id":  rx_id,
        "tx_id":  tx_id,
        "length": length,
        "cmd":    cmd,
        "data":   data,
        "cs":     cs,
        "cs_ok":  cs_ok,
    }


# ---------------------------------------------------------------------------
# Int4E1 (amplitude) codec
# ---------------------------------------------------------------------------

def int4e1_encode(value: float, exp: int) -> bytes:
    """Encode *value* as a 5-byte Int4E1 field with the given exponent.

    mantissa = round(value / 10**exp)  ->  4-byte LE signed int32
    exponent ->  1 signed byte
    """
    mantissa = round(value / (10 ** exp))
    return struct.pack("<iB", mantissa, exp & 0xFF)   # store exp as unsigned byte (two's complement)


def int4e1_decode(b: bytes) -> float:
    """Decode a 5-byte Int4E1 field, returning the float value."""
    mantissa = struct.unpack_from("<i", b, 0)[0]
    # Exponent byte is signed (two's complement)
    exp = struct.unpack_from("<b", b, 4)[0]
    return mantissa * (10 ** exp)


# ---------------------------------------------------------------------------
# Scaled int (phase / frequency) codec
# ---------------------------------------------------------------------------

def scaled_encode(value: float) -> bytes:
    """Encode *value* as a 4-byte LE signed int32 (value * 10000)."""
    return struct.pack("<i", round(value * 10000))


def scaled_decode(b: bytes) -> float:
    """Decode a 4-byte LE signed int32 back to float (/ 10000)."""
    return struct.unpack_from("<i", b, 0)[0] / 10000
