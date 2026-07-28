import struct


def _decode(words, word_order, fmt):
    if word_order == "little":
        words = list(reversed(words))
    elif word_order != "big":
        raise ValueError(f"word_order must be 'big' or 'little', got {word_order!r}")
    raw = struct.pack(">%dH" % len(words), *words)
    return struct.unpack(">" + fmt, raw)[0]


def decode_float(words, word_order="big"):
    assert len(words) == 2, "float needs 2 registers"
    return _decode(words, word_order, "f")


def decode_double(words, word_order="big"):
    assert len(words) == 4, "double needs 4 registers"
    return _decode(words, word_order, "d")


def decode_u32(words, word_order="big"):
    """2 个寄存器 → 32 位无符号整数(用于读回脉冲常数寄存器做校验)。"""
    assert len(words) == 2, "u32 needs 2 registers"
    return _decode(words, word_order, "I")


def encode_u32(value, word_order="big"):
    """32 位无符号整数 → 2 个 16 位寄存器(用于写脉冲常数等)。
    word_order='big'=高字在前(默认)；'little'=低字在前。"""
    v = int(round(value))
    if v < 0 or v > 0xFFFFFFFF:
        raise ValueError("encode_u32 超出 32 位无符号范围: %s" % value)
    hi, lo = (v >> 16) & 0xFFFF, v & 0xFFFF
    if word_order == "big":
        return [hi, lo]
    elif word_order == "little":
        return [lo, hi]
    raise ValueError("word_order must be 'big' or 'little', got %r" % word_order)
