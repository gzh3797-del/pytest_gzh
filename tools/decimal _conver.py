def dec_to_signed_decimal(dec_str, bit_width):
    if dec_str & (1 << (bit_width - 1)):
        # 计算补码形式的负数
        dec_str -= (1 << bit_width)

    return dec_str
