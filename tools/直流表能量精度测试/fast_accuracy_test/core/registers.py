_SINGLE = {
    "voltage": (12288, "float"),
    "current": (12290, "float"),
    "power": (12292, "float"),
    "import_energy": (16384, "double"),
    "export_energy": (16388, "double"),
}

_DUAL = {
    "voltage": (12288, "float"),
    "current_1": (12290, "float"),
    "current_2": (12292, "float"),
    "current_sum": (12294, "float"),
    "power_1": (12296, "float"),
    "power_2": (12298, "float"),
    "power_sum": (12300, "float"),
    "import_energy_1": (16384, "double"),
    "export_energy_1": (16388, "double"),
    "import_energy_2": (16400, "double"),
    "export_energy_2": (16404, "double"),
}


def regmap(is_dual):
    return dict(_DUAL if is_dual else _SINGLE)


# 各型号"能量脉冲常数"寄存器：(寄存器地址Dec, 写入换算系数, 物理常数上限)
# 写入寄存器值 = 物理脉冲常数 × 系数，占 2 个寄存器(32位整数)。
#   "260"(=261,双路): 0x1032=4146, ×10000, 上限 5000
#   "320"           : 0x101A=4122, ×1000 , 上限 100000
#   "300"(=301)     : 0x101A=4122, ×1000 , 上限 100000
_PULSE_CONST_REG = {
    "260": (4146, 10000, 5000.0),
    "320": (4122, 1000, 100000.0),
    "300": (4122, 1000, 100000.0),
}


def pulse_const_reg(device_model):
    """按型号取脉冲常数寄存器配置 (addr, scale, max_const)；未知型号返回 None。"""
    return _PULSE_CONST_REG.get(device_model)

