#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:memory_addrs.py
功能描述:定义寄存器地址
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
    2026-07-02 适配 RPP(MH主控): demand 路径用到的寄存器已按
    "RPP Modbus Address Table v1.00 20260617.xlsx" 替换——实时量取 10|12 Cycle
    sheet 的 VMM1 参数, 需量配置取 Basic Settings sheet(0x17C3~0x17C5/0x17DB),
    接线方式为 VMM1 Service Configuration(0x10AA, 枚举 0~5 与原 1320 wire_type
    编码一致)。能量类寄存器(*_energy_addr)仍为 AcuRev1320 残留, demand 路径
    不使用, 在 RPP 上不可直接使用。RPP 无系统时间/毫秒寄存器(时间走 NTP/PTP),
    sys_time_rms_addr / sys_millisecond 置 None, 时间触发需量重置不可用。
"""


class SlaveId:
    slave_id = 1  # RPP Modbus Slave ID(0x1000), RTU/TCP 共用, 默认 1


class MemoryReg:
    reg_single = 1
    reg_double = 2
    reg_uint32 = 2
    reg_int32 = 2
    reg_uint16 = 6
    reg_float32 = 2
    reg_uint16_t = 1
    reg_timestamp_active_energy = 146 # 114 + 32


class MemoryAddr:
    phase_order_rms_addr = 0x10AE  # RPP: VMM1 Phase Order (0:ABC, 1:ACB)
    set_freq_rms_addr = 0x17C1  # RPP: Frequency Selection, 枚举 0:50Hz / 1:60Hz(1320 为直接写频率值, 语义不同)
    sys_time_rms_addr = None  # RPP 无系统时间寄存器(NTP/PTP 同步), 不可经 Modbus 读写
    # freq_rms_addr = 0x3000
    # ua_rms_addr = 0x3002
    # ub_rms_addr = 0x3004
    # uc_rms_addr = 0x3006
    # uv_avg_rms_addr = 0x3008
    #
    # uab_rms_addr = 0x300A
    # ubc_rms_addr = 0x300C
    # uca_rms_addr = 0x300E
    # ul_avg_rms_addr = 0x3010
    #
    # ia_rms_addr = 0x3012
    # ib_rms_addr = 0x3014
    # ic_rms_addr = 0x3016
    # iv_avg_rms_addr = 0x3018
    #
    # in_rms_addr = 0x301A
    # pa_rms_addr = 0x301C
    # pb_rms_addr = 0x301E
    # pc_rms_addr = 0x3020
    # qa_rms_addr = 0x3024
    # qb_rms_addr = 0x3026
    # qc_rms_addr = 0x3028
    # sa_rms_addr = 0x302C
    # sb_rms_addr = 0x302E
    # sc_rms_addr = 0x3030
    # p_total_rms_addr = 0x3022
    # q_total_rms_addr = 0x302A
    # s_total_rms_addr = 0x3032
    #
    # pf_a_rms_addr = 0x3034
    # pf_b_rms_addr = 0x3036
    # pf_c_rms_addr = 0x3038
    # pf_total_rms_addr = 0x303A
    #
    # ua_phase_angle_rms_addr = 0x0000
    # ub_phase_angle_rms_addr = 0x42A0
    # uc_phase_angle_rms_addr = 0x42A1
    # ia_phase_angle_rms_addr = 0x42A2
    # ib_phase_angle_rms_addr = 0x42A3
    # ic_phase_angle_rms_addr = 0x42A4

    #  ================AcuRev1320 2hr update========================== #
    # freq_rms_addr = 0xB2A2
    # ua_rms_addr = 0xB2A4
    # ub_rms_addr = 0xB2A6
    # uc_rms_addr = 0xB2A8
    # uv_avg_rms_addr = 0xB2B0
    #
    # uab_rms_addr = 0xB2B2
    # ubc_rms_addr = 0xB2B4
    # uca_rms_addr = 0xB2B6
    # ul_avg_rms_addr = 0xB2BE
    #
    # ia_rms_addr = 0xB2C0
    # ib_rms_addr = 0xB2C2
    # ic_rms_addr = 0xB2C4
    # in_rms_addr = 0xB2C6
    # iv_avg_rms_addr = 0xB2D0
    #
    # pa_rms_addr = 0xB2D2
    # pb_rms_addr = 0xB2D4
    # pc_rms_addr = 0xB2D6
    #
    # qa_rms_addr = 0xB2DA
    # qb_rms_addr = 0xB2DC
    # qc_rms_addr = 0xB2DE
    #
    # sa_rms_addr = 0xB2E2
    # sb_rms_addr = 0xB2E4
    # sc_rms_addr = 0xB2E6
    #
    # p_total_rms_addr = 0xB2D8
    # q_total_rms_addr = 0xB2E0
    # s_total_rms_addr = 0xB2E8
    #
    # pf_a_rms_addr = 0xB2EA
    # pf_b_rms_addr = 0xB2EC
    # pf_c_rms_addr = 0xB2EE
    # pf_total_rms_addr = 0xB2F0
    #
    # ua_phase_angle_rms_addr = 0xB2AA
    # ub_phase_angle_rms_addr = 0xB2AC
    # uc_phase_angle_rms_addr = 0xB2AE
    #
    # ia_phase_angle_rms_addr = 0xB2C8
    # ib_phase_angle_rms_addr = 0xB2CA
    # ic_phase_angle_rms_addr = 0xB2CC
    # in_p_rms_addr = 0xB2CE

    #  ================AcuRev1320 10min update========================== #
    # freq_rms_addr = 0xA6FF
    # ua_rms_addr = 0xA701
    # ub_rms_addr = 0xA703
    # uc_rms_addr = 0xA705
    # uv_avg_rms_addr = 0xA70D
    #
    # uab_rms_addr = 0xA70F
    # ubc_rms_addr = 0xA711
    # uca_rms_addr = 0xA713
    # ul_avg_rms_addr = 0xA71B
    #
    # ia_rms_addr = 0xA71D
    # ib_rms_addr = 0xA71F
    # ic_rms_addr = 0xA721
    # in_rms_addr = 0xA723
    # iv_avg_rms_addr = 0xA72D
    #
    # pa_rms_addr = 0xA72F
    # pb_rms_addr = 0xA731
    # pc_rms_addr = 0xA733
    #
    # qa_rms_addr = 0xA737
    # qb_rms_addr = 0xA739
    # qc_rms_addr = 0xA73B
    #
    # sa_rms_addr = 0xA73F
    # sb_rms_addr = 0xA741
    # sc_rms_addr = 0xA743
    #
    # p_total_rms_addr = 0xA735
    # q_total_rms_addr = 0xA73D
    # s_total_rms_addr = 0xA745
    #
    # pf_a_rms_addr = 0xA747
    # pf_b_rms_addr = 0xA749
    # pf_c_rms_addr = 0xA74B
    # pf_total_rms_addr = 0xA74D
    #
    # ua_phase_angle_rms_addr = 0xA707
    # ub_phase_angle_rms_addr = 0xA709
    # uc_phase_angle_rms_addr = 0xA70B
    #
    # ia_phase_angle_rms_addr = 0xA725
    # ib_phase_angle_rms_addr = 0xA727
    # ic_phase_angle_rms_addr = 0xA729
    # in_p_rms_addr = 0xA72B

    #  ================3s update============================ #
    # freq_rms_addr = 0x9B80
    # ua_rms_addr = 0x9B82
    # ub_rms_addr = 0x9B84
    # uc_rms_addr = 0x9B86
    # uv_avg_rms_addr = 0x9B8E
    #
    # uab_rms_addr = 0x9B90
    # ubc_rms_addr = 0x9B92
    # uca_rms_addr = 0x9B94
    # ul_avg_rms_addr = 0x9B9C
    #
    # ia_rms_addr = 0x9B9E
    # ib_rms_addr = 0x9BA0
    # ic_rms_addr = 0x9BA2
    # in_rms_addr = 0x9BA4
    # iv_avg_rms_addr = 0x9BAE
    #
    # pa_rms_addr = 0x9BB0
    # pb_rms_addr = 0x9BB2
    # pc_rms_addr = 0x9BB4
    #
    # qa_rms_addr = 0x9BB8
    # qb_rms_addr = 0x9BBA
    # qc_rms_addr = 0x9BBC
    #
    # sa_rms_addr = 0x9BC0
    # sb_rms_addr = 0x9BC2
    # sc_rms_addr = 0x9BC4
    #
    # p_total_rms_addr = 0x9BB6
    # q_total_rms_addr = 0x9BBE
    # s_total_rms_addr = 0x9BC6
    #
    # pf_a_rms_addr = 0x9BC8
    # pf_b_rms_addr = 0x9BCA
    # pf_c_rms_addr = 0x9BCC
    # pf_total_rms_addr = 0x9BCE
    #
    # ua_phase_angle_rms_addr = 0x9B88
    # ub_phase_angle_rms_addr = 0x9B8A
    # uc_phase_angle_rms_addr = 0x9B8C
    #
    # ia_phase_angle_rms_addr = 0x9BA6
    # ib_phase_angle_rms_addr = 0x9BA8
    # ic_phase_angle_rms_addr = 0x9BAA
    # in_p_rms_addr = 0x9BAC
    #  ========== RPP 10|12 Cycle sheet(实时量, VMM1, float32) ========== #
    freq_rms_addr = 0x2907   # VMM1 System Frequency
    ua_rms_addr = 0x2909     # VMM1 Phase A Line-to-Neutral Voltage
    ub_rms_addr = 0x290B
    uc_rms_addr = 0x290D

    ua_phase_angle_rms_addr = 0x290F
    ub_phase_angle_rms_addr = 0x2911
    uc_phase_angle_rms_addr = 0x2913

    uab_phase_angle_rms_addr = 0x291D
    ubc_phase_angle_rms_addr = 0x291F
    uca_phase_angle_rms_addr = 0x2921

    uv_avg_rms_addr = 0x2915  # VMM1 Average Line-to-Neutral Voltage

    uab_rms_addr = 0x2917
    ubc_rms_addr = 0x2919
    uca_rms_addr = 0x291B
    ul_avg_rms_addr = 0x2923  # VMM1 Average Line-to-Line Voltage

    ia_rms_addr = 0x2927     # VMM1 Phase A Current
    ib_rms_addr = 0x2929
    ic_rms_addr = 0x292B
    in_rms_addr = 0x292D     # VMM1 Neutral Current

    ia_phase_angle_rms_addr = 0x2931
    ib_phase_angle_rms_addr = 0x2933
    ic_phase_angle_rms_addr = 0x2935
    in_p_rms_addr = 0x2937   # VMM1 Neutral Current Phase Angle

    iv_avg_rms_addr = 0x292F  # VMM1 System Average Current

    pa_rms_addr = 0x2BB2     # VMM1 Phase A Active Power
    pb_rms_addr = 0x2BBC
    pc_rms_addr = 0x2BC6

    qa_rms_addr = 0x2BB4     # VMM1 Phase A Reactive Power
    qb_rms_addr = 0x2BBE
    qc_rms_addr = 0x2BC8

    sa_rms_addr = 0x2BB6     # VMM1 Phase A Apparent Power
    sb_rms_addr = 0x2BC0
    sc_rms_addr = 0x2BCA

    p_total_rms_addr = 0x2BD0  # VMM1 System Active Power
    q_total_rms_addr = 0x2BD2
    s_total_rms_addr = 0x2BD4

    pf_a_rms_addr = 0x2BBA   # VMM1 Phase A Power Factor
    pf_b_rms_addr = 0x2BC4
    pf_c_rms_addr = 0x2BCE
    pf_total_rms_addr = 0x2BD8



    #  ================Moving Average 20ms update============================ #
    # freq_rms_addr = 0xC5AC
    #
    # ua_rms_addr = 0xC5AE
    # ub_rms_addr = 0xC5B0
    # uc_rms_addr = 0xC5B2
    # uv_avg_rms_addr =
    #
    # uab_rms_addr = 0xC5B4
    # ubc_rms_addr = 0xC5B6
    # uca_rms_addr = 0xC5B8
    # ul_avg_rms_addr =
    #
    # ia_rms_addr = 0xC5BA
    # ib_rms_addr = 0xC5BC
    # ic_rms_addr = 0xC5BE
    # in_rms_addr = 0xC5C0
    # iv_avg_rms_addr =
    #
    # pa_rms_addr = 0xC5C2
    # pb_rms_addr = 0xC5C4
    # pc_rms_addr = 0xC5C6
    #
    #
    # qa_rms_addr = 0xC5CA
    # qb_rms_addr = 0xC5CC
    # qc_rms_addr = 0xC5CE
    #
    #
    # sa_rms_addr = 0xC5D2
    # sb_rms_addr = 0xC5D4
    # sc_rms_addr = 0xC5D6
    #
    #
    # p_total_rms_addr = 0xC5C8
    # q_total_rms_addr = 0xC5D0
    # s_total_rms_addr = 0xC5D8
    #
    #
    # pf_a_rms_addr =
    # pf_b_rms_addr =
    # pf_c_rms_addr =
    # pf_total_rms_addr =
    #
    # ua_phase_angle_rms_addr =
    # ub_phase_angle_rms_addr =
    # uc_phase_angle_rms_addr =
    #
    # ia_phase_angle_rms_addr =
    # ib_phase_angle_rms_addr =
    # ic_phase_angle_rms_addr =
    # in_p_rms_addr =

    #  ================HS 20ms update============================ #
    # freq_rms_addr = 0xC47E
    # ua_rms_addr = 0xC480
    # ub_rms_addr = 0xC482
    # uc_rms_addr = 0xC484
    # uv_avg_rms_addr = 0xC486
    #
    # uab_rms_addr = 0xC492
    # ubc_rms_addr = 0xC494
    # uca_rms_addr = 0xC496
    # ul_avg_rms_addr = 0xC498
    #
    # ia_rms_addr = 0xC488
    # ib_rms_addr = 0xC48A
    # ic_rms_addr = 0xC48C
    # in_rms_addr = 0xC48E
    # iv_avg_rms_addr = 0xC490
    #
    # pa_rms_addr = 0xC4AC
    # pb_rms_addr = 0xC4AE
    # pc_rms_addr = 0xC4B0
    #
    # qa_rms_addr = 0xC4B4
    # qb_rms_addr = 0xC4B6
    # qc_rms_addr = 0xC4B8
    #
    # sa_rms_addr = 0xC4BC
    # sb_rms_addr = 0xC4BE
    # sc_rms_addr = 0xC4C0
    #
    # p_total_rms_addr = 0xC4B2
    # q_total_rms_addr = 0xC4BA
    # s_total_rms_addr = 0xC4C2
    #
    # pf_a_rms_addr = 0xC4C4
    # pf_b_rms_addr = 0xC4C6
    # pf_c_rms_addr = 0xC4C8
    # pf_total_rms_addr = 0xC4CA
    #
    # ua_phase_angle_rms_addr = 0xC4A0
    # ub_phase_angle_rms_addr = 0xC4A2
    # uc_phase_angle_rms_addr = 0xC4A4
    #
    # ia_phase_angle_rms_addr = 0xC4A6
    # ib_phase_angle_rms_addr = 0xC4A8
    # ic_phase_angle_rms_addr = 0xC4AA
    # in_p_rms_addr =

    #  =======================Energy 20ms================================= #
    # time_stamp_rms_addr = 0xC47A
    time_stamp_rms_addr = 0x9050
    clear_energy_addr = 0x1130

    pa_imp_energy_addr = 0x90E2
    pb_imp_energy_addr = 0x90E4
    pc_imp_energy_addr = 0x90E6
    p_sys_imp_energy_addr = 0x90E8

    pa_exp_energy_addr = 0x90EA
    pb_exp_energy_addr = 0x90EC
    pc_exp_energy_addr = 0x90EE
    p_sys_exp_energy_addr = 0x90F0

    p_sys_total_energy_addr = 0x90F8
    p_sys_net_energy_addr = 0x9100

    qa_imp_energy_addr = 0x9102
    qb_imp_energy_addr = 0x9104
    qc_imp_energy_addr = 0x9106
    q_sys_imp_energy_addr = 0x9108

    qa_exp_energy_addr = 0x910A
    qb_exp_energy_addr = 0x910C
    qc_exp_energy_addr = 0x910E
    q_sys_exp_energy_addr = 0x9110

    q_sys_total_energy_addr = 0x9118
    q_sys_net_energy_addr = 0x9120

    # phaseA phaseB phaseC视在能量
    sa_total_energy_addr = 0x9122
    sb_total_energy_addr = 0x9124
    sc_total_energy_addr = 0x9126
    s_sys_total_energy_addr = 0x9128

    # sa_imp_energy_addr = 0xC52C
    # sb_imp_energy_addr = 0xC52E
    # sc_imp_energy_addr = 0xC530
    # s_sys_imp_energy_addr = 0xC532

    # sa_exp_energy_addr = 0xC534
    # sb_exp_energy_addr = 0xC536
    # sc_exp_energy_addr = 0xC538
    # s_sys_exp_energy_addr = 0xC53A

    # acc_start_time_energy_addr =
    # acc_end_time_energy_addr =

    #  =======================Energy 20ms================================= #

    # RPP modbus addr(Basic Settings / Demand sheet)
    # 接线方式: VMM1 Service Configuration, 枚举与 1320 wire_type 一致
    # (0:1E2W 1:2E3W1P 2:2E3WD 3:2E3WN 4:3E4WY 5:3E4WD; RPP 官方不支持 3E4WD)
    voltage_wire_addr = 0x10AA
    current_wire_addr = voltage_wire_addr
    demand_algorithm_addr = 0x17C3   # Fixed Window: 0 / Sliding Window: 1
    demand_interval_addr = 0x17C4    # 1~30 min
    demand_update_rate_addr = 0x17C5  # 1~30 min
    # 需量读寄存器(VMM1, float32); 运行时会被 demand_addr_reader.load_demand_addr()
    # 从知识库 RPP 地址表重新解析覆盖。功率需量取 Import 口径。
    # RPP 无中线电流需量寄存器, 故无 phase_n_current 键(read_demand_in 返回 None)。
    demand_addr = {
        "system_active_power": 0x2726,
        "system_reactive_power": 0x272A,
        "system_apparent_power": 0x272E,
        "phase_a_current": 0x2700,
        "phase_b_current": 0x270C,
        "phase_c_current": 0x2718,
    }
    sys_millisecond = None  # RPP 无毫秒/系统时间寄存器 -> demand_trigger=1(时间触发)不可用
    clear_max_demand = 0x17DB  # Clear MaxDemand
