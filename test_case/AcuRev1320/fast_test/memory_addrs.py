#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:memory_addrs.py
功能描述:定义寄存器地址
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""


class SlaveId:
    slave_id = 1


class MemoryReg:
    reg_single = 1
    reg_double = 2
    reg_uint32 = 2
    reg_int32 = 2
    reg_uint16 = 6
    reg_float32 = 2


class MemoryAddr:

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
    freq_rms_addr = 0xA6FF
    ua_rms_addr = 0xA701
    ub_rms_addr = 0xA703
    uc_rms_addr = 0xA705
    uv_avg_rms_addr = 0xA70D

    uab_rms_addr = 0xA70F
    ubc_rms_addr = 0xA711
    uca_rms_addr = 0xA713
    ul_avg_rms_addr = 0xA71B

    ia_rms_addr = 0xA71D
    ib_rms_addr = 0xA71F
    ic_rms_addr = 0xA721
    in_rms_addr = 0xA723
    iv_avg_rms_addr = 0xA72D

    pa_rms_addr = 0xA72F
    pb_rms_addr = 0xA731
    pc_rms_addr = 0xA733

    qa_rms_addr = 0xA737
    qb_rms_addr = 0xA739
    qc_rms_addr = 0xA73B

    sa_rms_addr = 0xA73F
    sb_rms_addr = 0xA741
    sc_rms_addr = 0xA743

    p_total_rms_addr = 0xA735
    q_total_rms_addr = 0xA73D
    s_total_rms_addr = 0xA745

    pf_a_rms_addr = 0xA747
    pf_b_rms_addr = 0xA749
    pf_c_rms_addr = 0xA74B
    pf_total_rms_addr = 0xA74D

    ua_phase_angle_rms_addr = 0xA707
    ub_phase_angle_rms_addr = 0xA709
    uc_phase_angle_rms_addr = 0xA70B

    ia_phase_angle_rms_addr = 0xA725
    ib_phase_angle_rms_addr = 0xA727
    ic_phase_angle_rms_addr = 0xA729
    in_p_rms_addr = 0xA72B

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
    #  ================200ms update============================ #
    # freq_rms_addr = 0x9007
    # ua_rms_addr = 0x9009
    # ub_rms_addr = 0x900B
    # uc_rms_addr = 0x900D
    # uv_avg_rms_addr = 0x9015
    #
    # uab_rms_addr = 0x9017
    # ubc_rms_addr = 0x9019
    # uca_rms_addr = 0x901B
    # ul_avg_rms_addr = 0x9023
    #
    # ia_rms_addr = 0x9025
    # ib_rms_addr = 0x9027
    # ic_rms_addr = 0x9029
    # in_rms_addr = 0x902B
    # iv_avg_rms_addr = 0x9035
    #
    # pa_rms_addr = 0x9037
    # pb_rms_addr = 0x9039
    # pc_rms_addr = 0x903B
    #
    # qa_rms_addr = 0x903F
    # qb_rms_addr = 0x9041
    # qc_rms_addr = 0x9043
    #
    # sa_rms_addr = 0x9047
    # sb_rms_addr = 0x9049
    # sc_rms_addr = 0x904B
    #
    # p_total_rms_addr = 0x903D
    # q_total_rms_addr = 0x9045
    # s_total_rms_addr = 0x904D
    #
    # pf_a_rms_addr = 0x904F
    # pf_b_rms_addr = 0x9051
    # pf_c_rms_addr = 0x9053
    # pf_total_rms_addr = 0x9055
    #
    # ua_phase_angle_rms_addr = 0x900F
    # ub_phase_angle_rms_addr = 0x9011
    # uc_phase_angle_rms_addr = 0x9013
    #
    # ia_phase_angle_rms_addr = 0x902D
    # ib_phase_angle_rms_addr = 0x902F
    # ic_phase_angle_rms_addr = 0x9031
    # in_p_rms_addr = 0x9033

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
    #
    # pa_rms_addr = 0xC4AC
    # pb_rms_addr = 0xC4AE
    # pc_rms_addr = 0xC4B0
    #
    #
    # qa_rms_addr = 0xC4B4
    # qb_rms_addr = 0xC4B6
    # qc_rms_addr = 0xC4B8
    #
    #
    # sa_rms_addr = 0xC4BC
    # sb_rms_addr = 0xC4BE
    # sc_rms_addr = 0xC4C0
    #
    #
    # p_total_rms_addr = 0xC4B2
    # q_total_rms_addr = 0xC4BA
    # s_total_rms_addr = 0xC4C2
    #
    #
    # pf_a_rms_addr = 0xC4C4
    # pf_b_rms_addr = 0xC4C6
    # pf_c_rms_addr = 0xC4C8
    # pf_total_rms_addr = 0xC4CA
    #
    #
    # ua_phase_angle_rms_addr = 0xC49A
    # ub_phase_angle_rms_addr = 0xC49C
    # uc_phase_angle_rms_addr = 0xC49E
    #
    #
    # ia_phase_angle_rms_addr = 0xC4A6
    # ib_phase_angle_rms_addr = 0xC4A8
    # ic_phase_angle_rms_addr = 0xC4AA
    # in_p_rms_addr =

    #  =======================Energy 20ms================================= #
    clear_energy_addr = 0x1130

    pa_imp_energy_addr = 0xC4EC
    pa_exp_energy_addr = 0xC4F4

    pb_imp_energy_addr = 0xC4EE
    pb_exp_energy_addr = 0xC4F6

    pc_imp_energy_addr = 0xC4F0
    pc_exp_energy_addr = 0xC4F8

    qa_imp_energy_addr = 0xC50C
    qa_exp_energy_addr = 0xC514

    qb_imp_energy_addr = 0xC50E
    qb_exp_energy_addr = 0xC516

    qc_imp_energy_addr = 0xC510
    qc_exp_energy_addr = 0xC518

    sa_imp_energy_addr = 0xC52C
    sa_exp_energy_addr = 0xC534

    sb_imp_energy_addr = 0xC52E
    sb_exp_energy_addr = 0xC536

    sc_imp_energy_addr = 0xC530
    sc_exp_energy_addr = 0xC538

    # phaseA phaseB phaseC视在能量
    sa_app_energy_addr = 0xC53C
    sb_app_energy_addr = 0xC53E
    sc_app_energy_addr = 0xC540

    # acc_start_time_energy_addr =
    # acc_end_time_energy_addr =

    p_sys_imp_energy_addr = 0xC4F2
    p_sys_exp_energy_addr = 0xC4FA
    p_sys_total_energy_addr = 0xC502
    p_sys_net_energy_addr = 0xC50A

    q_sys_imp_energy_addr = 0xC512
    q_sys_exp_energy_addr = 0xC51A
    q_sys_total_energy_addr = 0xC522
    q_sys_net_energy_addr = 0xC52A

    s_sys_imp_energy_addr = 0xC532
    s_sys_exp_energy_addr = 0xC53A

    s_sys_total_energy_addr = 0xC542
    s_sys_net_energy_addr = 0xC54A

    #  =======================Energy 20ms================================= #



    # AcuRev1320 modbus addr
    voltage_wire_addr = 0x1042
    current_wire_addr = voltage_wire_addr
    demand_algorithm_addr = 0X1059
    demand_interval_addr = 0x105A
    demand_update_rate_addr = 0x105B
    demand_addr = {
        "system_active_power": 0xC466,
        "system_reactive_power": 0xC468,
        "system_apparent_power": 0xC46A,
        "phase_a_current": 0xC46C,
        "phase_b_current": 0xC46E,
        "phase_c_current": 0xC470,
        "phase_n_current": 0xC472,
    }
    sys_millisecond = 0x1026
    clear_max_demand = 0X1131



