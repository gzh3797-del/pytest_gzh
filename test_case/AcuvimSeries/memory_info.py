#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:memory_info.py
功能描述:获取寄存器地址
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""

import json
from tools.excel_operate import HandleExcel, json_to_dict, OUTPUT_JSON_PATH


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
    memory_data_dict = json_to_dict(json_path=OUTPUT_JSON_PATH)
    # memory_data_dict = HandleExcel().data_dict

    real_time_parameters_by_start_addr_by_100ms = 0x1003
    phase_angle_by_start_addr_by_100ms = 0x42A0
    clear_energy_addr = memory_data_dict["system_settings"].get("Clear Energy", 0)
    voltage_wire_addr = memory_data_dict["system_settings"].get("Voltage Input Wiring Type", 0)
    current_wire_addr = memory_data_dict["system_settings"].get("Current Input Wiring Type", 0)
    freq_rms_addr = memory_data_dict["real_time_parameters"].get("Freq_rms", 0)
    ua_rms_addr = memory_data_dict["real_time_parameters"].get("Ua_rms", 0)
    ub_rms_addr = memory_data_dict["real_time_parameters"].get("Ub_rms", 0)
    uc_rms_addr = memory_data_dict["real_time_parameters"].get("Uc_rms", 0)
    uv_avg_rms_addr = memory_data_dict["real_time_parameters"].get("Uvag_rms", 0)
    uab_rms_addr = memory_data_dict["real_time_parameters"].get("Uab_rms", 0)
    ubc_rms_addr = memory_data_dict["real_time_parameters"].get("Ubc_rms", 0)
    uca_rms_addr = memory_data_dict["real_time_parameters"].get("Uca_rms", 0)
    ul_avg_rms_addr = memory_data_dict["real_time_parameters"].get("Ulag_rms", 0)
    ia_rms_addr = memory_data_dict["real_time_parameters"].get("Ia_rms", 0)
    ib_rms_addr = memory_data_dict["real_time_parameters"].get("Ib_rms", 0)
    ic_rms_addr = memory_data_dict["real_time_parameters"].get("Ic_rms", 0)
    iv_avg_rms_addr = memory_data_dict["real_time_parameters"].get("Ivag_rms", 0)
    in_rms_addr = memory_data_dict["real_time_parameters"].get("In_rms", 0)
    pa_rms_addr = memory_data_dict["real_time_parameters"].get("Pa_rms", 0)
    pb_rms_addr = memory_data_dict["real_time_parameters"].get("Pb_rms", 0)
    pc_rms_addr = memory_data_dict["real_time_parameters"].get("Pc_rms", 0)
    p_total_rms_addr = memory_data_dict["real_time_parameters"].get("P_rms", 0)
    qa_rms_addr = memory_data_dict["real_time_parameters"].get("Qa_rms", 0)
    qb_rms_addr = memory_data_dict["real_time_parameters"].get("Qb_rms", 0)
    qc_rms_addr = memory_data_dict["real_time_parameters"].get("Qc_rms", 0)
    q_total_rms_addr = memory_data_dict["real_time_parameters"].get("Q_rms", 0)
    sa_rms_addr = memory_data_dict["real_time_parameters"].get("Sa_rms", 0)
    sb_rms_addr = memory_data_dict["real_time_parameters"].get("Sb_rms", 0)
    sc_rms_addr = memory_data_dict["real_time_parameters"].get("Sc_rms", 0)
    s_total_rms_addr = memory_data_dict["real_time_parameters"].get("S_rms", 0)
    pf_a_rms_addr = memory_data_dict["real_time_parameters"].get("PFa_rms", 0)
    pf_b_rms_addr = memory_data_dict["real_time_parameters"].get("PFb_rms", 0)
    pf_c_rms_addr = memory_data_dict["real_time_parameters"].get("PFc_rms", 0)
    pf_total_rms_addr = memory_data_dict["real_time_parameters"].get("PF_rms", 0)
    ua_phase_angle_rms_addr = 0x0000
    ub_phase_angle_rms_addr = memory_data_dict["real_time_parameters"].get("Phase Angle of V2 to V1", 0)
    uc_phase_angle_rms_addr = memory_data_dict["real_time_parameters"].get("Phase Angle of V3 to V1", 0)
    ia_phase_angle_rms_addr = memory_data_dict["real_time_parameters"].get("Phase Angle of I1 to V1", 0)
    ib_phase_angle_rms_addr = memory_data_dict["real_time_parameters"].get("Phase Angle of I2 to V1", 0)
    ic_phase_angle_rms_addr = memory_data_dict["real_time_parameters"].get("Phase Angle of I3 to V1", 0)

    pa_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Epa_imp Phase A import active energy", 0)
    pa_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Epa_exp Phase A expot active energy", 0)
    pb_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Epb_imp Phase B import active energy", 0)
    pb_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Epb_exp Phase B expot active energy", 0)
    pc_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Epc_imp Phase C import active energy", 0)
    pc_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Epc_exp Phase C expot active energy", 0)
    qa_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Eqa_imp Phase A import reactive energy", 0)
    qa_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Eqa_exp Phase A expot reactive energy", 0)
    qb_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Eqb_imp Phase B import reactive energy", 0)
    qb_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Eqb_exp Phase B expot reactive energy", 0)
    qc_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Eqc_imp Phase C import reactive energy", 0)
    qc_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Eqc_exp Phase C expot reactive energy", 0)
    sa_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Esa_imp Phase A import apparent energy", 0)
    sa_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Esa_exp Phase A export apparent energy", 0)
    sb_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Esb_imp Phase B import apparent energy", 0)
    sb_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Esb_exp Phase B export apparent energy", 0)
    sc_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Esc_imp Phase C import apparent energy", 0)
    sc_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Esc_exp Phase C export apparent energy", 0)
    sa_app_energy_addr = memory_data_dict["real_time_parameters"].get("Esa Phase A Apparent energy", 0)
    sb_app_energy_addr = memory_data_dict["real_time_parameters"].get("Esb Phase B Apparent energy", 0)
    sc_app_energy_addr = memory_data_dict["real_time_parameters"].get("Esc Phase C Apparent energy", 0)
    acc_start_time_energy_addr = memory_data_dict["real_time_parameters"].get("Energy accumulated start time", 0)
    acc_end_time_energy_addr = memory_data_dict["real_time_parameters"].get("Energy accumulated end time", 0)
    p_sys_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Ep_imp", 0)
    p_sys_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Ep_exp", 0)
    p_sys_total_energy_addr = memory_data_dict["real_time_parameters"].get("Ep_total", 0)
    p_sys_net_energy_addr = memory_data_dict["real_time_parameters"].get("Ep_net", 0)
    q_sys_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Eq_imp", 0)
    q_sys_exp_energy_addr = memory_data_dict["real_time_parameters"].get("Eq_exp", 0)
    q_sys_total_energy_addr = memory_data_dict["real_time_parameters"].get("Eq_total", 0)
    q_sys_net_energy_addr = memory_data_dict["real_time_parameters"].get("Eq_net", 0)
    s_sys_imp_energy_addr = memory_data_dict["real_time_parameters"].get("Es_imp", 0)
    s_sys_exp_energy_addr = memory_data_dict["real_time_parameters"].get("ES_exp", 0)
    s_sys_total_energy_addr = memory_data_dict["real_time_parameters"].get("Es_total", 0)

# class MemoryAddr:
#     real_time_parameters_by_start_addr_by_100ms = 0x1003
#     phase_angle_by_start_addr_by_100ms = 0x42A0
#     voltage_wire_addr = 0x1003
#     current_wire_addr = 0x1004
#     freq_rms_addr = 0x3000
#     ua_rms_addr = 0x3002
#     ub_rms_addr = 0x3004
#     uc_rms_addr = 0x3006
#     uv_avg_rms_addr = 0x3008
#     uab_rms_addr = 0x300A
#     ubc_rms_addr = 0x300C
#     uca_rms_addr = 0x300E
#     ul_avg_rms_addr = 0x3010
#     ia_rms_addr = 0x3012
#     ib_rms_addr = 0x3014
#     ic_rms_addr = 0x3016
#     iv_avg_rms_addr = 0x3018
#     in_rms_addr = 0x301A
#     pa_rms_addr = 0x301C
#     pb_rms_addr = 0x301E
#     pc_rms_addr = 0x3020
#     qa_rms_addr = 0x3024
#     qb_rms_addr = 0x3026
#     qc_rms_addr = 0x3028
#     sa_rms_addr = 0x302C
#     sb_rms_addr = 0x302E
#     sc_rms_addr = 0x3030
#     p_total_rms_addr = 0x3022
#     q_total_rms_addr = 0x302A
#     s_total_rms_addr = 0x3032
#     pf_a_rms_addr = 0x3034
#     pf_b_rms_addr = 0x3036
#     pf_c_rms_addr = 0x3038
#     pf_total_rms_addr = 0x303A
#     ua_phase_angle_rms_addr = 0x0000
#     ub_phase_angle_rms_addr = 0x42A0
#     uc_phase_angle_rms_addr = 0x42A1
#     ia_phase_angle_rms_addr = 0x42A2
#     ib_phase_angle_rms_addr = 0x42A3
#     ic_phase_angle_rms_addr = 0x42A4
