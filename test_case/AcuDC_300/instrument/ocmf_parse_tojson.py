"""
文件名称: ocmf_parse_tojson.py
功能描述: 解析交易日志并存在json中
创建日期: 2025-08-08
作者: 苏博
版本: v1.0
修改记录:
"""
import json
from comm.modbus_set_attr import read_ocmf_transactionlog
import time
from tools.mkDir import mk_dir
from tools.log import root_path


def ocmf_parse_tojson():
    ret = read_ocmf_transactionlog()
    log_path = root_path + '/test_case/AcuDC_300/instrument/data/log/'
    mk_dir(log_path)
    log_path = root_path + '/test_case/AcuDC_300/instrument/data/sd/'
    mk_dir(log_path)
    with open('./data/log/ocmf_log_{}.json'.format(time.strftime('%Y%m%d%H%M%S')), 'w') as f:
        mes = json.loads(ret.split('|')[1])
        json.dump(mes, f, indent=4)
    with open('./data/sd/ocmf_sd_{}.json'.format(time.strftime('%Y%m%d%H%M%S')), 'w') as f1:
        mes1 = json.loads(ret.split('|')[2])
        json.dump(mes1, f1, indent=4)
