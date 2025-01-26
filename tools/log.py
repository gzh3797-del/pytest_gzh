# coding:utf-8
import logging
import time
import os
from tools.mkDir import mk_dir

root_path = str(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))).replace("\\", "/")
logging.getLogger("faker").setLevel(logging.ERROR)
logging.getLogger("urllib3").setLevel(logging.ERROR)


class Log(object):

    def __init__(self, current_exec_filename):
        """
        日志配置
        """
        now = time.strftime('%Y-%m-%d')
        details_time = time.strftime('%Y%m%d%H%M%S')
        log_path = root_path + '/logs' + "/" + now + "/"
        mk_dir(log_path)
        logfile = log_path + "{}_".format(current_exec_filename.split('.')[0]) + "{}.logs".format(details_time)
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.DEBUG)
        self.logger.handlers = []
        fh = logging.FileHandler(logfile, mode='a+', encoding='utf-8')
        formatter = logging.Formatter("%(levelname)-8s%(asctime)s  %(name)s:%(filename)s:%(lineno)d %(message)s")
        fh.setFormatter(formatter)
        self.logger.addHandler(fh)

# if __name__ == '__main__':
#     Log(str(__file__).split("\\")[-1])
#     logging.info("111222")
#     logging.error("111222")
#     logging.debug("111222")
#     logging.warning("111222")
