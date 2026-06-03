import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_003_08_case13
# 用例标题：虚拟设备通过SFTP Channel推送CSV格式datalog
# 预置条件：
#   1. 服务启动正常，账号登录成功
#   2. 已添加多个虚拟设备
# 测试步骤：
#   1. 配置Channel1为SFTP，选择虚拟设备，LogFileFormat选择CSV，
#      LogFileLength选择5 mins，LogFileNameFormat选择Time interval Format，
#      执行Post操作
# 预期结果：
#   1. Post成功推送到远端SFTP服务器，SFTP服务器收到CSV文件
@pytest.mark.skip(
    reason=(
        "需要搭建外部FTP/SFTP/HTTP服务器并配合虚拟设备历史数据推送，"
        "当前测试环境不满足"
    )
)
def test_TestCase_AcuHMI_003_08_case13(login_page: LoginPage):
    pass
