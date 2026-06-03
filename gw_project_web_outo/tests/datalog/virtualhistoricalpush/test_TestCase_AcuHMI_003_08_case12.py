import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_003_08_case12
# 用例标题：虚拟设备通过FTP Channel推送CSV格式datalog（1min间隔，UTC时间戳文件名）
# 预置条件：
#   1. 服务启动正常，账号登录成功
#   2. 已添加多个虚拟设备
# 测试步骤：
#   1. 配置Channel1为FTP，选择虚拟设备，LogFileFormat选择CSV，
#      LogFileLength选择1 min，LogFileNameFormat选择UTC Timestamp，
#      执行Post操作
# 预期结果：
#   1. Post成功推送到远端FTP服务器，FTP服务器收到CSV文件，文件命名符合UTC时间戳格式
@pytest.mark.skip(
    reason=(
        "需要搭建外部FTP/SFTP/HTTP服务器并配合虚拟设备历史数据推送，"
        "当前测试环境不满足"
    )
)
def test_TestCase_AcuHMI_003_08_case12(login_page: LoginPage):
    pass
