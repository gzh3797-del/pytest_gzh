# Acuview QT 自动化测试框架使用说明

## 项目概述

这是一个基于Python的自动化测试框架，专门用于测试Acuview上位机。框架集成了UI自动化、Modbus通信、图像识别等功能，支持完整的充电交易流程测试。

## 项目结构

```
Acuview_QT_Auto/
├── config/                    # 配置文件目录
│   ├── devices_config.json    # 设备配置
│   └── pos_config.json        # 坐标配置
├── page_elements/             # 页面元素截图
├── test_case/                 # 测试用例
│   └── test_transaction.py    # 交易相关测试用例
├── utils/                     # 工具类
│   ├── common_utils.py        # 公共工具类
│   ├── ModbusClient.py        # Modbus客户端
│   └── QT_auto_utils.py       # UI自动化工具
├── test_reports/              # 测试报告
└── run_test.py               # 测试运行入口
```

## 环境要求

### Python版本
- Python 3.7+

### 第三方库依赖

```bash
# 核心自动化库
pip install pyautogui
pip install opencv-python
pip install pytesseract
pip install pillow

# 测试框架
pip install pytest
pip install pytest-html

# 系统工具
pip install psutil
pip install pyperclip

# 串口通信
pip install pyserial

# 数据处理
pip install pandas
```

### 额外软件要求

1. **Tesseract OCR**
   - 下载地址: https://github.com/UB-Mannheim/tesseract/wiki
   - 安装后配置路径: `C:\Program Files\Tesseract-OCR\tesseract.exe`

2. **Acuview 2 应用程序**
   - 默认路径: `C:\Users\YiSong\Acuview2\Acuview 2.exe`

## 快速开始

### 1. 环境配置

1. 安装所有必需的Python库
2. 安装Tesseract OCR并配置正确路径
3. 确保Acuview 2应用程序已安装
4. 检查配置文件中的设备连接参数
5. 检查页面元素是否存放在page_elements

### 2. 运行测试

```bash
# 运行所有测试
python run_test.py

# 运行特定测试文件
pytest test_case/test_transaction.py -v

# 生成HTML报告
pytest test_case/test_transaction.py --html=report.html --self-contained-html
```

## 核心工具类说明

### 1. AutoHelper (QT_auto_utils.py)

UI自动化核心工具类，提供屏幕操作、图像识别等功能。

#### 常用方法

**应用管理**
```python
# 启动应用
helper.launch_app(app_path, timeout=30)

# 关闭Acuview进程
helper.kill_acuview_apps()

# 重启应用
helper.restart_application()
```

**设备连接**
```python
# 连接设备
helper.connect_device(device_image_path, timeout=5)

# 检查设备连接状态
helper.check_image_exists(image_path)
```

**图像识别与点击**
```python
# 点击指定图片
helper.click_image(image_path, index=0, offset_x=0, offset_y=0)

# 检查图片是否存在
helper.check_image_exists(image_path, timeout=2)

# 点击指定坐标
helper.click_pos((x, y))
helper.double_click_pos((x, y))
```

**文本操作**
```python
# 输入文本
helper.type_text("text")

# 粘贴文本
helper.paste_text("text")

# 快捷键操作
helper.hotkey('ctrl', 'c')
```

**OCR识别**
```python
# 区域OCR识别
result = helper.extract_and_ocr_by_coordinates(
    top_left=(x1, y1), 
    bottom_right=(x2, y2)
)

# 基于配置的OCR
text = helper.quick_ocr_by_config('配置名称')
```

**等待操作**
```python
# 等待指定时间
helper.wait(seconds)
```

### 2. CommonUtils (common_utils.py)

测试业务逻辑封装类，提供具体的配置和测试操作。

#### 常用方法

**设备配置**
```python
# 连接设备
utils.connect_device()

# 配置识别状态
utils.configure_identification_status(True/False)

# 配置识别级别
utils.configure_identification_level('TRUSTED')

# 配置识别标志
utils.configure_identification_flag1('RFID_PLAIN')
utils.configure_identification_flag2('OCPP_RS')
utils.configure_identification_flag3('ISO15118_NONE')
utils.configure_identification_flag4('PLMN_RING')

# 配置识别类型
utils.configure_identification_type('UNDEFINED')

# 配置识别数据
utils.configure_identification_data('4525asddf$%@!')

# 配置费率文本
utils.configure_Tariff_Text('4525asddf$%@!')
```

**交易操作**
```python
# 执行完整充电循环
start_time, end_time = utils.perform_charging_cycle()

# 开始充电
utils.start_charging()

# 结束充电
utils.end_charging()

# 终止充电
utils.abort_charging()
```

**交易日志管理**
```python
# 读取并解析交易日志
log_data = utils.read_and_parse_transaction_log()

# 构造交易日志
utils.construct_transaction_logs(数量)

# 清除交易日志
utils.clear_transaction_logs()

# 读取交易日志
utils.read_transaction_logs()
```

**系统操作**
```python
# 重启应用
utils.restart_application()

# 重启设备
utils.reboot_device()

# 时间同步
utils.configure_time_Sync_status()
```

### 3. ModbusClient (ModbusClient.py)

Modbus通信客户端，支持TCP和RTU协议。

#### 常用方法

```python
# 创建客户端
with ModbusClient(ModbusProtocol.TCP) as client:
    # 寄存器操作
    response = client.validate_register_value('寄存器描述', 值)
    
    # 自定义报文
    response = client.send_custom_message('00 01 00 00 00 09 01 03 ...')
    
    # 批量操作
    results = client.batch_validate(commands_list)
```

## 配置文件说明

### devices_config.json
配置Modbus连接参数：
```json
{
    "tcp": {
        "host": "192.168.1.100",
        "port": 502,
        "timeout": 5,
        "slave_id": 1
    },
    "rtu": {
        "port": "COM1",
        "baudrate": 9600,
        "bytesize": 8,
        "parity": "N",
        "stopbits": 1,
        "timeout": 5,
        "slave_id": 1
    }
}
```

### pos_config.json
配置屏幕坐标，用于OCR识别：
```json
{
    "Max Transaction Id": {
        "top_left": [100, 200],
        "bottom_right": [300, 250],
        "lang": "eng",
        "description": "最大交易ID"
    }
}
```

## 测试用例编写示例

```python
import pytest
from utils.QT_auto_utils import AutoHelper
from utils.common_utils import CommonUtils

class TestExample:
    @pytest.fixture(autouse=True)
    def setup(self):
        self.helper = AutoHelper()
        self.utils = CommonUtils(self.helper, app_path, device_image_path)
        
    def test_example(self):
        # 连接设备
        self.utils.connect_device()
        
        # 配置参数
        self.utils.configure_identification_status(True)
        
        # 执行充电
        self.utils.perform_charging_cycle()
        
        # 验证结果
        log_data = self.utils.read_and_parse_transaction_log()
        result = self.helper.parse_ocmf(log_data)
        
        assert result['IS'] == True
```

## 故障排除

### 常见问题

1. **图像识别失败**
   - 检查屏幕分辨率是否匹配
   - 调整置信度参数 `confidence`
   - 确保页面元素截图准确

2. **Modbus连接失败**
   - 检查设备IP地址和端口
   - 验证从站ID配置
   - 检查网络连接

3. **OCR识别不准确**
   - 调整识别区域坐标
   - 尝试不同的语言包
   - 预处理图像提高对比度

4. **应用启动失败**
   - 检查应用路径是否正确
   - 确保没有其他Acuview进程运行
   - 检查系统权限

### 日志查看

框架使用Python logging模块，可以通过修改日志级别来获取更详细的调试信息：

```python
import logging
logging.getLogger().setLevel(logging.DEBUG)
```

## 注意事项

1. 运行测试前确保Acuview应用程序已关闭
2. 测试过程中不要移动鼠标，以免干扰自动化操作
3. 确保测试环境网络稳定
4. 重要的测试数据建议提前备份
5. 修改配置后记得验证配置的正确性

## 技术支持

如有问题请检查：
1. 所有依赖库是否安装正确
2. 配置文件路径和参数是否正确
3. 设备连接状态是否正常
4. 查看生成的日志文件获取详细错误信息