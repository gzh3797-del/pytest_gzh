import json, sys
sys.stdout.reconfigure(encoding='utf-8')

# MQTT data push cases (rows 420-429): code-build broker, UI configure device params
mqtt_data_rows = [420, 421, 422, 423, 424, 425, 426, 427, 428, 429]
mqtt_config_rows = [430, 431, 432, 433, 434, 435, 436, 437, 438, 439, 441, 442, 443, 444, 445, 446]
snmp_ui_rows = [448, 450]
snmp_nms_rows = [451, 452, 453, 454, 455, 456, 457, 458, 459, 460, 461, 462, 463, 466]
snmp_net_row = [464]   # network recovery - ORANGE
snmp_restart_row = [465]  # restart - RED
snmp_trap_rows = [467, 468, 469]  # trip alarm - RED
snmp_nms_tool_row = [449]  # third-party NMS - RED

aws_param_rows = [520, 521]
aws_cert_legal = [522]
aws_cert_illegal = [523]
aws_device_push = [524, 525, 526]
aws_longrun_rows = [529, 530, 531]  # RED
aws_disable_enable = [533]
aws_both_azure = [534]  # RED - no Azure
aws_multidevice = [535]
aws_offline = [536]  # RED

azure_secondary = [540]
azure_interval = [541]
azure_cert_ssl = [543]
azure_cert_invalid = [544, 545, 546]
azure_modbus = [547]
azure_bacnet = [548]  # RED - deprecated
azure_virtual = [549]
azure_multi = [550]
azure_twin_legal = [552]
azure_twin_illegal = [553]
azure_disconnect = [554]  # RED
azure_primary_fail = [555, 556]
azure_disable_enable = [557, 558, 559]
azure_offline = [560]  # RED

write_data = []

# MQTT data push - GREEN
for row in mqtt_data_rows:
    write_data.append({'row': row, 'color': '006400',
        'text': '已理解：UI配置MQTT连接设备和参数后，验证客户端能收到推送数据 | 断言方式：代码搭建MQTT Broker(paho-mqtt/mosquitto)，UI Topic and Parameter Selection页面选择设备和参数，配置连接并保存，断言Broker订阅的topic收到包含该设备参数的消息且数值合理'})

# MQTT config validation - GREEN
for row in mqtt_config_rows:
    write_data.append({'row': row, 'color': '006400',
        'text': '已理解：MQTT配置参数验证（有效/无效/非法值），通过Test MQTT按钮验证连接结果 | 断言方式：代码搭建MQTT Broker，在General/User Credential/SSL页面填写参数，点击Test MQTT按钮，断言页面提示连接成功或失败消息符合预期（有效值→成功，无效/非法→失败+错误提示）'})

# SNMP UI-only - GREEN
for row in snmp_ui_rows:
    write_data.append({'row': row, 'color': '006400',
        'text': '已理解：仅需UI页面参数输入验证，无需NMS工具 | 断言方式：在SNMP UI页面（Report Buffer Size/Hold Time字段）输入边界值和非法值，断言输入框显示错误提示或保存被阻止，合法值正常保存'})

# SNMP NMS manager cases - GREEN
for row in snmp_nms_rows:
    write_data.append({'row': row, 'color': '006400',
        'text': '已理解：开启SNMP并配置版本/端口/Community/Auth参数，通过代码搭建NMS管理端验证数据接收 | 断言方式：代码搭建NMS(python-netsnmp/pysnmp)，UI配置SNMP连接参数，断言NMS端能成功接收设备MIB数据（或端口/Community不匹配时请求失败）'})

# SNMP network recovery - ORANGE
for row in snmp_net_row:
    write_data.append({'row': row, 'color': 'FFA500',
        'text': '[半自动化] 需人工介入步骤：需模拟网络异常（断开网络连接）。可自动化部分：代码搭建NMS管理端，UI配置SNMP，NMS正常接收数据后模拟网络断开（如停止NMS服务），网络恢复后重启NMS，断言NMS重新请求成功且数据正常返回。'})

# SNMP restart - RED
for row in snmp_restart_row:
    write_data.append({'row': row, 'color': 'FF0000',
        'text': '[用户决定不实现自动化] 用户答复：需要模拟断电，暂不做自动化（重启设备破坏测试环境）。'})

# SNMP Trap - RED
for row in snmp_trap_rows:
    write_data.append({'row': row, 'color': 'FF0000',
        'text': '[用户决定不实现自动化] 用户答复：Trip告警信息需要模拟设备断网或下点，暂不做自动化。'})

# SNMP NMS tool - RED
for row in snmp_nms_tool_row:
    write_data.append({'row': row, 'color': 'FF0000',
        'text': '[用户决定不实现自动化] 用户答复：需使用第三方NMS工具，暂不做自动化。'})

# AWS IoT param validation - GREEN
for row in aws_param_rows:
    write_data.append({'row': row, 'color': '006400',
        'text': '已理解：AWS IoT Topic/Interval参数合法性校验（有效值保存、非法值被拒绝） | 断言方式：测试环境已配置AWS IoT连接参数，在Topic/Interval输入框填写合法和非法值，断言合法值保存成功，非法值显示错误提示或保存被拒'})

# AWS legal cert - GREEN
write_data.append({'row': 522, 'color': '006400',
    'text': '已理解：合法证书和密钥连接AWS IoT Core成功 | 断言方式：测试环境已配置好证书和AWS连接参数，扫描当前页面参数确认已配置，点击Test Connection按钮，断言页面显示连接成功提示'})

# AWS illegal cert - GREEN
write_data.append({'row': 523, 'color': '006400',
    'text': '已理解：非法证书或错误密钥连接AWS IoT失败 | 断言方式：代码生成错误证书/密钥文件，通过UI上传，点击Test Connection，断言页面显示连接失败错误提示（而非连接成功）'})

# AWS device data push - GREEN
for row in aws_device_push:
    write_data.append({'row': row, 'color': '006400',
        'text': '已理解：选择物理/虚拟设备连接AWS IoT并上报参数数据 | 断言方式：测试环境已配置AWS IoT连接参数，UI勾选设备配置参数，执行shell脚本（用户可提供）模拟客户端接收数据，断言数据中UTC时间戳符合Interval间隔，JSON字段包含设备参数'})

# AWS long-running - RED
for row in aws_longrun_rows:
    write_data.append({'row': row, 'color': 'FF0000',
        'text': '[用户决定不实现自动化] 用户答复：长时间运行脚本（24h/72h断网测试）暂不做自动化，后期可实现专项挂测脚本。'})

# AWS disable/enable - GREEN
write_data.append({'row': 533, 'color': '006400',
    'text': '已理解：禁用AWS IoT后停止发布，重新启用后恢复上报 | 断言方式：执行shell脚本监控数据，在UI将AWS IoT Enable切换为Disable，断言脚本收不到新消息；再切回Enable，断言恢复接收数据'})

# AWS + Azure - RED
write_data.append({'row': 534, 'color': 'FF0000',
    'text': '[用户决定不实现自动化] 用户答复：暂无Azure IoT平台账号，暂不做自动化（AWS IoT与Azure IoT同时启用互不干扰测试需两个平台）。'})

# AWS multi-device - GREEN
write_data.append({'row': 535, 'color': '006400',
    'text': '已理解：多设备多参数同时发布，系统性能正常 | 断言方式：UI勾选全部设备配置参数，执行shell脚本监控数据，断言脚本收到所有勾选设备的数据推送且系统无错误提示'})

# AWS device offline - RED
write_data.append({'row': 536, 'color': 'FF0000',
    'text': '[用户决定不实现自动化] 用户答复：需要模拟设备离线状态，暂不做自动化。'})

# Azure secondary connection string - GREEN
write_data.append({'row': 540, 'color': '006400',
    'text': '已理解：Secondary Connection String配置及Primary失效时切换 | 断言方式：测试环境已配置好Azure IoT连接参数，UI配置Primary和Secondary Connection String，清空Primary模拟失效，断言系统自动切换到Secondary连接正常'})

# Azure interval - GREEN
write_data.append({'row': 541, 'color': '006400',
    'text': '已理解：Azure IoT Interval最小值（10s）配置保存及数据上报间隔验证 | 断言方式：UI设置Interval=10s并保存，python脚本客户端（用户提供）监控设备上传数据，断言数据中UTC时间戳间隔约为10s'})

# Azure X509 cert SSL - GREEN
write_data.append({'row': 543, 'color': '006400',
    'text': '已理解：合法X509证书启用SSL连接Azure IoT Hub成功 | 断言方式：测试环境准备好证书并在UI配置，验证SSL连接状态为成功（Test Connection或查看连接状态指示）'})

# Azure cert invalid cases - GREEN
for row in azure_cert_invalid:
    write_data.append({'row': row, 'color': '006400',
        'text': '已理解：非法证书/密钥格式阻止上传或证书密钥不匹配连接失败 | 断言方式：python脚本生成格式错误的证书/密钥文件，UI上传时断言系统拒绝（提示格式错误）；或上传不匹配证书，点击连接，断言连接失败错误提示'})

# Azure modbus device - GREEN
write_data.append({'row': 547, 'color': '006400',
    'text': '已理解：选择Modbus设备发布数据按Interval间隔上报Azure IoT | 断言方式：测试环境配置好Azure连接，UI点击connection连接成功，勾选AcuRev4100设备配置参数，执行python客户端脚本，断言收到的JSON数据中设备参数存在且时间戳间隔符合Interval配置'})

# Azure BACnet - RED (deprecated)
write_data.append({'row': 548, 'color': 'FF0000',
    'text': '[用户决定废弃] 用户答复：废弃，该用例不再实现。'})

# Azure virtual device - GREEN
write_data.append({'row': 549, 'color': '006400',
    'text': '已理解：选择Virtual设备发布数据按间隔上报Azure IoT | 断言方式：测试环境配置好Azure连接，UI勾选Virtual设备配置参数，执行python客户端脚本，断言收到的JSON数据包含Virtual设备参数'})

# Azure multi-type devices - GREEN
write_data.append({'row': 550, 'color': '006400',
    'text': '已理解：多类型设备批量发布，各设备数据独立接收 | 断言方式：测试环境配置好Azure连接，UI勾选全部设备（Devices Selection）配置参数，python客户端脚本监控，断言每类设备的数据独立接收且互不干扰'})

# Azure device twin legal - GREEN
write_data.append({'row': 552, 'color': '006400',
    'text': '已理解：设备孪生下发合法配置变更（Interval值），验证变更生效 | 断言方式：测试环境配置好Azure连接，下发孪生配置变更Interval值，执行python监控脚本，断言设备推送数据的UTC时间戳间隔变为新的Interval值'})

# Azure device twin illegal - GREEN
write_data.append({'row': 553, 'color': '006400',
    'text': '已理解：设备孪生下发非法配置值，设备拒绝变更 | 断言方式：下发非法Interval值（如负数或超出范围），断言设备孪生状态不更新为非法值，设备继续以原Interval推送数据'})

# Azure disconnect - RED
write_data.append({'row': 554, 'color': 'FF0000',
    'text': '[用户决定不实现自动化] 用户答复：需要模拟断网，暂不做自动化。'})

# Azure primary fail switch - GREEN
write_data.append({'row': 555, 'color': '006400',
    'text': '已理解：Primary失效时自动切换Secondary继续上报 | 断言方式：配置Primary和Secondary连接字符串，清空Primary Connection String模拟失效，断言设备自动切换至Secondary并继续正常推送数据（python脚本监控）'})

# Azure both fail - GREEN
write_data.append({'row': 556, 'color': '006400',
    'text': '已理解：Primary/Secondary均失效时连接失败，本地缓存保留 | 断言方式：清空Primary和Secondary两个Connection String，断言连接失败提示；验证本地缓存数据保留（配置不丢失），重新配置连接字符串后恢复正常'})

# Azure disable/enable + multi + performance - GREEN
for row in azure_disable_enable:
    write_data.append({'row': row, 'color': '006400',
        'text': '已理解：Azure IoT禁用/启用/多设备性能场景，测试环境已配置连接参数，按步骤执行 | 断言方式：测试环境配置好Azure IoT连接，UI操作Enable/Disable开关或多设备配置，python客户端脚本监控，断言Enable时收到数据、Disable时停止收到数据，操作符合预期'})

# Azure device offline - RED
write_data.append({'row': 560, 'color': 'FF0000',
    'text': '[用户决定不实现自动化] 用户答复：需要模拟设备离线，影响其他业务，暂不做自动化。'})

with open('tmp_protocol_write.json', 'w', encoding='utf-8') as f:
    json.dump(write_data, f, ensure_ascii=False, indent=2)
print(f'Prepared {len(write_data)} items for 设备数据协议转换')
