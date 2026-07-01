# Azure IoT 凭据目录

## 与 AWS IoT 的区别

AWS IoT 使用 X.509 证书（.pem / .key 文件）进行 TLS 双向认证。  
Azure IoT Hub 默认使用 **Connection String**（共享访问密钥）认证，无需本地证书文件。

## 需要填写的配置

认证信息存放在 `../config.yaml`，填写以下字段即可：

```yaml
azure_iot:
  # 设备连接串（从 Azure Portal > IoT Hub > 设备 > 连接字符串 获取）
  primary_conn_str: "HostName=<hub>.azure-devices.net;DeviceId=<id>;SharedAccessKey=<key>"
  secondary_conn_str: ""

  # Event Hub 兼容端点连接串（从 Azure Portal > IoT Hub > 内置端点 获取）
  # 用于本地订阅脚本 subscribe_messages.py 接收设备消息
  eventhub_conn_str: "Endpoint=sb://<namespace>.servicebus.windows.net/;SharedAccessKeyName=iothubowner;SharedAccessKey=<key>;EntityPath=<hub>"
```

## 如果使用 X.509 证书认证（可选）

若 Azure IoT Hub 设备配置了 X.509 证书认证，将证书文件放在本目录：

| 文件名 | 说明 |
|--------|------|
| `device.pem` | 设备证书（与 AWS `client.pem` 对应） |
| `device.key` | 设备私钥（与 AWS `key.pem` 对应） |
| `DigiCertGlobalRootG2.crt.pem` | Azure IoT 根 CA（与 AWS `AmazonRootCA1.pem` 对应） |

然后在 `config.yaml` 中增加：

```yaml
azure_iot:
  cert_file: "tests/protocols/azure_iot/certs/device.pem"
  key_file:  "tests/protocols/azure_iot/certs/device.key"
  ca_file:   "tests/protocols/azure_iot/certs/DigiCertGlobalRootG2.crt.pem"
```

## 注意事项

- **不要将真实密钥提交到 Git**。本目录下的凭据文件已在 `.gitignore` 中排除。
- Connection String 包含 SharedAccessKey，属于敏感信息，请勿明文写入版本控制系统。
