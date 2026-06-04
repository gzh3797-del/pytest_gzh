用户要在当前项目中适配一个新的电表设备。

请先询问以下信息（可一次提问）：
1. 设备型号（如 AcuRev5000）
2. Modbus TCP：IP 地址、端口、Unit ID
3. 寄存器类型：FC03（Holding，float32）还是 FC02（Discrete，DI型）
4. 是否支持 AcuCloud 比对

然后按以下步骤执行：

**第一步（必做）：列出地址表全部 Sheet**
用户提供 Modbus 地址表 Excel 路径后，先用 Python 列出所有 sheet 名称，再逐一判断哪些包含可读寄存器数据（FC01/FC02/FC03），确认全部覆盖后再写设备文件。禁止跳过此步骤直接写代码。

示例：
```python
import openpyxl, warnings
warnings.filterwarnings('ignore')
wb = openpyxl.load_workbook('<路径>', data_only=True)
print(wb.sheetnames)
wb.close()
```

**FC03 设备（正常适配）：**
- 读取 devices/acuiom01.py 或 devices/acurev4100.py 作为参考
- 创建 devices/<name>.py，实现 build_param_map()
- 若支持 Cloud，同时实现 build_cloud_col_map()
- 更新 config.py 的 MODBUS_DEVICE_MAP
- 更新 comparator.py 的 _DEVICE_MAP
- 若支持 Cloud，更新 cloud_comparator.py 的 _DEVICE_MAP
- 更新 README.md 支持设备表

**FC02 设备（DI型，暂不支持）：**
- 创建 devices/<name>.py，build_param_map() 返回空 {} 并注释说明原因
- 更新上述配置文件，标注 ⚠️

完成代码后，提示用户手动更新知识库：
- shared/devices/<name>.md（新建设备知识文件）
- shared/modbus_tables/INDEX.md（追加一行）
- shared/templates/INDEX.md（若有模板，追加一行）
