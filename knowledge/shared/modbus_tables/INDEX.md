# Modbus 地址表索引

原始 Excel 文件存放于 raw/ 目录，代码运行时直接读取（权威来源）。
Claude 参考本文件了解各设备寄存器范围，无需读原始 Excel。

| 设备 | Excel 文件（raw/） | FC | 数据类型 | 特殊说明 |
|------|------------------|----|---------|---------|
| AcuRev4100 | AcuRev4100 Modbus Address Table v1.02 20260202.xlsx | FC03 | Float32 | 端口2000，非标准 |
| AcuRev2100 | AcuRev2100_ Modbus Address_v1.02_20260406.xlsx | FC03 | Float32 | |
| AcuVIM3 | Acuvim3 User Modbus Address Table v1.08_Hongjian Zhu_260203.xlsx | FC03 | Float32 | |
| AcuvimIIW | Acuvim IIW&IIR&CL&EL Modbus Address v1.27_Haibo Song_260323.xlsx | FC03 | Float32 | 同文件含 IIR / CL(PXE1) / EL(PXE2)，后两者目前不适配 |
| AcuvimIIR | 同上 | FC03 | Float32 | 与IIW完全相同，共用文件 |
| AcuRev1300 | AcuRev1310_PXM350_ Modbus Address_v1.01_Sam Xu_260305.xlsx | FC03 | Float32 | 文件名中 AcuRev1310 为笔误（实为1300）；与 PXM350 为同一设备，共用文件 |
| AcuIOM-01 | AcuIOM Modbus Address Table v1.01 20260228 的副本.xlsx | FC03 | Float32 | 01~04 共用同一文件 |
| AcuIOM-02 | 同上 | FC03 | Float32 | |
| AcuIOM-03 | 同上 | FC02 | Bit | DI型号，框架暂不支持 |
| AcuIOM-04 | 同上 | FC02 | Bit | DI型号，框架暂不支持 |

## 更新说明
新增设备后在本表追加一行，并将原始 Excel 放入 raw/ 目录。
