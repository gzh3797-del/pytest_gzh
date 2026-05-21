# web2 Bug 索引

原始 Jira 导出 CSV 存放于 raw/ 目录（权威来源）。
本文件为精简索引，供快速了解历史问题与当前状态。

**最后同步：2026-05-13 | 覆盖：A4WS-135 ~ A4WS-197 | 缺陷共 60 条**

---

## 产品 Bug（Jira）

### 统计概览

| 状态 | 数量 |
|------|------|
| CLOSED 已关闭 | 29 |
| IN PROGRESS 修复中 | 10 |
| CREATED 未分配 | 10 |
| TO BE VERIFIED 待验证 | 8 |
| SELF-TESTING 自测中 | 1 |
| REJECTED 已拒绝 | 2 |
| **未关闭合计** | **29**（排除 REJECTED） |

所有缺陷均为 Medium 优先级。

### 各模块活跃 Bug 数

| 模块 | 总计 | 已关闭 | 活跃 |
|------|------|--------|------|
| AWS IoT / Azure IoT / AcuCloud | 15 | 11 | 4 |
| BACnet/IP | 12 | 9 | 3 |
| 接线检查（Wiring Check） | 8 | 2 | 6 |
| 固件升级 / 降级 | 6 | 3 | 3 |
| 设备配置 / User Channel / Max-Min | 5 | 3 | 2 |
| 密码 / 安全 / EULA | 3 | 0 | 3 |
| Troubleshooting / Remote Access | 3 | 2 | 1 |
| MQTT | 2 | 0 | 2 |
| SNMP | 2 | 0 | 2 |
| Post Channel / Datalog | 1 | 0 | 1 |
| EtherNet/IP | 1 | 1 | 0 |
| Modbus | 1 | 0 | 1 |
| 系统稳定性 / Coredump | 1 | 0 | 1 |

### 重点关注：CREATED（未分配，尚未开始修复）

| ID | 模块 | 摘要 |
|----|------|------|
| A4WS-196 | AcuCloud | 同时上传 4100+IOM 参数导致全部参数无法上传 |
| A4WS-193 | AWS IoT | 全选参数后客户端未收到任何数据 |
| A4WS-190 | SNMP | MIB 文件参数数量（685）与模板（1059）不一致 |
| A4WS-184 | 固件升级 | 升级 system log 中 SN 号乱码 |
| A4WS-181 | BACnet/IP | 上传 1869 个参数耗时 40 分钟（性能严重异常） |
| A4WS-178 | Modbus | 默认端口显示 6000，预期应为 502 |
| A4WS-169 | 接线检查 | 2E3W1 接线方式下检查结果错误 |
| A4WS-168 | 接线检查 | 2E3WD 接线方式下检查结果错误 |
| A4WS-167 | 接线检查 | 2E3WN 接线方式下检查结果错误 |
| A4WS-164 | 接线检查 | 3E4WY 接线方式下检查结果错误（注：4条接线检查 bug 建议合并排查同一底层逻辑） |

### 待验证（TO BE VERIFIED，已修复待测试确认）

A4WS-192、A4WS-191、A4WS-189、A4WS-188、A4WS-185、A4WS-161、A4WS-160、A4WS-140

### 全量索引（60 条，按密钥降序）

| ID | 模块 | 状态 | 摘要 |
|----|------|------|------|
| A4WS-197 | MQTT | IN PROGRESS | MQTT 参数选择界面概率显示异常，无法选择参数 |
| A4WS-196 | AcuCloud | CREATED | 同时上传 4100+IOM 参数，所有参数无法上传 |
| A4WS-195 | 固件降级 | IN PROGRESS | 固件降级页面无任何提示，切换后显示"正在升级"无法操作 |
| A4WS-194 | 密码/安全 | IN PROGRESS | 临时密码提示语 "for today" 误写为 "fortoday" |
| A4WS-193 | AWS IoT | CREATED | 全选 4100 参数，客户端未收到任何数据 |
| A4WS-192 | 设备配置 | TO BE VERIFIED | User and CT 配置后 Realtime userChannel 描述不一致 |
| A4WS-191 | 密码/安全 | TO BE VERIFIED | 默认密码登录后点 YES 未自动跳转修改密码页面 |
| A4WS-190 | SNMP | CREATED | MIB 文件参数 685 个，与模板 1059 个不一致 |
| A4WS-189 | EULA | TO BE VERIFIED | EULA 章节编号错误，"related application" 单复数错误 |
| A4WS-188 | 设备配置 | TO BE VERIFIED | Max Min 数据时间日期与上位机不一致 |
| A4WS-186 | SNMP | IN PROGRESS | 配置页面提示语 "Invali" 拼写错误（应为 Invalid） |
| A4WS-185 | BACnet/IP | TO BE VERIFIED | COV Increment 非法值时切换 Parameter Type 提示信息错误 |
| A4WS-184 | 固件升级 | CREATED | 设备升级 system log 中 SN 号乱码 |
| A4WS-183 | AWS IoT | IN PROGRESS | 切换 WEB2 模式后亚马逊平台已保存主表未显示选中 |
| A4WS-182 | 固件升级 | IN PROGRESS | 一台设备升级包校验异常后整个升级界面无法恢复 |
| A4WS-181 | BACnet/IP | CREATED | 上传 1869 个参数耗时 40 分钟，速度严重异常 |
| A4WS-180 | AWS IoT | SELF-TESTING | AWS IoT 推送数据中 name/model 字段未显示 |
| A4WS-179 | Post Channel | IN PROGRESS | SFTP 选项下提示语误显示 "FTP" |
| A4WS-178 | Modbus | CREATED | 默认端口显示 6000（范围 2000-5999），预期应为 502 |
| A4WS-177 | MQTT | IN PROGRESS | MQTT 推送数据客户端收到的 SN 显示不正确 |
| A4WS-176 | 接线检查 | IN PROGRESS | 接线检查结果 Wiring Configuration 列未合并 |
| A4WS-175 | Troubleshooting | IN PROGRESS | Troubleshooting Remote Access URL 无法正常访问 |
| A4WS-174 | 固件升级 | CLOSED | 下挂表升级可上传 IOM 包给 4100 导致失败 |
| A4WS-173 | 固件升级 | CLOSED | 升级过程切换页面提示 SN 号乱码 |
| A4WS-171 | 接线检查 | CLOSED | Show Only Issues 过滤下 Voltage 仍显示 pass 设备 |
| A4WS-170 | 接线检查 | REJECTED | 离线设备显示历史结果而非 "Offline" |
| A4WS-169 | 接线检查 | CREATED | 2E3W1 接线方式下检查结果错误 |
| A4WS-168 | 接线检查 | CREATED | 2E3WD 接线方式下检查结果错误 |
| A4WS-167 | 接线检查 | CREATED | 2E3WN 接线方式下检查结果错误 |
| A4WS-166 | 固件升级 | CLOSED | 升级页面 URL 参数可暴露并修改升级 URL |
| A4WS-165 | Troubleshooting | CLOSED | 下载 PDF 邮箱与页面显示不一致 |
| A4WS-164 | 接线检查 | CREATED | 3E4WY 接线方式下检查结果错误 |
| A4WS-163 | Troubleshooting | CLOSED | AIO 设备 Download Current 未下载 PDF |
| A4WS-162 | User Channel | CLOSED | Max Min 页面 User Channel 名称重复显示 |
| A4WS-161 | 接线检查 | TO BE VERIFIED | WEB Module 模式下设备未隐藏 |
| A4WS-160 | BACnet/IP | TO BE VERIFIED | YABE 接收不到 COV 告警 |
| A4WS-159 | User Channel | REJECTED | THD/Harmonic 页面 User Channel 名称未显示 |
| A4WS-158 | Azure IoT | CLOSED | 取消设备上传后数据仍继续推送 |
| A4WS-157 | Azure IoT | CLOSED | 二次上传证书后仅保留第一次 |
| A4WS-156 | Azure IoT | CLOSED | 数据重复上报 |
| A4WS-155 | BACnet/IP | CLOSED | YABE 参数列表存在历史参数未同步 |
| A4WS-154 | BACnet/IP | CLOSED | EPICS 文件设备内容未随配置同步更新 |
| A4WS-153 | BACnet/IP | CLOSED | 开启 BACnet/IP 后参数上传值全为 0 |
| A4WS-152 | AWS IoT | CLOSED | Disable 后仍缓存并重新推送数据 |
| A4WS-151 | BACnet/IP | CLOSED | YABE Device Object Name/Instance 与界面不一致 |
| A4WS-150 | BACnet/IP | CLOSED | 保存配置后需重启才能与 YABE 通信 |
| A4WS-149 | BACnet/IP | CLOSED | Device Instance=4194302 时页面报错 |
| A4WS-148 | BACnet/IP | CLOSED | 首次开启后 YABE 扫描不到设备，需重启 WEB2 |
| A4WS-147 | AWS IoT | CLOSED | 多设备只上传了一个设备参数 |
| A4WS-146 | BACnet/IP | CLOSED | BACnet Port/BBMD Port 范围与需求不一致 |
| A4WS-145 | AWS IoT | CLOSED | 虚拟设备参数未上传 |
| A4WS-144 | AWS IoT | CLOSED | 取消上传后数据仍继续推送 |
| A4WS-143 | AWS IoT | CLOSED | 二次上传证书后仅保留第一次 |
| A4WS-142 | AWS IoT | CLOSED | 数据重复上报 |
| A4WS-141 | AWS IoT | CLOSED | Test Connection 提示连接失败 |
| A4WS-140 | 系统稳定性 | TO BE VERIFIED | 页面报错后台出现 coredump |
| A4WS-139 | AWS IoT | CLOSED | Parameters 列表初次打开不显示 |
| A4WS-138 | EtherNet/IP | CLOSED | EDS 文件参数范围与模板不一致 |
| A4WS-136 | 设备配置 | CLOSED | PT1/PT2 配置时 PT1>PT2 未校验可保存 |
| A4WS-135 | BACnet/IP | CLOSED | COV Increment 默认不显示 0.000 |

---

## 测试框架 Bug（测试代码层，非产品）

| 模块 | 标题 | 状态 | 根因摘要 |
|------|------|------|---------|
| AcuvimIIR | cloud_col_map 未导入导致 Cloud 比对报错 | 已修复 | acuvimiir.py 漏掉 build_cloud_col_map；改为同时导入两个函数 |
| AcuIOM-03/04 | DI 型号 FC02 不支持，比对直接跳过 | 已知限制 | FC02 未实现；创建空 stub 模块返回 {}，待后续支持 |
| 网关连接 | BAC0 启动后网关响应慢，频繁超时 | 已修复 | 网关性能差；READ_TIMEOUT 10→30s，MAX_RETRIES 2→4，CONNECT_WAIT 0.5→3.0s |
| 模板解析 | IOM 模板无 BACnetIP 列，get_bacnet_params 抛 ValueError | 已处理 | IOM 模板格式不同；try/except 捕获后 log warning 跳过范围检查继续比对 |
