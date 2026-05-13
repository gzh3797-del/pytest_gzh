# Home 首页

**截图**：`screenshots/03_home.png`

## 页面布局

| 区域 | 内容 |
|------|------|
| 顶部 | 组织 Logo（空） + 组织名称 + 右上组织选择器 |
| 告警横幅 | 离线设备检测警告（黄色横幅） |
| 主体 | 设施地图（Google Maps）|

## 离线设备告警

首页显示以下告警（测试环境数据）：

> **Offline Devices Detected**  
> There are **1878 devices** currently offline (you can check the details in the Device list), and **1875** of them do not have offline alerts configured.  
> Please go to **Configure Offline Alerts** to ensure timely handling.

- 告警类型：组织级别离线设备汇总
- 提示用户前往设备列表查看，并配置离线告警

## 设施地图

- 使用 Google Maps API（key: AIzaSyCPAngAAJdG4MBOMthkQ-jYsGn4G55REvM）
- 显示当前组织旗下所有设施的地理位置标点
- 测试环境地图视角：南美洲/非洲区域（反映测试设施分布）

## Organization View 切换

右上角 **Organization View** 开关：
- **开启**：显示组织视角（当前页面）
- **关闭**：切换为设施视角（单个设施的数据视图）

## 当前组织选择器

右上角下拉框可切换当前操作的组织（多组织账号）：
- 当前：AG PROYECTOS Y SERVICIOS, S.A.

## 示例设施列表（来自权限 API /permission/current）

部分可访问设施：

| facilityId | facilityName | 城市 | 时区 |
|------------|-------------|------|------|
| 11 | Whitestar | Stillwater | UTC |
| 15 | Tesla | Johnston | UTC |
| 16 | CLIENSOL OFFICE | Caldetas | UTC |
| 18 | Clements Ferry Warehouse | Charleston | America/Toronto |
| 409 | fdsgfdgdf | - | America/Toronto |
| 5341 | receiver_20250807 | n | Etc/GMT-8 |
