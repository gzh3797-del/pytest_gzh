# API 接口文档

## 基本信息

| 项目 | 值 |
|------|-----|
| Base URL | `http://acucloud-test-451397146.cn-northwest-1.elb.amazonaws.com.cn` |
| API 前缀 | `/api/v1` |
| 报告 API 前缀 | `/api/report` |
| 内容类型 | `application/json` |

## 认证方式

### 登录获取 Token

```http
POST /api/v1/login/do
Content-Type: application/json

{
  "email": "user@example.com",
  "password": "your_password"
}
```

### 使用 Token

登录成功后，token 存储在 `sessionStorage['common'].token`，格式为：

```
Bearer xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

在后续请求中通过 Header 携带：

```http
Authorization: Bearer xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

### 多组织请求

需要在 Header 中携带组织 ID：

```http
Orgid: 431
```

### 语言设置

```http
Accept-Language: en
```

---

## 通用响应格式

```json
{
  "code": 200,
  "data": { ... },
  "msg": "success"
}
```

### 响应码说明

| code | 含义 |
|------|------|
| 200 | 成功 |
| 400 | 请求错误（如密码错误） |
| 401 | 未登录/Token 过期 |
| 500 | 服务器内部错误 |
| 501 | 参数错误（缺少必填参数） |
| 603 | 未订阅服务 |
| 800 | 特殊状态（业务相关） |

---

## 核心 API 端点

### 认证相关

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/login/do` | 登录 |
| POST | `/login/info` | 获取登录信息 |
| POST | `/login/changePwd` | 修改密码 |
| DELETE | `/login/logout` | 退出登录 |
| GET | `/login/forget/password` | 忘记密码 |
| GET | `/login/{x}/{y}` | 登录相关操作 |

### 权限相关

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/permission/current` | 获取当前用户可访问的设施权限列表 |
| GET | `/system-menu/tree` | 获取系统菜单树（需订阅） |

### 用户相关

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/user/info` | 获取当前用户信息 |
| GET | `/user/lastOrgViewed` | 获取最近查看的组织 |

### 组织相关

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/organization/current/list` | 获取当前用户可访问的组织列表 |
| GET | `/organization/list` | 获取组织列表 |

### 设备相关

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/devices/list` | 设备列表 |
| POST | `/devices/createByType` | 新增设备 |
| POST | `/devices/editDeviceByType` | 编辑设备 |
| GET | `/devices/meterpoint/create` | 创建计量点 |
| GET | `/devices/meterpoint/update` | 更新计量点 |
| POST | `/devices/meterpoint/updateMeterPointTotal` | 更新计量点汇总 |

### 设施相关

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/facility/list` | 设施列表 |

### 告警相关

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/alert/list` | 告警列表 |

### 日志相关

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/log/security` | 安全日志 |

### Zoho 集成

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/zoho/allow_portal` | Zoho 门户授权 |

---

## 报告 API（/api/report）

报告相关 API 使用单独的 Base Path `/api/report`，认证方式不同：

- 请求时携带额外 Header `authentication`（值为 `MD5(biz + type + requestId)` 的哈希）
- 示例报告路径：`/mv/report/executor`、`/report/PowerQualityConfig/harmonicLimits`

---

## 登录账户锁定策略

- 连续 **5次** 密码错误 → 账户锁定
- 错误时返回：`"Incorrect password. N attempts left before account lockout."`
- 锁定后需管理员解锁（Super Admin 可操作）

---

## 示例：获取权限列表

```bash
curl -X GET \
  "http://acucloud-test-451397146.cn-northwest-1.elb.amazonaws.com.cn/api/v1/permission/current" \
  -H "Authorization: Bearer <REDACTED>" \
  -H "Content-Type: application/json" \
  -H "Orgid: 431"
```

响应示例（部分）：
```json
{
  "code": 200,
  "data": [
    {
      "facilityId": 11,
      "facilityName": "Whitestar",
      "permission": "EDIT",
      "location": { "x": -97.0645916, "y": 36.1298222 },
      "city": "Stillwater",
      "timezone": "Etc/UTC"
    },
    ...
  ]
}
```

---

## 示例：获取组织列表

```bash
curl -X POST \
  "http://acucloud-test-451397146.cn-northwest-1.elb.amazonaws.com.cn/api/v1/organization/current/list" \
  -H "Authorization: Bearer <REDACTED>" \
  -H "Content-Type: application/json" \
  -H "Orgid: 431" \
  -d '{}'
```
