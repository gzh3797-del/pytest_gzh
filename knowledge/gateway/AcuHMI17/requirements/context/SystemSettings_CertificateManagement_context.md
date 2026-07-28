# System Settings / Certificate Management — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/systemSettings/certificateManagement` |
| 路由名 | `certificateManagement` |
| 面包屑 | AcuHMI-1-7 / System Settings / Certificate Management |
| 顶级模块 | System Settings |

## 2. 页面用途

管理网关 HTTPS/服务证书：导入、生成自签名证书、生成 CSR、导出，并展示当前证书详情（颁发者/主体/有效期/详情）。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Import | button | `getByRole('button',{name:'Import'})` | 导入证书（打开上传弹框/表单） |
| Generate New Self-Signed Certificate | button | `getByRole('button',{name:'Generate New Self-Signed Certificate'})` | 生成自签名证书（打开填写主体信息表单） |
| Generate CSR | button | `getByRole('button',{name:'Generate CSR'})` | 生成证书签名请求 |
| Export | button | `getByRole('button',{name:'Export'})` | 导出当前证书 |

## 4. 只读信息区（读取断言用）

| 分组 | 字段 | 示例 |
|------|------|------|
| Certificate Issuer | Common Name / Company Name / Division Name / City / State / Country Code | AHI260110002 / Accuenergy (CANADA) Inc. / — / Toronto / ON / CA |
| Certificate Subject | 同上 | 同上 |
| Validity | Valid From / Expiration | Apr 10 2026 GMT / Apr 9 2056 GMT |
| Details | Public Key Size / Serial Number / Public Key Type / Certificate Version / Signature Algorithm / Extensions | 2048 / 26:f0:... / RSA / 3 / sha256WithRSAEncryption / — |

## 5. 自动化测试要点

- 4 个操作按钮的弹框/流程（Import 上传文件、Generate Self-Signed 填主体信息表单、Generate CSR、Export 下载）。
- 只读证书详情字段的呈现断言。
- 生成新证书后详情面板刷新验证。

## 6. 机器可解析摘要

```json
{
  "route": "/systemSettings/certificateManagement",
  "name": "certificateManagement",
  "title": "Certificate Management",
  "module": "System Settings",
  "buttons": ["Import","Generate New Self-Signed Certificate","Generate CSR","Export"],
  "readonly_sections": ["Certificate Issuer","Certificate Subject","Validity","Details"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/systemsettings/`。

### 加载态 API
- `GET /api/command/getCertInfo`

### pytest 选择器与控件
- 只读信息 group：`page.get_by_role("group", name="Common Name")` 等；**Issuer/Subject 各有一组同名 group（Common Name 等），需 `.first`/`.nth()` 或父容器 scope**（自签名证书 Subject=Issuer）。
- Details：Public Key Size(2048)/Serial Number/Public Key Type(RSA)/Certificate Version(3)/Signature Algorithm(sha256WithRSAEncryption)/Extensions。
- 按钮：Import / Generate New Self-Signed Certificate / Generate CSR / Export。

### 高危
- ⚠️ **Generate 类（生成自签名/CSR）会导致 web 服务重载**——执行前须确认，脚本需带重载等待+重登，禁无人值守连跑。
