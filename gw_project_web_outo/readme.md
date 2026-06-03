# 框架已建好，结构如下：                                                                                                                                                                                                            
                                                                                                                                                                                                                                    
  gw_project_web_outo/                                                                                                                                                                                                              
  ├── config/                                                                                                                                                                                                                       
  │   └── settings.py          # BASE_URL、浏览器、超时等配置，支持 .env 覆盖                                                                                                                                                       
  ├── pages/                                                                                                                                                                                                                        
  │   ├── base_page.py         # 所有页面的基类（click/fill/screenshot 封装）                                                                                                                                                     
  │   └── login_page.py        # 登录页 Page Object（选择器待按实际 UI 调整）                                                                                                                                                       
  ├── tests/
  │   └── test_login.py        # 登录测试示例（成功/错误密码/空凭据）
  ├── fixtures/                # 自定义 fixture（按需扩展）
  ├── utils/
  │   └── helpers.py           # 通用工具（时间戳、JSON 读写）
  ├── test_data/               # 测试数据（JSON/CSV 等）
  ├── reports/                 # HTML 测试报告（运行后生成）
  ├── screenshots/             # 失败自动截图
  ├── conftest.py              # Pytest fixture + 失败自动截图 hook
  ├── pytest.ini               # 测试配置（浏览器、报告路径、marker）
  ├── requirements.txt         # 依赖包
  └── .env.example             # 环境变量模板



  快速开始步骤：

  cd gw_project_web_outo

  # 1. 安装依赖
  pip install -r requirements.txt
  playwright install chromium

  # 2. 配置目标地址
  cp .env.example .env
  # 编辑 .env，设置 BASE_URL / WEB_USERNAME / WEB_PASSWORD

  # 3. 运行测试
  pytest                        # 全部测试
  pytest -m smoke               # 仅 smoke 用例
  pytest --headed               # 显示浏览器窗口
  

# 其他:
    创建虚拟环境 .venv              │ gw_project_web_outo\        │ 完成 ✅   │                                                                                                                                           
  ├─────────────────────────────────┼─────────────────────────────┼───────────┤                                                                                                                                          
  │ pip install -r requirements.txt │ .venv 内                    │ 完成 ✅   │                                                                                                                                           
  ├─────────────────────────────────┼─────────────────────────────┼───────────┤
  │ playwright install chromium     │ .venv 内                    │ 完成 ✅   │
  ├─────────────────────────────────┼─────────────────────────────┼───────────┤
  │ 创建 .vscode/settings.json      │ gw_project_web_outo\   
  
                  │  主项目（其他测试用例）  │            gw_project_web_outo                │
  ├───────────────┼──────────────────────────┼───────────────────────────────────────────┤
  │ 使用的 Python  │ Python311\python.exe     │ .venv\Scripts\python.exe（同一个 Python） │
  ├───────────────┼──────────────────────────┼───────────────────────────────────────────┤
  │ 使用的包       │ Python311\site-packages\ │ .venv\site-packages\（独立）              │
  ├───────────────┼──────────────────────────┼───────────────────────────────────────────┤
  │ 互相影响？      │ 否                       │ 否
 # 安装单个依赖包
  看到 (.venv) 前缀才说明在虚拟环境中，此时再执行 pip install 就会装到 .venv 里。
  ---
  安装单个额外的包（示例）
  先激活
  .\.venv\Scripts\Activate.ps1
  再安装
  pip install requests
  ---
  退出虚拟环境
  deactivate
# 安装requirements中的包
  .\.venv\Scripts\Activate.ps1
  pip install -r requirements.txt
  实际上这个已经执行过了，之前配置环境时我已经帮你装好了。你可以验证一下：
  .\.venv\Scripts\Activate.ps1
  pip list
  会看到 pytest、playwright、allure-pytest 等所有包都已经在里面了。
  退出虚拟环境
  deactivate
  

# 用例运行
直接运行 tests/usermanagement/ 目录下所有已调试的用例：

  python -m pytest tests/usermanagement/ -v

  加上 HTML 报告（pytest.ini 已配置自动生成，但可以显式指定）：

  python -m pytest tests/usermanagement/ -v --html=reports/report.html

  如果只想运行密码策略相关的 20 个新用例：

  python -m pytest tests/usermanagement/passwordpolicy/ -v


# 问题QA：
 我执行了没响应
  PS C:\autotest_local\autotest\gw_project_web_outo> python -m pytest tests/usermanagement/ -v --html=reports/report.html
  PS C:\autotest_local\autotest\gw_project_web_outo>

● 立即返回空提示符，说明 pytest 没有收集到任何测试或者 python 命令指向了错误的路径。先排查：

● PowerShell(python --version 2>&1; Get-Command python | Select-Object -ExpandProperty Source)
  ⎿  Error: Exit code 49
     C:\Users\YanLingCao\AppData\Local\Microsoft\WindowsApps\python.exe

● 问题找到了。python 命令指向的是 Windows 应用商店的 Python 占位符（WindowsApps\python.exe），不是真正的 Python，所以静默退出没有任何输出。

  需要用完整路径，改用以下命令：

  C:\Users\YanLingCao\AppData\Local\Programs\Python\Python312\python.exe -m pytest tests/usermanagement/ -v --html=reports/report.html

  或者直接调用 pytest 脚本（更简洁）：