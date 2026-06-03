import json, sys
sys.stdout.reconfigure(encoding='utf-8')

# 用户管理 55条红色待处理用例分析
write_data = [
    # ── 澄清已明确不实现自动化 ────────────────────────
    {'row': 745, 'color': 'FF0000',
     'text': '[澄清#13已明确不实现] 步骤1依赖新固件烧录后的初始设备状态，无法通过UI操作重置；auto=是标记与澄清决定冲突，建议将auto改为否。'},
    {'row': 747, 'color': 'FF0000',
     'text': '[澄清#14已明确不实现] admin忘记密码临时密码功能依赖外部工具（需根据时间+SN生成临时密码），Playwright无法自动化生成临时密码。'},
    {'row': 749, 'color': 'FF0000',
     'text': '[澄清#14已明确不实现] 同上，临时密码+系统时间修改组合，依赖外部工具生成临时密码，不实现自动化。'},
    {'row': 750, 'color': 'FF0000',
     'text': '[澄清#14已明确不实现] 非admin用户密码过期后使用临时密码，依赖外部工具生成临时密码，不实现自动化。'},
    {'row': 751, 'color': 'FF0000',
     'text': '[澄清#14已明确不实现] 锁定用户使用临时密码，依赖外部工具生成临时密码，不实现自动化。'},
    {'row': 752, 'color': 'FF0000',
     'text': '[澄清#14已明确不实现] 密码一天内不可修改+临时密码，依赖外部工具生成临时密码，不实现自动化。'},

    # ── EULA/固件升级依赖 ────────────────────────────
    {'row': 756, 'color': 'FF0000',
     'text': '疑问：步骤2"本地升级到新软件版本且EULA版本变更"依赖固件升级操作，无法通过Playwright执行；固件升级需要物理文件上传且设备重启，非纯UI自动化可覆盖范围。'},

    # ── 权限验证 — 缺具体UI断言 ─────────────────────
    {'row': 762, 'color': 'FF0000',
     'text': '疑问：1. 预期"系统权限为视图可编辑"措辞矛盾（view=只读，不可编辑）。2. 各模块view权限的具体UI表现未说明：哪些按钮可见/隐藏？User Management中Add User/Lock/Delete按钮是否可见？'},
    {'row': 763, 'color': 'FF0000',
     'text': '疑问：预期"系统权限为视图可编辑"与全edit权限不符。请确认：全edit权限用户登录后，所有模块的操作按钮均可见可用？还是仅特定模块？断言标准是什么（以User Management的Add User按钮是否可见为例）？'},
    {'row': 764, 'color': 'FF0000',
     'text': '疑问：User=view权限时，User Management页面具体UI状态？例如：Add User按钮是否可见？Lock按钮是否可见？编辑/删除按钮是否可见？还是仅能浏览列表（显示"no operation permission"提示）？'},
    {'row': 765, 'color': 'FF0000',
     'text': '疑问：User=edit权限时，User Management页面具体操作范围？例如：Add User/Delete/Lock是否都可操作？与admin权限有何区别？'},
    {'row': 766, 'color': 'FF0000',
     'text': '疑问：Device=view权限时，Physical Devices页面具体UI状态：Add Device/Edit/Delete按钮是否可见？或仅可查看列表？'},
    {'row': 767, 'color': 'FF0000',
     'text': '疑问：Device=edit权限时，Physical Devices的具体操作范围？Add/Edit/Delete是否全部可用？'},
    {'row': 768, 'color': 'FF0000',
     'text': '疑问：Data Log=view权限时，Data Log各子页（Data Loggers/Management/Post/AcuCloud）具体可访问范围？哪些配置操作被限制？'},
    {'row': 769, 'color': 'FF0000',
     'text': '疑问：Data Log=edit权限时，具体可操作的功能范围？与view权限的UI差异是什么？'},
    {'row': 770, 'color': 'FF0000',
     'text': '疑问：System Settings=view权限时，8个Tab页（Date&Time/Network/Access Control等）哪些Save按钮不可点击？或页面显示"只读"状态？具体可断言现象？'},
    {'row': 771, 'color': 'FF0000',
     'text': '疑问：System Settings=edit权限时，与admin权限的UI差异？所有8个Tab的Save按钮均可用？'},
    {'row': 772, 'color': 'FF0000',
     'text': '疑问：Protocol=view权限时，Protocols各页面（Modbus/SNMP/MQTT等）的Save/修改按钮是否可见？或仅可查看配置值？'},
    {'row': 773, 'color': 'FF0000',
     'text': '疑问：Protocol=edit权限时，Protocols各页面操作范围？'},
    {'row': 774, 'color': 'FF0000',
     'text': '疑问：标题"模板-视图"与步骤"告警(Alarm Log)-视图"不一致，请确认测试的是Templates还是Alarm Log权限？Alarm Log=view时的具体UI状态（Active Alarms确认按钮是否可见？）'},
    {'row': 775, 'color': 'FF0000',
     'text': '疑问：同case01_13，标题"模板-编辑"与步骤"告警-编辑"不一致。请确认测试的权限模块；Alarm Log=edit时的可操作范围？'},
    {'row': 776, 'color': 'FF0000',
     'text': '疑问：Maintenance=view权限时，System Status页面Reboot按钮是否可见？Event Log是否只读（无Clear按钮）？请提供具体UI状态。'},
    {'row': 777, 'color': 'FF0000',
     'text': '疑问：Maintenance=edit权限时，与view权限的具体UI差异（Reboot按钮可用？Event Log可操作）？'},
    {'row': 778, 'color': 'FF0000',
     'text': '疑问：Diagnostics=view权限时，8个诊断工具（Host Lookup/Connection Test/NTP Sync等）各自的操作按钮是否可见？只读浏览还是也可操作？'},
    {'row': 779, 'color': 'FF0000',
     'text': '疑问：Diagnostics=edit权限时，与view权限差异？各诊断工具的操作（如Ping测试、Host Lookup）均可执行？'},

    # ── 无效用例（重新启动权限不存在）────────────────
    {'row': 780, 'color': 'FF0000',
     'text': '[澄清#15：无效用例] Role Configuration中不存在"重新启动"权限列（struct.md §5.1列出的10个权限列中无此项），该用例无法实现，建议标记auto=否。'},
    {'row': 781, 'color': 'FF0000',
     'text': '[澄清#15：无效用例] 同case01_19，Role Configuration中"重新启动"权限列不存在，无效用例，建议标记auto=否。'},

    # ── 固件更新权限 ─────────────────────────────────
    {'row': 782, 'color': 'FF0000',
     'text': '疑问：Firmware Update=view权限时，Firmware Update页面具体UI状态？Manual Update的Upload按钮是否可见但不可点击？还是整个页面不可访问？请提供具体可断言现象。'},
    {'row': 783, 'color': 'FF0000',
     'text': '疑问：Firmware Update=edit权限时，Upload按钮是否可用？与admin权限的UI差异？'},

    # ── 删除角色 ─────────────────────────────────────
    {'row': 792, 'color': 'FF0000',
     'text': '疑问：1. 删除已有用户关联的角色时，具体错误提示文本是什么（"该角色已有用户，无法删除"或类似）？2. 步骤3（用户已删除后登录）的预期缺失——已删除的用户登录是什么错误提示？'},
    {'row': 793, 'color': 'FF0000',
     'text': '疑问：1. "删除全部角色"是否包括内置admin/view角色（struct.md §5.2显示admin行无删除按钮）？2. 测试步骤和完整预期结果缺失，请补充：删除所有自定义角色后，系统和用户的具体行为。'},

    # ── 权限变更验证 ─────────────────────────────────
    {'row': 785, 'color': 'FF0000',
     'text': '疑问：角色权限从"全无"改为"全视图"后，各模块具体UI变化未说明：是否各模块页面从无法访问变为可只读访问？具体以哪个模块/按钮的状态作为断言标准？'},
    {'row': 786, 'color': 'FF0000',
     'text': '疑问：角色权限从"全无"改为"全编辑"后，具体UI变化断言标准？'},
    {'row': 787, 'color': 'FF0000',
     'text': '疑问：角色权限从"全视图"改为"全编辑"后，"权限均为编辑"的具体UI断言标准（以哪个模块哪个按钮为例）？'},
    {'row': 788, 'color': 'FF0000',
     'text': '疑问：角色权限从"全视图"改为"全无"后，"权限均为无"的具体UI断言标准（页面是否完全无法访问？显示何种提示？）'},
    {'row': 789, 'color': 'FF0000',
     'text': '疑问：角色权限从"全编辑"改为"全无"后，"权限均为无"的具体UI断言标准？'},
    {'row': 790, 'color': 'FF0000',
     'text': '疑问：角色权限从"全编辑"改为"全视图"后，"权限均为视图"的具体UI断言标准？'},

    # ── 时间依赖 — 密码策略 ──────────────────────────
    {'row': 818, 'color': 'FF0000',
     'text': '疑问：Minimum Password Age=50天，验证50天内不可修改密码，时间跨度超出自动化限制（struct §2.6：30分钟以上不实现）。若需系统时间跳转50天来实现，请确认此方法是否批准。'},
    {'row': 819, 'color': 'FF0000',
     'text': '疑问：Minimum Password Age=90天，同上，时间跨度超出自动化限制。请确认是否通过系统时间跳转实现。'},
    {'row': 824, 'color': 'FF0000',
     'text': '疑问：Password Expires=1天验证（澄清#17批准Min Password Age=1天通过系统时间实现，同样逻辑是否适用于Password Expires=1天？）。若可通过修改系统时间实现，请确认此方案，并提供设备系统时间修改方法（SSH/API/其他）。'},
    {'row': 825, 'color': 'FF0000',
     'text': '疑问：Password Expires=50天，时间跨度超出自动化限制（50天跳转）。'},
    {'row': 826, 'color': 'FF0000',
     'text': '疑问：Password Expires=90天，时间跨度超出自动化限制（90天跳转）。'},
    {'row': 835, 'color': 'FF0000',
     'text': '疑问：Grace Period=0验证需先使密码过期（Password Expires=1天后），依赖时间跳转。请确认：1. 是否批准通过系统时间跳转实现？2. 提供设备时间修改方法。'},
    {'row': 836, 'color': 'FF0000',
     'text': '疑问：Grace Period=1天验证（1天内可登录改密，1天后锁定），需系统时间操作，时间跨度较大。是否批准通过时间跳转实现？'},

    # ── 登录失败策略 ─────────────────────────────────
    {'row': 847, 'color': 'FF0000',
     'text': '疑问：标题说"30分钟"窗口，但步骤配置为Window=30秒/Wait=5秒（时间短可自动化）。请确认实际测试配置是30秒还是30分钟？另外，预期结果前后矛盾（"均登录失败用户未锁定" vs "第3次被锁等待5秒解锁"），请统一预期逻辑。'},
    {'row': 849, 'color': 'FF0000',
     'text': '疑问：标题说"40000秒"，但步骤配置为MaxFailed=5/Window=120秒/Wait=5秒（短时间，可自动化）。请确认：1. 实际测试Window是120秒还是40000秒？2. 完整预期结果（120秒内5次失败→锁定5秒→解锁；120秒后计数重置）是否正确？'},

    # ── Session Timeout 时间依赖 ─────────────────────
    {'row': 861, 'color': 'FF0000',
     'text': '疑问：Session Timeout=30分钟，验证需等待31分钟，超出自动化时间限制（struct §2.6）。若可接受测试时间较长，请确认是否纳入自动化（可用wait_for_timeout(31*60*1000)实现，约需31分钟执行）。'},
    {'row': 862, 'color': 'FF0000',
     'text': '疑问：Session Timeout=60分钟，验证需等待61分钟，超出时间限制。是否接受超长执行时间？'},
    {'row': 865, 'color': 'FF0000',
     'text': '疑问：与case01_3重复（Session Timeout=60分钟），同上时间限制问题。另请确认是否为重复用例。'},

    # ── 用户权限修改验证 ─────────────────────────────
    {'row': 877, 'color': 'FF0000',
     'text': '疑问：1. 修改的目标Role未指定（"修改Role"改为什么？admin/view/自定义角色？）。2. 预期"用户权限正确"缺少具体UI断言（改为哪个角色后能看到/做什么操作）？密码修改部分可自动化，但权限验证部分需补充具体断言。'},
    {'row': 880, 'color': 'FF0000',
     'text': '疑问：步骤"修改Role为edit"——"edit"不是有效角色名（有效角色名应为admin、view或自定义角色名）。是指将用户角色改为某个具有edit权限的自定义角色？还是Role Configuration中"edit"权限的角色？请提供具体目标角色名称。'},
    {'row': 881, 'color': 'FF0000',
     'text': '疑问：修改9个用户权限，预期"被修改用户权限均正确"缺少具体UI断言。修改为哪些目标角色？每个角色对应的UI验证标准？'},
    {'row': 882, 'color': 'FF0000',
     'text': '疑问：修改10个用户权限，同上，缺少具体目标角色和UI断言标准。'},

    # ── 用户名/密码长度边界 ──────────────────────────
    {'row': 886, 'color': 'FF0000',
     'text': '疑问：用户名40字符（有效范围1-40，应成功）+ 密码65字符。预期"保存失败"，但struct.md §4.4说密码有效范围6-128字符，65应在范围内。请确认：1. 密码实际上限是64还是128？2. 若上限64，请更新struct.md §4.4。'},
    {'row': 887, 'color': 'FF0000',
     'text': '疑问：用户名41字符（超过40上限）。预期"保存失败"但备注说"会被自动截断，也可保存成功"——两种结果互相矛盾。请确认：超过40字符时，系统是①拒绝保存（显示错误）还是②自动截断后保存成功？'},
]

with open('tmp_um_write.json', 'w', encoding='utf-8') as f:
    json.dump(write_data, f, ensure_ascii=False, indent=2)
print(f'Prepared {len(write_data)} items for 用户管理')
