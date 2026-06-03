"""
Fix enable toggle pattern in SNMP, AWS IoT, Azure IoT test files.
Replaces the fragile .el-radio/.el-switch click with the MQTT-style
.el-radio__inner click + is-checked guard.
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8')

OLD = """    enable_item = page.locator(".el-form-item").filter(has_text="Enable").first
    enable_item.locator(".el-radio, .el-switch").filter(has_text="Enable").click()
    page.wait_for_timeout(300)"""

NEW = """    enable_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and 'is-checked' not in (enable_radio.get_attribute("class") or ""):
        enable_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)"""

dirs = [
    'tests/protocols/snmp',
    'tests/protocols/awsiot',
    'tests/protocols/azureiot',
]

fixed = 0
for d in dirs:
    for fn in sorted(os.listdir(d)):
        if not fn.endswith('.py') or fn == '__init__.py':
            continue
        fpath = os.path.join(d, fn)
        with open(fpath, encoding='utf-8') as f:
            content = f.read()
        if OLD in content:
            content = content.replace(OLD, NEW)
            with open(fpath, 'w', encoding='utf-8') as f:
                f.write(content)
            fixed += 1
            print(f"  Fixed: {fpath}")

print(f'\nTotal fixed: {fixed} files')
