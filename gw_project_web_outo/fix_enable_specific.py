"""
Fix enable locators to be more specific:
- SNMP: "Enable" → "SNMP Enable"
- Azure IoT: "Enable" → "Azure IoT Enable"
- AWS IoT: "Enable" → "AWS IoT Enable" or "Enable" (AWS IoT page may not have SSL Enable)
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8')

# SNMP: fix SNMP Enable and also fix Trap Enable if separate
OLD_SNMP = (
    '    enable_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable")\n'
    '    if enable_radio.count() > 0 and \'is-checked\' not in (enable_radio.get_attribute("class") or ""):\n'
    '        enable_radio.locator(".el-radio__inner").click()\n'
    '        page.wait_for_timeout(500)\n'
)
NEW_SNMP = (
    '    enable_radio = page.locator(".el-form-item").filter(has_text="SNMP Enable").locator(".el-radio").filter(has_text="Enable")\n'
    '    if enable_radio.count() > 0 and \'is-checked\' not in (enable_radio.first.get_attribute("class") or ""):\n'
    '        enable_radio.first.locator(".el-radio__inner").click()\n'
    '        page.wait_for_timeout(500)\n'
)

# Azure IoT: fix to use "Azure IoT Enable"
OLD_AZU = (
    '    enable_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable")\n'
    '    if enable_radio.count() > 0 and \'is-checked\' not in (enable_radio.get_attribute("class") or ""):\n'
    '        enable_radio.locator(".el-radio__inner").click()\n'
    '        page.wait_for_timeout(500)\n'
)
NEW_AZU = (
    '    enable_radio = page.locator(".el-form-item").filter(has_text="Azure IoT Enable").locator(".el-radio").filter(has_text="Enable")\n'
    '    if enable_radio.count() > 0 and \'is-checked\' not in (enable_radio.first.get_attribute("class") or ""):\n'
    '        enable_radio.first.locator(".el-radio__inner").click()\n'
    '        page.wait_for_timeout(500)\n'
)

# AWS IoT: fix to use "AWS IoT Enable" (may or may not have SSL on page)
# Actually let's use "Enable" but with .first to avoid strict mode
OLD_AWS = OLD_AZU  # same pattern
NEW_AWS = (
    '    enable_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable").first\n'
    '    if \'is-checked\' not in (enable_radio.get_attribute("class") or ""):\n'
    '        enable_radio.locator(".el-radio__inner").click()\n'
    '        page.wait_for_timeout(500)\n'
)

snmp_dir = 'tests/protocols/snmp'
azu_dir = 'tests/protocols/azureiot'
aws_dir = 'tests/protocols/awsiot'

total_fixed = 0

for d, old, new in [(snmp_dir, OLD_SNMP, NEW_SNMP), (azu_dir, OLD_AZU, NEW_AZU), (aws_dir, OLD_AWS, NEW_AWS)]:
    for fn in sorted(os.listdir(d)):
        if not fn.endswith('.py') or fn == '__init__.py':
            continue
        fpath = os.path.join(d, fn)
        with open(fpath, encoding='utf-8') as f:
            content = f.read()
        if old in content:
            content = content.replace(old, new)
            with open(fpath, 'w', encoding='utf-8') as f:
                f.write(content)
            total_fixed += 1
            print(f'  Fixed: {fpath}')

print(f'\nTotal fixed: {total_fixed} files')
