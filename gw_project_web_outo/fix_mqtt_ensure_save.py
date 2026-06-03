"""
Update _ensure_mqtt_enabled in all MQTT test files to:
1. Save MQTT Enable state after clicking (ensures form fully renders)
2. Increase wait from 300ms to 800ms after clicking Enable

This prevents race conditions where form fields (Clean Session, QoS, etc.)
haven't rendered yet when the test tries to interact with them.
"""
import glob

OLD = '''    enable_item = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
            enable_radio.locator(".el-radio__inner").click()
            page.wait_for_timeout(300)'''

NEW = '''    enable_item = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
            enable_radio.locator(".el-radio__inner").click()
            page.wait_for_timeout(800)
            page.get_by_role("button", name="Save").click()
            page.wait_for_timeout(500)'''

fixed = 0
for fpath in glob.glob('tests/protocols/mqtt/**/*.py', recursive=True):
    with open(fpath, encoding='utf-8') as f:
        content = f.read()
    if OLD in content:
        new_content = content.replace(OLD, NEW)
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f'FIXED: {fpath}')
        fixed += 1

print(f'\nDone: fixed={fixed}')
