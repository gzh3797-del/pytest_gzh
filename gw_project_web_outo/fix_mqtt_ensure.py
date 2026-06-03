"""
Fix _ensure_mqtt_enabled in all MQTT test files to return to original sub-page.
"""
import os, sys, re
sys.stdout.reconfigure(encoding='utf-8')

OLD_ENSURE = '''def _ensure_mqtt_enabled(page):
    """Enable MQTT if currently disabled so config fields are visible."""
    # Navigate to General to check/set enable state
    if "/protocols/mqtt/general" not in page.url:
        page.get_by_role("menuitem", name="MQTT").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
        page.get_by_role("menuitem", name="General").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
    enable_item = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
            enable_radio.locator(".el-radio__inner").click()
            page.get_by_role("button", name="Save").click()
            page.wait_for_timeout(1000)'''

NEW_ENSURE = '''def _ensure_mqtt_enabled(page):
    """Enable MQTT if disabled. Returns to original MQTT sub-page afterward."""
    current_url = page.url
    sub_map = {
        "credential": "User Credential",
        "ssl": "SSL",
        "testament": "Last Will and Testament",
        "deviceToPublish": "Topic and Parameter Selection",
    }
    original_sub = None
    for suffix, name in sub_map.items():
        if f"/protocols/mqtt/{suffix}" in current_url:
            original_sub = name
            break
    if "/protocols/mqtt/general" not in current_url:
        page.get_by_role("menuitem", name="MQTT").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
        page.get_by_role("menuitem", name="General").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
    enable_item = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
            enable_radio.locator(".el-radio__inner").click()
            page.wait_for_timeout(300)
    if original_sub:
        page.get_by_role("menuitem", name=original_sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)'''

mqtt_dir = 'tests/protocols/mqtt'
fixed = 0
for fn in sorted(os.listdir(mqtt_dir)):
    if not fn.endswith('.py') or fn == '__init__.py':
        continue
    fpath = os.path.join(mqtt_dir, fn)
    with open(fpath, encoding='utf-8') as f:
        content = f.read()
    if OLD_ENSURE in content:
        content = content.replace(OLD_ENSURE, NEW_ENSURE)
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(content)
        fixed += 1

print(f'Fixed _ensure_mqtt_enabled in {fixed} files')
