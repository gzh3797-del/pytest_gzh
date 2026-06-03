"""
Replace _ensure_mqtt_enabled menu-click navigation with direct URL navigation.

The menu-based approach fails because ElementUI collapses/hides sub-items
in certain states, making get_by_role("menuitem", name="General") time out.
Direct URL navigation (page.goto) is more reliable for SPA hash routing.
"""
import glob
import os

OLD = '''def _ensure_mqtt_enabled(page):
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
        if "/protocols/mqtt/" not in current_url:
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

NEW = '''def _ensure_mqtt_enabled(page):
    """Enable MQTT if disabled. Returns to original MQTT sub-page afterward."""
    current_url = page.url
    sub_url_map = {
        "credential": "User Credential",
        "ssl": "SSL",
        "testament": "Last Will and Testament",
        "deviceToPublish": "Topic and Parameter Selection",
    }
    sub_name_to_path = {v: k for k, v in sub_url_map.items()}
    original_sub = None
    for suffix, name in sub_url_map.items():
        if f"/protocols/mqtt/{suffix}" in current_url:
            original_sub = name
            break
    if "/protocols/mqtt/general" not in current_url:
        base = current_url.split("#")[0]
        page.goto(base + "#/protocols/mqtt/general")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    enable_item = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
            enable_radio.locator(".el-radio__inner").click()
            page.wait_for_timeout(300)
    if original_sub:
        path = sub_name_to_path.get(original_sub, "")
        if path:
            base = page.url.split("#")[0]
            page.goto(base + f"#/protocols/mqtt/{path}")
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)'''

fixed = 0
no_match = []

for fpath in glob.glob('tests/protocols/mqtt/**/*.py', recursive=True):
    with open(fpath, encoding='utf-8') as f:
        content = f.read()
    if OLD in content:
        new_content = content.replace(OLD, NEW)
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f'FIXED: {fpath}')
        fixed += 1
    elif NEW in content:
        pass  # already fixed
    elif '_ensure_mqtt_enabled' in content:
        no_match.append(fpath)

print(f'\nDone: fixed={fixed}')
if no_match:
    print('Files with _ensure_mqtt_enabled but pattern NOT matched:')
    for f in no_match:
        print(f'  {f}')
