"""
Add _ensure_mqtt_enabled helper and call it in all MQTT test files.
MQTT form fields (Broker Address, Port, etc.) only appear when MQTT is enabled.
"""
import os, sys, re
sys.stdout.reconfigure(encoding='utf-8')

HELPER_FUNC = '''

def _ensure_mqtt_enabled(page):
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
            page.wait_for_timeout(1000)
'''

mqtt_dir = 'tests/protocols/mqtt'
fixed = 0

for fn in sorted(os.listdir(mqtt_dir)):
    if not fn.endswith('.py') or fn == '__init__.py':
        continue
    fpath = os.path.join(mqtt_dir, fn)
    with open(fpath, encoding='utf-8') as f:
        content = f.read()

    # Skip if already has _ensure_mqtt_enabled
    if '_ensure_mqtt_enabled' in content:
        continue

    # Insert helper after _nav_protocol function
    # The function ends with the last "page.wait_for_timeout(500)" inside the if sub: block
    # Find the end of _nav_protocol by looking for the pattern
    nav_end_pattern = r'(def _nav_protocol\(.*?\n(?:.*\n)*?.*page\.wait_for_timeout\(500\)\n\n)'

    # Simpler: insert the helper after the closing of _nav_protocol (before the test function)
    # Find "def test_" and insert before it
    test_func_match = re.search(r'\ndef test_', content)
    if not test_func_match:
        continue

    insert_pos = test_func_match.start()
    new_content = content[:insert_pos] + HELPER_FUNC + content[insert_pos:]

    # Now add _ensure_mqtt_enabled() call after _nav_protocol call in the test function
    # Find the _nav_protocol call and add the helper call after it
    # Pattern: after "_nav_protocol(page, ..." line
    new_content = re.sub(
        r'(_nav_protocol\(page, [^\n]+\n)(    page\.wait_for_timeout\(500\)\n\n    (?!_ensure))',
        r'\1    _ensure_mqtt_enabled(page)\n    \2',
        new_content
    )
    # Also handle case where _nav_protocol call is not followed by page.wait_for_timeout(500)
    # (the function itself handles waits, so the call is just one line)
    new_content = re.sub(
        r'(_nav_protocol\(page, [^\n]+\n)(\n    (?![_\n]))',
        r'\1    _ensure_mqtt_enabled(page)\n\2',
        new_content
    )

    with open(fpath, 'w', encoding='utf-8') as f:
        f.write(new_content)
    fixed += 1
    # print(f'  Fixed: {fn}')

print(f'Fixed {fixed} MQTT test files')
