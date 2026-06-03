"""
Fix _ensure_mqtt_enabled forward navigation bug in all MQTT test files.

OLD (BUGGY): When already on a MQTT sub-page, clicking "MQTT" collapses the sub-menu
before "General" can be clicked.

NEW (FIXED): Only click "MQTT" if NOT already on a MQTT sub-page.
"""
import glob
import os

OLD = '''    if "/protocols/mqtt/general" not in current_url:
        page.get_by_role("menuitem", name="MQTT").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
        page.get_by_role("menuitem", name="General").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)'''

NEW = '''    if "/protocols/mqtt/general" not in current_url:
        if "/protocols/mqtt/" not in current_url:
            page.get_by_role("menuitem", name="MQTT").click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(300)
        page.get_by_role("menuitem", name="General").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)'''

fixed = 0
skipped = 0

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
        skipped += 1
    # else: file doesn't have _ensure_mqtt_enabled, skip

print(f'\nDone: fixed={fixed}, already-fixed={skipped}')
