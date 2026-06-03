"""
Fix _ensure_mqtt_enabled: add MQTT click to expand sub-menu before navigating back.
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8')

OLD_RETURN_BLOCK = '''    if original_sub:
        page.get_by_role("menuitem", name=original_sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)'''

NEW_RETURN_BLOCK = '''    if original_sub:
        # Re-expand MQTT sub-menu (it collapses when on a sub-page)
        page.get_by_role("menuitem", name="MQTT").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
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
    if OLD_RETURN_BLOCK in content:
        content = content.replace(OLD_RETURN_BLOCK, NEW_RETURN_BLOCK)
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(content)
        fixed += 1

print(f'Fixed {fixed} files')
