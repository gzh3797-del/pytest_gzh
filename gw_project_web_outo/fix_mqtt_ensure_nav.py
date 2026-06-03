"""
Fix _ensure_mqtt_enabled return navigation: remove the MQTT re-click
that collapses the sub-menu when we're already on a MQTT sub-page.
"""
import glob
import sys
sys.stdout.reconfigure(encoding='utf-8')

OLD = '''    if original_sub:
        # Re-expand MQTT sub-menu (it collapses when on a sub-page)
        page.get_by_role("menuitem", name="MQTT").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
        page.get_by_role("menuitem", name=original_sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)'''

NEW = '''    if original_sub:
        page.get_by_role("menuitem", name=original_sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)'''

fixed = 0
for fpath in glob.glob('tests/protocols/mqtt/**/*.py', recursive=True):
    with open(fpath, encoding='utf-8') as f:
        content = f.read()
    if OLD in content:
        content = content.replace(OLD, NEW)
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(content)
        fixed += 1
        print(f'  Fixed: {fpath}')

print(f'\nTotal fixed: {fixed} files')
