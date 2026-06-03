import os, sys
sys.stdout.reconfigure(encoding='utf-8')

OLD_NAV = '''def _nav_protocol(page, protocol: str, sub: str = None):
    if "/protocols/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_role("menuitem", name="Protocols").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=protocol).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    if sub:
        page.get_by_role("menuitem", name=sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)'''

NEW_NAV = '''def _nav_protocol(page, protocol: str, sub: str = None):
    if "/protocols/" not in page.url:
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/maintenance",
            "/templates", "/firmwareUpdate", "/diagnostics",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Protocols").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=protocol).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    if sub:
        page.get_by_role("menuitem", name=sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)'''

fixed = 0
for root, dirs, files in os.walk('tests/protocols'):
    for fn in files:
        if not fn.endswith('.py') or fn == '__init__.py':
            continue
        fpath = os.path.join(root, fn)
        with open(fpath, encoding='utf-8') as f:
            content = f.read()
        if OLD_NAV in content:
            content = content.replace(OLD_NAV, NEW_NAV)
            with open(fpath, 'w', encoding='utf-8') as f:
                f.write(content)
            fixed += 1

print(f'Fixed {fixed} files')
