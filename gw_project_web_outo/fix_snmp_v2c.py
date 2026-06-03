"""
Fix SNMP v2c tests (case01-07):
1. Add SNMP version selection (v2c)
2. Fix Community locator from 'Enter Community' to 'Enter RO Community'
3. Fix community value per test description
4. Remove `if "<value>":` wrapper
"""
import os, sys, re
sys.stdout.reconfigure(encoding='utf-8')

# Maps: case_number -> (port, community_value)
CASE_DATA = {
    'case01': ('161', 'public'),
    'case02': ('16100', 'private'),
    'case03': ('16159', '@@@###'),
    'case04': ('16199', '123456'),
    'case05': ('161', ''),
    'case06': ('161', 'wrong_community'),
    'case07': ('16200', 'public'),
}

VERSION_SELECT_CODE = """
    # Select SNMP Version v2c
    version_select = page.locator(".el-form-item").filter(has_text="Version").locator(".el-select")
    if version_select.count() > 0:
        version_select.click()
        page.wait_for_timeout(200)
        v2c_opt = page.get_by_role("option").filter(has_text=re.compile(r"v2c", re.IGNORECASE))
        if v2c_opt.count() > 0:
            v2c_opt.first.click()
        page.wait_for_timeout(200)
"""

snmp_dir = 'tests/protocols/snmp'
fixed = 0

for case_key, (port, community) in CASE_DATA.items():
    fn = f'test_TestCase_AcuRev4100_WEB2_008_003_{case_key}.py'
    fpath = os.path.join(snmp_dir, fn)
    if not os.path.exists(fpath):
        print(f'  SKIP (not found): {fn}')
        continue
    with open(fpath, encoding='utf-8') as f:
        content = f.read()

    # Remove stale 'import re' if already present
    content = re.sub(r'^import re\n', '', content, flags=re.MULTILINE)

    # Add 'import re' after 'import pytest'
    content = content.replace('import pytest\n', 'import pytest\nimport re\n', 1)

    # Replace community block (handles both truthy and falsy if conditions)
    # Pattern: if "<anything>":
    #     community_field = ...("Enter Community")
    #     community_field.fill("...")
    old_community_pattern = re.compile(
        r'    if "[^"]*":\n'
        r'        community_field = page\.get_by_label\("Community".*?\n'
        r'            page\.get_by_placeholder\("Enter Community"\)\)\n'
        r'        community_field\.fill\("[^"]*"\)\n',
        re.DOTALL
    )
    new_community = (
        f'    community_field = page.get_by_placeholder("Enter RO Community").or_(\n'
        f'        page.get_by_label("RO Community", exact=False))\n'
        f'    community_field.fill("{community}")\n'
    )
    content, n = old_community_pattern.subn(new_community, content)
    if n == 0:
        print(f'  WARN: community pattern not found in {fn}')

    # Insert version select code before port_field.fill()
    port_fill_line = f'    port_field.fill("{port}")\n'
    if port_fill_line in content and VERSION_SELECT_CODE not in content:
        content = content.replace(port_fill_line, VERSION_SELECT_CODE + port_fill_line)

    with open(fpath, 'w', encoding='utf-8') as f:
        f.write(content)
    fixed += 1
    print(f'  Fixed: {fn}')

print(f'\nTotal fixed: {fixed} files')
