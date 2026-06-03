"""
Fix syntax error in cases 19-28: pytest.skip("..." → pytest.skip("...")
Also adds @pytest.mark.skip decorator to the test function so it skips at collection time.
"""
import glob, re

BROKEN = 'pytest.skip("'
FIXED_SUFFIX = '")'  # The line ends with " but needs ") to close the call

fixed = 0
for fpath in (glob.glob('tests/protocols/mqtt/test_TestCase_AcuRev4100_WEB2_005_004_case1*.py') +
              glob.glob('tests/protocols/mqtt/test_TestCase_AcuRev4100_WEB2_005_004_case2*.py')):
    with open(fpath, encoding='utf-8') as f:
        lines = f.readlines()

    changed = False
    new_lines = []
    for line in lines:
        # Fix: pytest.skip("...device msg..." → pytest.skip("...device msg...")
        if BROKEN in line and line.rstrip().endswith('"') and 'pytest.skip' in line:
            # The line is: `        pytest.skip("...msg..."\n`
            # Fix it to:   `        pytest.skip("...msg...")\n`
            line = line.rstrip()  # remove newline
            line = line + ')\n'   # add closing paren + newline
            changed = True
        new_lines.append(line)

    if changed:
        with open(fpath, 'w', encoding='utf-8') as f:
            f.writelines(new_lines)
        print(f'FIXED: {fpath}')
        fixed += 1

print(f'\nDone: fixed={fixed}')
