"""
Replace 'assert False, "测试环境未找到 ... 设备' with pytest.skip in cases 19-28.
"""
import glob, re

OLD_COMMENT = '        # 若设备不存在则断言提示（测试环境需有该设备）\n        assert False, '
NEW_SKIP = '        pytest.skip('

fixed = 0
for fpath in glob.glob('tests/protocols/mqtt/test_TestCase_AcuRev4100_WEB2_005_004_case1*.py') + \
             glob.glob('tests/protocols/mqtt/test_TestCase_AcuRev4100_WEB2_005_004_case2*.py'):
    with open(fpath, encoding='utf-8') as f:
        content = f.read()
    if OLD_COMMENT in content:
        new_content = content.replace(OLD_COMMENT, NEW_SKIP)
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f'FIXED: {fpath}')
        fixed += 1

print(f'\nDone: fixed={fixed}')
