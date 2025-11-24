import pytest
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item, call):
    """pytest钩子，用于在测试报告中添加截图（成功和失败都添加）"""
    pytest_html = item.config.pluginmanager.getplugin('html')
    outcome = yield
    report = outcome.get_result()
    extra = getattr(report, 'extra', [])

    # 在测试执行阶段（call）添加截图，无论是成功还是失败
    if report.when == 'call':
        try:
            # 获取测试类实例
            test_instance = item.instance
            if hasattr(test_instance, 'helper'):
                # 根据测试结果设置截图描述
                if report.failed:
                    description = f"测试失败: {item.name}"
                    screenshot_type = "失败截图"
                else:
                    description = f"测试成功: {item.name}"
                    screenshot_type = "成功截图"

                # 获取base64格式的截图
                screenshot_base64 = test_instance.helper.take_screenshot_base64(description)

                if screenshot_base64:
                    # 添加到HTML报告
                    extra.append(pytest_html.extras.image(screenshot_base64, screenshot_type))
                    report.extra = extra

        except Exception as e:
            print(f"添加截图到报告时出错: {e}")


def pytest_configure(config):
    """pytest配置钩子"""
    # 确保html插件已加载
    if not config.pluginmanager.hasplugin('html'):
        config.pluginmanager.import_plugin('pytest_html.plugin')