import os
import pytest
import sys
import time

def main():
    # 项目根目录
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # 测试用例目录
    test_dir = os.path.join(base_dir, "test_case")

    # 报告目录
    report_dir = os.path.join(base_dir, "reports")
    os.makedirs(report_dir, exist_ok=True)

    localtime = time.strftime('%Y%m%d%H%M%S', time.localtime())
    report_path = os.path.join(report_dir, f"report{localtime}.html")

    # pytest 参数
    pytest_args = [
        test_dir,                 # 执行整个 test_case 目录
        "-v",                      # 详细输出
        "--html=" + report_path,   # 生成 HTML 报告
        "--self-contained-html",   # 报告不依赖外部资源
    ]

    # 运行 pytest
    exit_code = pytest.main(pytest_args)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
