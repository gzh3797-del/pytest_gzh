import os
import time
import pytest
from tools.mkDir import mk_dir

project_path = os.path.dirname(os.path.abspath(__file__))
sheet = ['basic_configuration']


def run():
    localtime = time.strftime('%Y%m%d%H%M%S', time.localtime())
    for i in sheet:
        test_case_path = project_path + '/test_case/{}/'.format(i).replace('/', '\\')
        test_case_path.replace('/', '\\')
        html_path = project_path + '/report/html/'.replace('/', '\\') + ""
        mk_dir(html_path)
        args = ['-s', '-q', test_case_path, "--html=./report/html/{}_test_report_{}.html".format(i, localtime),
                "--self-contained-html", "--disable-warnings", "--css=report.css", "--capture=sys"]
        pytest.main(args)


if __name__ == '__main__':
    run()
