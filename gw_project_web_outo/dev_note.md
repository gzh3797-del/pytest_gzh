# 调试自动化时候，使用如下命令进行调试
playwright codegen --ignore-https-errors https://192.168.2.199/ 

# 自动执行
python -m pytest tests/test_login.py -v -s
