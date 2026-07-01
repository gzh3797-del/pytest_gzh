Set shell = CreateObject("WScript.Shell")
shell.Run "C:\work\tools\python311\python.exe -m http.server 8899 --directory ""C:\work\autotest\autotest\AcuHMI-1-7\tests\protocols\aws_iot""", 0, False
WScript.Sleep 1500
shell.Run "chrome.exe http://localhost:8899", 1, False
