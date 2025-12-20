#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

; ---------------- 配置 ----------------
LoopCount := 2              ; 总执行次数
ClickDelay := 100           ; 点击后等待（毫秒）
TimeoutSec := 300            ; 等待图片最长秒数
ButtonsDir := A_ScriptDir "\buttons"
InputFile := A_ScriptDir "\inputs.ini"
LogFile := A_ScriptDir "\log.txt"

; ---------------- GDI+ 设置（可选） ----------------
; 【注意】要使用 GDI+ 精确定位，请确保在脚本目录中放置了 GDIP.ahk 文件
; 并取消下面 #Include 的注释！
#Include GDIP.ahk
; 如果 GDIP.ahk 被包含，下面的变量和函数调用才会成功。

GDIP_Avail := False
pGdipToken := ""

; 尝试启动 GDI+ 并在成功时标记可用
if IsFunc("Gdip_Startup")  ; 检查函数是否存在，仅当 #Include 成功时才执行
{
    try {
        pGdipToken := Gdip_Startup()
        if (pGdipToken != 0)
            GDIP_Avail := True
    } catch {
        ; 启动失败，保持不可用状态
    }
}

; ---------------- 读取 buttons 目录下 PNG ----------------
Buttons := []
Loop, Files, %ButtonsDir%\*.png, FR
{
    Buttons.Push(A_LoopFileFullPath)
}
if (Buttons.MaxIndex() = "")
{
    MsgBox, 48, 提示, 未找到任何 PNG 文件，请把图片放到 %ButtonsDir% 目录下。
    ; 在 AHK v1.1 中，需要先 Shutdown GDI+
    if (pGdipToken)
        Gdip_Shutdown(pGdipToken)
    ExitApp
}
TotalSteps := Buttons.MaxIndex()

; ---------------- 写日志函数 ----------------
WriteLog(Text)
{
    global LogFile
    Line := A_Now . " | " . Text . "`n"
    FileAppend, %Line%, %LogFile%    ; <--- 默认写入
}

; ---------------- 主循环 ----------------
WriteLog("脚本启动。总步骤: " . TotalSteps . "，循环次数: " . LoopCount)

Loop, %LoopCount%
{
    RunIndex := A_Index
    WriteLog("开始执行流程 #" . RunIndex)

    Loop, %TotalSteps%
    {
        Step := A_Index
        ImgPath := Buttons[Step]

        ; 安全拼接日志
        Text := "步骤 " . Step . " | 查找图像: " . ImgPath
        WriteLog(Text)

        ; 查找并点击
        Result := FindAndClick(ImgPath, ClickDelay, TimeoutSec, GDIP_Avail, pGdipToken)
        if (Result = 0)
        {
            WriteLog("步骤 " . Step . " | 找不到图像，流程终止: " . ImgPath)
            MsgBox, 16, 错误, 找不到图片：`n%ImgPath%`n脚本终止！
            ; 在 AHK v1.1 中，需要先 Shutdown GDI+
            if (pGdipToken)
                Gdip_Shutdown(pGdipToken)
            ExitApp
        }
        WriteLog("步骤 " . Step . " | 点击成功: " . ImgPath)

        ; 检查输入步骤
        IniRead, InputTxt, %InputFile%, InputSteps, %Step%, ERROR
        if (InputTxt != "ERROR")
        {
            Text := "步骤 " . Step . " | 输入文字: " . InputTxt
            WriteLog(Text)
            Sleep, 200
            SendInput, % InputTxt  ; 表达式风格，兼容 v1.1
            Sleep, 200
        }
    }

    WriteLog("流程 #" . RunIndex . " 完成")
}

WriteLog("全部流程完成")
MsgBox, 64, 完成, 全部 %LoopCount% 次自动流程执行结束！

if (pGdipToken)
    Gdip_Shutdown(pGdipToken)

ExitApp

; ---------------- FindAndClick ----------------
FindAndClick(ImagePath, ClickDelay, TimeoutSec, GDIP_Avail, pToken)
{
    ; 确保坐标模式是 Screen，防止依赖用户的设置
    CoordMode, Pixel, Screen
    CoordMode, Mouse, Screen

    StartTime := A_TickCount
    Fuzz := 30 ; 容错率

    Loop
    {
        ; 查找图片
        ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *%Fuzz% %ImagePath%
        if (ErrorLevel = 0)
        {
            ; 默认点击位置 (左上角 + 20x20)
            CX := FoundX + 20
            CY := FoundY + 20

            ; GDI+ 精确计算中心点
            if (GDIP_Avail && pToken)
            {
                ; 仅当 Gdip_CreateBitmapFromFile 函数存在时才尝试执行
                if IsFunc("Gdip_CreateBitmapFromFile")
                {
                    W := 0, H := 0
                    try {
                        pBitmap := Gdip_CreateBitmapFromFile(ImagePath)
                        if (pBitmap)
                        {
                            W := Gdip_GetImageWidth(pBitmap)
                            H := Gdip_GetImageHeight(pBitmap)
                            Gdip_DisposeImage(pBitmap)
                            ; 计算中心点
                            CX := FoundX + (W // 2)
                            CY := FoundY + (H // 2)
                        }
                    } catch {} ; 【核心修改】消除 'e {}' 错误，且兼容 v1.1
                }
            }

            Click, %CX%, %CY%
            Sleep, %ClickDelay%
            return 1 ; 成功找到并点击
        }

        ; 检查是否超时
        if ((A_TickCount - StartTime) > TimeoutSec * 1000)
            return 0 ; 查找失败，超时

        Sleep, 200 ; 暂停后继续查找
    }
}