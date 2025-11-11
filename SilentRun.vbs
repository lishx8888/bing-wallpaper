' 简单的静默执行VBS脚本
Option Explicit
Dim shell, cmd

' 检查是否提供了参数
If WScript.Arguments.Count = 0 Then
    WScript.Quit 1
End If

Set shell = CreateObject("WScript.Shell")
cmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -NoLogo -NoProfile -File """ & WScript.Arguments(0) & """"
shell.Run cmd, 0, True
Set shell = Nothing