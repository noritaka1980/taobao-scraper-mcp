Set fso = CreateObject("Scripting.FileSystemObject")
Dim dir : dir = fso.GetParentFolderName(WScript.ScriptFullName)
Set ws = CreateObject("WScript.Shell")
ws.CurrentDirectory = dir

' 优先用 pythonw（无黑窗），找不到就用 python
Dim pyExe
If ws.Run("where pythonw", 0, True) = 0 Then
    pyExe = "pythonw"
Else
    pyExe = "python"
End If

' 第三个参数 False = 不等待，wscript 启动后立刻退出，GUI 独立运行
ws.Run pyExe & " """ & dir & "\scraper_gui.py""", 0, False
