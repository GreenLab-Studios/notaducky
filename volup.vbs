Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

do
    DoEvents
    DoEvents
    Sleep(2000)
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys(chr(&hAF))
loop