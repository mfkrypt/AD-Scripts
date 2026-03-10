Do While True
    CreateObject("WScript.Shell").Run "cmd /c net use \\server\share >nul 2>&1", 0, False
    WScript.Sleep 60000
Loop