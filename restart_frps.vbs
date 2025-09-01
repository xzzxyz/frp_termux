Dim Wsh
Set Wsh = WScript.CreateObject("WScript.Shell")
Wsh.Run "taskkill /f /im frps.exe",0
Set Wsh=NoThing
WScript.Sleep(1000)
Set ws = CreateObject("Wscript.Shell") 
ws.run "cmd /c .\frps.exe -c .\frps.toml",vbhide