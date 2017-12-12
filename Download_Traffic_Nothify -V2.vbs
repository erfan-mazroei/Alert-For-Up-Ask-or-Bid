Set oShell = CreateObject("Wscript.Shell")    
oShell.Run "%comspec% /c echo " & Chr(7) ,0,false   
WScript.Sleep 1
oShell.Run "%comspec% /c echo " & Chr(7) ,0,false  
WScript.Sleep 1
oShell.Run "%comspec% /c echo " & Chr(7) ,0,false  
Set oShell = nothing