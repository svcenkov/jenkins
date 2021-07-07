$wshell = New-Object -ComObject Wscript.Shell
$result = $wshell.Popup("Operation Completed",0,"Done",0x1)