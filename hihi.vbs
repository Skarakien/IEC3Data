MsgBox "Look at your keyboard for a cool light show!" 
set wshShell = wscript.CreateObject("wscript.shell")
wscript.Createobject("WScript.Shell")
do
wscript.sleep 100
wshShell.sendkeys"{NUMLOCK}"
wshShell.sendkeys"{CAPSLOCK}"
wshShell.sendkeys"{SCROLLLOCK}"
loop