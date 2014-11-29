Set WEBCHILLER =wscript.CreateObject("WScript.Shell")
do
wscript.sleep 100
WEBCHILLER.sendkeys "{CAPSLOCK}"
WEBCHILLER.sendkeys "{NUMLOCK}"
WEBCHILLER.sendkeys "{SCROLLLOCK}"
loop