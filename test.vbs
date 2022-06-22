Dim i
i = 0
set shell = createobject("wscript.Shell")

Do while i = 0
shell.Sendkeys"{NUMLOCK}"
wscript.echo time
wscript.sleep 60000
Loop