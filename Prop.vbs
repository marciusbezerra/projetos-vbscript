msg = "Propriedades do Windows Scripting Host" + vbcrlf + vbcrlf
msg = msg + "Application:" + chr(9) + wscript.application + vbcrlf
msg = msg + "FullName:" + chr(9) + chr(9) + wscript.fullname + vbcrlf
msg = msg + "Name:" + chr(9) + chr(9) + wscript.name + vbcrlf
msg = msg + "Path:" + chr(9) + chr(9) + wscript.path + vbcrlf
msg = msg + "ScriptFullName:" + chr(9) + wscript.scriptfullname + vbcrlf
msg = msg + "Version:" + chr(9) + chr(9) + wscript.version + vbcrlf
msg = msg + "ScriptName:" + chr(9) + wscript.scriptname

set s = wscript.createobject("wscript.shell")

s.popup msg, 5, "VBScript", 0