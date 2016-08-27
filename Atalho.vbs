'Criar atalhos com Windows Scripting Host

set s = wscript.createobject("wscript.shell")
set fsys = createobject("Scripting.FileSystemObject")
desk = s.specialfolders("desktop")
locacao = inputbox("Localização do Visual Basic","Reponda", _
          "c:\arquivos de programas\microsoft Visual Studio\vb98\vb6.exe")
if trim(locacao) = "" then wscript.quit 1

locacao1 = inputbox("Que projeto deve ser aberto com o Visual Basic","Reponda")

atalho = desk & "\Visual Basic 6 - Interprise Ediction.lnk"

if fsys.fileexists(atalho) then fsys.deletefile atalho, true

set ob_atalho = s.createshortcut(atalho)
ob_atalho.targetpath = locacao
ob_atalho.arguments = locacao1
ob_atalho.hotkey = "ALT+V"
ob_atalho.WindowStyle = 2
ob_atalho.save

msg = "Foi criado um atalho para o Visual Basic na área de trabalho." & vbcrlf 
msg = msg & "Use ALT+V para acessa-lo."
wscript.echo msg   