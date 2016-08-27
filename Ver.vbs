
set arg = wscript.arguments

if arg.count = 0 then
	msgbox "Sintaxe: VER Argumentos ...",,"Nenhum argumento"
	wscript.quit
end if

if arg.count >= 1 then
	for each a in arg
		txt = txt & a & vbcrlf
	next
	msgbox "Você digitou (todos) esse(s) arqumento(s):" & vbcrlf & vbcrlf & _
		txt & vbcrlf & "Sintaxe: VER ArquivoDeTexto",,"Muitos argumentos"
	wscript.quit
end if
