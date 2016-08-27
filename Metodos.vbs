'Métodos do Windows Scripting Host

wscript.echo "Marcius", "Carneiro", "Bezerra","-", _
     "Idade:",datediff("yyyy", datevalue("30/11/1975"), date())

set s = wscript.createobject("wscript.shell")

set vars = s.environment

for each var in vars
     msg = msg & var & vbcrlf     
next

msgbox msg, vbinformation, "Variáveis de ambiente"



msg = ""

set s = wscript.createobject("wscript.shell")
set pes = s.specialfolders

for each pe in pes
     msg = msg & pe & vbcrlf   
next

msgbox msg, vbinformation, "Pastas especiais"




set fsys = createobject("Scripting.FileSystemObject")
set s = wscript.createobject("wscript.shell")
desk = s.specialfolders("desktop")

set desktop = fsys.getfolder(desk)
set arqs_desk = desktop.files

msg = ""

for each arq in arqs_desk
     msg = msg & arq.name & " (" & arq.type & ")" & vbcrlf
next

msgbox msg, vbinformation, "Itens do Desktop"




set s = wscript.createobject("wscript.shell")
msg = "O Seu nome é Marcius ? (resposta em 5 segundos)"
resp = s.popup(msg, 5, "Pergunta", 4+32)

if resp = 6 then
     msg = "Sim, meu nome é Marcius."
elseif resp = 7 then
     msg = "Não, meu nome não é Marcius."
else
     msg = "Como nada foi respondido, vai ser assumido como sim." & vbcrlf & _
           "Meu nome é Marcius."
end if

wscript.echo msg




wscript.quit 1      'O código do erro pode ser qualquer 
                    'número, o padrão é zero

    