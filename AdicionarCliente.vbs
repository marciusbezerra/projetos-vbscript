on error resume next
Set Classe = createobject("ClientesRemotos.clsClientes")
if err.number <> 0 then
	msgbox "Ocorreu um erro, para que este script funcione, voc� deve registrar o COM ClientesRemotos.dll e copiar o banco de dados NWIND.MDB para c:\Servicos\Temp.", vbcritical, "Aten��o"
	wscript.quit
end if
ret = Classe.AdicionarCliente("NOVO", "nova companhia", "novo contato", _
    "novo t�t. do contato", "novo endere�o", "nova cidade", "NovoRegi�o", _
    "NovoCEP", "NovoPa�s", "NovoFone", "NovoFax")
if err.number <> 0 then
	msgbox err.description, vbcritical, "Aten��o"
	wscript.quit
end if
If ret = True Then
    MsgBox "Novo cliente adicionado com �xito!", vbInformation, "Aten��o"
Else
    MsgBox "O cliente n�o pode ser adicionado. Talvez o Talvez a chave prim�ria esteja repetida.", vbCritical, "Aten��o"
End If