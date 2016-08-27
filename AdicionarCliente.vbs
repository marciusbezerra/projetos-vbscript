on error resume next
Set Classe = createobject("ClientesRemotos.clsClientes")
if err.number <> 0 then
	msgbox "Ocorreu um erro, para que este script funcione, você deve registrar o COM ClientesRemotos.dll e copiar o banco de dados NWIND.MDB para c:\Servicos\Temp.", vbcritical, "Atenção"
	wscript.quit
end if
ret = Classe.AdicionarCliente("NOVO", "nova companhia", "novo contato", _
    "novo tít. do contato", "novo endereço", "nova cidade", "NovoRegião", _
    "NovoCEP", "NovoPaís", "NovoFone", "NovoFax")
if err.number <> 0 then
	msgbox err.description, vbcritical, "Atenção"
	wscript.quit
end if
If ret = True Then
    MsgBox "Novo cliente adicionado com êxito!", vbInformation, "Atenção"
Else
    MsgBox "O cliente não pode ser adicionado. Talvez o Talvez a chave primária esteja repetida.", vbCritical, "Atenção"
End If