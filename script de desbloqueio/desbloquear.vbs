'UNIDADE ORGANIZACIONAL QUE VAI SER CHECADA
Ou = "ou=management,dc=redes,dc=br"

'ARQUIVO CSV QUE VAI SER CHECADO, NA MESMA PASTA OU PASSAR O CAMINHO
arquivoCSV ="planilha.csv"

'Logs
ArquivoLOG ="DesBloqueados.log"
ArquivoLOG2 ="NaoDesBloqueados.log"


dim fs,objTextFile
set fs= CreateObject("Scripting.FileSystemObject")
Set Shell = CreateObject("wscript.shell")
dim userArq
dim status


'Gravar no arquivo de log
Dim Log,Log2
Set objFSO = CreateObject("Scripting.FileSystemObject")
Const ForAppending = 8
Set Log = objFSO.OpenTextFile(ArquivoLOG, ForAppending, True)
Set Log2 = objFSO.OpenTextFile(ArquivoLOG2, ForAppending, True)

Set objTextFile = fs.OpenTextFile(arquivoCSV)    
 
Do while NOT objTextFile.AtEndOfStream
    userArq = split(objTextFile.ReadLine,",")
    status = false 
    Set ListUsuarios = GetObject("LDAP://"+Ou+"")
    For Each usuarioAD in ListUsuarios    
         If  CStr(usuarioAD.userPrincipalName) = CStr(userArq(0)) Then
      
                Set objUser = GetObject _
                ("LDAP://cn="+usuarioAD.Get("cn")+","+Ou+"")
                objUser.AccountDisabled = FALSE
                objUser.SetInfo   

                status = True
                Log.WriteLine ("DesBloqueado:"&usuarioAD.Get("cn")&" NomeAD:"&usuarioAD.userPrincipalName&" NomeArquivo:"&userArq(0)&" "& Date &" " & Time)
 
        
        End If 
    Next 
    If status = false  Then
        Log2.WriteLine ("Usuario: "&userArq(0))
    End if
    
Loop
objTextFile.Close  
Log.Close
Log2.Close
set objTextFile = Nothing
set fs = Nothing






