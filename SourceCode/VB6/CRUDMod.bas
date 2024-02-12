Attribute VB_Name = "CRUDMod"
Public Sub ImportarArquivoLote(ArqImp As String)

Dim rs As Object
Dim filePath As String
Dim ssql As String
Dim line As String
Dim fields() As String
Dim isFirstLine As Boolean

isFirstLine = True

    If ArqImp <> "" Then
        OpenConnection
        Set rs = CreateObject("ADODB.Recordset")
        
        filePath = "" & ArqImp & ""
    
        Open filePath For Input As #1
    
        Do Until EOF(1)
            Line Input #1, line
    
            If isFirstLine Then
                isFirstLine = False
                GoTo SkipLine
                
            End If
    
            fields = Split(line, ";")
            
            ssql = "INSERT INTO psf_clientes (nome_str, situacao_str, cpf_str, data_nascimento, endereco_str, telefone_str, email_str) VALUES ('" & fields(1) & "', '" & fields(2) & "', '" & fields(3) & "', '" & fields(4) & "', '" & fields(5) & "', '" & fields(6) & "', '" & fields(7) & "')"
            conn.Execute ssql
            
SkipLine:
        Loop
    
        Close #1
        conn.Close
    
        Set rs = Nothing
        Set conn = Nothing
    
        MsgBox "Arquivo importado com sucesso ao Estoque!"
        
        EscreveLog ("Arquivo '" & ArqImp & "' foi importado com sucesso ao Estoque!")
        
        LimpaDadosgridEstoque
        ArqImp = ""
        
    Else
        MsgBox "Não foi selecionado o arquivo para importação!"
        
    End If

End Sub

Public Sub AtualizaCliente(IdCliente As String)

Dim ssql As String
Dim rsId As Object
    
    GeralMod.OpenConnection
    
    If IdCliente <> "" Then
    
        ssql = "SELECT pv.id, pv.situacao_str FROM psf_clientes pv " & _
               "WHERE pv.id = " & IdCliente & ""
        Set rsId = conn.Execute(ssql)
        
        If Not rsId.EOF Then
            If UCase(rsId.fields("situacao_str").Value) <> "BAIXADO" Then
                ssql = "UPDATE psf_clientes SET situacao_str = 'BAIXADO' " & _
                " WHERE id = " & IdCliente & ";"
                conn.Execute (ssql)
                
                MsgBox "Foi alterada a situacao do cliente com o Id " & IdCliente & "!"
                
                EscreveLog ("Foi alterado o status do cliente com o Id " & IdCliente & "!")
                
                LimpaDadosgridEstoque
                IdCliente = ""
            
            End If
        
        Else
            MsgBox "Não existe nenhum cliente com esse Id em seu Estoque!"
            IdCliente = ""
            
        End If
    
    Else
        MsgBox "Digite o ID do Cliente para atualizar seu Status!"

    End If

End Sub

Public Sub DeletarCliente(Index As Integer, IdCliente As String)

Dim ssql As String
Dim rsId As Object
    
    GeralMod.OpenConnection
    
    If IdCliente <> "" Then
    
        ssql = "SELECT pv.id FROM psf_clientes pv " & _
               "WHERE pv.id = " & IdCliente & ""
        Set rsId = conn.Execute(ssql)
        
        If Not rsId.EOF Then
            ssql = "DELETE FROM psf_clientes pv WHERE pv.id = " & IdCliente & ";"
            conn.Execute (ssql)
            
            MsgBox "O cliente com Id IdCliente foi deletado de seu Estoque!"
            
            EscreveLog ("O cliente com Id IdCliente foi deletado de seu Estoque!")
            
            AjusteMod.LimpaDadosgridEstoque
            IdCliente = ""
            
        Else
            MsgBox "Não existe nenhum cliente com esse Id em seu Estoque!"
            IdCliente = ""
        End If
    
    Else
        MsgBox "Digite o ID do cliente para deletá-lo do Estoque!"
    End If

End Sub

Public Sub DelTodosClientes(Index As Integer)

    Dim ssql As String
      
      Dim retval
      retval = MsgBox("Tem certeza que deseja deletar todos os clientes do Estoque?", vbYesNo)
      If retval = 6 Then
    
        GeralMod.OpenConnection
        
            ssql = "DELETE FROM psf_clientes;"
            conn.Execute (ssql)

        conn.Close
        
        MsgBox "Todos os clientes foram deletados do seu Estoque!"
        
        EscreveLog ("Todos os clientes foram deletados de seu Estoque!")
        
        LimpaDadosgridEstoque
        
      Else
        Exit Sub
        
      End If

End Sub

Public Function EscreveLog(Textlog As String)

Dim logFilePath As String
Dim logText As String

logFilePath = Environ("USERPROFILE") & "\Desktop\log.txt"

logText = Format(Time(), "hh:mm:ss") & " - " & Textlog & vbCrLf & "*****************************************************************"

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

Dim logFile As Object
Set logFile = fso.OpenTextFile(logFilePath, 8, True)

logFile.WriteLine logText

logFile.Close

Set logFile = Nothing
Set fso = Nothing

End Function
