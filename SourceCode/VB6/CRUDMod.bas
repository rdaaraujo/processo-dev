Attribute VB_Name = "CRUDMod"
Public Sub ImportarArquivoLote(ArqImp As String)

Dim rs As Object
Dim filePath As String
Dim ssql As String
Dim line As String
Dim fields() As String
Dim isFirstLine As Boolean

Dim url As String
Dim jsonData As String
Dim WinHttpReq As Object

Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

isFirstLine = True

    If ArqImp <> "" Then
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
            
            url = "http://localhost:7500/api/Values"

            jsonData = "{""nome"": """ & fields(1) & """, ""situacao"": """ & fields(2) & """, ""cpf"": """ & fields(3) & """, ""dataNasc"": """ & fields(4) & """, ""endereco"": """ & fields(5) & """, ""telefone"": """ & fields(6) & """, ""email"": """ & fields(7) & """}"

            WinHttpReq.Open "Post", url, False
            WinHttpReq.setRequestHeader "Content-Type", "application/json"
            WinHttpReq.send jsonData
            
SkipLine:
        Loop
    
        Close #1
    
        Set rs = Nothing
    
        If WinHttpReq.Status = 200 Then
            MsgBox "Arquivo importado com sucesso ao Estoque!"
            
            EscreveLog ("Arquivo '" & ArqImp & "' foi importado com sucesso ao Estoque!")
    
            LimpaDadosgridEstoque
            ArqImp = ""
            
        Else
            MsgBox "Erro na requisição: " & WinHttpReq.Status & " - " & WinHttpReq.StatusText
            
        End If
        
        Set WinHttpReq = Nothing
    
    Else
        MsgBox "Não foi selecionado o arquivo para importação!"
        
    End If

End Sub

Public Sub AtualizaCliente(idCliente As String)

Dim ssql As String
Dim rsId As Object

Dim url As String
Dim WinHttpReq As Object

    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

    GeralMod.OpenConnection
    
    If idCliente <> "" Then
    
        ssql = "SELECT pv.id, pv.situacao_str FROM psf_clientes pv " & _
               "WHERE pv.id = " & idCliente & ""
        Set rsId = conn.Execute(ssql)
        
        If Not rsId.EOF Then
            If UCase(rsId.fields("situacao_str").Value) <> "BAIXADO" Then
                
                url = "http://localhost:7500/api/Values?Id=" & idCliente & ""
        
                WinHttpReq.Open "Put", url, False
                WinHttpReq.setRequestHeader "Content-Type", "application/json"
                WinHttpReq.send
    
                If WinHttpReq.Status = 200 Then
                    MsgBox "Foi alterada a situacao do cliente com o Id " & idCliente & "!"
                
                Else
                    MsgBox "Erro na requisição: " & WinHttpReq.Status & " - " & WinHttpReq.StatusText
                
                End If
                
                Set WinHttpReq = Nothing

                EscreveLog ("Foi alterado o status do cliente com o Id " & idCliente & "!")
                
                LimpaDadosgridEstoque
                idCliente = ""
            
            End If
        
        Else
            MsgBox "Não existe nenhum cliente com esse Id em seu Estoque!"
            idCliente = ""
            
        End If
    
    Else
        MsgBox "Digite o ID do Cliente para atualizar seu Status!"

    End If

End Sub

Public Sub DeletarCliente(idCliente As String)

Dim ssql As String
Dim rsId As Object

Dim url As String
Dim WinHttpReq As Object

    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    GeralMod.OpenConnection
    
    If idCliente <> "" Then
    
        ssql = "SELECT pv.id FROM psf_clientes pv " & _
               "WHERE pv.id = " & idCliente & ""
        Set rsId = conn.Execute(ssql)
        
        If Not rsId.EOF Then
            
            url = "http://localhost:7500/api/Values?Id=" & idCliente & ""
        
            WinHttpReq.Open "Delete", url, False
            WinHttpReq.setRequestHeader "Content-Type", "application/json"
            WinHttpReq.send
            
            If WinHttpReq.Status = 200 Then
                    MsgBox "O cliente com Id IdCliente foi deletado de seu Estoque!"
                
                Else
                    MsgBox "Erro na requisição: " & WinHttpReq.Status & " - " & WinHttpReq.StatusText
                
                End If
                
                Set WinHttpReq = Nothing

            EscreveLog ("O cliente com Id IdCliente foi deletado de seu Estoque!")
            
            AjusteMod.LimpaDadosgridEstoque
            idCliente = ""
            
        Else
            MsgBox "Não existe nenhum cliente com esse Id em seu Estoque!"
            idCliente = ""
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
