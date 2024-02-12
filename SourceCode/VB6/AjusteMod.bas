Attribute VB_Name = "AjusteMod"
Option Explicit

Private Declare Function SetWindowLong Lib "user32" _
        Alias "SetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long, ByVal dwNewLong As _
        Long) As Long
Private Declare Function GetWindowLong Lib "user32" _
        Alias "GetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long) As Long

Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const GWL_STYLE As Long = (-16)

Public Sub ConfigJanela(ByVal hwnd As Long)
    ' Configura o estilo da janela para remover botões Minimizar e Maximizar
    Dim lWnd As Long
    lWnd = GetWindowLong(hwnd, GWL_STYLE)
    lWnd = lWnd And Not WS_MINIMIZEBOX
    lWnd = lWnd And Not WS_MAXIMIZEBOX
    Call SetWindowLong(hwnd, GWL_STYLE, lWnd)
    
End Sub

Public Sub CarregaDadosgridEstoque()

Dim ssql As String
Dim rsDados As Object
Dim row As Long
    
    OpenConnection
    
    ssql = "SELECT * FROM psf_clientes ORDER BY 1"
    conn.Execute (ssql)
    
    Set rsDados = CreateObject("ADODB.Recordset")
    rsDados.Open ssql, conn
    
    row = 1
    While Not rsDados.EOF
        ClientesFrm.gridEstoque.row = row
        ClientesFrm.gridEstoque.col = 1
        ClientesFrm.gridEstoque.Text = rsDados.fields("id").Value
        
        ClientesFrm.gridEstoque.col = 2
        ClientesFrm.gridEstoque.Text = rsDados.fields("nome_str").Value
        
        ClientesFrm.gridEstoque.col = 3
        ClientesFrm.gridEstoque.Text = rsDados.fields("situacao_str").Value
        
        ClientesFrm.gridEstoque.col = 4
        ClientesFrm.gridEstoque.Text = rsDados.fields("cpf_str").Value
        
        ClientesFrm.gridEstoque.col = 5
        ClientesFrm.gridEstoque.Text = rsDados.fields("data_nascimento").Value
        
        ClientesFrm.gridEstoque.col = 6
        ClientesFrm.gridEstoque.Text = rsDados.fields("endereco_str").Value
        
        ClientesFrm.gridEstoque.col = 7
        ClientesFrm.gridEstoque.Text = rsDados.fields("telefone_str").Value
        
        ClientesFrm.gridEstoque.col = 8
        ClientesFrm.gridEstoque.Text = rsDados.fields("email_str").Value
        
        row = row + 1
        rsDados.MoveNext
    Wend

    rsDados.Close
    conn.Close

End Sub

Public Sub LimpaDadosgridEstoque()

Dim row As Integer
Dim col As Integer
    
    For row = 0 To ClientesFrm.gridEstoque.MaxRows - 1
        For col = 0 To ClientesFrm.gridEstoque.MaxCols - 1
            ClientesFrm.gridEstoque.SetText row, col, ""
        Next col
    Next row
    
CarregaDadosgridEstoque

End Sub

Public Function ValidaEntradaTxt(KeyAscii As Integer) As Boolean
    
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        ValidaEntradaTxt = False
    
    Else
        ValidaEntradaTxt = True
    
    End If
    
End Function

Public Sub FecharForm(ClientesForm As Object)
  
  Unload ClientesForm

End Sub
