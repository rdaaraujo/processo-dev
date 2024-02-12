VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ClientesFrm 
   Caption         =   "PSF Clientes"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   ScaleHeight     =   5895
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraCliente 
      Caption         =   "Cliente"
      Height          =   1335
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   9185
      Begin VB.TextBox txtIdCliente 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdDeletarCliente 
         Caption         =   "Deletar Cliente"
         Height          =   360
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   795
         Width           =   1455
      End
      Begin VB.CommandButton cmdAtualizarCliente 
         Caption         =   "Atualizar Cliente"
         Height          =   360
         Left            =   1920
         TabIndex        =   6
         Top             =   795
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelAllCli 
         Caption         =   "Deletar todos os clientes!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   4800
         TabIndex        =   7
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lblIDCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Cliente: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   3720
         X2              =   3720
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Label lblCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   3090
      End
      Begin VB.Label lblDesejaDELETAR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deseja DELETAR todos os clientes?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   360
         Width           =   4305
      End
   End
   Begin VB.Frame FraLote 
      Caption         =   "Lote"
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   9185
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6480
         Top             =   -480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "..."
         Height          =   315
         Left            =   6000
         TabIndex        =   1
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtArqImp 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   5295
      End
      Begin VB.CommandButton cmdImportar 
         Caption         =   "Importar"
         Height          =   315
         Left            =   6480
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   7800
         TabIndex        =   3
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label lblArquivo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivo: "
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   660
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "14:47"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "12/02/2024"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Versão 1.0"
            TextSave        =   "Versão 1.0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   9172
            MinWidth        =   9172
            Text            =   "Rafael de Almeida Araújo"
            TextSave        =   "Rafael de Almeida Araújo"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread gridEstoque 
      Height          =   1935
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   9165
      _Version        =   393216
      _ExtentX        =   16166
      _ExtentY        =   3413
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   10
      SpreadDesigner  =   "Form1.frx":1084A
   End
   Begin MSComctlLib.Toolbar BarMenu 
      Height          =   600
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   9185
      _ExtentX        =   16193
      _ExtentY        =   1058
      ButtonWidth     =   1455
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgBarra"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Novo"
            Key             =   "Novo"
            Description     =   "Novo Registro"
            Object.ToolTipText     =   "Novo Registro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Editar"
            Key             =   "Editar"
            Description     =   "Modificar Registro"
            Object.ToolTipText     =   "Modificar Registro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Apagar"
            Key             =   "Apagar"
            Description     =   "Apagar Registro"
            Object.ToolTipText     =   "Apagar Registro"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Fechar"
            Key             =   "Fechar"
            Description     =   "Fechar Programa"
            Object.ToolTipText     =   "Fechar Programa"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   1
      Begin MSComctlLib.ImageList ImgBarra 
         Left            =   8400
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":10CC7
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":21521
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":31D7B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":425D5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraCli 
      Height          =   5775
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "ClientesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NomeArq As String

Private Sub Form_Load()
    
    AjusteMod.ConfigJanela Me.hwnd
    EscreveLog ("Serviço Iniciado.")
    
    CarregaDadosgridEstoque

End Sub

Private Sub Form_Resize()
    
    Me.Height = 6480
    Me.Width = 9900
    
End Sub

Public Sub barMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "Novo"
      cmdImport_Click
    
    Case "Editar"
      MsgBox "Coloque o Id do Cliente desejado e clique em Atualizar Cliente!"
      txtIdCliente.SetFocus
    
    Case "Apagar"
      MsgBox "Coloque o Id do Cliente desejado e clique em Deletar Cliente!"
      txtIdCliente.SetFocus
    
    Case "Fechar"
      FecharForm Me
  
  End Select
End Sub

Public Sub cmdImport_Click()

    CommonDialog1.Filter = "Apps (*.csv, *.txt, *.xlsx)|*.csv;*.txt;*.xlsx|All Files (*.*)|*.*"
    CommonDialog1.DefaultExt = "csv"
    CommonDialog1.DialogTitle = "Importar Arquivo"
    CommonDialog1.ShowOpen
    
    NomeArq = CommonDialog1.FileName
    txtArqImp = NomeArq

End Sub

Public Sub cmdImportar_Click()

Dim ArqImp As String
    
    ArqImp = txtArqImp.Text
    ImportarArquivoLote ArqImp
    
    txtArqImp = ""

End Sub

Private Sub cmdCancelar_Click()

    If txtArqImp.Text <> "" Then
    Dim retval
      retval = MsgBox("Você tem certeza que deseja cancelar a importação?", vbYesNo)
      If retval = 6 Then
         txtArqImp.Text = ""
      End If
    End If
  
End Sub

Private Sub txtIdCliente_KeyPress(KeyAscii As Integer)

    KeyAscii = IIf(ValidaEntradaTxt(KeyAscii), KeyAscii, 0)

End Sub

Public Sub cmdDeletarCliente_Click(Index As Integer)

Dim IdCliente As String
    
    IdCliente = Val(txtIdCliente.Text)
    DeletarCliente Index, IdCliente
    
    txtIdCliente = ""

End Sub
Public Sub cmdAtualizarCliente_Click()

Dim IdCliente As String
    
    IdCliente = Val(txtIdCliente.Text)
    AtualizaCliente IdCliente
    
    txtIdCliente = ""

End Sub

Public Sub cmdDelAllCli_Click(Index As Integer)

    DelTodosClientes Index

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    On Local Error Resume Next
    If MsgBox("Confirma saída?", vbYesNo + vbDefaultButton2, "PSF Clientes") = vbNo Then Cancel = 1
    EscreveLog ("Serviço Finalizado.")
   
End Sub
