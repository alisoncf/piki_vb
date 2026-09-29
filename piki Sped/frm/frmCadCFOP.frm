VERSION 5.00
Begin VB.Form frmCadCFOP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inclusão Alteração de CFOP"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.OptionButton optnEntrada 
         Caption         =   "Entrada"
         Height          =   195
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtbxCFOP 
         Height          =   285
         Left            =   840
         MaxLength       =   18
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtbxNome 
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1200
         Width           =   4455
      End
      Begin VB.OptionButton optnSaida 
         Caption         =   "Saída"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCpf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   885
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   465
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fecha&r"
      Height          =   735
      Left            =   3960
      Picture         =   "frmCadCFOP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Sal&var"
      Height          =   735
      Left            =   2400
      Picture         =   "frmCadCFOP.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmCadCFOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCNPJ As Boolean 'para saber se o CNPJ é válido


Private Sub cmndbtnSair_Click()
    Me.Hide
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxCFOP.Text = "" Then
        MsgBox "Informe o código CFOP!", vbExclamation
        txtbxCFOP.SetFocus
        Exit Sub
    End If
    If txtbxNome.Text = "" Then
        MsgBox "Informe o nome do CFOP!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    
    If Me.tag = "" Then
        
    
        P_GravaCFOP txtbxCFOP, txtbxNome, IIf(optnSaida.Value = True, "S", "E")
        MsgBox "Gravado!", vbInformation
        txtbxNome.Text = ""
        txtbxCFOP.Text = ""
        txtbxCFOP.SetFocus

        
    Else
    
        P_AlteraCFOP txtbxCFOP, txtbxNome, IIf(optnSaida.Value = True, "S", "E")
        MsgBox "CFOP Alterado!", vbInformation
        
    End If
    
End Sub



Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
    
End Sub



Private Sub optnPessoa_Click(Index As Integer)
    If Index = 0 Then
        lblcpf.caption = "&CPF:"
        'lblRg.caption = "&RG:"
        txtbxCNPJ.MaxLength = 14
    Else
        lblcpf.caption = "&CNPJ:"
        'lblRg.caption = "&Insc Est:"
        txtbxCNPJ.MaxLength = 18
    End If
    'txtbxCNPJ.Text = ""
End Sub





Private Sub txtbxCFOganhafoco()
    ganhaFoco txtbxCFOP
End Sub

Private Sub txtbxCFOP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCFOP_LostFocus()
    P_AjustaNota txtbxCFOP, 4
End Sub

Private Sub txtbxNome_GotFocus()
    ganhaFoco txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    AjustaTexto txtbxNome
End Sub


