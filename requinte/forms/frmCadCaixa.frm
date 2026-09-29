VERSION 5.00
Begin VB.Form frmCadCaixa 
   Caption         =   "Cadastro de Caixas"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   Icon            =   "frmCadCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   6840
      TabIndex        =   4
      Top             =   2475
      Width           =   6870
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "F7 - Sal&var"
         Height          =   735
         Left            =   3240
         Picture         =   "frmCadCaixa.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "ESC - &Fecha&r"
         Height          =   735
         Left            =   5040
         Picture         =   "frmCadCaixa.frx":0914
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox txtIp 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ComboBox cmbxCaixa 
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Caixa:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   435
   End
End
Attribute VB_Name = "frmCadCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    On Error GoTo trata_ERRO
    If txtIp.text = "" Then
        MsgBox "Informe o IP", vbExclamation
        txtIp.SetFocus
        Exit Sub
    End If
    If cmbxCaixa.text = "" Then
        MsgBox "Informe o IP", vbExclamation
        txtIp.SetFocus
        Exit Sub
    End If
    'If strExclcai_usu <> "S" Then
    '    MsgBox "Você não tem permissão para esta operação!", vbExclamation
    '    Exit Sub
    'End If
    
    strSQL = "delete from tab_caixasip "
    
    strSQL = strSQL & " where ip_cai='" & txtIp.text & "'"
    Conn.Execute strSQL
    strSQL = "insert into tab_caixasip(cod_cai,ip_cai)values("
    strSQL = strSQL & "'" & Left(cmbxCaixa.text, 2) & "',"
    strSQL = strSQL & "'" & txtIp.text & "')"
    Conn.Execute strSQL
    
    
    MsgBox "Dados gravados com sucesso! ", vbInformation
    
    IniciaVariaveis
    
    Unload Me
    Exit Sub
trata_ERRO:
    P_MostraErro
    
    
End Sub

Private Sub Form_Load()
    P_PreencheCaixas cmbxCaixa
    txtIp.text = strHost_Usu
End Sub
