VERSION 5.00
Begin VB.Form frmCadPlano 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Receita/Despesa"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5700
   Icon            =   "frmCadPla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Gravar"
      Height          =   735
      Left            =   1920
      Picture         =   "frmCadPla.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   3600
      Picture         =   "frmCadPla.frx":05F2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.OptionButton optnDeb 
      Caption         =   "&Débito"
      Height          =   195
      Left            =   4200
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.OptionButton optnCred 
      Caption         =   "&Crédito"
      Height          =   195
      Left            =   3240
      TabIndex        =   2
      Top             =   1800
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label lblTipo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   540
      Left            =   3540
      TabIndex        =   4
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome da Conta:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1155
   End
End
Attribute VB_Name = "frmCadPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()

End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If F_Vazio(txtbxNome.Text) Then
        MsgBox "Informe o nome da Conta que deseja incluir!", vbExclamation
        Exit Sub
    End If
    If Me.tag = "" Then
        If F_BuscaUmCampo("Cod_Pla", "tab_planocontas", "Nome_Pla='" & _
            txtbxNome.Text & "'") <> "" Then
            MsgBox "Já existe uma conta com este Nome!", vbExclamation
            Exit Sub
        End If
        grava
        txtbxNome.Text = ""
        txtbxNome.SetFocus
    Else
        If F_BuscaUmCampo("Cod_Pla", "tab_planocontas", "Nome_Pla='" & _
            txtbxNome.Text & "' and cod_pla<>'" & Me.tag & "'") <> "" Then
            MsgBox "Você está tentanto mudar o nome desta conta para um que já existe!", vbExclamation
            Exit Sub
        End If
        altera
    End If
    If IsLoad("frmEdiPlanoContas") Then
        RefreshDataPosition frmEdiPlanoContas.dtPla
    End If
    
End Sub
Sub grava()
    Dim strSQL As String
    strSQL = "insert into tab_planocontas(cod_pla,nome_pla,tipo_pla)Values("
    strSQL = strSQL & "'" & MaxKeySQL("tab_planocontas", "Cod_Pla", "tipo_pla='" & Left(lblTipo.caption, 1) & "'", "00#") & "',"
    strSQL = strSQL & "'" & txtbxNome.Text & "',"
    strSQL = strSQL & "'" & Left(lblTipo.caption, 1) & "')"
    Banco.Execute strSQL
    MsgBox "Conta incluída com sucesso!", vbExclamation
    If IsLoad("frmEdiPlanoContas") Then
        frmEdiPlanoContas.dtPla.Refresh
    End If
End Sub
Sub altera()
    Dim strSQL As String
    strSQL = "Update tab_planocontas set "
    
    strSQL = strSQL & "Nome_Pla='" & txtbxNome.Text & "' "
    strSQL = strSQL & " where cod_pla='" & Me.tag & "'     "
    Banco.Execute strSQL
    MsgBox "Conta alterada!", vbInformation
    If IsLoad("frmEdiPlanoContas") Then
        RefreshDataPosition frmEdiPlanoContas.dtPla
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    P_CentraForm Me
End Sub

Private Sub optnCred_Click()
    lblTipo.caption = "Crédito"
    lblTipo.ForeColor = vbBlue
End Sub

Private Sub optnDeb_Click()
    lblTipo.caption = "Débito"
    lblTipo.ForeColor = vbRed
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub
