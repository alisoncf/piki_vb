VERSION 5.00
Begin VB.Form frmCadBan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Bancos"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
   Icon            =   "frmCadBan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox txtbxCod 
      Height          =   285
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   615
      Left            =   2520
      Picture         =   "frmCadBan.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   4440
      Picture         =   "frmCadBan.frx":1034
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome do Banco"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C&odigo:"
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   540
   End
End
Attribute VB_Name = "frmCadBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxCod.Text = "" Then
        MsgBox "Informe o código!", vbExclamation
         
        Exit Sub
    End If
    If txtbxNome.Text = "" Then
        MsgBox "Informe o código!", vbExclamation
         
        Exit Sub
    End If
    
    If Me.tag = "" Then
        If F_BuscaUmCampo("Cod_Bco", "Tab_bancos", "cod_bco='" & _
            txtbxCod.Text & "'") <> "" Then
            MsgBox "Já existe um banco com este código!", vbExclamation
            Exit Sub
        End If
        If F_BuscaUmCampo("Cod_Bco", "Tab_bancos", "Nome_bco='" & _
            txtbxNome.Text & "'") <> "" Then
            MsgBox "Já existe um banco com este Nome!", vbExclamation
            Exit Sub
        End If
        Grava
    Else
        If F_BuscaUmCampo("Cod_Bco", "Tab_bancos", "Nome_bco='" & _
            txtbxNome.Text & "' And cod_bco <> '" & Me.tag & "'") <> "" Then
            MsgBox "Já existe um banco com este Nome!", vbExclamation
            Exit Sub
        End If
        Altera
    End If
End Sub
Sub Grava()
    Dim strSQL As String
    strSQL = "insert into tab_bancos(Cod_bco,Nome_Bco)Values("
    strSQL = strSQL & "'" & txtbxCod.Text & "',"
    strSQL = strSQL & "'" & txtbxNome.Text & "')"
    Conn.Execute strSQL
    MsgBox "Banco gravado!", vbInformation
    txtbxCod.Text = ""
    txtbxNome.Text = ""
    txtbxNome.SetFocus
End Sub
Sub Altera()
    Dim strSQL As String
    strSQL = "update tab_bancos set "
    
    strSQL = strSQL & " nome_bco = '" & txtbxNome.Text & "'"
    strSQL = strSQL & " where cod_bco = '" & Me.tag & "' "
    Conn.Execute strSQL
    MsgBox "Banco alterado!", vbInformation
    Unload Me
End Sub


Private Sub txtbxCod_GotFocus()
    P_SelecTexto txtbxCod
End Sub

Private Sub txtbxCod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCod_LostFocus()
    P_AjustaTexto txtbxCod
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
