VERSION 5.00
Begin VB.Form frmCadBan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Bancos"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "frmCadBan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4035
   Begin VB.Frame frControle 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3735
         Begin VB.TextBox txtbxCod 
            Height          =   285
            Left            =   3120
            MaxLength       =   3
            TabIndex        =   5
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtbxNome 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   3
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C&odigo:"
            Height          =   195
            Left            =   3120
            TabIndex        =   4
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Nome do Banco"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   1155
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   3735
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   615
            Left            =   1920
            Picture         =   "frmCadBan.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnSalvar 
            Caption         =   "&Confirmar"
            Height          =   615
            Left            =   120
            Picture         =   "frmCadBan.frx":0A14
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
      End
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
        grava
    Else
        If F_BuscaUmCampo("Cod_Bco", "Tab_bancos", "Nome_bco='" & _
            txtbxNome.Text & "' And cod_bco <> '" & Me.tag & "'") <> "" Then
            MsgBox "Já existe um banco com este Nome!", vbExclamation
            Exit Sub
        End If
        altera
    End If
End Sub
Sub grava()
    Dim strSQL As String
    strSQL = "insert into tab_bancos(Cod_bco,Nome_Bco)Values("
    strSQL = strSQL & "'" & txtbxCod.Text & "',"
    strSQL = strSQL & "'" & txtbxNome.Text & "'"
    conn.Execute strSQL
    MsgBox "Banco gravado!", vbInformation
    txtbxCod.Text = ""
    txtbxNome.Text = ""
    txtbxNome.SetFocus
End Sub
Sub altera()
    Dim strSQL As String
    strSQL = "update tab_bancos set "
    
    strSQL = strSQL & " nome_bco = '" & txtbxNome.Text & "'"
    strSQL = strSQL & " where cod_bco = '" & Me.tag & "' "
    conn.Execute strSQL
    MsgBox "Banco alterado!", vbInformation
    Unload Me
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxCod_GotFocus()
    ganhaFoco txtbxCod
End Sub

Private Sub txtbxCod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCod_LostFocus()
    AjustaTexto txtbxCod
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
