VERSION 5.00
Begin VB.Form frmCodTotGravar 
   Caption         =   "sped_totalizador_redz"
   ClientHeight    =   3600
   ClientLeft      =   1170
   ClientTop       =   450
   ClientWidth     =   9720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   9720
   Begin VB.ComboBox cmbxImposto 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2160
      Width           =   7335
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   9690
      TabIndex        =   8
      Top             =   2745
      Width           =   9720
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Salvar"
         Height          =   735
         Left            =   6600
         Picture         =   "frmCodTotGravar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   8040
         Picture         =   "frmCodTotGravar.frx":0487
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.TextBox txtDescricao 
      DataField       =   "descricao"
      Height          =   1005
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   7335
   End
   Begin VB.TextBox txtNome 
      DataField       =   "nome"
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "codigo"
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   1935
   End
   Begin VB.TextBox txtId 
      DataField       =   "id"
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Imposto:"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Width           =   600
   End
   Begin VB.Label lblLabels 
      Caption         =   "descricao:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "nome:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "codigo:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "id:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmCodTotGravar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNovo_Click()
    Dim strSQL As String
    If Me.Tag = "" Then
        strSQL = "insert into sped_totalizador_redz(codigo,nome,descricao,cod_imp)values("
        strSQL = strSQL & "'" & txtCodigo & "',"
        strSQL = strSQL & "'" & txtNome & "',"
        strSQL = strSQL & "'" & F_AjustaTexto(txtDescricao) & "',"
        strSQL = strSQL & "'" & Left(cmbxImposto.Text, 3) & "')"
        conn.Execute strSQL
        MsgBox "Gravado com sucesso!", vbInformation
    Else
        strSQL = "update sped_totalizador_redz set "
        strSQL = strSQL & "codigo='" & txtCodigo & "',"
        strSQL = strSQL & "nome='" & txtNome & "',"
        strSQL = strSQL & "descricao='" & txtDescricao & "',"
        strSQL = strSQL & "cod_imp='" & Left(cmbxImposto.Text, 3) & "'"
        strSQL = strSQL & " where id=" & Me.Tag
        conn.Execute strSQL
        MsgBox "alterado com sucesso!", vbInformation
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    P_PreencheImpostos cmbxImposto
End Sub
