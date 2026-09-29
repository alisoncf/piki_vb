VERSION 5.00
Begin VB.Form frmBolInstr 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Texto de Instrução no Boleto"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Salvar"
      Height          =   735
      Left            =   3840
      Picture         =   "frmBolInstr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   5280
      Picture         =   "frmBolInstr.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtbxInstr 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1814
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmBolInstr.frx":0294
      Top             =   120
      Width           =   6349
   End
End
Attribute VB_Name = "frmBolInstr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    On Error Resume Next
    strSQL = "update tab_bolconfig set instr_bol='" & txtbxInstr.Text & "'"
    conn.Execute strSQL
End Sub

Private Sub Form_Activate()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select instr_bol from tab_bolconfig "
    rPesq.Open strSQL, conn_client
    With rPesq
        txtbxInstr.Text = Format(!instr_bol)
        .Close
    End With
End Sub
