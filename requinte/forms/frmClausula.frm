VERSION 5.00
Begin VB.Form frmClausula 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cláusula"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbxCab 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   840
      Width           =   5655
   End
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "Confirmar"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtbxTexto 
      Height          =   1245
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Título:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   465
   End
   Begin VB.Label lblClausula 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cláusula Contratual"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Texto:"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   1680
      Width           =   450
   End
End
Attribute VB_Name = "frmClausula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnOk_Click()
    If txtbxCab.tag = "" Then
        Grava
    Else
        Altera
    End If
End Sub
Sub Grava()
    Dim sSeq As String
    sSeq = MaxKeySQL("tab_clausulas", "seq_cla", "seq_CON=" & Me.tag, "")
    strSQL = "insert into tab_clausulas(seq_con,seq_cla,cab_cla,texto_cla)values("
    strSQL = strSQL & Me.tag & ","
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & txtbxCab.Text & "',"
    strSQL = strSQL & "'" & txtbxTexto.Text & "')"
    Conn.Execute strSQL
    Me.tag = sSeq
    MsgBox "Cláusula Gravada!", vbInformation
    
    Unload Me
    
    
End Sub
Sub Altera()
    Dim strSQL As String
    strSQL = "update tab_clausulas set "
    strSQL = strSQL & "cab_cla='" & txtbxCab.Text & "',"
    strSQL = strSQL & "texto_cla='" & txtbxTexto.Text & "' "
    strSQL = strSQL & " where seq_con=" & Me.tag & " "
    strSQL = strSQL & " and seq_cla=" & txtbxCab.tag & " "

    Conn.Execute strSQL
    MsgBox "Cláusula Alterada!", vbInformation
    Unload Me
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub txtbxCab_LostFocus()
    P_AjustaTexto txtbxCab, -1
End Sub

Private Sub txtbxTexto_LostFocus()
    P_AjustaTexto txtbxTexto, -1
End Sub
