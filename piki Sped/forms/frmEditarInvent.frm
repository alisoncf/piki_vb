VERSION 5.00
Begin VB.Form frmEditarInvent 
   Caption         =   "Editar"
   ClientHeight    =   1335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtAte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtDe 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Até:"
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   285
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "De:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblId 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   210
   End
End
Attribute VB_Name = "frmEditarInvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterar_Click()
    Dim strSQL As String
    inventario.AlteraInventario Me.tag, txtDe.Text, txtAte.Text
    Unload Me
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub
