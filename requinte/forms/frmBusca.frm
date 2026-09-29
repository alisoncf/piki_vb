VERSION 5.00
Begin VB.Form frmBusca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizar Produtos"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmBusca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtbxExpr 
      Height          =   285
      Left            =   120
      MaxLength       =   40
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite a Expressão e Tecle <Enter>"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2505
   End
End
Attribute VB_Name = "frmBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Me.Tag = ""
    txtbxExpr.SetFocus
        
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
    
    
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.Tag = txtbxExpr.Text
        Me.Hide
    ElseIf KeyCode = vbKeyEscape Then
        Me.Tag = ""
        Unload Me
        
    End If
End Sub
