VERSION 5.00
Begin VB.Form frmCFOP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CFOP"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCFOP 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmCFOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Me.Hide
End If
End Sub

Private Sub Form_Load()
    P_PreencheNotaSaida lstCFOP
End Sub

Private Sub lstCFOP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.Tag = lstCFOP.Text
        Me.Hide
    End If
End Sub
