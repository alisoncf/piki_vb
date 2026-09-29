VERSION 5.00
Begin VB.Form frmMsg 
   Appearance      =   0  'Flat
   BackColor       =   &H00EF9C7C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mensagem"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9180
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancela (Esc)"
      Height          =   735
      Left            =   6960
      Picture         =   "frmMsg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   2175
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6495
      LargeChange     =   100
      Left            =   8880
      SmallChange     =   50
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mensagem de ECF"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6465
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8745
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnCancela_Click()
'    MsgBox "aqui"
    Me.tag = "1"
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
'        MsgBox ""
    End If
End Sub

Private Sub VScroll1_Change()
    lblMsg.Top = 0 - VScroll1.Value
End Sub
