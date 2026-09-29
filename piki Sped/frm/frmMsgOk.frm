VERSION 5.00
Begin VB.Form frmMsgOk 
   BackColor       =   &H00EF9C7C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mensagem"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5505
      ScaleWidth      =   8145
      TabIndex        =   1
      Top             =   120
      Width           =   8175
      Begin VB.VScrollBar VScroll1 
         Height          =   5520
         LargeChange     =   500
         Left            =   7920
         SmallChange     =   100
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblMsg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Mensagem de ECF"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   13680
         Left            =   105
         TabIndex        =   2
         Top             =   120
         Width           =   7575
      End
   End
   Begin VB.CommandButton cmndbtnOk 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Ok (Enter)"
      Default         =   -1  'True
      Height          =   735
      Left            =   6120
      Picture         =   "frmMsgOk.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   2175
   End
End
Attribute VB_Name = "frmMsgOk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnOk_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.Refresh
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Refresh
End Sub

Private Sub VScroll1_Change()
    lblMsg.Top = 0 - VScroll1.Value
End Sub
