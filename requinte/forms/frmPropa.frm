VERSION 5.00
Object = "{C18424FD-277B-4362-B7A8-2788E7DBF8B4}#1.0#0"; "QPROGIF.OCX"
Begin VB.Form frmPropa 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   240
      ScaleHeight     =   6735
      ScaleWidth      =   9495
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      Begin prjQProGIF.QProGIF QProGIF1 
         Height          =   5175
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9128
         Filename        =   "C:\market\Logo_Descendo.gif"
      End
   End
End
Attribute VB_Name = "frmPropa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub Form_Activate()
    bActive = True
End Sub

Private Sub Form_Click()
    QProGIF1.FileName = ""
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    QProGIF1.FileName = ""
    Unload Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    QProGIF1.FileName = ""
    Unload Me
End Sub

Private Sub QProGIF1_Click()
    QProGIF1.FileName = ""
    Unload Me
End Sub
