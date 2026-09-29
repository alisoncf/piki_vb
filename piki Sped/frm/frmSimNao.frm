VERSION 5.00
Begin VB.Form frmSimNao 
   BackColor       =   &H00EF9C7C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Escolha Sim ou Não"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnSim 
      Caption         =   "&Sim"
      Default         =   -1  'True
      Height          =   735
      Left            =   4680
      Picture         =   "frmSimNao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnNao 
      Cancel          =   -1  'True
      Caption         =   "&Não"
      Height          =   735
      Left            =   6960
      Picture         =   "frmSimNao.frx":04A8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnOk 
      Appearance      =   0  'Flat
      Caption         =   "Ok (Enter)"
      Height          =   735
      Left            =   7080
      Picture         =   "frmSimNao.frx":286F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2745
      ScaleWidth      =   8985
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.VScrollBar VScroll1 
         Height          =   2775
         LargeChange     =   10
         Left            =   8760
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Mensagem de ECF"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2790
         TabIndex        =   2
         Top             =   0
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frmSimNao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnNao_Click()
    Me.Tag = "1"
    Me.Hide
End Sub

Private Sub cmndbtnSim_Click()
    Me.Tag = "0"
    Me.Hide
End Sub

Private Sub VScroll1_Change()
    lblMsg.Top = 0 - VScroll1.Value
End Sub
