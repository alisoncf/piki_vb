VERSION 5.00
Begin VB.Form frmStatusBar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   495
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   FillColor       =   &H00FF0000&
   ForeColor       =   &H8000000C&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmStatusBar.frx":0000
   ScaleHeight     =   495
   ScaleWidth      =   11340
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   10200
      Y1              =   470
      Y2              =   470
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000A&
      X1              =   0
      X2              =   10200
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   10200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   10200
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   9000
      X2              =   9000
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   6600
      X2              =   6600
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   1920
      X2              =   1920
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   15
      Width           =   225
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   225
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   7
      Top             =   15
      Width           =   225
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   240
      Width           =   225
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   4440
      TabIndex        =   5
      Top             =   15
      Width           =   225
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   4440
      TabIndex        =   4
      Top             =   240
      Width           =   225
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   6720
      TabIndex        =   3
      Top             =   15
      Width           =   225
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   225
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   9120
      TabIndex        =   1
      Top             =   15
      Width           =   225
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   9120
      TabIndex        =   0
      Top             =   240
      Width           =   225
   End
End
Attribute VB_Name = "frmStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    Line5.X2 = Me.Width
    Line6.X2 = Me.Width
    Line7.X2 = Me.Width
    Line8.X2 = Me.Width
End Sub
