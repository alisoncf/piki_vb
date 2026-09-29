VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   6000
      TabIndex        =   3
      Top             =   0
      Width           =   6000
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Piki Box"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   1560
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   1920
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   1005
      ScaleWidth      =   2010
      TabIndex        =   2
      Top             =   1320
      Width           =   2010
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versão Alfa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   600
      X2              =   5445
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblAviso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscando Informações..."
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2160
      TabIndex        =   0
      Top             =   2640
      Width           =   1770
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iVezes As Integer

Private Sub Form_Activate()
    'busca
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
    
End Sub


