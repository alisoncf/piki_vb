VERSION 5.00
Begin VB.Form frmLogo 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "X-Manager"
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2100
   DrawStyle       =   5  'Transparent
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2115
   ScaleWidth      =   2100
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   555
      Left            =   720
      Picture         =   "frmLogo.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X-pert Softwares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   1320
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   ' Me.Top = 0 'mdiFrm.Height - 2000
   ' Me.Left = 0 '
   'MudaAparencia Me
   P_CentraForm Me, 0
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub
