VERSION 5.00
Begin VB.Form frmCred 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Créditos de Projeto e Desenvolvimento"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6030
   ForeColor       =   &H80000013&
   Icon            =   "frmCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   1200
      Picture         =   "frmCred.frx":0442
      ScaleHeight     =   1560
      ScaleWidth      =   3900
      TabIndex        =   2
      Top             =   600
      Width           =   3900
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Retornar"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versão 2017 - 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alison Filgueiras - alisoncf@gmail.com"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   2670
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Projetos de SI 2003-2017"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2850
   End
End
Attribute VB_Name = "frmCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

