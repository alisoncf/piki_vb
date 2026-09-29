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
      Left            =   120
      Picture         =   "frmCred.frx":0442
      ScaleHeight     =   1560
      ScaleWidth      =   3900
      TabIndex        =   2
      Top             =   840
      Width           =   3900
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soluções Inteligentes para Empresas Inteligentes"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "www.xpertsoftwares.com.br"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   2190
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anápolis-GO - Fone: (0..62) 3098-3005"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2730
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Retornar"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Analista Responsável: Alison Filgueiras - E.ls. S.I"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   3420
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Projetado e Desenvolvido por:"
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
      Width           =   3405
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

