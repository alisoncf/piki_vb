VERSION 5.00
Begin VB.Form frmOptUsu 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções de Usuário"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2280
   Icon            =   "frmoptusu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   2280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnFEchar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   165
      Picture         =   "frmoptusu.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Relogar no Sistema"
      Height          =   735
      Left            =   165
      Picture         =   "frmoptusu.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aparência"
      Height          =   735
      Left            =   165
      Picture         =   "frmoptusu.frx":2BF6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Mudar Senha"
      Height          =   735
      Left            =   165
      Picture         =   "frmoptusu.frx":5398
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmOptUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    frmMudaSenha.Show 1
End Sub

Private Sub Command2_Click()
    frmConfig.Show 0
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    frmLogin.tag = "x"
    frmLogin.Show 1
    frmSplash.Show 1
    'Load frmSplash
    Unload frmLoginal
End Sub

Private Sub Form_Load()
    Me.Top = 1000
    Me.Left = 0
End Sub

