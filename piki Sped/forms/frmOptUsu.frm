VERSION 5.00
Begin VB.Form frmOptUsu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções de Usuário"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3285
   Icon            =   "frmOptUsu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnFEchar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   165
      Picture         =   "frmOptUsu.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Relogar no Sistema"
      Height          =   975
      Left            =   165
      Picture         =   "frmOptUsu.frx":0898
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Mudar Visão"
      Height          =   975
      Left            =   165
      Picture         =   "frmOptUsu.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Mudar Senha"
      Height          =   975
      Left            =   165
      Picture         =   "frmOptUsu.frx":57DC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   3015
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
    'frmvis
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    frmLogin.Show 1
    frmSplash.Show 1
    'Load frmSplash
End Sub

Private Sub Form_Load()
    Me.Top = 1000
    Me.Left = 0
End Sub

