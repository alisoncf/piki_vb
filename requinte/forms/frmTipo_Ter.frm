VERSION 5.00
Begin VB.Form frmTipo_Ter 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indique o Tipo do Cliente"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   2265
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optnTipo 
      BackColor       =   &H80000001&
      Caption         =   "&Ambos"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.OptionButton optnTipo 
      BackColor       =   &H80000001&
      Caption         =   "&Fornecedor"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.OptionButton optnTipo 
      BackColor       =   &H80000001&
      Caption         =   "&Cliente"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnConfirm 
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      Picture         =   "frmTipo_Ter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo do Terceiros"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   510
      TabIndex        =   0
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "frmTipo_Ter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmndbtnConfirm_Click()
    If Me.Tag = "" Then
        MsgBox "Clique numa das opções para escolha do tipo do terceiro", vbCritical
    Else
        Me.Hide
    End If
End Sub

Private Sub OptnTipo_Click(Index As Integer)
    If Index = 0 Then
        Me.Tag = "C"
    ElseIf Index = 1 Then
        Me.Tag = "F"
    ElseIf Index = 1 Then
        Me.Tag = "A"
    End If
    
End Sub
