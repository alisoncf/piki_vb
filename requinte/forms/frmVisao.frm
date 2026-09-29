VERSION 5.00
Begin VB.Form frmVisao 
   Caption         =   "Visões"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUsuario 
      Caption         =   "..."
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtUsu 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "frmVisao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
