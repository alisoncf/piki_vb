VERSION 5.00
Begin VB.Form frmPerguntaUsuario 
   BorderStyle     =   0  'None
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   ScaleHeight     =   2010
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblMensagem 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Informa aqui sua mensagem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3600
   End
End
Attribute VB_Name = "frmPerguntaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
