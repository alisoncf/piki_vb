VERSION 5.00
Begin VB.Form frmVendaEstZer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produto com Estoque Zerado:"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "&Não Lançar"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnPrev 
      Caption         =   "Lançar &Pré-Venda"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnExt 
      Caption         =   "&Lançar Produto Externo"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Está com o estoque zerado,  o que deseja fazer?"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label lblNome_pro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "frmVendaEstZer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnExt_Click()
    Me.Tag = "E"
    Me.Hide
End Sub

Private Sub cmndbtnPrev_Click()
    Me.Tag = "P"
    Me.Hide
End Sub

Private Sub cmndbtnSair_Click()
    Me.Tag = ""
    Me.Hide
End Sub
