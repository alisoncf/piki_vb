VERSION 5.00
Begin VB.Form frmPadrao 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcTool 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6930
      TabIndex        =   2
      Top             =   4095
      Width           =   6930
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   615
         Left            =   5520
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Gravar"
         Height          =   615
         Left            =   4080
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.TextBox txtCampo1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblCampo1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campo1:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalvar_Click()
    
End Sub

Private Sub txtCampo1_Change()

End Sub

Private Sub txtCampo1_GotFocus()
    ganhaFoco txtCampo1
End Sub

Private Sub txtCampo1_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtCampo1_LostFocus()
    AjustaTexto txtCampo1
End Sub
