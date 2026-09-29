VERSION 5.00
Begin VB.Form frmEscolheBanco 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Escolha Um Banco"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox lst 
      Height          =   1425
      ItemData        =   "frmEscolheBanco.frx":0000
      Left            =   240
      List            =   "frmEscolheBanco.frx":000A
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmEscolheBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        strNomeBanco = lst.Text
        If F_AbreBanco() Then
           MsgBox banco.Connect, vbInformation
           Unload Me
           
        End If
    End If
End Sub
