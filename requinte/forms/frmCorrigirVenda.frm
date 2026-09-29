VERSION 5.00
Begin VB.Form frmCorrigirVenda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Corrigir Venda"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxNumero 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venda:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "frmCorrigirVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtbxNumero_GotFocus()
    P_SelecTexto txtbxNumero
End Sub
