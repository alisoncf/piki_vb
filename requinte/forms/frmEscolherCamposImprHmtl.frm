VERSION 5.00
Begin VB.Form frmEscolherCamposImprHmtl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Escolha os Campos que Deseja Imprimir"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnTodos 
      Caption         =   "Todos"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.ListBox lst 
      Height          =   5685
      ItemData        =   "frmEscolherCamposImprHmtl.frx":0000
      Left            =   120
      List            =   "frmEscolherCamposImprHmtl.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmEscolherCamposImprHmtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnOk_Click()
    
    
    
    Me.Hide
End Sub

Private Sub cmndbtnTodos_Click()
    mudaTodos
End Sub
Sub mudaTodos()
    Dim valor As Boolean
    If cmndbtnTodos.caption = "Todos" Then
        cmndbtnTodos.caption = "Nenhum"
        valor = True
    Else
        cmndbtnTodos.caption = "Todos"
        valor = False
    End If
    For iCount = 0 To lst.ListCount - 1
        
        lst.Selected(iCount) = valor
    Next
End Sub

