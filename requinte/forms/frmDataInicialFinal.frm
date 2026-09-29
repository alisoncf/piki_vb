VERSION 5.00
Begin VB.Form frmDataInicialFinal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datas:"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "frmDataInicialFinal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   615
      Left            =   360
      Picture         =   "frmDataInicialFinal.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   1680
      Picture         =   "frmDataInicialFinal.frx":07CC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtbxDataFinal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtbxDataInicial 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data &Final:"
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   765
   End
   Begin VB.Label lblDataIni 
      AutoSize        =   -1  'True
      Caption         =   "Data &Inicial:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmDataInicialFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub cmndbtnSair_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxDataFinal.Text = "" Then
        MsgBox "Informe a Data Final!", vbExclamation
        txtbxDataFinal.SetFocus
        Exit Sub
    End If
    If txtbxDataInicial.Text = "" Then
        MsgBox "Informe a Data Inicial!", vbExclamation
        txtbxDataInicial.SetFocus
        Exit Sub
    End If
    
    If F_DataMaior(txtbxDataInicial.Text, txtbxDataFinal.Text) Then
        MsgBox "A Data Inicial não pode ser maior que a Final!", vbExclamation
        txtbxDataFinal.SetFocus
        Exit Sub
    End If
    Me.tag = Format(txtbxDataInicial.Text, "dd/mm/yyyy") & Format(txtbxDataFinal.Text, "dd/mm/yyyy")
    Me.Hide
End Sub

Private Sub txtbxDataFinal_GotFocus()
   P_SelecTexto txtbxDataFinal
End Sub

Private Sub txtbxDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataFinal_LostFocus()
    P_AjustaData txtbxDataFinal
End Sub

Private Sub txtbxDataInicial_GotFocus()
    P_SelecTexto txtbxDataInicial
End Sub

Private Sub txtbxDataInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataInicial_LostFocus()
    P_AjustaData txtbxDataInicial
End Sub
