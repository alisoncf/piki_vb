VERSION 5.00
Begin VB.Form frmTipo_Imp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo do Imposto"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   2790
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Default         =   -1  'True
      Height          =   615
      Left            =   120
      Picture         =   "frmTipo_Imp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Frame frImp 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2535
      Begin VB.TextBox txtbxData_imp 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxValor_Imp 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   12
         TabIndex        =   5
         Text            =   "0,0000"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Venc. Padrão:"
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "&Valor: (%)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo do Imposto:"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   100
      Width           =   2535
      Begin VB.OptionButton optnTipo 
         Caption         =   "&Especial"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optnTipo 
         Caption         =   "&Normal"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmTipo_Imp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSalvar_Click()
    If optnTipo(0).Value = True Then
        Me.Tag = ""
        Me.Hide
    Else
        If CCur(txtbxValor_Imp.Text) <= 0 Then
            MsgBox "Informe um valor válido para o Imposto!", vbExclamation
            txtbxValor_Imp.SetFocus
            Exit Sub
        End If
        Me.Tag = txtbxValor_Imp.Text
        Me.Hide
    End If
End Sub

Private Sub OptnTipo_Click(Index As Integer)
    If Index = 0 Then
        frImp.Enabled = False
    Else
        frImp.Enabled = True
    End If
End Sub

Private Sub txtbxData_imp_GotFocus()
    P_SelecTexto txtbxData_imp
End Sub

Private Sub txtbxData_imp_LostFocus()
    txtbxData_imp.Text = Format(txtbxData_imp.Text, "dd/mm")
End Sub

Private Sub txtbxValor_Imp_gotFocus()
    P_SelecTexto txtbxValor_Imp
End Sub

Private Sub txtbxValor_Imp_LostFocus()
    P_ValorLostFocus txtbxValor_Imp, "###,##0.0000"
End Sub
