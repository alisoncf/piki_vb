VERSION 5.00
Begin VB.Form frmCPFVenda 
   Caption         =   "CPF Venda"
   ClientHeight    =   1440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5610
   Icon            =   "frmCPFVenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCPF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CPF:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   660
   End
End
Attribute VB_Name = "frmCPFVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCPF As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtCPF.Text = ""
        Me.Hide
        
    End If
End Sub

Private Sub txtCPF_GotFocus()
    ganhaFoco txtCPF
End Sub

Private Sub txtCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtCPF_LostFocus
        If bCPF Then
            Me.Hide
        End If
    End If
End Sub

Private Sub txtCPF_LostFocus()
    If txtCPF = "" Then
        bCPF = True
        Exit Sub
    End If
    If Not F_CPF(txtCPF.Text) Then
        MsgBox "CPF Inválido!", vbExclamation
        txtCPF.SetFocus
    Else
        bCPF = True
    End If
End Sub
