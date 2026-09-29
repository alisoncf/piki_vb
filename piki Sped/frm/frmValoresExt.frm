VERSION 5.00
Begin VB.Form frmValoresExt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Produto Externo"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   915
      Width           =   1335
   End
   Begin VB.ComboBox cmbxLoja 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   945
      Width           =   3375
   End
   Begin VB.TextBox txtbxQtd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   4845
      TabIndex        =   5
      Text            =   "0,00"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtbxPreco 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "0,00"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar:"
      Height          =   375
      Left            =   6165
      TabIndex        =   6
      Top             =   915
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Loja e Origem:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label lblNome_Pro 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   1320
      TabIndex        =   11
      Top             =   360
      Width           =   1665
   End
   Begin VB.Label lblCod_Pro 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9999999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço &Unit:"
      Height          =   195
      Left            =   3720
      TabIndex        =   2
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Qtd:"
      Height          =   195
      Left            =   4920
      TabIndex        =   4
      Top             =   720
      Width           =   300
   End
End
Attribute VB_Name = "frmValoresExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxLoja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnConfirma_Click()
    If cmbxLoja.Text = "" Then
        MsgBox "Informe a loja!", vbExclamation
        'cmbxLoja.SetFocus
        Exit Sub
    End If
    If txtbxQtd.Text = "0,00" Then
        MsgBox "Informe a quantidade!", vbExclamation
        txtbxQtd.SetFocus
        Exit Sub
    End If
    Me.Tag = "X"
    Me.Hide
End Sub

Private Sub Command1_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub Form_Load()
    P_PreencheLojas cmbxLoja, False
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd
End Sub
