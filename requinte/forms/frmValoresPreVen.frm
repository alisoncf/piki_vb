VERSION 5.00
Begin VB.Form frmValoresPReVEn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Produto de Pré-Venda"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar:"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   795
      Width           =   1335
   End
   Begin VB.TextBox txtbxPreco 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "0,00"
      Top             =   840
      Width           =   975
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
      Left            =   1365
      TabIndex        =   1
      Text            =   "0,00"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   795
      Width           =   1335
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Qtd:"
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço &Unit:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   1440
      TabIndex        =   8
      Top             =   0
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   330
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
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   840
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
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   1665
   End
End
Attribute VB_Name = "frmValoresPReVEn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnConfirma_Click()
    Me.Tag = "X"
    Me.Hide
End Sub

Private Sub Command1_Click()
    Me.Tag = ""
    Me.Hide
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
