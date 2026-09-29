VERSION 5.00
Begin VB.Form frmVendaCartao 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cartão de Crédito"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbxParc 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "1"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "C&onfirmar (F7)"
      Height          =   735
      Left            =   2160
      Picture         =   "frmVendaCartao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Ca&ncelar (F12)"
      Height          =   735
      Left            =   3840
      Picture         =   "frmVendaCartao.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtbxAutorizacao 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmndbtnBuscarControl 
      Height          =   285
      Left            =   4320
      Picture         =   "frmVendaCartao.frx":0698
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtbxFinanceira 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   5490
      TabIndex        =   9
      Top             =   0
      Width           =   5520
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
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
         Left            =   840
         TabIndex        =   13
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblNumVen 
         AutoSize        =   -1  'True
         Caption         =   "0000000"
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
         Left            =   840
         TabIndex        =   11
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   510
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Autorização:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financeira:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   780
   End
End
Attribute VB_Name = "frmVendaCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscarControl_Click()
    Dim f As Financeira.Financeira
    f = F_BuscaFinanceira(txtbxFinanceira.Text)
    If f.Cod_Fin <> "" Then
        txtbxFinanceira.Text = f.nome_fin
        txtbxFinanceira.tag = f.Cod_Fin
        
        
    End If
    
    
    
End Sub

Private Sub cmndbtnCancel_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmndbtnOk_Click()
    If txtbxFinanceira.tag = "" Then
        MsgBox "Informe uma financeira!", vbExclamation
        Exit Sub
        
    End If
    If txtbxAutorizacao.Text = "" Then
        MsgBox "Informe o número do comprovante/autorização da transação!", vbExclamation
        Exit Sub
    End If
    Me.tag = txtbxFinanceira.tag
    Me.Hide
End Sub

Private Sub txtbxAutorizacao_GotFocus()
    P_SelecTexto txtbxAutorizacao
End Sub

Private Sub txtbxAutorizacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAutorizacao_LostFocus()
    P_AjustaTexto txtbxAutorizacao
End Sub

Private Sub txtbxParc_GotFocus()
    P_SelecTexto txtbxParc
End Sub

Private Sub txtbxParc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxParc_LostFocus()
    P_ValorLostFocus txtbxParc
End Sub
