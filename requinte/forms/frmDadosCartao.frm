VERSION 5.00
Begin VB.Form frmDadosCartao 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cartão de Crédito"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6645
      TabIndex        =   11
      Top             =   0
      Width           =   6675
      Begin VB.Label lblNumVen 
         AutoSize        =   -1  'True
         Caption         =   "Informe o Cartão e os Valores"
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
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.TextBox txtbxValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   8
      Text            =   "0,00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtbxParc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   6
      Text            =   "1"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "C&onfirmar (F7)"
      Height          =   735
      Left            =   2280
      Picture         =   "frmDadosCartao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Ca&ncelar (F12)"
      Height          =   735
      Left            =   3960
      Picture         =   "frmDadosCartao.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtbxAutorizacao 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtbxFinanceira 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4695
   End
   Begin VB.CommandButton cmndbtnBuscarControl 
      Height          =   360
      Left            =   4920
      Picture         =   "frmDadosCartao.frx":0698
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblCNPJ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ"
      Height          =   195
      Left            =   1080
      TabIndex        =   16
      Top             =   840
      Width           =   405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      Height          =   195
      Left            =   3360
      TabIndex        =   14
      Top             =   2400
      Width           =   420
   End
   Begin VB.Label lblOperacao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   195
      Left            =   4200
      TabIndex        =   13
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Autorização:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financeira:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   780
   End
End
Attribute VB_Name = "frmDadosCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscarControl_Click()
    Dim f As financeira.financeira
    f = F_BuscaFinanceira(txtbxFinanceira.Text)
    If f.cod_fin <> "" Then
        txtbxFinanceira.Text = f.nome_fin
        txtbxFinanceira.tag = f.cod_fin
        lblOperacao.caption = f.operacao_fin
        lblTipo.caption = f.tipo_fin
        lblCNPJ.caption = f.CNPJ
        txtbxAutorizacao.SetFocus
    End If
    
    
    
End Sub

Private Sub cmndbtnCancel_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmndbtnOk_Click()
    
    Me.tag = txtbxFinanceira.tag
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnOk_Click
    End If
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

Private Sub txtbxFinanceira_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscarControl_Click
    End If
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
    P_ValorLostFocus txtbxParc, "#"
End Sub

Private Sub txtbxValor_GotFocus()
    P_SelecTexto txtbxValor
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    P_ValorLostFocus txtbxValor
End Sub
