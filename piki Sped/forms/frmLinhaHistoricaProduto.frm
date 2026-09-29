VERSION 5.00
Begin VB.Form frmLinhaHistoricaProduto 
   Caption         =   "Linha Histórica do Produto"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   12270
      TabIndex        =   0
      Top             =   0
      Width           =   12300
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5535
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7320
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdPesquisar 
         Height          =   285
         Left            =   5640
         Picture         =   "frmLinhaHistoricaProduto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de:"
         Height          =   195
         Left            =   6240
         TabIndex        =   6
         Top             =   120
         Width           =   225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até:"
         Height          =   195
         Left            =   7320
         TabIndex        =   5
         Top             =   120
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmLinhaHistoricaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPesquisar_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As Produto
    
    p = F_BuscaProdSaida(txtProduto.Text)
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
    End If
    
End Sub

Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub
Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtProduto_Change()
    txtProduto.tag = ""
End Sub

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
    
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdPesquisar_Click
    End If
End Sub

