VERSION 5.00
Begin VB.Form frmContagemNova 
   Caption         =   "Contagem"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReservado 
      Caption         =   "Ajustar Estoque reservado"
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtReservado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   6060
      TabIndex        =   9
      Top             =   2790
      Width           =   6090
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   4320
         Picture         =   "frmContagemNova.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   615
         Left            =   2760
         Picture         =   "frmContagemNova.frx":016B
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.TextBox txtEst 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtProd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.TextBox txtQtd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdBusca 
      Caption         =   "..."
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reservado"
      Height          =   195
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque"
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label lblExpr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Produto"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Tag             =   "B"
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Qtd:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   300
   End
End
Attribute VB_Name = "frmContagemNova"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBusca_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As Produto
    
    p = Produtos.F_BuscaProdSaida(txtProd.Text)
    If p.cod_pro <> "" Then
        txtProd.Text = p.nome_pro
        txtProd.tag = p.cod_pro
        txtEst.Text = p.est.er_pro
        txtQtd.Text = p.est.er_pro
        txtQtd.SetFocus
    End If
    
    
End Sub

Private Sub cmndbtnSalvar_Click()
    If CCur(txtQtd.Text) < 0 Then
        MsgBox "Informa uma Quantidade correta", vbExclamation
        Exit Sub
    End If
    
    If txtProd.tag = "" Then
        MsgBox "Informe o produto!", vbExclamation
        txtProd.SetFocus
        Exit Sub
    End If
    If strAbrenota_usu <> "S" Then
        MsgBox "Para lançar contagem, você precisa ter a permissão de AbreNotaUsu", vbExclamation
        Exit Sub
    End If
    
    If GravaContagem(txtProd.tag, txtEst.Text, txtQtd.Text) Then
        strSQL = "update tab_produtos set er_pro=" & F_GravaDec(txtQtd.Text)
        strSQL = strSQL & " where cod_pro='" & txtProd.tag & "' "
        Conn.Execute strSQL
        
        MsgBox "Contagem Gravada", vbInformation
        txtProd.SetFocus
        
        
        
        
    End If
    
End Sub

Private Sub txtProd_GotFocus()
    P_SelecTexto txtProd
End Sub

Private Sub txtProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
    End If
    
End Sub

Private Sub txtProd_LostFocus()
    P_AjustaTexto txtProd
End Sub

Private Sub txtQtd_GotFocus()
    P_SelecTexto txtQtd
End Sub

Private Sub txtQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtQtd_LostFocus()
    P_ValorLostFocus txtQtd
End Sub
