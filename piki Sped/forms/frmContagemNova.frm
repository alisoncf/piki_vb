VERSION 5.00
Begin VB.Form frmContagemNova 
   Caption         =   "Contagem"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReservado 
      Caption         =   "Ajusta reservado"
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtReservado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   11
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
      ScaleWidth      =   7950
      TabIndex        =   9
      Top             =   2775
      Width           =   7980
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   6360
         Picture         =   "frmContagemNova.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   615
         Left            =   4680
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
      Width           =   6855
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
      Left            =   7080
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblMarca 
      AutoSize        =   -1  'True
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
      Left            =   2880
      TabIndex        =   17
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Tag             =   "B"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCodFab 
      AutoSize        =   -1  'True
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
      Left            =   960
      TabIndex        =   15
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód:Fab."
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Tag             =   "B"
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reservado"
      Height          =   195
      Left            =   1920
      TabIndex        =   12
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label lblEmpresa 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   615
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
    
    p = Produtos.F_BuscaProdSaida(txtProd.text)
    If p.cod_pro <> "" Then
        txtProd.text = p.nome_pro
        txtProd.tag = p.cod_pro
        txtEst.text = p.er_pro
        txtQtd.text = p.er_pro
        lblEmpresa.tag = p.cod_emp
        txtReservado.text = p.r_pro
        lblEmpresa.caption = p.cod_emp & "-" & F_BuscaUmCampo("nome", "sped_empresa", "codigo='" & p.cod_emp & "'")
        lblCodFab.caption = p.codfab_pro
        lblMarca.caption = p.nome_mar
        txtQtd.SetFocus
    End If
    
    
End Sub

Private Sub cmdReservado_Click()
    P_AjustaEstoqueReservado txtProd.tag
    buscaProd
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If CCur(txtQtd.text) < 0 Then
        MsgBox "Informa uma Quantidade correta", vbExclamation
        Exit Sub
    End If
    
    If txtProd.tag = "" Then
        MsgBox "Informe o produto!", vbExclamation
        txtProd.SetFocus
        Exit Sub
    End If
    If strAltpreco_usu <> "S" Then
        MsgBox "Para lançar contagem, você precisa ter a permissão de Alterar Preço", vbExclamation
        Exit Sub
    End If
    
    If GravaContagem(txtProd.tag, txtEst.text, txtQtd.text, lblEmpresa.tag) Then
        strSQL = "update estoque set er=" & F_GravaDec(txtQtd.text)
        strSQL = strSQL & " where cod_pro='" & txtProd.tag & "' "
        strSQL = strSQL & " and cod_emp='" & lblEmpresa.tag & "'"
        conn.Execute strSQL
        
        strSQL = "update " & TAB_PRODUTOS & "  set er_pro= (select sum(er) from estoque"
        strSQL = strSQL & " where cod_pro='" & txtProd.tag & "') where cod_pro='" & txtProd.tag & "' "
        conn.Execute strSQL
        

        
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
