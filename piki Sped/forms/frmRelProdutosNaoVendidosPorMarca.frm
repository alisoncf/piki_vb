VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelProdutosNaoVendidosPorMarca 
   Caption         =   "Relatório de Produtos Não Vendidos"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   12435
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade 
      Height          =   2175
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   3836
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   12405
      TabIndex        =   15
      Top             =   6735
      Width           =   12435
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   1680
         Picture         =   "frmRelProdutosNaoVendidosPorMarca.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   720
         Left            =   120
         Picture         =   "frmRelProdutosNaoVendidosPorMarca.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   10560
         Picture         =   "frmRelProdutosNaoVendidosPorMarca.frx":0579
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   12435
      TabIndex        =   0
      Top             =   0
      Width           =   12435
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   495
         Left            =   10800
         Picture         =   "frmRelProdutosNaoVendidosPorMarca.frx":0AC7
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtForn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdProd 
         Height          =   285
         Left            =   5400
         Picture         =   "frmRelProdutosNaoVendidosPorMarca.frx":0E72
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdCat 
         Height          =   285
         Left            =   7920
         Picture         =   "frmRelProdutosNaoVendidosPorMarca.frx":121D
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Height          =   285
         Left            =   10440
         Picture         =   "frmRelProdutosNaoVendidosPorMarca.frx":15C8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         Height          =   195
         Left            =   8280
         TabIndex        =   14
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         Height          =   195
         Left            =   5760
         TabIndex        =   13
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   2520
         TabIndex        =   12
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   210
      End
   End
End
Attribute VB_Name = "frmRelProdutosNaoVendidosPorMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdCat_Click()
    Dim c As categoria.categoria
    
    If c.codigo <> "" Then
        txtCat.Text = c.nome
        txtCat.tag = c.codigo
        busca
    End If
    
End Sub

Private Sub cmdMarca_Click()
    buscaMarca
End Sub
Sub buscaMarca()
    Dim m As Marca.Marca
    If m.codigo <> "" Then
        txtMarca.Text = m.nome
        txtMarca.tag = m.codigo
        busca
    End If
    
    
    
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    
    If txtDe.Text = "" Then
        MsgBox "Informe uma data inicial", vbExclamation
        txtDe.SetFocus
        Exit Sub
    End If
    If txtAte.Text = "" Then
        MsgBox "Informe uma data final", vbExclamation
        txtAte.SetFocus
        Exit Sub
    End If
    
    strSQL = "select p.cod_pro,p.codfab_pro,p.nome_pro, m.nome_mar, c.nome_cat, tp.er_pro "
    strSQL = strSQL & " from " & TAB_PRODUTOS & " as p "
    strSQL = strSQL & " inner join " & TAB_MARCAS & " as m on m.cod_mar=p.cod_mar "
    strSQL = strSQL & " inner join " & TAB_CATEGORIAS & " as c on c.cod_CAT=p.cod_CAT "
    strSQL = strSQL & " where (select sum(qtd_ite) from " & TAB_VENDAS & " )"
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and left(dhemi,10)>=" & F_GravaData(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and left(dhemi,10)<=" & F_GravaData(txtAte.Text)
    End If
    
    
    abreObjetoAdo ado, strSQL
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    txtDe.Text = F_PrimDiaMes
    txtAte.Text = Date
    
    
    
End Sub

Private Sub Form_Resize()
    PosicionaGrade grade, Me, pcBot
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

