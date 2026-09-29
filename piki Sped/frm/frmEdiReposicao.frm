VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEdiReposicao 
   Caption         =   "Reposição de Estoque"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13695
   Icon            =   "frmEdiReposicao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3855
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6800
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   13695
      TabIndex        =   0
      Top             =   5835
      Width           =   13695
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   9360
         Picture         =   "frmEdiReposicao.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrintLista 
         Caption         =   "Imprimir Lista"
         Height          =   615
         Left            =   7800
         Picture         =   "frmEdiReposicao.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   6240
         Picture         =   "frmEdiReposicao.frx":0852
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   4680
         Picture         =   "frmEdiReposicao.frx":09AE
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Editar"
         Height          =   615
         Left            =   3120
         Picture         =   "frmEdiReposicao.frx":0B0A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdNova 
         Caption         =   "Nova"
         Height          =   615
         Left            =   1560
         Picture         =   "frmEdiReposicao.frx":150C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdPesq 
         Caption         =   "Pesquisar"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiReposicao.frx":1668
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtFun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10680
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         ItemData        =   "frmEdiReposicao.frx":1A13
         Left            =   120
         List            =   "frmEdiReposicao.frx":1A20
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscaProd 
         Height          =   285
         Left            =   6720
         Picture         =   "frmEdiReposicao.frx":1A50
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdFun 
         Height          =   285
         Left            =   12120
         Picture         =   "frmEdiReposicao.frx":1DFB
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdCat 
         Height          =   285
         Left            =   8520
         Picture         =   "frmEdiReposicao.frx":21A6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Height          =   285
         Left            =   10320
         Picture         =   "frmEdiReposicao.frx":2551
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         Height          =   195
         Left            =   8880
         TabIndex        =   18
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         Height          =   195
         Left            =   7080
         TabIndex        =   15
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário:"
         Height          =   195
         Left            =   10680
         TabIndex        =   11
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   3960
         TabIndex        =   8
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até:"
         Height          =   195
         Left            =   2760
         TabIndex        =   6
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmEdiReposicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ado As ADODB.Recordset
Private Sub cmdBuscaProd_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As Produtos.Produto
    p = Produtos.F_BuscaProdSaida(txtProd.Text)
    If p.cod_pro <> "" Then
        txtProd.Text = p.nome_pro
        txtProd.tag = p.cod_pro
        
    End If
    
End Sub

Private Sub cmdCat_Click()
    CarregaCategoria cmdCat
End Sub

Private Sub cmdFun_Click()
    Dim f As funcionario.funcionario
    
    carregaFunc txtFun
End Sub

Private Sub cmdMarca_Click()
    CarregaMarca txtMarca
    
    
End Sub

Private Sub cmdPesq_Click()
    busca
    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = reposicao.selectReposicao(Left(cmbStatus.Text, 1), txtDe.Text, txtAte.Text, _
    txtProd.tag, txtMarca.tag, txtCat.tag, txtFun.tag)
    abreObjetoAdo ado, strSQL, grid
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    'ajustaGrid
    Set ado = New ADODB.Recordset
    txtDe.Text = F_PrimDiaMes
    txtAte.Text = F_UltDiaMes
    
    resizeForm Me
End Sub
Sub ajustaGrid()
    With grid
        .FixedCols = 0
        .FixedRows = 1
        .rows = 1
        .Cols = 6
        .col = 0
        .Text = "id"
       ' .DataField = "id"
        .col = 1: .Text = "abertura"
        '.DataField "abertura"
        .col = 1: .Text = "abertura"
        '.DataField "abertura"
    End With
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
    posicionaADireita cmdSair, pcBot
End Sub

Private Sub txtDe_GotFocus()
    ganhaFoco txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub
Private Sub txtAte_GotFocus()
    ganhaFoco txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdBuscaProd_Click
    End If
    
End Sub
