VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMediaConsumo_Mer 
   Caption         =   "Calcular Média de Movimentação Diária de Produtos"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12870
   Icon            =   "frmMediaConsumo_Mer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   12870
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   12870
      TabIndex        =   21
      Top             =   4050
      Width           =   12870
      Begin VB.CommandButton cmndbtnCalcula 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   9120
         Picture         =   "frmMediaConsumo_Mer.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   735
         Left            =   10920
         Picture         =   "frmMediaConsumo_Mer.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   8520
      ScaleHeight     =   2655
      ScaleWidth      =   4095
      TabIndex        =   16
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox txtFor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   2895
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridForn 
         Height          =   1695
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2990
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdFor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         Picture         =   "frmMediaConsumo_Mer.frx":07D6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fornecedor"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   4320
      ScaleHeight     =   2655
      ScaleWidth      =   4095
      TabIndex        =   11
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2895
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridCat 
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2990
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         Picture         =   "frmMediaConsumo_Mer.frx":0B81
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   675
      End
   End
   Begin VB.PictureBox pcMarcas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCB683&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2655
      ScaleWidth      =   4095
      TabIndex        =   6
      Top             =   1200
      Width           =   4095
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMarca 
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2990
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         Picture         =   "frmMediaConsumo_Mer.frx":0F2C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12495
      Begin VB.CheckBox ChkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Zerar MMD de todos os produtos"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMediaConsumo_Mer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxGru_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxGru.ListIndex = -1
    End If
End Sub

Private Sub cmbxCat_Click()
    cmbxCat.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxForn_Click()
    cmbxForn.tag = F_BuscaUmCampo("cod_for", TAB_FORNECEDORES, "nome_for='" & cmbxForn.Text & "'")
End Sub

Private Sub cmbxForn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxForn.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxMar_Click()
    cmbxMar.tag = F_BuscaUmCampo("cod_mar", TAB_MARCAS, "nome_mar='" & cmbxMar.Text & "'")
End Sub

Private Sub cmbxMar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxMar.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxSub.ListIndex = -1
    End If
End Sub

Private Sub cmdMarca_Click()
    buscaMarca
End Sub
Sub buscaMarca()
    Dim c As Marca.Marca
    c = Marca.buscaMarca(txtMarca.Text)
    If c.codigo <> "" Then
        txtMarca.Text = c.nome
        txtMarca.tag = c.codigo
        addMarca
    End If
End Sub
Private Sub cmdCat_Click()
    buscaCat
End Sub
Sub buscaCat()
    Dim c As categoria.categoria
    c = categoria.BuscaCategoria(txtCat.Text)
    If c.codigo <> "" Then
        txtCat.Text = c.nome
        txtCat.tag = c.codigo
        addCat
    End If
End Sub
Private Sub cmdFor_Click()
    buscaFor
End Sub
Sub buscaFor()
    Dim c As fornecedor.fornecedor
    c = fornecedor.F_BuscaFornecedor(txtFor.Text)
    If c.cod_For <> "" Then
        txtFor.Text = c.nome_For
        txtFor.tag = c.cod_For
        addFor
    End If
End Sub
Sub ajustaGrid()
    Dim obj As Object
 
    AjustaGridData gridMarca, 2, 1, 0, "Código,Nome", "", "800,3000", Nothing, "", False
    AjustaGridData gridCat, 2, 1, 0, "Código,Nome", "", "800,3000", Nothing, "", False
    AjustaGridData gridForn, 2, 1, 0, "Código,Nome", "", "800,3000", Nothing, "", False
End Sub
Sub addMarca()
    If F_InFlxgrd(gridMarca, txtMarca.tag, 0) = -1 Then
        gridMarca.AddItem txtMarca.tag & Chr(9) & txtMarca.Text
    End If
    
End Sub
Sub addCat()
    If F_InFlxgrd(gridCat, txtCat.tag, 0) = -1 Then
        gridCat.AddItem txtCat.tag & Chr(9) & txtCat.Text
    End If
    
End Sub

Sub addFor()
    If F_InFlxgrd(gridForn, txtFor.tag, 0) = -1 Then
        gridForn.AddItem txtFor.tag & Chr(9) & txtFor.Text
    End If
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnCalcula_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCat As New Collection
    Dim sFor As New Collection
    Dim sMar As New Collection
    Dim sPro As New Collection
    
    
    
    If F_Vazio(txtbxDataFin.Text) Then
        MsgBox "Informe a data final!", vbExclamation
        txtbxDataFin.SetFocus
        Exit Sub
    ElseIf F_Vazio(txtbxDataIni.Text) Then
        MsgBox "Informe a data Inicial!", vbExclamation
        txtbxDataIni.SetFocus
        Exit Sub
    End If
    
    If DateDiff("D", txtbxDataIni.Text, txtbxDataFin.Text) < 14 Then
        MsgBox "A Diferença entre as Datas deve ser maior ou superior a 15", vbExclamation
        txtbxDataFin.SetFocus
        Exit Sub
    End If
    
    Label3.Visible = True
    'Image1.Visible = True
    Set sCat = GridUtil.ColunaGrid(gridCat, 0)
    Set sMar = GridUtil.ColunaGrid(gridMarca, 0)
    Set sFor = GridUtil.ColunaGrid(gridForn, 0)
    
    P_MediaDiaria txtbxDataIni.Text, txtbxDataFin.Text, sMar, sCat, sPro, sFor, IIf(ChkTodos.Value = 1, True, False)
    
    Unload Me
    
End Sub



Private Sub Form_Load()
    txtbxDataIni.Text = F_PrimDiaMes
    txtbxDataFin.Text = Date
    ajustaGrid
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxDataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDataFin
    
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtCat_Change()
    txtCat.tag = ""
End Sub

Private Sub txtCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCat
    End If
End Sub

Private Sub txtFor_Change()
    txtFor.tag = ""
End Sub

Private Sub txtFor_GotFocus()
    P_SelecTexto txtFor
End Sub

Private Sub txtFor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFor
    End If
End Sub

Private Sub txtMarca_Change()
    txtMarca.tag = ""
End Sub

Private Sub txtMarca_GotFocus()
    P_SelecTexto txtMarca
End Sub

Private Sub txtMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaMarca
    End If
    
    
    
End Sub
