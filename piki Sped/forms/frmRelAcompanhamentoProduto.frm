VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelAcompanhamentoProduto 
   Caption         =   "Relatório de Acompanhamento de Clientes por Período:"
   ClientHeight    =   8715
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14655
   Icon            =   "frmRelAcompanhamentoProduto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   14655
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   6735
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   11880
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   14625
      TabIndex        =   5
      Top             =   7740
      Width           =   14655
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   1680
         Picture         =   "frmRelAcompanhamentoProduto.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   720
         Left            =   120
         Picture         =   "frmRelAcompanhamentoProduto.frx":10994
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   12600
         Picture         =   "frmRelAcompanhamentoProduto.frx":10DC3
         Style           =   1  'Graphical
         TabIndex        =   6
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
      ScaleWidth      =   14655
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Agrupar por:"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5760
         TabIndex        =   12
         Top             =   0
         Width           =   4455
         Begin VB.OptionButton optnProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Produto"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2280
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optnMarca 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Marca"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optCat 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Categoria"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtAno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   495
         Left            =   13320
         Picture         =   "frmRelAcompanhamentoProduto.frx":11311
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCli 
         Height          =   285
         Left            =   5280
         Picture         =   "frmRelAcompanhamentoProduto.frx":116BC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnDetalhar 
         Caption         =   "Detalhar"
      End
      Begin VB.Menu mnSair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmRelAcompanhamentoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub busca()
    Dim strSQL As String
    Dim ado As New ADODB.Recordset
    Dim cod_cli As String
    Dim nome As String
    Dim cod As String
    If optnMarca.Value Then
        nome = "tm.nome_mar"
        cod = "tp.cod_mar"
    End If
    If optCat.Value Then
        nome = "nome_cat"
        cod = "tp.cod_cat"
    End If
    If optnProduto.Value Then
        nome = "nome_pro"
        cod = "tp.cod_pro"
    End If
    
    
  '  If txtCliente.tag = "" Then
 '       MsgBox "É necessário informar o cliente!", vbExclamation
'        txtCliente.SetFocus
    '    Exit Sub
   ' End If
    
    
    
    P_Status "Aguarde..."
    
    
    
    
    grid.Visible = False
    
    
    
    strSQL = "select mid(tv.dataefet_ven,6,2) as mes, " & nome & " as nome," & cod & " as codigo, "
    strSQL = strSQL & "   sum((final_ite/qtd_ite)*(qtd_ite-devolvido_ite)) As total "
    strSQL = strSQL & " from " & TAB_CLIENTES & " as cli inner join "
    strSQL = strSQL & " vendas As tv on cli.cod_cli=tv.cod_cli    inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & "  as tp on ti.cod_pro=tp.cod_pro"
    strSQL = strSQL & " inner join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar"
    strSQL = strSQL & " inner join " & TAB_CATEGORIAS & " as tc on tp.cod_Cat=tc.cod_cat "
    strSQL = strSQL & " "
    
    
    strSQL = strSQL & " Where 1>0 "
    
    
    
    
    strSQL = strSQL & " and tv.dataefet_ven>='" & txtAno.Text & "-01-01'"
    strSQL = strSQL & " and tv.dataefet_ven<='" & txtAno.Text & "-12-31'"
    
    strSQL = strSQL & " and situacao_ven=4"
    
    
    
    
    If txtCliente.tag <> "" Then
        strSQL = strSQL & " and tv.cod_cli='" & txtCliente.tag & "' "
    End If
  
    strSQL = strSQL & " group by " & cod & ",  mes"
  
    strSQL = strSQL & " order by " & cod & ",  mes"
    
    abreObjetoAdo ado, strSQL, grid


    AjustaGrid

    With ado
        Do While Not .EOF
            If cod_cli <> CStr(!codigo) Then
                grid.AddItem !codigo & Chr(9) & !nome
                cod_cli = !codigo
            End If
            grid.row = grid.rows - 1
            
            grid.col = !mes + 1
            grid.Text = FormatVal(!total)
            .MoveNext
        Loop
    End With

soma
ZebraGrid grid, 1
grid.Visible = True
P_Status
End Sub
Private Sub cmdCat_Click()
    Dim c As categoria.categoria
    c = BuscaCategoria(txtCat.Text)
    If c.codigo <> "" Then
        txtCat.Text = c.nome
        txtCat.tag = c.codigo
        busca
    End If
End Sub

Private Sub cmdExportar_Click()
    ExportaMSFLXGridToCSV grid
   ' P_MsflxToHtml
    
End Sub

Private Sub cmdMarca_Click()
    Dim m As Marca.Marca
    m = buscaMarca(txtMarca.Text)
    If m.codigo <> "" Then
        txtMarca.Text = m.nome
        txtMarca.tag = m.codigo
        busca
    End If

End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub
Sub soma()
    Dim cTotal As Currency
    Dim i As Integer
    Dim j As Integer
    Dim cGeral As Currency
    Dim totalCol(15) As Currency
    With grid
        For i = 1 To .rows - 1
            cTotal = 0
            For j = 2 To .Cols - 2
                .col = j
                .row = i
                If .Text = "" Then .Text = "0"
                cTotal = cTotal + CCur(.Text)
                totalCol(j) = totalCol(j) + CCur(.Text)
            Next
            .col = .Cols - 1
            .Text = FormatVal(cTotal)
            cGeral = cGeral + cTotal
        Next
        grid.AddItem "total"
        .row = .rows - 1
        For i = 2 To .Cols - 1
            .col = i
            .Text = FormatVal(totalCol(i))
        Next
        
        
        For i = 1 To .rows - 1
            .col = .Cols - 1
            .row = i
            
        Next
        
        
        
        .col = .Cols - 1
        .row = .rows - 1
        .Text = FormatVal(cGeral)
        
    End With
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmdCli_Click()
    Dim c As cliente
    c = Clientes.F_BuscaCliente(txtCliente.Text)
    If c.cod_cli <> "" Then
        txtCliente.Text = c.nome_Cli
        txtCliente.tag = c.cod_cli
        busca
    End If

End Sub

Private Sub Form_Load()
    AjustaGrid

    
End Sub
Sub AjustaGrid()
    Dim i As Integer
    With grid
        .FixedCols = 0
      '  .FixedRows = 1
        .rows = 1
        .Cols = 15
        .row = 0
        .col = 0
        .Text = "Código"
        .col = 1
        .Text = "Descrição"
        .ColWidth(1) = 2500
        
        
        
        For i = 2 To 13
            .col = i
            .Text = MonthName(i - 1)
            
            .ColWidth(i) = 1200
        Next
        .col = 14
        
        .Text = "Total"


    End With
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
End Sub

Private Sub grid_DblClick()
    Dim mes As String
    Dim i As Integer
    Dim dataI As String
    Dim dataF As String
    i = grid.col
    
    
    
    
    If i > 1 And i < 13 Then
        i = i - 1
        dataI = "01/" & Format(i, "0#") & "/" & txtAno.Text
        dataF = F_UltDiaMes("01/" & Format(i, "0#") & "/" & txtAno.Text)
    End If
    grid.col = 0
    
    P_histCli grid.Text, dataI, dataF
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnMenu
    End If
End Sub

Private Sub mnDetalhar_Click()
    grid_DblClick
End Sub

Private Sub mnSair_Click()
    Unload Me
End Sub

Private Sub optCat_Click()
    busca
End Sub

Private Sub optnMarca_Click()
    busca
End Sub

Private Sub optnProduto_Click()
    busca
End Sub

Private Sub txtCliente_Change()
    txtCliente.tag = ""
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCli_Click
    End If
End Sub



Private Sub txtDe2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub txtAte2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte2_LostFocus()
    P_AjustaData txtAte2
End Sub


Private Sub txtProd_Change()
    txtProd.tag = ""
End Sub

Private Sub txtProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
    End If
End Sub
