VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelVendasFrequencia 
   Caption         =   "Relatório de Frequência de Vendas"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15255
   Icon            =   "frmRelVendasFrequencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   15255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcEsq 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7395
      Left            =   0
      ScaleHeight     =   7395
      ScaleWidth      =   2175
      TabIndex        =   34
      Top             =   1455
      Width           =   2175
      Begin VB.ListBox lstDia 
         Height          =   1635
         ItemData        =   "frmRelVendasFrequencia.frx":1084A
         Left            =   120
         List            =   "frmRelVendasFrequencia.frx":10851
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   4080
         Width           =   1935
      End
      Begin VB.CommandButton cmdTodosDia 
         Caption         =   "Todos"
         Height          =   255
         Left            =   960
         TabIndex        =   28
         Top             =   3735
         Width           =   1095
      End
      Begin VB.CommandButton cmdTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
      Begin VB.ListBox lstPeriodo 
         Height          =   3210
         ItemData        =   "frmRelVendasFrequencia.frx":1085D
         Left            =   120
         List            =   "frmRelVendasFrequencia.frx":1085F
         Style           =   1  'Checkbox
         TabIndex        =   26
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblMenorFreq 
         AutoSize        =   -1  'True
         Caption         =   "Menor frequência: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   6120
         Width           =   1425
      End
      Begin VB.Label lblMaiorFreq 
         AutoSize        =   -1  'True
         Caption         =   "Maior frequencia: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   5880
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   3735
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Períodos:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   690
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grade 
      Height          =   5895
      Left            =   2280
      TabIndex        =   31
      Top             =   1560
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10398
      _Version        =   393216
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   15225
      TabIndex        =   30
      Top             =   0
      Width           =   15255
      Begin VB.TextBox txtCli 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9360
         TabIndex        =   15
         Top             =   960
         Width           =   2895
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Buscar por"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   2400
         TabIndex        =   4
         Top             =   0
         Width           =   3615
         Begin VB.OptionButton optnVendas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Vendas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2400
            TabIndex        =   38
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optnValor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Valor"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1440
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optnQtd 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Quantidade"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox txtVendedor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   12600
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         TabIndex        =   13
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3120
         TabIndex        =   11
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   615
         Left            =   8760
         Picture         =   "frmRelVendasFrequencia.frx":10861
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   10080
         Picture         =   "frmRelVendasFrequencia.frx":10BB1
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   12720
         Picture         =   "frmRelVendasFrequencia.frx":1109F
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   595
         Left            =   11400
         Picture         =   "frmRelVendasFrequencia.frx":1162D
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   315
         Left            =   2760
         Picture         =   "frmRelVendasFrequencia.frx":11A5C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Height          =   315
         Left            =   9000
         Picture         =   "frmRelVendasFrequencia.frx":11DE6
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton cmdCat 
         Height          =   315
         Left            =   5760
         Picture         =   "frmRelVendasFrequencia.frx":12170
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton cmdCli 
         Height          =   315
         Left            =   12240
         Picture         =   "frmRelVendasFrequencia.frx":124FA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton cmdVen 
         Height          =   315
         Left            =   13800
         Picture         =   "frmRelVendasFrequencia.frx":12884
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblCodFab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód.Fab"
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
         TabIndex        =   37
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   9360
         TabIndex        =   14
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   12600
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   3120
         TabIndex        =   10
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   195
         Left            =   6120
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmRelVendasFrequencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim maiorFreq As Currency
Dim menorFreq As Currency

Private Sub cmdExportar_Click()
    ExportaMSFLXGridToCSV grade
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdTodosDia_Click()
    Dim b As Boolean
    Dim i As Integer
    If cmdTodosDia.tag = "" Then
        cmdTodosDia.tag = "x"
        cmdTodosDia.caption = "Nenhum"
        b = True
    Else
        cmdTodosDia.tag = ""
        cmdTodosDia.caption = "Todos"
        b = False
    End If
    For i = 0 To lstDia.ListCount - 1
        lstDia.Selected(i) = b
    Next
End Sub

Private Sub cmdCli_Click()
    buscaCli
End Sub

Private Sub cmdMarca_Click()
    buscaMarca
End Sub

Private Sub cmdTodos_Click()
    Dim b As Boolean
    Dim i As Integer
    If cmdTodos.tag = "" Then
        cmdTodos.tag = "x"
        cmdTodos.caption = "Nenhum"
        b = True
    Else
        cmdTodos.tag = ""
        cmdTodos.caption = "Todos"
        b = False
    End If
    For i = 0 To lstPeriodo.ListCount - 1
        lstPeriodo.Selected(i) = b
    Next
    
End Sub

Private Sub cmdVen_Click()
    buscaVendedor
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub soma()
    Dim i As Integer
    Dim j As Integer
    Dim c2 As Currency
    
    With grade
        For i = 1 To .rows - 2
            For j = 1 To .Cols - 2
                .col = j
                .row = i
                If .Text = "" Then .Text = "0"
                c2 = c2 + CCur(.Text)
            Next
            .col = .Cols - 1
            .row = i
            If optnQtd.Value Then
                .Text = (c2)
            Else
                .Text = FormatVal(c2)
            End If
            c2 = 0
          
        Next
    End With
    c2 = 0
    
    With grade
        For i = 1 To .Cols - 1
            For j = 1 To .rows - 2
                .row = j
                .col = i
                If .Text = "" Then .Text = "0"
                c2 = c2 + CCur(.Text)
            Next
            .row = .rows - 1
            .col = i
            If optnQtd.Value Then
                .Text = (c2)
            Else
                .Text = FormatVal(c2)
            End If
            c2 = 0
          
        Next
    End With
    lblMaiorFreq.caption = "Maior freq: " & maiorFreq
    lblMenorFreq.caption = "Menor freq: " & menorFreq
End Sub
Sub busca()
    P_Status "Pesquisando..."
    grade.Visible = False
    
    menorFreq = 0
    maiorFreq = 0
    ajustaGrid
    Dim valor As Currency
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select substring(data_ven,12,2) as hora,weekday(data_ven) as dia,"
    strSQL = strSQL & " sum(ti.qtd_ite-ti.devolvido_ite) as qtd, sum((ti.final_ite/qtd_ite)*(qtd_ite-devolvido_ite)) as total, count(distinct tv.num_ven) as vendas "
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv inner join " & TAB_ITVENDAS & " as ti"
    strSQL = strSQL & " on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " inner join " & TAB_CLIENTES & " as tc on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " inner join " & TAB_CATEGORIAS & " as tcat on tp.cod_cat=tcat.cod_cat "
    strSQL = strSQL & " inner join " & TAB_MARCAS & " as tm on tm.cod_mar=tp.cod_mar "
    strSQL = strSQL & " inner join " & TAB_FUNCIONARIOS & " as tf on tf.cod_fun=tv.abrevenda_usu "
    strSQL = strSQL & " where situacao_ven=4"
    If txtDataIni.Text <> "" Then
        strSQL = strSQL & " and left(tv.data_ven,10)>=" & F_GravaData(txtDataIni.Text)
    End If
    If txtDataFin.Text <> "" Then
        strSQL = strSQL & " and left(tv.data_ven,10)<=" & F_GravaData(txtDataFin.Text)
    End If
    If txtCat.tag <> "" Then
        strSQL = strSQL & " and tp.cod_cat='" & txtCat.tag & "'"
    End If
    If txtMarca.tag <> "" Then
        strSQL = strSQL & " and tp.cod_mar='" & txtmar.tag & "'"
    End If
    If txtProduto.tag <> "" Then
        strSQL = strSQL & " and tp.cod_cod='" & txtProduto.tag & "'"
    End If
    If txtCli.tag <> "" Then
        strSQL = strSQL & " and tv.cod_cli='" & txtCli.tag & "'"
    End If
    If txtVendedor.tag <> "" Then
        strSQL = strSQL & " and tv.abrevenda_usu='" & txtVendedor.tag & "'"
    End If
    strSQL = strSQL & " group by dia, hora"
    strSQL = strSQL & " order by dia, hora"
    abreObjetoAdo rPesq, strSQL

    With rPesq
        
        Do While Not .EOF
            
            grade.col = !dia + 1
            grade.row = !hora + 1
            If grade.Text = "" Then
                valor = 0
            Else
                valor = CCur(grade.Text)
            End If
            If optnQtd.Value = True Then
                valor = valor + !qtd
                grade.Text = valor
            ElseIf optnVendas.Value = True Then
                valor = valor + !Vendas
                grade.Text = valor
            Else
                valor = valor + !total
                grade.Text = FormatVal(valor)
            End If
            If menorFreq = 0 Then
                menorFreq = valor
            End If
            If valor > maiorFreq Then
                maiorFreq = valor
            End If
            If valor > 0 Then
                If valor < menorFreq Then
                    menorFreq = valor
                End If
            End If
            
            .MoveNext
        Loop
    End With


    soma
    BuscarMedia
    grade.Visible = True
    P_Status
End Sub

Private Sub cmndbtnImprimir_Click()
    Dim extra As String
    extra = "De: " & txtDataIni.Text & " até: " & txtDataFin.Text & " produto: " & txtProduto.Text
    
    P_MsflxToHtml "Relatório de frequência de vendas", grade, extra
End Sub

Private Sub cmndbtnLocaliza_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As Produto
    p = Produtos.F_BuscaProdSaida(txtProduto.Text)
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
        lblCodFab.caption = p.codfab_pro
    End If
    
End Sub
Sub buscaMarca()
    Dim p As Marca.Marca
    p = Marca.buscaMarca(txtMarca.Text)
    If p.codigo <> "" Then
        txtMarca.Text = p.nome
        txtMarca.tag = p.codigo
        
    End If
    
End Sub
Sub buscaCat()
    Dim p As categoria.categoria
    p = categoria.BuscaCategoria(txtCat.Text)
    If p.codigo <> "" Then
        txtCat.Text = p.nome
        txtCat.tag = p.codigo
        
    End If
    
End Sub
Sub buscaCli()
    Dim p As cliente
    p = Clientes.F_BuscaCliente(txtCli.Text)
    If p.cod_cli <> "" Then
        txtCli.Text = p.nome_Cli
        txtCli.tag = p.cod_cli
        
    End If
    
End Sub
Sub buscaVendedor()
    Dim p As funcionario.funcionario
    p = funcionario.f_BuscaFun(txtVendedor.Text)
    If p.cod_fun <> "" Then
        txtVendedor.Text = p.nome_fun
        txtVendedor.tag = p.cod_fun
        
    End If
    
End Sub

Private Sub cmdCat_Click()
    buscaCat
End Sub


Private Sub Form_Load()
    txtDataIni.Text = F_PrimDiaMes(Date)
    txtDataFin.Text = F_UltDiaMes(Date)
    
    PreencheDia
    PreencheHora
       ajustaGrid
End Sub
Sub MostrarNiveis()
    Dim pico As Single
    Dim media As Single
    Dim baixa As Single
    pico = maiorFreq * 0.7
    pico = baixa * 0.7
    
    
    
End Sub
Sub BuscarMedia()
    Dim i As Integer
    Dim j As Integer
    Dim m As Integer
    Dim val As String
    With grade
        .Cols = .Cols + 1
        .row = 0
        .col = .Cols - 1
        .Text = "Média"
        .ColWidth(.Cols - 1) = 1200
        For i = 1 To .rows - 1
            .row = i
            m = 0
            For j = 1 To .Cols - 3
                .col = j
                If CCur(.Text) > 0 And .ColWidth(j) > 0 Then
                    m = m + 1
                End If
            Next
            
            
            
            .col = .Cols - 2
            val = FormatVal(.Text)
            .col = .Cols - 1
            If m > 0 Then
                .Text = FormatVal(CCur(val) / m)
            Else
                .Text = "0"
            End If
            .CellFontItalic = True
            .CellForeColor = vbBlue
        Next
    End With
    With grade
        .rows = .rows + 1
        .col = 0
        .row = .rows - 1
        .Text = "Média"
        .RowHeight(.rows - 1) = 500
        For i = 1 To .Cols - 1
            .col = i
            m = 0
            For j = 1 To .rows - 3
                .row = j
                If CCur(.Text) > 0 And .RowHeight(j) > 0 Then
                    m = m + 1
                End If
            Next
            
            
            
            .row = .rows - 2
            val = FormatVal(.Text)
            .row = .rows - 1
            If m > 0 Then
                .Text = FormatVal(CCur(val) / m)
            Else
                .Text = "0"
            End If
            .CellFontItalic = True
            .CellForeColor = vbBlue
        Next
    End With
    MostrarNiveis
End Sub
Sub PreencheHora()
    Dim i As Integer
    lstPeriodo.Clear
    For i = 1 To 24
            
        lstPeriodo.AddItem (Format(i - 1, "0#") & ":00")
        If i > 7 And i <= 20 Then
            lstPeriodo.Selected(i - 1) = True
        End If
    Next
End Sub
Sub PreencheDia()
    Dim i As Integer
    lstDia.Clear
    lstDia.AddItem "Segunda"
    lstDia.AddItem "Terça"
    lstDia.AddItem "Quarta"
    lstDia.AddItem "Quinta"
    lstDia.AddItem "Sexta"
    lstDia.AddItem "Sábado"
    lstDia.AddItem "Domingo"
    For i = 1 To 6
        lstDia.Selected(i - 1) = True
        
    Next
    
End Sub

Sub ajustaGrid()
    Dim i As Integer
    Dim j As Integer
    
    grade.Clear
    
    
    With grade
        .Cols = 9
        .FixedCols = 1
        .FixedRows = 1
        .rows = 25
        .row = 0
        .Font = "Arial"
        .col = 1: .Text = "Segunda"
        .col = 2: .Text = "Terça"
        .col = 3: .Text = "Quarta"
        .col = 4: .Text = "Quinta"
        .col = 5: .Text = "Sexta"
        .col = 6: .Text = "Sábado"
        .col = 7: .Text = "Domingo"
        .col = 8: .Text = "Total"
            
        For i = 1 To 8
            grade.ColWidth(i) = 1200
        Next
        
            
        For i = 1 To 24
            .col = 0
            
            .row = i
            .Text = Format(i - 1, "0#") & ":00"
            .RowHeight(i) = 300
            If lstPeriodo.Selected(i - 1) = False Then
                .RowHeight(i) = 0
            
            End If
            
        Next
        .AddItem "Total"
        .ColAlignment(0) = 3
        For i = 1 To 7
            If lstDia.Selected(i - 1) = False Then
                .ColWidth(i) = 0
            End If
    Next
    End With
    
End Sub

Private Sub txtCat_Change()
    txtCat.tag = ""
End Sub

Private Sub txtCat_GotFocus()
    P_SelecTexto txtCat
End Sub

Private Sub txtCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCat
    End If
End Sub

Private Sub txtCli_Change()
    txtCli.tag = ""
End Sub

Private Sub txtCli_GotFocus()
    P_SelecTexto txtCli
End Sub

Private Sub txtCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCli
    End If
End Sub

Private Sub txtDataIni_GotFocus()
    P_SelecTexto txtDataIni
End Sub

Private Sub txtDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDataIni_LostFocus()
    P_AjustaData txtDataIni
End Sub

Private Sub txtDataFin_GotFocus()
    P_SelecTexto txtDataFin
End Sub

Private Sub txtDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDataFin_LostFocus()
    P_AjustaData txtDataFin
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

Private Sub txtProduto_Change()
    txtProduto.tag = ""
End Sub

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
    End If
End Sub

Private Sub txtVendedor_Change()
    txtVendedor.tag = ""
End Sub

Private Sub txtVendedor_GotFocus()
    P_SelecTexto txtVendedor
End Sub

Private Sub txtVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaVendedor
    End If
End Sub
