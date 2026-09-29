VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelVendasVendedor 
   Caption         =   "Vendas por Vendedor"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14040
   Icon            =   "frmRelVendasVendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   14040
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid grid 
      Height          =   5175
      Left            =   120
      TabIndex        =   29
      Top             =   1320
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   9128
      _Version        =   393216
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "guerra_fun"
         Caption         =   "Vendedor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cod"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "nome"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "qtd"
         Caption         =   "Qtd"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "qtdini"
         Caption         =   "Inicial"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "devolvido"
         Caption         =   "Devolvido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "unit"
         Caption         =   "Unit. Méd."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   6015,118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1275,024
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   14010
      TabIndex        =   25
      Top             =   6735
      Width           =   14040
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   12360
         Picture         =   "frmRelVendasVendedor.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   720
         Left            =   120
         Picture         =   "frmRelVendasVendedor.frx":10D98
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   1680
         Picture         =   "frmRelVendasVendedor.frx":111C7
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total: 0"
         Height          =   195
         Left            =   4320
         TabIndex        =   30
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   14040
      TabIndex        =   0
      Top             =   0
      Width           =   14040
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   17
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtFun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   495
         Left            =   12000
         Picture         =   "frmRelVendasVendedor.frx":11311
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Agrupar por:"
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   8400
         TabIndex        =   19
         Top             =   120
         Width           =   2415
         Begin VB.OptionButton optnCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cliente"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1320
            TabIndex        =   23
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optCat 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Categoria"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optnMarca 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Marca"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1320
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optnProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Produto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdFun 
         Height          =   285
         Left            =   1920
         Picture         =   "frmRelVendasVendedor.frx":116BC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Height          =   285
         Left            =   4440
         Picture         =   "frmRelVendasVendedor.frx":11A67
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdCat 
         Height          =   285
         Left            =   4440
         Picture         =   "frmRelVendasVendedor.frx":11E12
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdProduto 
         Height          =   285
         Left            =   7800
         Picture         =   "frmRelVendasVendedor.frx":121BD
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdCliente 
         Height          =   285
         Left            =   7800
         Picture         =   "frmRelVendasVendedor.frx":12568
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   195
         Left            =   4800
         TabIndex        =   16
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
         Height          =   195
         Left            =   4800
         TabIndex        =   7
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         Height          =   195
         Left            =   2400
         TabIndex        =   13
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria"
         Height          =   195
         Left            =   2400
         TabIndex        =   5
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmRelVendasVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Sub buscaFun()
    Dim f As funcionario.funcionario
    f = f_BuscaFun(txtFun.Text)
    If f.cod_fun <> "" Then
        txtFun.Text = f.guerra_fun
        txtFun.tag = f.cod_fun
        busca
    End If
End Sub

Private Sub cmdCat_Click()
    buscaCat
End Sub

Private Sub cmdCliente_Click()
    buscaCli
End Sub
Sub buscaCli()
    Dim c As cliente
    c = Clientes.F_BuscaCliente(txtCliente.Text)
    If c.cod_cli <> "" Then
        txtCliente.Text = c.nome_Cli
        txtCliente.tag = c.cod_cli
        busca
    End If
End Sub
Sub buscaCat()
    Dim c As categoria.categoria
    c = categoria.BuscaCategoria(txtCat.Text)
    If c.codigo <> "" Then
        txtCat.Text = c.nome
        txtCat.tag = c.codigo
        busca
    End If
End Sub
Sub buscaMarca()
    Dim c As Marca.Marca
    c = Marca.buscaMarca(txtMarca.Text)
    If c.codigo <> "" Then
        txtMarca.Text = c.nome
        txtMarca.tag = c.codigo
        busca
    End If
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV grid, ado
End Sub

Private Sub cmdFun_Click()
    buscaFun
End Sub

Private Sub cmdMarca_Click()
    Me.buscaMarca
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub
Sub soma()
    Dim cTotal As Currency
    grid.Visible = False
    With ado
        Do While Not .EOF
            If Not IsNull(!total) Then
                cTotal = cTotal + !total
            End If
            .MoveNext
        Loop
    End With
    lblTotal.caption = "Total...: " & FormatVal(cTotal)
    grid.Visible = True
End Sub
Sub busca()
    Dim strSQL As String
    Dim cod As String
    Dim nome As String
    
    If optnCliente.Value Then
        cod = "tc.cod_cli"
        nome = "tc.nome_cli"
    ElseIf optnMarca.Value Then
        cod = "tm.cod_mar"
        nome = "tm.nome_mar"
    ElseIf optnProduto.Value Then
        cod = "tp.cod_pro"
        nome = "tp.nome_pro"
    Else
        cod = "c.cod_cat"
        nome = "c.nome_cat"
    End If
    
    strSQL = "select guerra_fun,sum(qtd_ite-devolvido_ite) as qtd, sum(qtd_ite) as qtdIni, sum(devolvido_ite) as devolvido,"
    'strSQL = strSQL & " sum(((qtd_ite*unit_ite)-desconto_ite)-((devolvido_ite*unit_ite)-desconto_ite)) As total,"
    strSQL = strSQL & " sum((final_ite/qtd_ite)*(qtd_ite-devolvido_ite)) as total, "
    strSQL = strSQL & cod & " as cod, " & nome & " as nome "
    strSQL = strSQL & " from " & TAB_CLIENTES & " as tc inner join " & TAB_VENDAS & " as tv on tc.cod_cli=tv.cod_Cli "
    strSQL = strSQL & " inner join  " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven"
    strSQL = strSQL & " inner join  " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " inner join  " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " inner join  " & TAB_CATEGORIAS & " as c on tp.cod_cat=c.cod_cat "
    strSQL = strSQL & " inner join " & TAB_FUNCIONARIOS & " as tf on tf.cod_fun=tv.abrevenda_usu "
    strSQL = strSQL & " where situacao_ven='4' "
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtAte.Text)
    End If
    
    If txtProd.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pro='" & txtProd.tag & "'"
    End If
    If txtMarca.tag <> "" Then
        strSQL = strSQL & " and tp.cod_mar='" & txtMarca.tag & "'"
    End If
    If txtCat.tag <> "" Then
        strSQL = strSQL & " and tp.cod_cat='" & txtCat.tag & "'"
    End If
    If txtCliente.tag <> "" Then
        strSQL = strSQL & " and tv.cod_cli='" & txtCliente.tag & "'"
    End If
    If txtFun.tag <> "" Then
        strSQL = strSQL & " and tv.abrevenda_usu='" & txtFun.tag & "'"
    End If
    
    
    strSQL = strSQL & " group by abrevenda_usu, " & cod
    
    strSQL = strSQL & " order by guerra_fun "
    
    
    
    
    abreObjetoAdo ado, strSQL, grid
    soma
End Sub

Private Sub cmdProduto_Click()
    buscaProd
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub
Sub buscaProd()
    Dim p As Produto
    p = F_BuscaProdSaida(txtProd.Text)
    If p.cod_pro <> "" Then
        txtProd.Text = p.nome_pro
        txtProd.tag = p.cod_pro
        busca
    End If
End Sub


Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    txtDe.Text = F_PrimDiaMes(Date)
    txtAte.Text = F_UltDiaMes(Date)
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

Private Sub txtCliente_Change()
    txtCliente.tag = ""
End Sub

Private Sub txtCliente_GotFocus()
    P_SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCli
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

Private Sub txtFun_Change()
    txtFun.tag = ""
End Sub

Private Sub txtFun_GotFocus()
    P_SelecTexto txtFun
End Sub

Private Sub txtFun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFun
    End If
    
End Sub
