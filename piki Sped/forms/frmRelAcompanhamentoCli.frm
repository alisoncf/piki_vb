VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelAcompanhamentoCli 
   Caption         =   "Relatório de Acompanhamento de Clientes por Período:"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   14655
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   14625
      TabIndex        =   25
      Top             =   7740
      Width           =   14655
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   1680
         Picture         =   "frmRelAcompanhamentoCli.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   720
         Left            =   120
         Picture         =   "frmRelAcompanhamentoCli.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   12600
         Picture         =   "frmRelAcompanhamentoCli.frx":0579
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4095
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      ScaleWidth      =   14655
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   21
         Top             =   840
         Width           =   4455
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Período 2:"
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   2880
         TabIndex        =   6
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtDe2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtAte2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   210
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E:"
            Height          =   195
            Left            =   1320
            TabIndex        =   9
            Top             =   240
            Width           =   150
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Período 1:"
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtAte 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   5
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtDe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E:"
            Height          =   195
            Left            =   1320
            TabIndex        =   4
            Top             =   240
            Width           =   150
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   210
         End
      End
      Begin VB.TextBox txtProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10560
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   495
         Left            =   13320
         Picture         =   "frmRelAcompanhamentoCli.frx":0AC7
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdProd 
         Height          =   285
         Left            =   12960
         Picture         =   "frmRelAcompanhamentoCli.frx":0E72
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdCat 
         Height          =   285
         Left            =   7920
         Picture         =   "frmRelAcompanhamentoCli.frx":121D
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Height          =   285
         Left            =   10200
         Picture         =   "frmRelAcompanhamentoCli.frx":15C8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdCli 
         Height          =   285
         Left            =   10200
         Picture         =   "frmRelAcompanhamentoCli.frx":1973
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   5760
         TabIndex        =   23
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   10560
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         Height          =   195
         Left            =   5760
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         Height          =   195
         Left            =   8280
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmRelAcompanhamentoCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Sub busca()
    Dim strSQL As String
    strSQL = "select tc.cod_cli as codigo, tc.nome_cli as nome,"
    strSQL = strSQL & " (select "
    strSQL = strSQL & "    sum(valor_ven) As total "
    strSQL = strSQL & " from "
    strSQL = strSQL & " vendas As tv " 'inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    'strSQL = strSQL & " inner join " & TAB_PRODUTOS & "  as tp on ti.cod_pro=tp.cod_pro"
    'strSQL = strSQL & " inner join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar"
    'strSQL = strSQL & " inner join " & TAB_CATEGORIAS & " as tc on tp.cod_Cat=tc.cod_cat "
    
    
    
    strSQL = strSQL & " Where "
    
    
    
    
    strSQL = strSQL & " tv.dataefet_ven>=" & F_GravaData(txtDe.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtAte.Text)
    strSQL = strSQL & " and situacao_ven=4"
    strSQL = strSQL & " and cod_cli=tc.cod_cli"
    
    If txtCat.tag <> "" Then
        strSQL = strSQL & " and tp.cod_cat='" & txtCat.Text & "' "
    End If
    If txtMarca.tag <> "" Then
        strSQL = strSQL & " and tp.cod_mar='" & txtmar.Text & "' "
    End If
    If txtProd.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pro='" & txtProd.Text & "' "
    End If
    
    
    strSQL = strSQL & ") as periodo1,"
    
    strSQL = strSQL & " (select "
    strSQL = strSQL & "    sum(valor_ven) As total "
    strSQL = strSQL & " from "
    strSQL = strSQL & " vendas As tv " 'inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    'strSQL = strSQL & " inner join " & TAB_PRODUTOS & "  as tp on ti.cod_pro=tp.cod_pro"
    'strSQL = strSQL & " inner join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar"
    'strSQL = strSQL & " inner join " & TAB_CATEGORIAS & " as tc on tp.cod_Cat=tc.cod_cat "
    
    
    strSQL = strSQL & " Where "
    
    
    
    strSQL = strSQL & " tv.dataefet_ven>=" & F_GravaData(txtDe2.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtAte2.Text)
    strSQL = strSQL & " and situacao_ven=4"
    strSQL = strSQL & " and cod_cli=tc.cod_cli"
    
    If txtCat.tag <> "" Then
        strSQL = strSQL & " and tp.cod_cat='" & txtCat.Text & "' "
    End If
    If txtMarca.tag <> "" Then
        strSQL = strSQL & " and tp.cod_mar='" & txtmar.Text & "' "
    End If
    If txtProd.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pro='" & txtProd.Text & "' "
    End If
    
    
    
    
    strSQL = strSQL & ") as periodo2 "
    
    strSQL = strSQL & " from " & TAB_CLIENTES & " as tc "
  
    If txtCliente.tag <> "" Then
        strSQL = strSQL & " where cod_cli='" & txtCliente.tag & "' "
    End If
  
  
    strSQL = strSQL & " order by nome_cli "
    abreObjetoAdo ado, strSQL, grid

End Sub
Private Sub cmdCat_Click()
    Dim c As categoria.categoria
    c = BuscaCategoria(txtCat.Text)
    If c.codigo <> "" Then
        txtCat.Text = c.nome
        txtCat.tag = c.codigo
        
    End If
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV grid, ado
End Sub

Private Sub cmdMarca_Click()
    Dim m As Marca.Marca
    m = BuscaMarca(txtMarca.Text)
    If m.codigo <> "" Then
        txtMarca.Text = m.nome
        txtMarca.tag = m.codigo
        
    End If

End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub

Private Sub cmdProd_Click()
    Dim p As Produto
    p = Produtos.F_BuscaProd(txtProd.Text)
    If p.cod_pro <> "" Then
        txtProd.Text = p.nome_pro
        txtProd.tag = p.cod_pro
        
    End If

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
'    ajustaGrid
    Set ado = New ADODB.Recordset
    
    
End Sub
Sub ajustaGrid()
'    With grid
'        .FixedCols = 0
'        .FixedRows = 1
'        .rows = 1
'        .Cols = 5
'        .row = 0
'        .col = 0
'        .Text = "Código"
'        .col = 1
'        .Text = "Cliente"
'        .col = 2
'        .Text = "Período 1"
'        .col = 3
'        .Text = "Período 2"
'        .col = 4
'        .Text = "Total"
'        .ColWidth(1) = 3000
'
'
'    End With
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
End Sub

Private Sub txtCat_Change()
    txtCat.tag = ""
End Sub

Private Sub txtCat_GotFocus()
    P_SelecTexto txtCat
End Sub

Private Sub txtCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCat_Click
    End If
End Sub

Private Sub txtCliente_Change()
    txtCliente.tag = ""
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCli_Click
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
Private Sub txtDe2_GotFocus()
    P_SelecTexto txtDe2
End Sub

Private Sub txtDe2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDe2_LostFocus()
    P_AjustaData txtDe2
End Sub

Private Sub txtAte2_GotFocus()
    P_SelecTexto txtAte2
End Sub

Private Sub txtAte2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte2_LostFocus()
    P_AjustaData txtAte2
End Sub

Private Sub txtMarca_Change()
    txtMarca.tag = ""
End Sub

Private Sub txtProd_Change()
    txtProd.tag = ""
End Sub
