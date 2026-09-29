VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCadSim 
   Caption         =   "Cadastro de Similares"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11595
   LinkTopic       =   "Form2"
   ScaleHeight     =   6015
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid grid 
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   3201
      _Version        =   393216
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cod_pro"
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
      BeginProperty Column01 
         DataField       =   "codfab_pro"
         Caption         =   "Referência"
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
         DataField       =   "nome_pro"
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
         DataField       =   "nome_mar"
         Caption         =   "Marca"
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
         DataField       =   "pv_pro"
         Caption         =   "Preço"
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
         DataField       =   "er_pro"
         Caption         =   "Estoque"
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
            ColumnWidth     =   1110,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1409,953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4515,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1934,929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair (ESC)"
      Height          =   855
      Left            =   9840
      Picture         =   "frmCadSim.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   11265
      TabIndex        =   4
      Top             =   3360
      Width           =   11295
      Begin VB.ComboBox cmbxSim 
         Height          =   315
         ItemData        =   "frmCadSim.frx":054E
         Left            =   6120
         List            =   "frmCadSim.frx":0558
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddSim 
         Caption         =   "Adicionar"
         Height          =   495
         Left            =   8040
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtSim 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   5655
      End
      Begin VB.CommandButton cmdSim 
         Height          =   285
         Left            =   5760
         Picture         =   "frmCadSim.frx":0582
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Height          =   285
         Left            =   5760
         Picture         =   "frmCadSim.frx":08C4
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblMarca 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Left            =   6240
         TabIndex        =   20
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblPrincipal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   5730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Escolha o Produto:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto Principal:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lblInfoSim 
         AutoSize        =   -1  'True
         Caption         =   "Ref"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de similaridade:"
         Height          =   195
         Left            =   6120
         TabIndex        =   10
         Top             =   600
         Width           =   1440
      End
   End
   Begin VB.PictureBox pcSim 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   11265
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin VB.TextBox txtPrinc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
      Begin VB.CommandButton cmdProd 
         Height          =   285
         Left            =   5760
         Picture         =   "frmCadSim.frx":0C06
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   825
      End
      Begin VB.Label lblInfoProdPrinc 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6240
         TabIndex        =   2
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   9360
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Adicionar Novo Produto Similar ao Principal:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   3780
   End
End
Attribute VB_Name = "frmCadSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdAddSim_Click()
    If txtSim.tag = "" Then
        MsgBox "informe o produto similar!", vbExclamation
        txtSim.SetFocus
        Exit Sub
    End If
    
    If txtPrinc.tag = "" Then
        MsgBox "informe o produto principal!", vbExclamation
        txtPrinc.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Deseja adicionar a similaridade do produto " & txtSim.Text & " para o produto principal " & _
        txtPrinc.Text & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    If gravaSimimilar(txtPrinc.tag, txtSim.tag, Left(cmbxSim.Text, 1)) Then
        buscaSim
    End If
    
    
    
    
End Sub

Private Sub cmdExcluir_Click()
    If ExcluirSimilar(txtPrinc.tag, ado!cod_pro) Then
        buscaSim
    End If
    
End Sub

Private Sub cmdProd_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As Produto
    
    
    
    p = Produtos.F_BuscaProdSaida(txtPrinc.Text)
    If p.cod_pro <> "" Then
        txtPrinc.Text = p.nome_pro
        txtPrinc.tag = p.cod_pro
        
        lblInfoProdPrinc.caption = "Cód:" & p.cod_pro & " Ref:" & p.codfab_pro & Chr(13) & " Marca: " & p.nome_mar
        lblPrincipal.caption = p.nome_pro
        lblMarca.caption = p.nome_mar
        buscaSim
    Else
        lblPrincipal.caption = "Escolha um produto"
    End If
    
    
End Sub
Sub BuscaProdSim()
    Dim p As Produto
    
    p = Produtos.F_BuscaProdSaida(txtSim.Text)
    If p.cod_pro <> "" Then
        txtSim.Text = p.nome_pro
        txtSim.tag = p.cod_pro
        
        lblInfoSim.caption = "Cód:" & p.cod_pro & " Ref:" & p.codfab_pro & " Marca: " & p.nome_mar
        
    End If
    
    
End Sub

Sub buscaSim()
    Dim strSQL As String
    strSQL = "select distinct tp.er_pro,tp.nome_pro,tp.codfab_pro,tp.cod_pro,tp.pv_pro,tp.cr_pro,tm.nome_mar,tc.nome_cat,ts.nome_sub "
    strSQL = strSQL & " from tab_similares as sim, " & TAB_PRODUTOS & " as tp left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " left join tab_categorias as tc on tp.cod_cat=tc.cod_Cat "
    strSQL = strSQL & " left join tab_subcat as ts on tp.cod_sub=ts.cod_sub "
    strSQL = strSQL & " where (sim.cod_pro='" & txtPrinc.tag & "' or sim.cod_sim='" & txtPrinc.tag & "') "
    
    strSQL = strSQL & " and (tp.cod_pro=sim.cod_pro or tp.cod_pro=sim.cod_sim) "
    
    abreObjetoAdo ado, strSQL, grid
    
    
End Sub

Private Sub cmdSim_Click()
    BuscaProdSim
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub Command1_Click()
    cmdProd_Click
End Sub


Private Sub Form_Activate()
    If bActive Then Exit Sub
    
    
    If Me.tag <> "" Then
        carrega
    End If
    
    bActive = True
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    
    strSQL = "select tp.nome_pro,tp.codfab_pro,tp.cod_pro,tm.nome_mar,tc.nome_cat,ts.nome_sub "
    strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " left join tab_categorias as tc on tp.cod_cat=tc.cod_Cat "
    strSQL = strSQL & " left join tab_subcat as ts on tp.cod_sub=ts.cod_sub "
    strSQL = strSQL & " where tp.cod_pro ='" & Me.tag & "' "
    
    
    
    
    
    
    
    abreObjetoAdo rPesq, strSQL
    
    
        
    With rPesq
        txtPrinc.Text = !nome_pro
        txtPrinc.tag = !cod_pro
        lblInfoProdPrinc.caption = "Cód:" & !cod_pro & " Ref:" & !codfab_pro & Chr(13) & " Marca: " & !nome_mar
        buscaSim
    End With
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Private Sub Form_Load()
    Set ado = New ADODB.Recordset
End Sub

Private Sub txtPrinc_Change()
    txtPrinc.tag = ""
    
    
    
    
End Sub

Private Sub txtPrinc_GotFocus()
    ganhaFoco txtPrinc
End Sub

Private Sub txtPrinc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        cmdProd_Click
    End If
End Sub
