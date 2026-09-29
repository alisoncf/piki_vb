VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProdutoSaida 
   Caption         =   "Selecionar um Produto"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12480
   ForeColor       =   &H00000000&
   Icon            =   "frmProdutoSaida.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcSimilares 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   12450
      TabIndex        =   20
      Top             =   7755
      Visible         =   0   'False
      Width           =   12480
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Trazer Produtos Zerados"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6960
         TabIndex        =   24
         Tag             =   "!"
         Top             =   0
         Width           =   2175
      End
      Begin VB.CommandButton cmdListSim 
         Caption         =   "Listar apenas similares (L)"
         Height          =   315
         Left            =   4560
         TabIndex        =   22
         Top             =   0
         Width           =   2055
      End
      Begin VB.CommandButton cmdAbrirSim 
         Caption         =   "Abrir (C)"
         Height          =   315
         Left            =   3720
         TabIndex        =   21
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Este produto possui registros similares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   12450
      TabIndex        =   0
      Top             =   0
      Width           =   12480
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         Picture         =   "frmProdutoSaida.frx":014A
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   1
         Top             =   150
         Width           =   240
      End
      Begin VB.Label lblProduto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione um Produto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmProdutoSaida.frx":049A
      Height          =   6495
      Left            =   0
      TabIndex        =   18
      Top             =   480
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11456
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
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
      Caption         =   "Lista de Produtos"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "ean_pro"
         Caption         =   "EAN"
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
      BeginProperty Column02 
         DataField       =   "codfab_pro"
         Caption         =   "Ref."
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
      BeginProperty Column04 
         DataField       =   "er_pro"
         Caption         =   "Qtde."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "pv_pro"
         Caption         =   "Preço Unitário"
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
      BeginProperty Column06 
         DataField       =   "Cod_Emp"
         Caption         =   "Emp."
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
         DataField       =   "er"
         Caption         =   "Er"
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
         DataField       =   "preco"
         Caption         =   "Prç Atacado"
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
      BeginProperty Column09 
         DataField       =   "qtd"
         Caption         =   "Emb. Atacado"
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
      BeginProperty Column10 
         DataField       =   "perc_imp"
         Caption         =   "Imposto"
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
      BeginProperty Column11 
         DataField       =   "ncm_pro"
         Caption         =   "NCM"
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
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1725,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4935,118
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   555,024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column11 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EF9C7C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12480
      TabIndex        =   3
      Top             =   8130
      Width           =   12480
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6480
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.PictureBox pcQtd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   8400
         ScaleHeight     =   825
         ScaleWidth      =   1545
         TabIndex        =   14
         Top             =   0
         Width           =   1575
         Begin VB.TextBox txtQtd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   15
            Text            =   "0,00"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd. (F3)"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   11280
         Picture         =   "frmProdutoSaida.frx":04AC
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Selecionar"
         Height          =   615
         Left            =   10080
         Picture         =   "frmProdutoSaida.frx":0A3A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtbxBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cmbxOpcao 
         Height          =   315
         ItemData        =   "frmProdutoSaida.frx":0EC1
         Left            =   120
         List            =   "frmProdutoSaida.frx":0ECB
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         Picture         =   "frmProdutoSaida.frx":0EE9
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8040
         Picture         =   "frmProdutoSaida.frx":1294
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   255
         Left            =   6480
         TabIndex        =   11
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   4440
         TabIndex        =   8
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expressão:"
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opções"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmProdutoSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    VerSim ado!cod_pro
    lblProduto.caption = ado!nome_pro
    
    
    
End Sub

Sub VerSim(cod_pro As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If cod_pro = "" Then
        pcSimilares.Visible = False
    End If
    strSQL = "select count(*) as total from tab_similares where cod_pro='" & cod_pro & "' "
   
    abreObjetoAdo rPesq, strSQL
   
    If rPesq!total > 0 Then
        pcSimilares.Visible = True
    Else
        pcSimilares.Visible = False
    End If
   
   
   
   
    
    
End Sub

Private Sub chkTodas_Click()
    Busca
End Sub

Private Sub cmdAbrirSim_Click()
    On Error Resume Next
    buscaSimilares ado!cod_pro
End Sub

Private Sub cmdCancel_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmdCat_Click()
    BuscaCat
End Sub

Private Sub cmdMarca_Click()
    BuscarMarca
End Sub

Private Sub cmdOk_Click()
    Dim m As Double
    On Error GoTo trata_Erro
    With ado
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If txtQtd.Text = "" Then
            Exit Sub
        End If
        If CCur(txtQtd.Text) > !qtd Then
            m = CCur(txtQtd.Text) Mod !qtd
            If m > 0 Then
                MsgBox "A quantidade informada deve ser múltipla de " & !qtd, vbExclamation
                txtQtd.SetFocus
                Exit Sub
           
            
            End If
        End If
        If txtbxBuscar.Text = !barra Then
            If !qtd <> CCur(txtQtd.Text) Then
        
                    'MsgBox "Para este produto ser vendido em preço de atacado, a quantidade informada deve ser " & !QTD, vbExclamation
                    'Exit Sub
                'txtQtd.Text = !QTD
            End If
        End If
        Me.tag = !cod_pro
        Me.Hide
    End With
trata_Erro:
End Sub

Private Sub dtgrd_DblClick()
    cmdOk_Click
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
    If KeyCode = vbKeyN Or KeyCode = vbKeyP Or KeyCode = vbKeyM Then
        frmBusca.tag = "N"
        frmBusca.Show 1
        Me.tag = frmBusca.tag
        Unload frmBusca
        Busca
    ElseIf KeyCode = vbKeyS Then
        Busca "", True
    End If
End Sub


Private Sub Form_Activate()
    ' 'MudaAparencia Me
    
    Busca
    On Error Resume Next
    dtgrd.SetFocus
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtQtd.SetFocus
    End If
    If KeyCode = vbKeyEscape Then
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    resizeForm Me
End Sub
Sub Busca(Optional limit As String = "", Optional bSim As Boolean = False)
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim iCount As Integer
    
        If txtbxBuscar.Text = "" Then
            MsgBox "Informe parte do nome ou descrição.", vbInformation
            txtbxBuscar.SetFocus
            Exit Sub
        End If
                
                
        strSQL = "select distinct tc.nome_Cat, tc.descmax_cat,tm.nome_mar, ti.perc_imp,"
        strSQL = strSQL & "tp.*,e.cod_emp,e.er,e.ef,tp.er_pro as estoque, tpa.qtd,tpa.preco,tpa.barra  "
        strSQL = strSQL & " from "
        strSQL = strSQL & " " & TAB_PRODUTOS & " as tp "
        strSQL = strSQL & " inner join estoque as e on tp.cod_pro=e.cod_pro "
        strSQL = strSQL & " left join tab_produtoatacado as tpa on tp.cod_pro=tpa.codigo "
        'strSQL = strSQL & " left join tab_codbarras as tb on tb.cod_pro=tp.cod_pro "
        strSQL = strSQL & " left join " & TAB_MARCAS & " as tm on tm.cod_mar=tp.cod_mar "
        strSQL = strSQL & " left join " & tab_categorias & " as tc on tc.cod_cat=tp.cod_cat "
        strSQL = strSQL & " left join " & TAB_IMPOSTOS & " as ti on ti.cod_imp=tp.cod_imp"
        If bSim Then
            strSQL = strSQL & " left join tab_similares as ts on tp.cod_pro=ts.cod_pro or tp.cod_pro=ts.cod_sim "
            strSQL = strSQL & " where ts.cod_pro='" & ado!cod_pro & "' or ts.cod_sim='" & ado!cod_pro & "' "
        Else
        strSQL = strSQL & " where 1>0 " 'tc.cod_cat=tp.cod_cat  and tm.cod_mar=tp.cod_mar "
       
        
                
            strSQL = strSQL & " and  (tp.codbarra_PRO = '"
            strSQL = strSQL & Me.tag & "' "
            strSQL = strSQL & " or  ean_pro = '"
            strSQL = strSQL & Me.tag & "' "
            
            strSQL = strSQL & " or  codfab_pro = '"
            strSQL = strSQL & Me.tag & "' "
            
            
            'strSQL = strSQL & " or  tb.codbarra_pro like '"
            'strSQL = strSQL & Me.tag & "' "
            
            If bUsaVendaSupermercado = False Then
            
                strSQL = strSQL & " or tp.COD_PRO = '"
                strSQL = strSQL & (Me.tag) & "' "
            
                strSQL = strSQL & " or tp.CODfab_PRO like '%"
                strSQL = strSQL & (Me.tag) & "%' "
            End If
            If bUsaVendaSupermercado = True Then
            
                strSQL = strSQL & " or (tp.CODbarra_PRO = '"
                strSQL = strSQL & Left(Me.tag, 7) & "' "
                strSQL = strSQL & " and tp.balanca_pro='S') "
                
                strSQL = strSQL & " or (ean_pro= '"
                strSQL = strSQL & Left(Me.tag, 7) & "' "
                strSQL = strSQL & " and tp.balanca_pro='S') "
                
                strSQL = strSQL & " or tp.COD_PRO = '"
                strSQL = strSQL & (Me.tag) & "' "
                
                strSQL = strSQL & " or tp.CODfab_PRO = '"
                strSQL = strSQL & (Me.tag) & "' "
                
            End If
            
            
            tipo_busca = "B"
        
            If bUsaVendaSupermercado = True Then
                If F_SoNumero(Me.tag) <> Me.tag Then
                    strSQL = strSQL & " or NOME_PRO like '"
                    strSQL = strSQL & Me.tag & "%'"
                End If
            
            Else
                strSQL = strSQL & " or NOME_PRO like '%"
                strSQL = strSQL & Me.tag & "%'"
            End If
            
            strSQL = strSQL & ")"
        End If
        strSQL = strSQL & " and tp.visivel_pro='S' "
        
        'strSQL = strSQL & " and tp.cod_pro=tb.cod_pro "
    '   strSQL = strSQL & " and tp.cod_pro=ts.cod_pro "
        'strSQL = strSQL & " and tc.cod_cat=tp.cod_Cat and tm.cod_mar=tp.cod_mar "
    
        If chkTodas.Value = 0 Then
    
    '        strSQL = strSQL & " and (ts.er_pro-ts.r_pro) > 0 "
        End If
    
        strSQL = strSQL & " order by tp.nome_pro"
    
        
        If txtMarca.tag <> "" Then
            strSQL = strSQL & " and tp.cod_mar='" & txtMarca.tag & "' "
        End If
        If txtCat.tag <> "" Then
            strSQL = strSQL & " and tp.cod_cat='" & txtCat.tag & "' "
        End If
        
        
        abreObjetoAdo ado, strSQL, dtgrd
        
        If ado.RecordCount = 1 Then
            cmdOk_Click
        Else
            
            frmProdutoSaida.Visible = True
        End If
End Sub

Private Sub Form_Resize()
    PosicionaGrade dtgrd, Me, pcBot
End Sub



Private Sub txtbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.tag = F_AjustaTexto(txtbxBuscar)
        Busca
    End If
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
End Sub

Private Sub txtCat_Change()
    txtCat.tag = ""
    
End Sub

Private Sub txtCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        BuscaCat
    End If
End Sub
Sub BuscaCat()
    Dim c As categoria.categoria
    c = BuscaCategoria(txtCat.Text)
    If c.codigo <> "" Then
        txtCat.Text = c.nome
        txtCat.tag = c.codigo
        Busca
    End If
End Sub
Sub BuscarMarca()
    Dim c As Marca.Marca
    c = BuscaMarca(txtMarca.Text)
    If c.codigo <> "" Then
        txtMarca.Text = c.nome
        txtMarca.tag = c.codigo
        Busca
    End If
End Sub

Private Sub txtMarca_Change()
    txtMarca.tag = ""
End Sub

Private Sub txtMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        BuscarMarca
    End If
End Sub

Private Sub txtQtd_GotFocus()
    ganhaFoco txtQtd
End Sub

Private Sub txtQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey 13
End Sub

Private Sub txtQtd_LostFocus()
On Error Resume Next
    If CCur(txtQtd.Text) Mod ado!qtd > 0 Then
        txtQtd.BackColor = vbRed
    Else
        txtQtd.BackColor = vbGreen
    End If
    
    PerdeFoco txtQtd
End Sub
