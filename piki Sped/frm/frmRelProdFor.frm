VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelProdFor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Produtos por Fornecedor"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "frmRelProdFor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9375
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   4680
      Width           =   2295
      Begin VB.Label lblQdade 
         AutoSize        =   -1  'True
         Caption         =   "Quant:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   585
         Width           =   615
      End
      Begin VB.Label lblCusto 
         AutoSize        =   -1  'True
         Caption         =   "Custo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   195
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.ComboBox cmbxForn 
         Height          =   315
         ItemData        =   "frmRelProdFor.frx":08CA
         Left            =   120
         List            =   "frmRelProdFor.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdFor.frx":08FA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2040
         Width           =   2055
      End
      Begin VB.ComboBox cmbxUF 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   2175
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxdataFin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Tag             =   "!"
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1200
            TabIndex        =   4
            Tag             =   "!"
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdFor.frx":0C04
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdFor.frx":14CE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lblMar 
         AutoSize        =   -1  'True
         Caption         =   "&Fornecedor:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&UF"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   210
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   3360
      Top             =   4920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmRelProdFor.frx":17D8
      Height          =   4575
      Left            =   2400
      TabIndex        =   16
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8070
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRelProdFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub soma()
    Dim strSQL As String
    Dim cCusto As Currency
    Dim cVen As Currency
    Dim iEst As Integer
    
    With dt.Recordset
    
        Do While Not .EOF
            cCusto = cCusto + (!unit_ite * !qtd_ite)
            iEst = iEst + !qtd_ite
            'cVen = cVen + (!pv_pro * !er_pro)
            .MoveNext
            
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        
        lblCusto.caption = "Custo: "
        'lblPreco.caption = "Preço: "
        lblQdade.caption = "Quantidade: "
        lblCusto.caption = lblCusto.caption & " " & FormatVal(cCusto)
        'lblPreco.caption = lblPreco.caption & " " & FormatVal(cVen)
        lblQdade.caption = lblQdade.caption & " " & FormatVal(iEst)
    End With

End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tc.*, ti.*, tf.nome_for, tp.*, tm.nome_mar, tp.unidade_pro,tca.nome_cat "
    strSQL = strSQL & " from tab_cmp as tc, tab_itcmp as ti,tab_Categorias as tca, "
    strSQL = strSQL & TAB_PRODUTOS & " as tp, " & TAB_FORNECEDORES & " as tf," & TAB_MARCAS & " as tm"
    
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and tca.cod_cat=tp.cod_Cat "
    strSQL = strSQL & " and tm.cod_mar=tp.cod_mar "
    strSQL = strSQL & " and tf.cod_for = tc.cod_for "
    
    If txtbxDataIni.Text <> "" And txtbxDAtaFIn.Text <> "" Then
        strSQL = strSQL & " and tc.emissao_not >=" & F_GravaData(txtbxDataIni.Text)
        strSQL = strSQL & " and tc.emissao_not <=" & F_GravaData(txtbxDAtaFIn.Text)
    End If
    If cmbxForn.Text <> "" Then
        strSQL = strSQL & " and tc.cod_for ='" & cmbxForn.tag & "'"
    End If
    
    If cmbxUf.Text <> "" Then
        strSQL = strSQL & " and tf.uf_for ='" & cmbxUf.Text & "'"
    End If
    strSQL = strSQL & " order by tc.emissao_not "
    
    dt.RecordSource = strSQL
    dt.Refresh
    soma
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

Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxUf.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnImprimir_Click()
    If dt.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que ser impresso!", vbExclamation
        Exit Sub
    End If
    If cmbxForn.tag = "" Then
        MsgBox "Escolha um fornecedor!", vbExclamation
        cmbxForn.SetFocus
        Exit Sub
    End If
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    Dim iPage As Integer
    Dim iCount As Integer
    P_CabRelTimbr "Listagem de Produtos por Fornecedor", obj
    obj.FontName = "arial"
    obj.fontsize = 12
    obj.FontBold = True
    P_ImprX obj, cmbxForn.Text, 10, False
    obj.fontsize = 8
    obj.FontBold = False
    P_ImprLinha obj, "_"
    With dt.Recordset
        .MoveFirst
        P_ImprX obj, "Cód", 0
        P_ImprX obj, "Nome", 15
        P_ImprX obj, "UN", 100
        P_ImprX obj, "Estoque", 115, True, True
        P_ImprX obj, "Custo", 130, True, True
        P_ImprX obj, "Preço", 145, True, True
        P_ImprX obj, "Categoria", 147
        P_ImprX obj, "Marca", 180, False
        P_ImprLinha obj, "_"
        iPage = 1
        Do While Not .EOF
            If iCount = 50 Then
                
                iCount = 0
                P_RodapeVis obj, strNomeSis & "-" & constVersao_Sis, "Pág:" & iPage
                iPage = iPage + 1
                P_ImprX obj, "Cód", 0
                P_ImprX obj, "Nome", 15
                P_ImprX obj, "UN", 100
                P_ImprX obj, "Estoque", 115, True, True
                P_ImprX obj, "Custo", 130, True, True
                P_ImprX obj, "Preço", 145, True, True
                P_ImprX obj, "Categoria", 147
                P_ImprX obj, "Marca", 180, False
                P_ImprLinha obj, "_"
            End If
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome_pro, 15
            P_ImprX obj, !unidade_pro, 100
            P_ImprX obj, FormatVal(!er_pro), 115, True, True
            P_ImprX obj, FormatVal(!cr_pro), 130, True, True
            P_ImprX obj, FormatVal(!pv_pro), 145, True, True
            P_ImprX obj, !nome_cat, 147
            P_ImprX obj, " " & !nome_mar, 179, False
            
            iCount = iCount + 1
            .MoveNext
        Loop
    End With
    P_RodapeVis obj, strNomeSis & "-" & constVersao_Sis, "Pág:" & iPage & "(Fim)"
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    P_PreencheUF cmbxUf
    P_PreencheFornecedores cmbxForn
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dt.ConnectionString = strConnect
    
        
End Sub


Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
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
