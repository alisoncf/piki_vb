VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiEspelhoNota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Notas Fiscais de Saída"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9600
   Begin VB.Frame frCOntrole 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.Frame Frame6 
         Caption         =   "E&xpressão:"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2295
         Begin VB.TextBox txtbxExpr 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   15
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnTerc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            Picture         =   "frmEdiEspelhoNota.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2295
         Begin VB.TextBox txtbxdataFin 
            Height          =   285
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxDataIni 
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1200
            TabIndex        =   12
            Top             =   0
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   2295
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   685
            Left            =   120
            Picture         =   "frmEdiEspelhoNota.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnImprimir 
            Caption         =   "&Imprimir"
            Height          =   685
            Left            =   120
            Picture         =   "frmEdiEspelhoNota.frx":04D4
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Novo Espelho de Nota"
            Height          =   685
            Left            =   120
            Picture         =   "frmEdiEspelhoNota.frx":061E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnCancelar 
            Caption         =   "&Cancelar"
            Height          =   685
            Left            =   120
            Picture         =   "frmEdiEspelhoNota.frx":09A8
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1680
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "&Buscar Por:"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2295
         Begin VB.ComboBox cmbxBuscar 
            Height          =   315
            ItemData        =   "frmEdiEspelhoNota.frx":0AF2
            Left            =   120
            List            =   "frmEdiEspelhoNota.frx":0AFC
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "&Ordenar Por:"
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   2295
         Begin VB.ComboBox cmbxOrder 
            Height          =   315
            ItemData        =   "frmEdiEspelhoNota.frx":0B11
            Left            =   120
            List            =   "frmEdiEspelhoNota.frx":0B21
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   240
            Width           =   2055
         End
      End
      Begin MSAdodcLib.Adodc adc 
         Height          =   570
         Left            =   2640
         Top             =   2760
         Visible         =   0   'False
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   1005
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
         Bindings        =   "frmEdiEspelhoNota.frx":0B49
         Height          =   6015
         Left            =   2520
         TabIndex        =   18
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   10610
         _Version        =   393216
         AllowUpdate     =   0   'False
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
End
Attribute VB_Name = "frmEdiEspelhoNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxBuscar_Click()
    If cmbxBuscar.ListIndex = 1 Then
        cmndbtnTerc.Visible = True
    Else
        cmndbtnTerc.Visible = False
    End If
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sOrd As String
    If cmbxOrder.ListIndex = 0 Then
        sOrd = "nota_sai"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrd = "Nome_cli"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrd = "Emissao_sai"
    Else
        sOrd = "emissao_sai"
    End If
    strSQL = "select ts.*,tc.nome_cli,tc.uf_cli from tab_clientes as tc, tab_saidas as ts "
    strSQL = strSQL & " where ts.cod_cli=tc.cod_cli "
    If cmndbtnTerc.Visible = True Then
        strSQL = strSQL & " and tc.nome_cli like '" & txtbxExpr.text & "%' "
    Else
        strSQL = strSQL & " and ts.nota_sai like '" & txtbxExpr.text & "%' "
    End If
    If txtbxdataFin.text <> "" Then
        strSQL = strSQL & " and ts.emissao_sai < " & F_PesqData(txtbxdataFin.text)
    End If
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and ts.emissao_sai > " & F_PesqData(txtbxDataIni.text)
    End If
    strSQL = strSQL & " and ts.serie_sai not in('CUPM','CANC')"
    strSQL = strSQL & " order by " & sOrd
    adc.RecordSource = strSQL
'    adc.Refresh
End Sub

Private Sub cmndbtnCancelar_Click()
    Dim strSQL As String
    With adc.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If MsgBox("Deseja Realmente Excluir a Nota Fiscal " & !nota_sai & _
            " ?", vbYesNo + vbCritical) = vbYes Then
            P_cancelaNotaFiscalSaida !nota_sai, !serie_sai
            
            
            adc.Refresh
        End If
    End With
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmEntradas
    fr.operacao = "I"
    If adc.Recordset.RecordCount <= 0 Or dbgrd.row = -1 Then
        MsgBox "Não há notas fiscais nesta pesquisa!", vbExclamation
        Exit Sub
    End If
    
    With adc.Recordset
        If !status_not <> "1" Then
            MsgBox "A Nota Não está Aberta!", vbExclamation
            Exit Sub
        End If
        
        fr.tag = !seq_not
        fr.grid.tag = fr.tag
        fr.caption = "Nota Fiscal: " & !Num_not & " de " & !nome_for
        fr.txtbxDesconto.text = FormatVal(!desconto_not)
        fr.txtbxNum_Not.text = !Num_not
        fr.txtbxNum_Not.tag = !Num_not
        fr.txtbxSerie_Not.text = !serie_not
        fr.txtbxSerie_Not.tag = !serie_not
        fr.txtbxBC_Not.text = Format(!bc_not, "##,###,##0.00")
        fr.txtbxValorProd.text = FormatVal(!valorprod_not)
        fr.txtbxDepesas_Not.text = Format(!despesas_not, "###,###,##0.00")
        fr.txtbxFrete.text = Format(!frete_not, "###,###,##0.00")
        fr.txtbxEmissao_Not.text = Format(!emissao_not)
        fr.txtbxEntrada_Not.text = Format(!entrada_not)
        fr.txtbxICMS_Not.text = Format(!icms_not, "###,###,##0.00")
        fr.txtbxIPI_Not.text = Format(!ipi_not, "###,###,##0.00")
        fr.txtbxNOme_Ter.text = !nome_for
        fr.txtbxNOme_Ter.tag = !cod_For
        fr.lblUF_Ter.caption = Format(!uf_for)
        
        fr.txtbxValor_Not.text = FormatVal(!valor_not)
        
        P_PreencheOperacoes fr.cmbxTipo_Op, "E"
        
         
        P_LocalizaOperacao !operacao_not, fr.cmbxTipo_Op
        
        
        fr.Show 0, mdiFrm
    End With
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_Erro
    With adc.Recordset
        imprEspelhoNota !nota_sai, !serie_sai
    
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNew_Click()
        frmEspelhoNota.Show 0
    frmEspelhoNota.SetFocus

End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    adc.Recordset.Close
    Unload Me
End Sub

Private Sub cmndbtnTerc_Click()
     
    frmSelectCli.txtbxNome.tag = "C"
    
    frmSelectCli.txtbxNome.text = txtbxExpr.text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxExpr.text = frmSelectCli.adoR!nome
    txtbxExpr.tag = frmSelectCli.adoR!cod
    Unload frmSelectCli
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub adc_Reposition()
    With adc.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        
        
    End With
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        
        
        
    End If
End Sub

Private Sub Form_Load()
   ' Me.Width = 10560
   ' Me.Height = 6585
    P_CentraForm Me
    adc.ConnectionString = Conn.ConnectionString
    cmbxBuscar.ListIndex = 0
    cmbxOrder.ListIndex = 0
    
    cmbxBuscar_Click
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And cmndbtnTerc.Visible = True Then
        cmndbtnTerc_Click
    End If
    cmndbtnBuscar_Click
End Sub


