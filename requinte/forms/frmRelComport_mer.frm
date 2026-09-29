VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelComport_mer 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Relatório de Comportamento Periódico de Mercadorias"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.ComboBox cmbxComport 
         BackColor       =   &H80000004&
         Height          =   315
         ItemData        =   "frmRelComport_mer.frx":0000
         Left            =   5640
         List            =   "frmRelComport_mer.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   9960
         Picture         =   "frmRelComport_mer.frx":001E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Salvar grid como texto"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   9360
         Picture         =   "frmRelComport_mer.frx":03A8
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Copiar Grid"
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cmbxComp 
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   420
         Width           =   2655
      End
      Begin VB.Frame Frame4 
         Caption         =   "Valor:"
         Height          =   580
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2295
         Begin VB.OptionButton optnBusc 
            Caption         =   "&Bruto"
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optnBusc 
            Caption         =   "&Líquido"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Buscar:"
         Height          =   580
         Left            =   2640
         TabIndex        =   14
         Top             =   840
         Width           =   2655
         Begin VB.OptionButton optnTipo 
            Caption         =   "&Mercadorias"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   200
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optnTipo 
            Caption         =   "&Serviços"
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   16
            Top             =   200
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmndbtnChart 
         Height          =   495
         Left            =   11160
         Picture         =   "frmRelComport_mer.frx":0732
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Visualizar Gráfico"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cmbxPor 
         BackColor       =   &H80000000&
         Height          =   315
         ItemData        =   "frmRelComport_mer.frx":13FC
         Left            =   2640
         List            =   "frmRelComport_mer.frx":140C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   420
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Orçs. &Fechados"
         Height          =   255
         Left            =   7440
         TabIndex        =   18
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnConfirma 
         Height          =   495
         Left            =   9360
         Picture         =   "frmRelComport_mer.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Confirma a efetivação do Orçamento."
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnFechar 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   11160
         Picture         =   "frmRelComport_mer.frx":1DBE
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Fecha o formulário atual."
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Height          =   495
         Left            =   10560
         Picture         =   "frmRelComport_mer.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Imprime o resultado apresentado."
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   450
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   450
         Width           =   975
      End
      Begin VB.TextBox txtbxNome_mer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4560
         TabIndex        =   9
         Top             =   450
         Width           =   3855
      End
      Begin VB.CheckBox chk3Eixo 
         Caption         =   "3º &Eixos"
         Height          =   255
         Left            =   7440
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "&Comportamento:"
         Height          =   195
         Left            =   5640
         TabIndex        =   35
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblMarca 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         ForeColor       =   &H8000000C&
         Height          =   195
         Left            =   7920
         TabIndex        =   10
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Expressão:"
         Height          =   195
         Left            =   4560
         TabIndex        =   7
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Buscar &Por:"
         Height          =   195
         Left            =   2640
         TabIndex        =   5
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data &Final"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "&Data Inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   25
      Top             =   1440
      Width           =   11775
      Begin VB.Data dtVendas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDBGrid.DBGrid dbgrdVendas 
         Bindings        =   "frmRelComport_mer.frx":24D2
         Height          =   4935
         Left            =   120
         OleObjectBlob   =   "frmRelComport_mer.frx":24E9
         TabIndex        =   26
         Top             =   120
         Width           =   11535
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lucro Médio:"
      Height          =   195
      Left            =   7920
      TabIndex        =   34
      Top             =   6840
      Width           =   930
   End
   Begin VB.Label lblLucroMed 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   8880
      TabIndex        =   33
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label lblLiq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6600
      TabIndex        =   32
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receita Liqüida:"
      Height          =   195
      Left            =   5280
      TabIndex        =   31
      Top             =   6840
      Width           =   1155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Itens Vendidos:"
      Height          =   195
      Left            =   3000
      TabIndex        =   30
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label lblItens 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4200
      TabIndex        =   29
      Top             =   6840
      Width           =   90
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   10560
      TabIndex        =   28
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   10080
      TabIndex        =   27
      Top             =   6840
      Width           =   405
   End
End
Attribute VB_Name = "frmRelComport_mer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Mer
    Codigo As String
    Descricao As String
    qtd_total As Currency
    valor_total As Currency
End Type
Sub P_buscaTotal()
    Dim cTotal As Currency
    Dim cLiq As Currency
    Dim cLucro As Currency
    With dtVendas.Recordset
        Do While Not .EOF
            cTotal = cTotal + !valor_total
            cLiq = cLiq + !Valor_Liq
            cLucro = cLucro + !Lucro_Medio
            .MoveNext
        Loop
        
    End With
    If dtVendas.Recordset.RecordCount = 0 Then Exit Sub
    lblTotal.caption = F_Format(cTotal)
    lblItens.caption = dtVendas.Recordset.RecordCount
    lblLiq.caption = F_Format(cLiq)
    lblLucroMed.caption = Format(cLucro / dtVendas.Recordset.RecordCount, "#0.0") & " %"
    If lblItens.caption > 0 Then
        dtVendas.Recordset.MoveLast
    End If
End Sub

Private Sub cmbxComp_Click()
    If cmbxComp.Text = "" Then
        cmbxComp.Tag = ""
    Else
        cmbxComp.Tag = F_BuscaComponente(1, cmbxComp.Text, Left(cmbxPor.Text, 1))
    End If
End Sub


Private Sub cmbxPor_Click()
    If cmbxPor.ListIndex = 0 Then
        txtbxNome_mer.Visible = True
        lblMarca.Visible = True
        txtbxNome_mer.Text = ""
        txtbxNome_mer.Tag = ""
        lblMarca.Tag = ""
        cmbxComp.Visible = False
    Else
        
        txtbxNome_mer.Visible = False
        lblMarca.Visible = False
        txtbxNome_mer.Text = ""
        txtbxNome_mer.Tag = ""
        lblMarca.Tag = ""
        If cmbxPor.ListIndex = 1 Then
            P_PreencheComboComp cmbxComp, "G"
        ElseIf cmbxPor.ListIndex = 2 Then
            P_PreencheComboComp cmbxComp, "S"
        Else
            P_PreencheComboComp cmbxComp, "M"
        End If
        cmbxComp.Visible = True
    End If
    
End Sub

Private Sub cmndbtnChart_Click()
    Dim bl As BlChart
    Dim cRest As Currency
    Dim iCount As Integer
    Dim iResult As Integer
    
        bl.count = dtVendas.Recordset.RecordCount
    
    With dtVendas.Recordset
    .MoveFirst
       
        For iCount = 0 To .RecordCount - 1
            bl.val(iCount).caption = F_AjustaNome(!Mes)
            If optnBusc(0).Value = True Then
                bl.val(iCount).Value = !valor_total
            Else
                bl.val(iCount).Value = !Valor_Liq
            End If
            .MoveNext
            iResult = iCount
            If .EOF Then
                iResult = 0
                Exit For
            End If
        Next
        cRest = 0
        Do While Not .EOF
            cRest = cRest + !valor_total
            .MoveNext
        Loop
'        If Check2.Value = 1 Then
'            If iResult > 0 Then
'                bl.val(15).caption = "Demais"
'                bl.val(15).Value = cRest
'            End If
'        End If
        .MovePrevious
    End With
    
    P_Chart dtVendas.Recordset!valor_total, 1, bl, 8, False, "Gráfico de Comportamento de Mercadorias - " & cmbxPor.Text & " :" & cmbxComp.Text & txtbxNome_mer.Text
End Sub

Private Sub cmndbtnConfirma_Click()
    P_Status "Buscando informações...   Aguarde..."
     strSQL = "select sum(ti.quantidade_ito) as Qtd_total, "
    strSQL = strSQL & " sum(ti.unitario_ito * ti.quantidade_ito) As Valor_total, "
    strSQL = strSQL & " sum((ti.unitario_ito-tr.custo_mer) * ti.quantidade_ito) as Valor_liq, "
    strSQL = strSQL & " (sum((ti.unitario_ito-tr.custo_mer) * ti.quantidade_ito) *100/  "
    strSQL = strSQL & " sum(ti.unitario_ito * ti.quantidade_ito)) as Lucro_Medio, "
    If cmbxComport.ListIndex = 0 Then
        strSQL = strSQL & " substring(tor.datafechamento_orc,1,7) as mes,"
        dbgrdVendas.Columns(0).NumberFormat = "mm/yyyy"
    Else
        strSQL = strSQL & " tor.datafechamento_orc as mes,"
        dbgrdVendas.Columns(0).NumberFormat = "dd/mm/yyyy"
    End If
    If cmbxPor.ListIndex = 0 Then
        strSQL = strSQL & "tm.nome_mer as nome,tm.cod_mer as codigo "
    Else
        strSQL = strSQL & "tc.nome_com as nome,tc.cod_com as codigo "
    End If
    strSQL = strSQL & " From  tab_mercadorias as tm, tab_itensorcamento as ti, "
    strSQL = strSQL & " tab_orcamentos As tor,TAB_COMPONENTES as tc,tab_real01 as tr  "
    strSQL = strSQL & " Where tm.cod_mer = ti.cod_mer and tm.marca_Mer = ti.marca_Mer "
    strSQL = strSQL & " and tm.cod_mer=tr.cod_mer and tm.marca_mer=tr.marca_mer "
    If cmbxPor.ListIndex = 1 Then
        strSQL = strSQL & " and tc.cod_com=tm.gru_mer and tc.tipo_com='G'"
    ElseIf cmbxPor.ListIndex = 2 Then
        strSQL = strSQL & " and tc.cod_com=tm.sub_mer and tc.tipo_com='S'"
    Else
        strSQL = strSQL & " and tc.cod_com=tm.Marca_mer and tc.tipo_com='M'"
    End If
    strSQL = strSQL & " and "
    strSQL = strSQL & " ti.orcamento_orc = tor.orcamento_orc"
    strSQL = strSQL & " and tor.datafechamento_orc <=  " & F_PesqData(txtbxDAtaFIn.Text)
    strSQL = strSQL & " and tor.datafechamento_orc >=  " & F_PesqData(txtbxDataIni.Text)
    If optnTipo(0).Value = True Then
        strSQL = strSQL & " And tm.tipo_mer='M' "
    Else
        strSQL = strSQL & " And tm.tipo_mer='S' "
    End If
    If chk3Eixo.Value = 0 Then
        strSQL = strSQL & " and tm.GRU_MER <> '24' "
    End If
    If Check1.Value = 1 Then
        strSQL = strSQL & " and tor.situacao_orc in('2','4')"
    Else
        strSQL = strSQL & " and tor.situacao_orc ='4'"
    End If
    
    If cmbxPor.ListIndex = 1 Then
        If cmbxComp.Tag <> "" Then
            strSQL = strSQL & " and tm.gru_mer='" & cmbxComp.Tag & "' "
        End If
        
        strSQL = strSQL & " Group By  tm.Gru_mer  "
        dbgrdVendas.TabIndex = 1
    ElseIf cmbxPor.ListIndex = 2 Then
        If cmbxComp.Tag <> "" Then
            strSQL = strSQL & " and tm.sub_mer='" & cmbxComp.Tag & "' "
        End If
        strSQL = strSQL & " Group By  tm.sub_mer "
    ElseIf cmbxPor.ListIndex = 3 Then
        If cmbxComp.Tag <> "" Then
            strSQL = strSQL & " and tm.marca_mer='" & cmbxComp.Tag & "' "
        End If
        strSQL = strSQL & " Group By  tm.Marca_mer "
    Else
        If txtbxNome_mer.Tag <> "" Then
            strSQL = strSQL & " and tm.cod_mer='" & txtbxNome_mer.Tag & "' "
        End If
        strSQL = strSQL & " Group By  tm.Cod_mer "
    End If
    strSQL = strSQL & ", mes "
  '  If optnBusc(0).Value = True Then
  '      strSQL = strSQL & " order by Valor_Liq desc"
  '  Else
  '      strSQL = strSQL & " order by Valor_total desc"
  '  End If
    strSQL = strSQL & " order by nome,mes"
    
    dtVendas.Connect = banco.Connect
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
    
    P_Status
    P_buscaTotal
    If dtVendas.Recordset.RecordCount > 0 Then
        dtVendas.Recordset.MoveFirst
    End If
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
    Dim mMer(10) As Mer
    Dim iCount As Integer
    Dim cQtd_Total As Currency, cValor_Total As Currency
    With dtVendas.Recordset
        .MoveFirst
        Do While Not .EOF
            For iCount = 0 To 9
                If !valor_total > mMer(iCount).valor_total Then
                    If iCount <> 0 Then
                        mMer(iCount - 1).valor_total = mMer(iCount).valor_total
                        mMer(iCount - 1).qtd_total = mMer(iCount).qtd_total
                        mMer(iCount - 1).Codigo = mMer(iCount).Codigo
                        mMer(iCount - 1).Descricao = mMer(iCount).Descricao
                    End If
                    mMer(iCount).valor_total = !valor_total
                    mMer(iCount).qtd_total = !qtd_total
                    mMer(iCount).Codigo = !Codigo
                    mMer(iCount).Descricao = !nome
                Else
                    Exit For
                End If
            Next
            cQtd_Total = cQtd_Total + !qtd_total
            cValor_Total = cValor_Total + !valor_total
            .MoveNext
        Loop
    End With
    
    
    With frmGraph.Mschart1
      .chartType = 1 'VtChChartType2dLine
      .ColumnCount = 1
      .RowCount = 10
      .Column = 1
      For Column = 1 To 1
         For Row = 0 To 9
            .Column = Column
            .Row = Row + 1
            .RowLabel = F_Format(mMer(Row).Codigo, "&&.&&.&&&-&")
            .Data = mMer(Row).valor_total / cValor_Total * 100
            '.FootnoteText = "teste"
            '.
            frmGraph.Label2.caption = frmGraph.Label2.caption & mMer(Row).Codigo & "-" & mMer(Row).Descricao & Chr(13)
        Next Row
      Next Column
      
      .ShowLegend = True
      .SelectPart VtChPartTypePlot, index1, index2, index3, index4
      .EditCopy
      .SelectPart VtChPartTypeLegend, index1, index2, index3, index4
      .EditPaste
      '.ColumnLabel = "$"
    End With
    frmGraph.Label1.caption = "Mercadorias Por Expressividade "
    frmGraph.lblFonte.caption = "CPD " & strEmp_Abrevia
    frmGraph.Show 1
    
End Sub

Private Sub Form_Load()
    cmbxPor.ListIndex = 0
End Sub

Private Sub txtbxDataFin_GotFocus()
    P_SelecTexto txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_LostFocus()
    P_AjustaData txtbxDAtaFIn
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxNome_Mer_DblClick()
    Dim m As Mercadoria
    m = F_Mercadoria_Localiza("NOME_MER", txtbxNome_mer.Text)
    txtbxNome_mer.Text = m.nome
    txtbxNome_mer.Tag = m.Codigo
    lblMarca.caption = m.NMarca
    lblMarca.Tag = m.CMarca
End Sub

Private Sub txtbxNome_Mer_GotFocus()
    P_SelecTexto txtbxNome_mer
End Sub

Private Sub txtbxNome_Mer_LostFocus()
    P_AjustaTexto txtbxNome_mer
End Sub
