VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelComiss 
   Caption         =   "Relatório de Comissão"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   13950
   Icon            =   "frmRelComiss.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   13950
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   873
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pcComissDif 
      Height          =   2055
      Left            =   2040
      ScaleHeight     =   1995
      ScaleWidth      =   13890
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   13950
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   13950
      TabIndex        =   12
      Top             =   7425
      Width           =   13950
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         ScaleHeight     =   255
         ScaleWidth      =   4335
         TabIndex        =   14
         Top             =   0
         Width           =   4335
         Begin VB.OptionButton optnGeral 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Valor"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1440
            TabIndex        =   16
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optnTipoPag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Comissao"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox txtFun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   11400
         Picture         =   "frmRelComiss.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Detalhes"
         Height          =   735
         Left            =   9360
         Picture         =   "frmRelComiss.frx":0537
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico de Vendas"
         Height          =   735
         Left            =   7320
         Picture         =   "frmRelComiss.frx":0621
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   5280
         Picture         =   "frmRelComiss.frx":076B
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   3240
         Picture         =   "frmRelComiss.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmnbuscarFun 
         Height          =   285
         Left            =   2400
         Picture         =   "frmRelComiss.frx":0F02
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmRelComiss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoCom As ADODB.Recordset
Attribute adoCom.VB_VarHelpID = -1
Dim WithEvents adoDif As ADODB.Recordset
Attribute adoDif.VB_VarHelpID = -1
Dim adoVendedores As New ADODB.Recordset
Dim AdoVendas As New ADODB.Recordset

Dim iQtdColFP As Integer
Dim iColTotalVendas As Integer
Dim iColDifTroca As Integer
Private Sub cmbxFunc_Click()
    
End Sub

Private Sub cmbxFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxFunc.ListIndex = -1
    End If
End Sub

Private Sub cmnbuscarFun_Click()
    buscaFun
    
End Sub
Sub buscaFun()
    Dim f As funcionario.funcionario
    
    f = funcionario.F_SelectFuncionario(txtFun.Text)
    If f.cod_fun <> "" Then
        
        txtFun.Text = f.guerra_fun
        txtFun.tag = f.cod_fun
        cmndbtnBuscar_Click
    End If
    
End Sub
Sub buscaVendedores()
    Dim strSQL As String
    
    strSQL = "select distinct tf.cod_fun,tf.guerra_fun,tf.nome_fun from tab_vendas as tv, tab_cargoscomiss as tc,tab_cargos as tca, "
        strSQL = strSQL & " tab_funcionarios as tf,tab_formaspag as tfp where tv.abrevenda_usu=tf.cod_fun "
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and tv.dataefet_ven>=" & F_PesqData(txtbxDataIni.Text)
        End If
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and tv.dataefet_ven<=" & F_PesqData(txtbxdataFin.Text)
        End If
        strSQL = strSQL & " and tfp.cod_for=tv.cod_For "
        strSQL = strSQL & " and tfp.tipo_cai=tc.tipo_cai "
        strSQL = strSQL & " and tc.cod_car=tf.cod_car"
        strSQL = strSQL & " and tc.cod_car=tca.cod_car "
        If txtFun.tag <> "" Then
            strSQL = strSQL & " and abrevenda_usu='" & txtFun.tag & "' "
        End If
        
        
        strSQL = strSQL & " and situacao_ven=4 "
        strSQL = strSQL & " order by guerra_fun "
        
        abreObjetoAdo adoVendedores, strSQL
        grid.rows = 1
        With adoVendedores
            If .RecordCount > 0 Then
            .MoveFirst
            End If
            Do While Not .EOF
                grid.AddItem !cod_fun & Chr(9) & !guerra_fun
                
                .MoveNext
            Loop
            
        End With
        
        
    
End Sub

Sub buscaVendas()
    Dim strSQL As String
    Dim iCol As Integer
    Dim iRow As Integer
    strSQL = "select distinct tca.nome_car,tc.comiss_car,tc.tipo_cai,sum(tx.valor_cai) as total,tx.esp_cai,(sum(valor_cai)*tc.comiss_car/100) as comiss,"
        strSQL = strSQL & " tf.cod_fun,tf.guerra_fun,tf.nome_fun from tab_vendas as tv  "
        strSQL = strSQL & " inner join tab_caixa as tx on tv.num_ven=tx.num_ven "
        strSQL = strSQL & " inner join tab_funcionarios as tf on tv.abrevenda_usu=tf.cod_fun "
        strSQL = strSQL & " inner join tab_cargos as tca on tca.cod_car=tf.cod_car "
        strSQL = strSQL & " inner join tab_cargoscomiss as tc on tc.cod_car=tca.cod_car and tx.esp_cai=tc.tipo_cai "
        
        strSQL = strSQL & " where 1>0 "
        
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
        End If
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
        End If
        If txtFun.tag <> "" Then
            strSQL = strSQL & " and abrevenda_usu='" & txtFun.tag & "' "
        End If
        strSQL = strSQL & " and tx.tipo_cai='C' "
        
        strSQL = strSQL & " and situacao_ven=4 "
       
            strSQL = strSQL & " group by esp_cai, tf.cod_fun order by tf.guerra_fun,tc.tipo_cai  "
        '    dbgrd.Columns(2).Visible = True
        
        
        abreObjetoAdo AdoVendas, strSQL
        
        With AdoVendas
            Do While Not .EOF
                iCol = F_InFlxgrdCol(grid, !esp_cai, 0, 2, iQtdColFP + 2)
                
                If iCol = -1 Then
                    MsgBox "Não foi encontrada a coluna com a espécie " & !esp_cai, vbExclamation
                    Exit Sub
                End If
                iRow = F_InFlxgrd(grid, !cod_fun, 0)
                grid.col = iCol
                grid.row = iRow
                
                
                If optnTipoPag.Value = True Then
                    grid.Text = FormatVal(!comiss)
                Else
                    grid.Text = FormatVal(!total)
                End If
                .MoveNext
            Loop
        End With
                
        
End Sub

Sub buscaDifTroca()
    Dim strSQL As String

    strSQL = "select distinct tca.nome_car,tc.comiss_car,tc.tipo_cai,sum(tt.acres_tro-tt.abat_tro-desconto_tro) as total,"
    strSQL = strSQL & " tfp.tipo_cai,(sum(tt.acres_tro-tt.abat_tro-desconto_tro)*tc.comiss_car/100) as comiss,"
    strSQL = strSQL & " tf.cod_fun,tf.guerra_fun,tf.nome_fun from tab_vendas as tv, tab_cargoscomiss as tc,tab_cargos as tca, tab_trocas as tt,"
    strSQL = strSQL & " tab_funcionarios as tf,tab_formaspag as tfp where tt.cod_usu=tf.cod_fun "
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_tro>=" & F_GravaDataTempo(txtbxDataIni.Text)
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_tro<=" & F_GravaDataTempo(txtbxdataFin.Text & " 23:59:59")
    End If
    strSQL = strSQL & " and tfp.cod_for=tv.cod_For "
    strSQL = strSQL & " and tfp.tipo_cai=tc.tipo_cai "
    strSQL = strSQL & " and tc.cod_car=tf.cod_car"
    
    strSQL = strSQL & " and tc.cod_car=tca.cod_car "
    strSQL = strSQL & " and tt.num_ven=tv.num_ven "
    If txtFun.Text <> "" Then
        strSQL = strSQL & " and tf.cod_fun='" & txtFun.tag & "' "
    End If
    
    
    strSQL = strSQL & " and situacao_ven=4 "
    strSQL = strSQL & " group by tfp.tipo_cai, tf.cod_fun order by tf.guerra_fun "
    abreObjetoAdo adoDif, strSQL
    
    With adoDif
            Do While Not .EOF
                iCol = F_InFlxgrdCol(grid, !tipo_cai, 0, iColDifTroca - iQtdColFP, grid.Cols - 2)
                iRow = F_InFlxgrd(grid, !cod_fun, 0)
                grid.col = iCol
                If iRow > -1 Then
                    grid.row = iRow
                    If optnGeral.Value = True Then
                        grid.Text = FormatVal(!total)
                    Else
                        grid.Text = FormatVal(!comiss)
                    End If
                End If
                .MoveNext
            Loop
        End With
    
    
    
    
End Sub
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    
    If txtFun.tag <> strCod_Usu Then
        If strVisrelest_usu <> "S" Then
            MsgBox "Você não pode visualizar uma comissão de outra pessoa!", vbExclamation
            Exit Sub
        End If
    End If
    
    buscaVendedores
    
    buscaVendas
    buscaDifTroca
    soma
    
    P_Total
    
    
End Sub
Sub P_Total()
    Dim i As Integer
    Dim j As Integer
    Dim total As Currency
    With grid
        .AddItem Chr(9) & "Totais...:"
        .col = 1: .row = .rows - 1
        .CellBackColor = vbMagenta
        .CellFontBold = True
        For i = 2 To .Cols - 1
            .col = i
            total = 0
            For j = 1 To .rows - 2
                .row = j
                If .Text = "" Then .Text = "0"
                total = total + CCur(.Text)
                
            Next
            .row = .rows - 1
            .Text = FormatVal(total)
            .CellForeColor = vbGreen
            .CellFontBold = True
            .CellBackColor = vbMagenta
        Next
    End With
End Sub
Sub soma()
    Dim i As Integer
    Dim j As Integer
    Dim cVendas As Currency
    Dim cDif As Currency
    Dim cTotalCol As Currency
    
    With grid
        .Visible = False
        For i = 1 To .rows - 1
            .row = i
            cTotalCol = 0
            For j = 2 To iColTotalVendas - 1
                .col = j
                If .Text = "" Then
                    .Text = "0"
                End If
                cTotalCol = cTotalCol + CCur(.Text)
                
            Next
            .col = iColTotalVendas
            .Text = FormatVal(cTotalCol)
            cTotalCol = 0
            .CellFontBold = True
            .CellBackColor = vbYellow
        Next
    
    End With
     With grid
        For i = 1 To .rows - 1
            .row = i
            cTotalCol = 0
            For j = iColTotalVendas + 1 To iColDifTroca - 1
                .col = j
                If .Text = "" Then
                    .Text = "0"
                End If
                cTotalCol = cTotalCol + CCur(.Text)
                
            Next
            .col = iColDifTroca
            .Text = FormatVal(cTotalCol)
            cTotalCol = 0
            .CellFontBold = True
            .CellBackColor = vbGreen
            .CellForeColor = vbWhite
        Next
    
    End With
    
    With grid
        For i = 1 To .rows - 1
            .row = i
            cTotalCol = 0
            .col = iColDifTroca
            cDif = .Text
            .col = iColTotalVendas
            cVendas = .Text
            
            .col = .Cols - 1
            .Text = FormatVal(cDif + cVendas)
            .CellFontBold = True
            .CellBackColor = vbRed
        Next
    
    End With
    grid.Visible = True
    
End Sub

Private Sub cmndbtnImprimir_Click()
    imprHTML
    Exit Sub
    imprime frmVisualiza.Pic
    
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprHTML()


    P_MsflxToHtml "Relatório de Comissões", grid, "Relatório de Comissões de " & txtbxDataIni.Text & " até " & txtbxdataFin.Text


Exit Sub

Exit Sub

    Dim sNomeArq As String
    Dim sCodFun As String
    Dim cTotal As Currency
    Dim cTotalComiss As Currency
    Dim cTotalTotalComiss As Currency
    Dim cTotalTotal As Currency
    sNomeArq = F_AbreHtml("rel_comiss")
    P_HtmlAbreHead
    P_HtmlAbreTitle 1, "Relatório de Comissões", True
    P_HtmlFechaHead
    P_HtmlAbreBody
    P_HtmlAbreTable 1, 1, 700
    
    P_HtmlAbreLinha
        P_HtmlAbreColuna 1, "Relatório de " & IIf(optnVendas.Value, "Comissões", "Diferenças de Trocas") & ". Período de " & txtbxDataIni.Text & " Até " & txtbxdataFin.Text, , True, , "4"
    P_HtmlFechaLinha
    
    With adc.Recordset
        .MoveFirst
        Do While Not .EOF
            If sCodFun <> !cod_fun Then
                If .AbsolutePosition > 0 Then
                    P_HtmlAbreLinha
                        P_HtmlAbreColuna 1, "Total"
                        P_HtmlAbreColuna 1, FormatVal(cTotal)
                        P_HtmlAbreColuna 1, FormatVal(cTotalComiss)
                        P_HtmlAbreColuna 1, "."
                    P_HtmlFechaLinha
                End If
                
                sCodFun = !cod_fun
                cTotal = 0
                cTotalComiss = 0
                P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, !nome_fun & " (Cargo: " & !nome_car & ")", , True, , "4"
                P_HtmlFechaLinha
                P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, "Tipo Pag."
                    P_HtmlAbreColuna 1, "Vlr Vendido."
                    P_HtmlAbreColuna 1, "Comissão"
                    P_HtmlAbreColuna 1, "%"
                P_HtmlFechaLinha
            End If
            P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, !tipo_cai
                    P_HtmlAbreColuna 1, FormatVal(!total)
                    P_HtmlAbreColuna 1, FormatVal(!comiss)
                    
                    P_HtmlAbreColuna 1, !comiss_car
            P_HtmlFechaLinha
            
            cTotal = cTotal + !total
            cTotalComiss = cTotalComiss + !comiss
            cTotalTotal = cTotalTotal + !total
            cTotalTotalComiss = cTotalTotalComiss + !comiss
            .MoveNext
        Loop
        
    End With
    
    P_HtmlAbreLinha
        P_HtmlAbreColuna 1, "Soma", , True, , "4"
    P_HtmlFechaLinha
    P_HtmlAbreLinha
        P_HtmlAbreColuna 1, "Geral"
        P_HtmlAbreColuna 1, FormatVal(cTotalTotal)
        P_HtmlAbreColuna 1, FormatVal(cTotalTotalComiss)
        P_HtmlAbreColuna 1, "."
    P_HtmlFechaLinha
    
    
    P_HtmlFechaTable 1
    P_HtmlFechaBody
    P_FechaHtml , True, sNomeArq
    
    
    
    
End Sub
    
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub
Sub imprime(obj As Object)
    P_CabRelTimbr "Relatório de Comissões de " & txtbxDataIni.Text & " até " & txtbxdataFin.Text, obj
    Dim sCodFun As String
    P_ImprLinha obj, "-"
    
    
    
    With adc.Recordset
        .MoveFirst
        Do While Not .EOF
            If !cod_fun <> sCodFun Then
                sCodFun = !cod_fun
                P_ImprX obj, "Nome", 0
                P_ImprLinha obj, "-"
                P_ImprX obj, "TipoCai", 10
                P_ImprX obj, "Vendido", 70, True, True
                P_ImprX obj, "Comissão", 100, True, True
                P_ImprX obj, "Perc.", 115, False, True
                P_ImprLinha obj, "-"
            End If
            
            P_ImprX obj, !guerra_fun, 10
            P_ImprX obj, FormatVal(!total), 70, True, True
            P_ImprX obj, FormatVal(!comiss), 100, True, True
            P_ImprX obj, FormatVal(!total / CCur(lblVendas.caption) * 100), 115, False, True

            .MoveNext
        Loop
        P_RodapeVis obj, constNome_sis
    End With
    
    
End Sub
Sub AjustaGridTotal()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    Dim sFP As String
    
    strSQL = "select distinct tipo_cai from tab_formaspag "
    abreObjetoAdo rPesq, strSQL
    
    iQtdColFP = rPesq.RecordCount
    
    
    
    
    
    With grid
        .Cols = 2
        .rows = 1
        
        
        
        
        .col = 0: .Text = "Código"
        .col = 1: .Text = "Nome"
        grid.ColWidth(1) = 2000
        
        With rPesq
            Do While Not .EOF
            
                grid.Cols = grid.Cols + 1
                grid.col = grid.Cols - 1
                grid.Text = !tipo_cai
                grid.CellBackColor = vbYellow
                .MoveNext
            Loop
            grid.Cols = grid.Cols + 1
            grid.col = grid.Cols - 1
            If optnGeral.Value Then
                grid.Text = "Total Vendas"
            Else
                grid.Text = "Total Comiss."
            End If
            iColTotalVendas = grid.col
            grid.ColWidth(grid.col) = 1200
            grid.CellBackColor = vbYellow
        End With
        With rPesq
            .MoveFirst
            Do While Not .EOF
            
                grid.Cols = grid.Cols + 1
                grid.col = grid.Cols - 1
                grid.Text = !tipo_cai
                grid.CellBackColor = vbGreen
                .MoveNext
            Loop
            grid.Cols = grid.Cols + 1
            grid.col = grid.Cols - 1
            grid.Text = "Dif. Troca"
            grid.ColWidth(grid.col) = 1200
            grid.CellBackColor = vbGreen
            iColDifTroca = grid.col
        End With
        grid.Cols = grid.Cols + 1
        grid.col = grid.Cols - 1
        grid.Text = "Total"
        grid.ColWidth(grid.col) = 1200
        grid.CellBackColor = vbRed
    End With
End Sub
Private Sub Form_Load()
    resizeForm Me
    
    Set adoCom = New ADODB.Recordset
    Set adoDif = New ADODB.Recordset
    
    'P_PreencheVendedor cmbxFunc, 1
    
    AjustaGridTotal
    If strVisrelfin_usu <> "S" Then
        txtFun.Text = strNome_Usu
        txtFun.tag = strCod_Usu
        
        txtFun.Enabled = False
        cmnbuscarFun.Enabled = False
        
    End If
    
End Sub

Private Sub Option1_Click()
    dbgrd.caption = "Diferenças de Troca no Período"
    cmndbtnBuscar_Click
End Sub

Private Sub optnVendas_Click()
    dbgrd.caption = "Comissões de Vendas no Período"
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
    posicionaADireita cmndbtnSair, pcBot
End Sub

Private Sub optnGeral_Click()
    AjustaGridTotal
End Sub

Private Sub optnTipoPag_Click()
    AjustaGridTotal
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtFun_Change()
    txtFun.tag = ""
End Sub

Private Sub txtFun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFun
    End If
End Sub
