VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelFatMensal 
   Caption         =   "Relatório Análise Mensal de Faturamento"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   11565
   Icon            =   "frmRelFatMensal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   11565
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   11535
      TabIndex        =   8
      Top             =   6975
      Width           =   11565
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   11535
      TabIndex        =   0
      Top             =   0
      Width           =   11565
      Begin VB.TextBox txtAno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   690
         Left            =   2520
         Picture         =   "frmRelFatMensal.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   690
         Left            =   4080
         Picture         =   "frmRelFatMensal.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnFechar 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   690
         Left            =   9240
         Picture         =   "frmRelFatMensal.frx":0A86
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   690
         Left            =   5760
         Picture         =   "frmRelFatMensal.frx":0CD0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   330
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   5535
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9763
      _Version        =   393216
      AllowUserResizing=   2
      Appearance      =   0
   End
End
Attribute VB_Name = "frmRelFatMensal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function F_Dev(mes As String) As Currency
    Dim rPesq As New ADODB.Recordset
    strSQL = " select sum(gerado_dev) as total "
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv, tab_devolucoes as td "
    strSQL = strSQL & " where tv.num_ven=td.num_ven "
    strSQL = strSQL & " and substring(dataefet_ven,1,7)='" & Format(mes, "yyyy-mm") & "' "
    
    'strSQL = strSQL & " and tv.situacao_ven>1 "
    rPesq.Open strSQL, conn_client
    F_Dev = rPesq!total
    rPesq.Close
End Function
Sub AjustaGrid()
    Dim i As Integer
    With grid
        .rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .rows = 1
        .Cols = 15
        .col = 0: .Text = "Esp"
        .col = 1: .Text = "Descrição"
        .ColWidth(1) = 2500
        For i = 2 To 13
            .col = i
            .Text = MonthName(i - 1)
            
            .ColWidth(i) = 1000
        Next
        .col = .Cols - 1
        .ColWidth(.Cols - 1) = 2000
        .Text = "Total"
        
    End With
End Sub

Sub busca()
    Dim strSQL  As String
    Dim rPesq   As New ADODB.Recordset
    Dim rDev    As New ADODB.Recordset
    Dim cDev    As Currency
    Dim cTot    As Currency
    Dim cTrocas   As Currency
    Dim cTotal  As Currency
    Dim cTotDev As Currency
    Dim cFinan  As Currency
    Dim cAprazo As Currency
    Dim cAvista As Currency
    Dim cTotTroc As Currency
    On Error GoTo trata_erro
    Dim sDataIni As String
    Dim sDataFin As String
    Dim sUltDia As String
    Dim sSigla As String
    Dim i As Integer
    AjustaGrid
    
    sDataIni = "01/01/" & txtAno.Text
    
    sDataFin = "31/12/" & txtAno.Text
    
    strSQL = " select tf.tipo_cai,mid(tv.dataefet_ven,6,2) as mes,sum(valor_ven-devolvido_ven) as total from  " & TAB_VENDAS
    strSQL = strSQL & " as tv inner join " & TAB_FORMASPAG & " as tf on tv.cod_For=tf.cod_for "
    strSQL = strSQL & " inner join tab_espcai as te on tf.tipo_cai=te.sigla_cai "
    strSQL = strSQL & " where 1 > 0 "
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(sDataIni)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(sDataFin)
    strSQL = strSQL & " and tv.situacao_ven=4"
    strSQL = strSQL & " group by tf.tipo_cai,mes order by tf.tipo_cai,mes "
    abreObjetoAdo rPesq, strSQL
    sigla = ""
    With rPesq
            grid.rows = 1
        Do While Not .EOF
            cDev = 0
            
            If sigla <> !tipo_cai Then
                sigla = !tipo_cai
                grid.AddItem !tipo_cai & Chr(9) & F_BuscaUmCampo("descricao_Cai", "tab_Espcai", "sigla_cai='" & !tipo_cai & "'")
            End If
            
            
                
                grid.col = !mes + 1
                
                grid.row = grid.rows - 1
                    grid.Text = FormatVal(!total)
                
                        .MoveNext
        Loop
     
                
        .Close
    End With
    Soma
    ZebraGrid grid, 1
  Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdExportar_Click()
    ExportaMSFLXGridToCSV grid
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnGraf_Click()
'    With frmGraph.grph
'
'        .GraphTitle = Tit
'        .BottomTitle = Bottom
'        .Labels = 1
'        .NumPoints = 5
'        .NumSets = MSFlxGrd.rows - 1
'        If toolVis = True Then
'            .Toolbar = graphRunTime
'        End If
'        For iRow = 1 To .NumSets
'            .ThisSet = iRow
'            MSFlxGrd.Row = iRow
'            MSFlxGrd.Col = 0
'            .Legend(iRow) = MSFlxGrd.Text
'
'            For iCount = 1 To 5
'                MSFlxGrd.Col = iCount
'                MSFlxGrd.Row = iRow
'                .Data(iCount) = CCur(MSFlxGrd.Text)
'
'                MSFlxGrd.Col = 0
'
'                MSFlxGrd.Col = iCount
'                MSFlxGrd.Row = 0
'
'
'
'
'            Next
'        Next
'        For iCount = 1 To 5
'            MSFlxGrd.Row = 0
'            MSFlxGrd.Col = iCount
'            .label(iCount) = MSFlxGrd.Text
'        Next
'    .GraphType = 6
'    .DrawMode = 3
'    End With
'    frmGraph.Show 0
    grafico
    
End Sub
Sub grafico()
    Dim iRow As Integer: Dim iCount As Integer
'    With frmGraph.grph
'        On Error Resume Next
'        .GraphTitle = "Gráfico de Faturamento Mensal entre " & txtbxDataIni.Text & " e " & txtbxdataFin.Text
'
'        .BottomTitle = Bottom
'        .Labels = 1
'        .NumSets = 5
'        .NumPoints = msflxgrd.rows - 2
'        If toolVis = True Then
'            .Toolbar = graphRunTime
'        End If
'        For iRow = 1 To 5
'           msflxgrd.Col = 0
'           msflxgrd.Row = iRow
'            .Label(iRow) = msflxgrd.Text
'            .ThisSet = iRow
'            msflxgrd.Col = iRow
'            msflxgrd.Row = 0
'            .Legend(iRow) = msflxgrd.Text
'            For iCount = 1 To .NumPoints
'                msflxgrd.Row = iCount
'                .data(iCount) = msflxgrd.Text
'            Next
'        Next
'    .Label(6) = "Total"
'    .GraphType = 3
'    .DrawMode = 3
'    End With
'    frmGraph.Show 0
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    Dim iCount As Integer
    obj.FontName = "arial"
    P_CabRelTimbr "Análise Mensal de Faturamento por Venda de " & txtbxDataIni.Text & " até " & txtbxDAtaFIn.Text, obj
    obj.fontsize = 11
    obj.FontBold = True
    P_Line obj, 0, obj.CurrentY, 170, obj.CurrentY, vbDot
    P_ImprX obj, "Mês/Esp", 0, True, False, 4
    P_Line obj, 20, obj.CurrentY, 20, obj.CurrentY + 5, vbDot
    P_ImprX obj, "À Vista", 50, True, True, 4
    P_ImprX obj, "A Prazo", 80, True, True, 4
    P_ImprX obj, "Financ.", 110, True, True, 4
    P_ImprX obj, "Devolv.", 140, True, True, 4
    P_ImprX obj, "Total", 170, False, True, 4
    P_Line obj, 0, obj.CurrentY, 170, obj.CurrentY, vbDot
    obj.FontBold = False
    With MSFlxGrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 0
            P_ImprX obj, .Text, 0, True, False, 4
            P_Line obj, 20, obj.CurrentY, 20, obj.CurrentY + 5, vbDot
            .col = 1
            P_ImprX obj, .Text, 50, True, True, 4
            .col = 2
            P_ImprX obj, .Text, 80, True, True, 4
            .col = 3
            P_ImprX obj, .Text, 110, True, True, 4
            .col = 4
            P_ImprX obj, .Text, 140, True, True, 4
            .col = 5
            P_ImprX obj, .Text, 170, False, True, 4
            P_Line obj, 0, obj.CurrentY, 170, obj.CurrentY, vbDot
        Next
    End With
    P_RodapeVis obj, strNomeSis & " " & constVersao_Sis
End Sub
Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_Load()
    resizeForm Me
   AjustaGrid
End Sub
Sub Soma()
    Dim cTotal As Currency
    Dim i As Integer
    Dim j As Integer
    Dim cGeral As Currency
    Dim totalCol(15) As Currency
    
    With grid
        For i = 1 To .rows - 1
            cTotal = 0
            For j = 2 To .Cols - 1
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
        
        .col = .Cols - 1
        .Text = FormatVal(cGeral)
    End With
End Sub
Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
    posicionaADireita cmndbtnFechar, pcTop
End Sub

Private Sub txtAno_GotFocus()
    P_SelecTexto txtAno
End Sub

Private Sub txtAno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAno_LostFocus()
    txtAno.Text = Format(txtAno.Text, "000#")
End Sub
