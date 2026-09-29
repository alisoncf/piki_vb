VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelFatMensal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório Análise Mensal de Faturamento"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9570
   Icon            =   "frmRelFatMensal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame1 
         Caption         =   "Tipo:"
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2055
         Begin VB.OptionButton optnDia 
            Caption         =   "Diário"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton optnMens 
            Caption         =   "Mensal"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2055
         Begin VB.TextBox txtbxDAtaFIn 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1080
            TabIndex        =   4
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D&e:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   255
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlxGrd 
         Height          =   5535
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   9763
         _Version        =   393216
         AllowUserResizing=   2
         Appearance      =   0
      End
      Begin VB.Frame Frame2 
         Height          =   3975
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   2055
         Begin VB.CommandButton cmndbtnExpor 
            Caption         =   "&Exportar"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelFatMensal.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnGraf 
            Caption         =   "&Gráfico"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelFatMensal.frx":059E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2400
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnFechar 
            Cancel          =   -1  'True
            Caption         =   "&Sair"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelFatMensal.frx":06E8
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   3120
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnImprimir 
            Caption         =   "Im&primir"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelFatMensal.frx":0853
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnBuscar 
            Caption         =   "&Buscar"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelFatMensal.frx":0C60
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmRelFatMensal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sDataIni As String
    Dim sDataFin As String
    Dim sUltDia As String
Private Function F_Dev(mes As String) As Currency
    Dim rPesq As New adodb.Recordset
    Dim strSQL As String
    strSQL = " select sum(gerado_dev) as total "
    strSQL = strSQL & " from tab_vendas as tv, tab_devolucoes as td "
    strSQL = strSQL & " where tv.num_ven=td.num_ven "
    strSQL = strSQL & " and substring(dataefet_ven,1,7)='" & Format(mes, "yyyy-mm") & "' "
    
    'strSQL = strSQL & " and tv.situacao_ven>1 "
    abreObjetoAdo rPesq, strSQL
    F_Dev = rPesq!total
    rPesq.Close
End Function
Sub busca()
    Dim strSQL  As String
    Dim rPesq   As New adodb.Recordset
    Dim rDev    As New adodb.Recordset
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
    
    Dim sMes As String
    Dim iCount As Integer
    sDataIni = "01/" & txtbxDataIni.Text
    sUltDia = DateAdd("m", 1, "01/" & txtbxdataFin.Text)
    sDataFin = DateAdd("d", -1, sUltDia)
    
    strSQL = " select sum(tf.valor_ven) as valor,tp.nome_for,tf.cod_for , "
    If optnMens.Value Then
        strSQL = strSQL & "substring(dataefet_ven,1,7) as mes "
    Else
        strSQL = strSQL & "dataefet_ven as mes "
    End If
    
    strSQL = strSQL & " from tab_vendas as tv, tab_venfec as tf,tab_formaspag as tp  "
    strSQL = strSQL & " where tv.num_ven=tf.num_ven "
    strSQL = strSQL & " and tf.cod_for=tp.cod_for"
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(sDataIni)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(sDataFin)
    strSQL = strSQL & " and tv.situacao_ven=4"
    strSQL = strSQL & " group by tp.cod_for,mes order by mes "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        msflxgrd.rows = 1
        
        Do While Not .EOF
            cDev = 0
            If sMes <> !mes Then
                msflxgrd.AddItem !mes
                
            End If
            sMes = !mes
            For iCount = 1 To msflxgrd.Cols - 1
                msflxgrd.row = 0
                msflxgrd.Col = iCount
                If Left(msflxgrd.Text, 2) = !cod_for Then
                    msflxgrd.row = msflxgrd.rows - 1
                    msflxgrd.Text = FormatVal(!valor)
                End If
            Next
            
            
            
            
            
            
            
            .MoveNext
        Loop
        buscaDev
        ajustaTamGrid msflxgrd, Me
        ZebraGrid msflxgrd, 1
        soma
                
        .Close
    End With
    
  Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub soma()
    Dim cTotal As Currency
    Dim cTotalV As Currency
    Dim iCount As Integer
    Dim jCount As Integer
    With msflxgrd
        
        For iCount = 1 To .rows - 1
            .row = iCount
            cTotal = 0
            For jCount = 1 To .Cols - 1
                .Col = jCount
                If .Text = "" Then .Text = "0,00"
                If jCount = msflxgrd.Cols - 2 Then
                    cTotal = cTotal - CCur(.Text)
                Else
                    cTotal = cTotal + CCur(.Text)
                End If
            Next
            msflxgrd.Col = msflxgrd.Cols - 1
            msflxgrd.Text = FormatVal(cTotal)
        Next
    End With
    
    
    
    msflxgrd.AddItem "Total"
    
    
    With msflxgrd
        
        For iCount = 1 To .Cols - 1
            .Col = iCount
            cTotal = 0
            For jCount = 1 To .rows - 1
                .row = jCount
                If .Text = "" Then .Text = "0,00"
                
                cTotal = cTotal + CCur(.Text)
            Next
            msflxgrd.row = msflxgrd.rows - 1
            msflxgrd.Text = FormatVal(cTotal)
        Next
    End With
    
End Sub
Sub buscaDev()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim iCount As Integer
    Dim bAchou As Boolean
    strSQL = "select sum(gerado_dev) as total, "
    If optnMens.Value Then
        strSQL = strSQL & " substring(data_dev,1,7) as mes "
    Else
        strSQL = strSQL & " substring(data_dev,1,10) as mes "
    End If
    strSQL = strSQL & " from tab_Devolucoes "
    strSQL = strSQL & " where left(data_Dev,10)>=" & F_GravaData(CDate(sDataIni))
    strSQL = strSQL & " and left(data_dev,10)<=" & F_GravaData(sDataFin)
    strSQL = strSQL & " group by mes order by mes "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            For iCount = 1 To msflxgrd.rows - 1
                msflxgrd.Col = 0
                msflxgrd.row = iCount
                If CDate(msflxgrd.Text) = CDate(!mes) Then
                    msflxgrd.Col = msflxgrd.Cols - 2
                    msflxgrd.Text = FormatVal(!total)
                    bAchou = True
                End If
            Next
            If Not bAchou Then
                msflxgrd.AddItem !mes
                msflxgrd.Col = msflxgrd.Cols - 2
                    msflxgrd.Text = FormatVal(!total)
            End If
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub
Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnExpor_Click()
    P_MsflxToHtml Me.caption, msflxgrd
    
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
'    Dim iRow As Integer: Dim iCount As Integer
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
'            .label(iRow) = msflxgrd.Text
'            .ThisSet = iRow
'            msflxgrd.Col = iRow
'            msflxgrd.Row = 0
'            .Legend(iRow) = msflxgrd.Text
'            For iCount = 1 To .NumPoints
'                msflxgrd.Row = iCount
'                .Data(iCount) = msflxgrd.Text
'            Next
'        Next
'    .label(6) = "Total"
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
    P_CabRelTimbr "Análise Mensal de Faturamento por Venda de " & txtbxDataIni.Text & " até " & txtbxdataFin.Text, obj
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
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .Col = 0
            P_ImprX obj, .Text, 0, True, False, 4
            P_Line obj, 20, obj.CurrentY, 20, obj.CurrentY + 5, vbDot
            .Col = 1
            P_ImprX obj, .Text, 50, True, True, 4
            .Col = 2
            P_ImprX obj, .Text, 80, True, True, 4
            .Col = 3
            P_ImprX obj, .Text, 110, True, True, 4
            .Col = 4
            P_ImprX obj, .Text, 140, True, True, 4
            .Col = 5
            P_ImprX obj, .Text, 170, False, True, 4
            P_Line obj, 0, obj.CurrentY, 170, obj.CurrentY, vbDot
        Next
    End With
    P_RodapeVis obj, constNome_sis & " " & constVersao_Sis
End Sub
Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    With msflxgrd
        .Cols = 1
        .rows = 1
        .Col = 0
        .Text = "§"
        strSQL = "select * from tab_formaspag order by cod_for "
        abreObjetoAdo rPesq, strSQL
        With rPesq
            Do While Not .EOF
                msflxgrd.Cols = msflxgrd.Cols + 1
                msflxgrd.Col = msflxgrd.Cols - 1
                msflxgrd.Text = !cod_for & "-" & StrConv(!nome_for, vbProperCase)
                
                
                .MoveNext
            Loop
            msflxgrd.Cols = msflxgrd.Cols + 1
            msflxgrd.Col = msflxgrd.Cols - 1
            msflxgrd.Text = "Devoluções"
            msflxgrd.Cols = msflxgrd.Cols + 1
            msflxgrd.Col = msflxgrd.Cols - 1
            msflxgrd.Text = "Total"
            
            .Close
        End With
        
        
        
    End With
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
    P_AjustaData txtbxdataFin, "mm/yyyy"
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
    P_AjustaData txtbxDataIni, "mm/yyyy"
End Sub
