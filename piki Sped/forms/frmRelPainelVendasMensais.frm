VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelPainelVendasMensais 
   Caption         =   "Painel Mensal de Vendas"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12570
   Icon            =   "frmRelPainelVendasMensais.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   12570
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   5895
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   10398
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12540
      TabIndex        =   10
      Top             =   6975
      Width           =   12570
      Begin VB.Frame frGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Agrupamento:"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8640
         TabIndex        =   12
         Top             =   0
         Width           =   4815
         Begin VB.OptionButton optnPro 
            Caption         =   "por &Produto."
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optnCat 
            Caption         =   "por &Categoria"
            Height          =   255
            Left            =   1320
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optnMar 
            Caption         =   "por &Marca"
            Height          =   255
            Left            =   2760
            TabIndex        =   13
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdProd 
         Caption         =   "..."
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Líqüido:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         TabIndex        =   22
         Top             =   45
         Width           =   585
      End
      Begin VB.Label lblLiq 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.000,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2520
         TabIndex        =   21
         Top             =   30
         Width           =   630
      End
      Begin VB.Label lblItens 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4200
         TabIndex        =   20
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblTotalGerado 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.000,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   600
         TabIndex        =   19
         Top             =   30
         Width           =   630
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gerado:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   18
         Top             =   45
         Width           =   570
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Itens:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3720
         TabIndex        =   17
         Top             =   45
         Width           =   390
      End
      Begin VB.Label lblProduto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
         Height          =   195
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12540
      TabIndex        =   0
      Top             =   0
      Width           =   12570
      Begin VB.ComboBox cmbxOrderBY 
         Height          =   315
         ItemData        =   "frmRelPainelVendasMensais.frx":1CFA
         Left            =   2280
         List            =   "frmRelPainelVendasMensais.frx":1D0D
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   345
         Width           =   1935
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   615
         Left            =   7200
         Picture         =   "frmRelPainelVendasMensais.frx":1D4C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   8520
         Picture         =   "frmRelPainelVendasMensais.frx":209C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico"
         Enabled         =   0   'False
         Height          =   615
         Left            =   9840
         Picture         =   "frmRelPainelVendasMensais.frx":3D96
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   11160
         Picture         =   "frmRelPainelVendasMensais.frx":41D8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1080
         TabIndex        =   8
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordem:"
         Height          =   195
         Left            =   2280
         TabIndex        =   7
         Top             =   120
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmRelPainelVendasMensais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Sub AjustaGrid()
    Dim i As Integer
    Dim meses As Integer
    With grid
        meses = DateDiff("M", txtbxDataIni.Text, txtbxDAtaFIn.Text) + 2
        .Cols = meses + 2
        .FixedCols = 2
        .AllowUserResizing = flexResizeColumns
        .Row = 0
        .col = 0: .Text = "Cód.Fab."
        .col = 1: .Text = "Nome"
        .ColWidth(1) = 3000
        .col = .Cols - 1: .Text = "Total"
        For i = 2 To meses
            .col = i
            .ColWidth(i) = 1000
            .Text = Left(MonthName(Month(DateAdd("M", i - 2, txtbxDataIni.Text))), 3) & "/" & Year(DateAdd("M", i - 2, txtbxDataIni.Text))
        Next
        
        
    End With
End Sub







Private Sub cmbxCat_Click()
    cmbxCat.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
End Sub

Private Sub cmbxMarca_Click()
    cmbxMarca.tag = F_BuscaUmCampo("cod_mar", TAB_MARCAS, "nome_mar='" & cmbxMarca.Text & "'")
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub

Private Sub cmbxOrderBY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnBuscar_Click()
    AjustaGrid
    P_Status "Buscando..."
    busca
    soma
    P_Status ""
End Sub

Private Sub cmndbtnGraf_Click()
    ReDim b(max) As valChart
    ReDim row2(max) As valChart
    Dim iCount As Integer
    Dim cPrim As Currency
    Dim cSoma As Currency
    'Dim max As Integer
    Dim strSQL As String
    With dt.Recordset
        .MoveFirst
        If .RecordCount < 15 Then
            max = .RecordCount - 1
        Else
            max = 15
        End If
        ReDim b(max + 1) As valChart
        ReDim row2(max + 1) As valChart
        .MoveFirst
        For iCount = 0 To max
    
            b(iCount).caption = !nome & " "
            If cmbxOrderBY.ListIndex = 2 Then
                b(iCount).Value = !gerado
                cPrim = cPrim + !gerado
            ElseIf cmbxOrderBY.ListIndex = 3 Then
                b(iCount).Value = !liquido
                cPrim = cPrim + !liquido
            ElseIf cmbxOrderBY.ListIndex = 4 Then
                b(iCount).Value = !soma
                cPrim = cPrim + !soma
            End If
            .MoveNext
        Next
        row2(0).caption = "15 Primeiros"
        row2(0).Value = cPrim
        
        Do While Not .EOF
            If cmbxOrderBY.ListIndex = 2 Then
                cSoma = cSoma + !gerado
            ElseIf cmbxOrderBY.ListIndex = 3 Then
                cSoma = cSoma + !liquido
            ElseIf cmbxOrderBY.ListIndex = 4 Then
                cSoma = cSoma + !soma
            End If
            
            .MoveNext
        Loop
        
        ' b.val(max).caption = "Restante"
       ' b.val(max).Value = cSoma
       ' row2.val(1).caption = "Restante"
       ' row2.val(1).Value = cSoma
    End With
    If cmbxOrderBY.ListIndex = 2 Then
    
    ElseIf cmbxOrderBY.ListIndex = 3 Then
    
    ElseIf cmbxOrderBY.ListIndex = 4 Then
    
    End If
    P_Graph b, Int(max), 1, True, "Gráfico de Giro de Produtos(" & cmbxOrderBY.Text & ")", "Giro Periódico", True
    
    
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_erro
    If dt.Recordset.RecordCount <= 0 Then
        MsgBox "Não há o quer ser impresso!", vbExclamation
        Exit Sub
    End If
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag = "" Then
        Unload frmVisualiza
        Exit Sub
    End If
    'If MsgBox("Confirma Impressão do Relatório?", vbYesNo + vbQuestion) = vbYes Then
        imprime Printer
    'End If
    Exit Sub
trata_erro:
    MostraErro Me.Name
End Sub
Sub imprime(obj As Object)
    Dim iItens As Integer
    Dim iPag As Integer
    Dim strSQL As String
    Dim cTotal As Currency
    Dim cQPage As Integer
    cQPage = 60
    P_CabRelTimbr "Relatório de Giro Periódico " & de & txtbxDataIni.Text & _
        " Até " & txtbxDAtaFIn.Text, obj
    obj.FontName = "arial"
    obj.fontsize = 9
    P_ImprX obj, "Cód.:", 0
    P_ImprX obj, "Descrição", 20
    P_ImprX obj, "UN", 120
    P_ImprX obj, "QTD", 140, True, True
    P_ImprX obj, "Ger.", 160, True, True
    P_ImprX obj, "Líquido", 175, True, True
    P_ImprX obj, "Estoque", 192, True, True
    P_ImprX obj, "Expr(%)", 205, False, True
    
    P_ImprLinha obj
    
    With dt.Recordset
        
        .MoveFirst
        strSQL = CStr(.RecordCount / cQPage)
        If InStr(1, strSQL, ",") <> 0 Then
            iPag = Int(strSQL) + 1
        End If
        
        Do While Not .EOF
            iItens = iItens + 1
            If iItens = cQPage Then
                If TypeName(obj) = "Printer" Then
                    P_Rodape "Continua...", True, "Pág:" & IIf(TypeName(obj) = "Printer", obj.Page, "%") & " de " & iPag
                
                End If
                If obj.Name = "Pic" Then
                    frmVisualiza.imglst.ListImages.Add , , frmVisualiza.Pic.image
                    frmVisualiza.Pic.Cls
                   
                    frmVisualiza.Pic.Picture = Nothing
                    obj.CurrentY = yIni
                End If
                iItens = 0
                P_ImprX obj, "Cód.:", 0
                P_ImprX obj, "Descrição", 20
                P_ImprX obj, "UN", 120
                P_ImprX obj, "QTD", 140, True, True
                P_ImprX obj, "Ger.", 160, True, True
                P_ImprX obj, "Líquido", 175, True, True
                P_ImprX obj, "Estoque", 192, True, True
                P_ImprX obj, "Expr(%)", 205, False, True
                P_ImprLinha obj
            End If
            obj.FontName = "arial"
            obj.fontsize = 9
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome, 20
            P_ImprX obj, !unidade_pro, 120
            P_ImprX obj, FormatVal(!soma), 140, True, True
            P_ImprX obj, FormatVal(!gerado), 160, True, True
            P_ImprX obj, FormatVal(!liquido), 175, True, True
            P_ImprX obj, FormatVal(!er_pro), 192, True, True
            P_ImprX obj, FormatVal(!gerado * 100 / CCur(lblTotalGerado.caption)) & "%", 205, False, True
            cTotal = cTotal + !gerado * 100 / CCur(lblTotalGerado.caption)
           
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "Total...: " & lblTotalGerado.caption, 80, True
        'P_ImprX obj, lblTotal.caption, 130, False
        On Error Resume Next
        P_RodapeVis obj, strNome_Emp, "Pág: " & obj.Page & " de " & obj.Page
        If obj.Name = "Pic" Then
                    frmVisualiza.imglst.ListImages.Add , , frmVisualiza.Pic.image
                    frmVisualiza.Pic.Cls
                   
                    frmVisualiza.Pic.Picture = Nothing
                    obj.CurrentY = yIni
                End If
    End With
    
    
    obj.EndDoc
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub dt_Reposition()
    If bActive = True Then Exit Sub
    If dt.Recordset.RecordCount > 0 Then
        If cmbxOrderBY.ListIndex >= 2 Then
            cmndbtnGraf.Enabled = True
        Else
            cmndbtnGraf.Enabled = False
        End If
    Else
        cmndbtnGraf.Enabled = False
    End If
    BuscaEnt
    bActive = True
End Sub
Sub BuscaEnt()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error Resume Next
    strSQL = "select ti.*,tc.entrada_not from tab_cmp as tc, tab_itcmp as ti "
    strSQL = strSQL & " where tc.num_not=ti.num_not "
    strSQL = strSQL & " and tc.serie_not=ti.serie_not "
    strSQL = strSQL & " and tc.cod_for=ti.cod_for "
    strSQL = strSQL & " and ti.cod_pro='" & dt.Recordset!cod_pro & "' "
    strSQL = strSQL & " order by entrada_not desc "
    rPesq.Open strSQL, conn_client
    If rPesq.RecordCount > 0 Then
        lblUltEnt.caption = rPesq!entrada_not
        lblUltEnt.caption = lblUltEnt.caption & "-Qdt:" & FormatVal(rPesq!qtd_ite)
    Else
        lblUltEnt.caption = "Ult. Entr: Indisponível"
    End If
    rPesq.Close
End Sub

Private Sub Form_Activate()
    bative = True
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    cmbxOrderBY.ListIndex = 0
    
    
    txtbxDataIni.Text = "01/01/" & Year(Date)
    txtbxDAtaFIn.Text = Date
    
End Sub
Sub soma()
    
    
    
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTime As String
    Dim sMes As String
    Dim sCod As String
    Dim iCol As Integer
    sTime = Now
    
    
    
    If optnPro.Value = True Then
        strSQL = "select tp.nome_pro as nome,tp.cod_pro as cod,tp.codfab_pro as codfab,"
    ElseIf optnCat.Value = True Then
        strSQL = "select tt.nome_cat as nome,tt.cod_cat as cod,tt.cod_cat as codfab"
    ElseIf optnMar.Value = True Then
        strSQL = "select tm.nome_mar as nome,tm.cod_mar as cod,tm.cod_mar as codfab"
    
    End If
    strSQL = strSQL & " substring(tv.dataefet_ven,1,7) as mes,"
    
    
    
    
    strSQL = strSQL & "sum(ti.qtd_ite) as soma,"
    
    strSQL = strSQL & " sum(ti.final_ite) as gerado, "
    strSQL = strSQL & " sum((ti.final_ite/ti.qtd_ite-tp.cr_pro)*ti.qtd_ite) as liquido "
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.nuM_ven "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro  "
    strSQL = strSQL & " left join tab_categorias as tt on tp.cod_cat=tt.cod_cat left join " & TAB_MARCAS & "  as tm on tm.cod_mar=tp.cod_mar"
    strSQL = strSQL & " where tv.situacao_ven = 4 "
    
    strSQL = strSQL & " and dataefet_ven >=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and dataefet_ven <=" & F_GravaData(txtbxDAtaFIn.Text)
    
    If optnPro.Value = True Then
        strSQL = strSQL & " group by ti.cod_pro "
    ElseIf optnCat.Value = True Then
        strSQL = strSQL & " group by tp.cod_cat "
    ElseIf optnMar.Value = True Then
        strSQL = strSQL & " group by tp.cod_mar "
    End If
    strSQL = strSQL & ",mes order by nome,mes "
    
    
    
    abreObjetoAdo rPesq, strSQL
    grid.rows = 1
    With rPesq
        Do While Not .EOF
                 
            If sCod <> !cod Then
                grid.AddItem !codfab & Chr(9) & !nome
            End If
            sMes = !mes
            iCol = BuscaColunaAnoMes(sMes)
            If iCol > 0 Then
                grid.col = iCol
                grid.Row = grid.rows - 1
                grid.Text = !soma
            End If
            
            .MoveNext
        Loop
    End With
    soma
    
End Sub



Function BuscaColunaAnoMes(AnoMes As String) As Integer
    Dim i As Integer
    With grid
        .Row = 0
        For i = 2 To .Cols - 2
            grid.col = i
            If Left(AnoMes, 4) = Right(.Text, 4) Then
                If Right(AnoMes, 2) = RetornaMes(.Text) Then
                    BuscaColunaAnoMes = i
                
                    Exit Function
                End If
            End If
        
        Next
    End With
    
    
End Function
Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade grid, Me, Picture1
End Sub

Private Sub optnCat_Click()
    frGroup.caption = "Agrupar por Categoria"
End Sub

Private Sub optnMar_Click()
    frGroup.caption = "Agrupar por Marca"
End Sub

Private Sub optnPro_Click()
    frGroup.caption = "Agrupar por Produto"
    BuscaProd
End Sub
Sub BuscaProd()
    Dim p As Produto
    p = Produtos.F_BuscaProdSaida()
    If p.cod_pro <> "" Then
        lblProduto.caption = p.codfab_pro & " - " & p.nome_pro
        lblProduto.tag = p.cod_pro
    End If
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
    txtbxDAtaFIn.Text = F_UltDiaMes(txtbxDAtaFIn.Text)
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
    txtbxDataIni.Text = F_PrimDiaMes(txtbxDataIni.Text)
End Sub




