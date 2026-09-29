VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelComportPro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Comportamento de Produtos"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12615
   Icon            =   "frmRelComportProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnGraf 
      Caption         =   "&Gráfico"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6840
      Picture         =   "frmRelComportProd.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   735
      Left            =   8760
      Picture         =   "frmRelComportProd.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   10680
      Picture         =   "frmRelComportProd.frx":0999
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.ComboBox cmbxComport 
         Height          =   315
         ItemData        =   "frmRelComportProd.frx":0B04
         Left            =   8040
         List            =   "frmRelComportProd.frx":0B0E
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   11280
         Picture         =   "frmRelComportProd.frx":0B22
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbxOrderBY 
         Height          =   315
         ItemData        =   "frmRelComportProd.frx":0EAC
         Left            =   9480
         List            =   "frmRelComportProd.frx":0EB9
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1545
      End
      Begin VB.TextBox txtbxNome_pro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         TabIndex        =   6
         Top             =   360
         Width           =   3945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   3960
         TabIndex        =   21
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Comportamento:"
         Height          =   195
         Left            =   8040
         TabIndex        =   10
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   9480
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
      Begin VB.Label lblBuscarPor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   2280
         TabIndex        =   5
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   6465
      TabIndex        =   14
      Top             =   6720
      Width           =   6495
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5520
         TabIndex        =   26
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Itens"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4080
         TabIndex        =   25
         Top             =   0
         Width           =   345
      End
      Begin VB.Label lblQtd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5640
         TabIndex        =   20
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Líqüido:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2400
         TabIndex        =   19
         Top             =   0
         Width           =   585
      End
      Begin VB.Label lblLiq 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2400
         TabIndex        =   18
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblItens 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   90
      End
      Begin VB.Label lblTotalGerado 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gerado:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   0
         Width           =   570
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelComportProd.frx":0EE7
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmRelComportProd.frx":0EF8
      TabIndex        =   13
      Top             =   960
      Width           =   12375
   End
End
Attribute VB_Name = "frmRelComportPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmbxCat_Click()
    If optnCat.Value = True Then
        cmbxCat.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
    Else
        cmbxCat.tag = F_BuscaUmCampo("cod_MAR", "tab_MARCAS", "nome_MAR='" & cmbxCat.Text & "'")
    End If
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
End Sub

Private Sub cmbxMarca_Click()
    cmbxMarca.tag = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxMarca.Text & "'")
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub

Private Sub cmbxOrderBY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    P_Status "Buscando..."
    busca
    soma
    P_Status ""
End Sub

Private Sub cmndbtnImprimir_Click()
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
End Sub
Sub imprime(obj As Object)
    Dim iItens As Integer
    Dim iPag As Integer
    Dim strSQL As String
    Dim cTotal As Currency
    P_CabRelTimbr "Relatório de Giro Periódico " & de & txtbxDataIni.Text & _
        " Até " & txtbxDAtaFIn.Text, obj
    obj.FontName = "arial"
    obj.fontsize = 9
    P_ImprX obj, "Código", 0
    P_ImprX obj, "Descrição", 20
    P_ImprX obj, "UN", 120
    P_ImprX obj, "QTD", 140, True, True
    P_ImprX obj, "Ger.", 160, True, True
    P_ImprX obj, "Líquido", 180, True, True
    P_ImprX obj, "Estoque", 195, True, True
    P_ImprX obj, "Expr(%)", 210, False, True
    
    P_ImprLinha obj
    
    With dt.Recordset
        
        .MoveFirst
        strSQL = CStr(.RecordCount / 70)
        If InStr(1, strSQL, ",") <> 0 Then
            iPag = Int(strSQL) + 1
        End If
        
        Do While Not .EOF
            iItens = iItens + 1
            If iItens = 70 Then
                If TypeName(obj) = "Printer" Then
                    P_Rodape "Continua...", True, "Pág:" & IIf(TypeName(obj) = "Printer", obj.Page, "%") & " de " & iPag
                End If
                iItens = 0
                P_ImprX obj, "Código", 0
                P_ImprX obj, "Descrição", 20
                P_ImprX obj, "UN", 120
                P_ImprX obj, "QTD", 140, True, True
                P_ImprX obj, "Ger.", 160, True, True
                P_ImprX obj, "Líquido", 180, True, True
                P_ImprX obj, "Estoque", 195, True, True
                P_ImprX obj, "Expr(%)", 210, False, True
                P_ImprLinha obj
            End If
            obj.FontName = "arial"
            obj.fontsize = 9
            P_ImprX obj, !cod, 0
            P_ImprX obj, !Nome, 20
            P_ImprX obj, !unidade_pro, 120
            P_ImprX obj, FormatVal(!soma), 140, True, True
            P_ImprX obj, FormatVal(!gerado), 160, True, True
            P_ImprX obj, FormatVal(!liquido), 180, True, True
            P_ImprX obj, FormatVal(!e2_pro), 195, True, True
            P_ImprX obj, FormatVal(!gerado * 100 / CCur(lblTotalGerado.caption)) & "%", 210, False, True
            cTotal = cTotal + !gerado * 100 / CCur(lblTotalGerado.caption)
           
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "Total...: " & lblTotalGerado.caption, 80, True
        'P_ImprX obj, lblTotal.caption, 130, False
        On Error Resume Next
        P_RodapeVis obj, strNome_Emp, "Pág: " & obj.Page & " de " & obj.Page
    End With
    
    
    obj.EndDoc
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub cmndbtnGraf_Click()
    ReDim b(max) As valChart
    Dim iCount As Integer
    Dim cSoma As Currency
    
    Dim n As Integer
    Dim strSQL As String
    With dt.Recordset
        If .RecordCount > 60 Then
            max = 60
        Else
            max = .RecordCount - 1
        End If
        ReDim b(max) As valChart
        .MoveFirst
        For iCount = 0 To max
            If cmbxComport.ListIndex = 0 Then
                b(iCount).caption = WeekdayName(Weekday(!mes)) & "-" & !mes
            Else
                b(iCount).caption = !mes & " "
            End If
            If cmbxOrderBY.ListIndex = 0 Then
                b(iCount).Value = !gerado
                        
            ElseIf cmbxOrderBY.ListIndex = 1 Then
                b(iCount).Value = !liquido
            ElseIf cmbxOrderBY.ListIndex = 2 Then
                b(iCount).Value = !soma
            End If
            .MoveNext
        Next
        
        Do While Not .EOF
            cSoma = cSoma + !liquido
            .MoveNext
        Loop
        'b.count = max
        
    
    End With
    If cmbxOrderBY.ListIndex = 2 Then
    
    ElseIf cmbxOrderBY.ListIndex = 3 Then
    
    ElseIf cmbxOrderBY.ListIndex = 4 Then
    
    End If
    P_Graph b, Int(max), 6, True  ', LINES, "Gráfico de Giro de Produtos(" & cmbxOrderBy.Text & ")"
    
End Sub

    


Private Sub dt_Validate(Action As Integer, Save As Integer)
    If dt.Recordset.RecordCount > 0 Then
        'If cmbxOrderBY.ListIndex >= 2 Then
            cmndbtnGraf.Enabled = True
        'Else
         '   cmndbtnGraf.Enabled = False
        'End If
    Else
        cmndbtnGraf.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    dt.Connect = Banco.Connect
    
    cmbxOrderBY.ListIndex = 0
    P_CentraForm Me
    'P_PreencheCateg cmbxCat
    'P_PreencheMarca cmbxMarca
    optnPro_Click
    cmbxComport.ListIndex = 0
    txtbxDAtaFIn.Text = Date
    txtbxDataIni.Text = Date
End Sub
Sub soma()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim cTotalGer As Currency
    Dim cSoma As Currency
    bActive = True
    With dt.Recordset
    
        Do While Not .EOF
           ' cTotal = cTotal + !Total
            cTotalGer = cTotalGer + !gerado
            cTotal = cTotal + !liquido
            cSoma = cSoma + !soma
            .MoveNext
            
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    
    
    lblTotalGerado.caption = FormatVal(cTotalGer)
    lblLiq.caption = FormatVal(cTotal)
    lblItens.caption = "Itens: " & dt.Recordset.RecordCount
    lblQtd.caption = "Qtd: " & cSoma
    bActive = False
    
End Sub
Sub busca()
    Dim strSQL As String
    
    Dim sTime As String
    sTime = Now
    
    If optnPro.Value = True Then
        strSQL = "select tp.nome_pro as nome,tp.cod_pro as cod,"
    ElseIf optnCat.Value = True Then
        strSQL = "select ts.nome_sec as nome,ts.cod_sec as cod,"
    ElseIf optnMar.Value = True Then
        strSQL = "select tu.nome_sub as nome,tu.cod_sub as cod,"
    
    End If
    If cmbxComport.ListIndex = 0 Then
        strSQL = strSQL & " substring(tv.dataefet_ven,1,11) as mes,"
    Else
        strSQL = strSQL & " substring(tv.dataefet_ven,1,7) as mes,"
    End If
    
    
    
    strSQL = strSQL & "tp.*,ti.*,sum(ti.qtd_ite) as soma,"
    
    strSQL = strSQL & " sum(ti.final_ite) as gerado, "
    strSQL = strSQL & " sum((ti.final_ite/ti.qtd_ite-tp.cr_pro)*ti.qtd_ite) as liquido "
    strSQL = strSQL & " from tab_vendas as tv, tab_itvendas as ti,"
    strSQL = strSQL & " tab_produtos as tp, "
    strSQL = strSQL & " tab_secoes as ts,tab_subsec as tu "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tp.cod_sub=tu.cod_sub and ts.cod_sec=tu.cod_sec  "
    strSQL = strSQL & " and ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and tv.situacao_ven = 4 "
    If optnCat.Value = True Then
        strSQL = strSQL & " and tp.cod_Cat='" & cmbxCat.tag & "' "
    End If
    If optnMar.Value = True Then
        strSQL = strSQL & " and tp.cod_mar='" & cmbxCat.tag & "' "
    End If
    If optnPro.Value = True Then
        strSQL = strSQL & " and tp.cod_pro='" & txtbxNome_pro.tag & "' "
    End If
    strSQL = strSQL & " and dataefet_ven >=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and dataefet_ven <=" & F_GravaData(txtbxDAtaFIn.Text)
    If optnPro.Value = True Then
        strSQL = strSQL & " group by ti.cod_pro "
    End If
    If optnCat.Value = True Then
        strSQL = strSQL & " group by tp.cod_cat "
    End If
    If optnMar.Value = True Then
        strSQL = strSQL & " group by tp.cod_mar "
    End If
    strSQL = strSQL & ",mes order by mes "
    If cmbxOrderBY.ListIndex = 0 Then
        
    ElseIf cmbxOrderBY.ListIndex = 1 Then
        
    ElseIf cmbxOrderBY.ListIndex = 2 Then
        
    ElseIf cmbxOrderBY.ListIndex = 3 Then
        
    Else
        
    End If
    
    
    
    dt.RecordSource = strSQL
    dt.Refresh
    MsgBox "Busca concluída em " & DateDiff("s", sTime, Now) & " segundos..."
End Sub

Private Sub optnCat_Click()
    lblCat.caption = "&Categoria:"
    'P_PreencheCateg cmbxCat
    cmbxCat.Visible = True
    txtbxNome_pro.Visible = False
End Sub

Private Sub optnMar_Click()
    lblCat.caption = "&Marca:"
    
    cmbxCat.Visible = True
    txtbxNome_pro.Visible = False
End Sub

Private Sub optnPro_Click()
    lblCat.caption = "&Nome:"
    
    cmbxCat.Visible = False
    txtbxNome_pro.Visible = True
    ' txtbxNome_pro.SetFocus
End Sub

Private Sub txtbxDatafin_GotFocus()
    P_SelecTexto txtbxDAtaFIn
End Sub

Private Sub txtbxDatafin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDatafin_LostFocus()
    P_AjustaData txtbxDAtaFIn
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub




Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
    End If
End Sub
Sub buscaProd()
    
    Dim p As ProdutoSaida
    p = F_BuscaProdSaida(txtbxNome_pro.Text)
    txtbxNome_pro.Text = p.nome_pro
    txtbxNome_pro.tag = p.cod_pro
End Sub
