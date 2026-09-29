VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmComport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Comportamento de Produtos"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   Icon            =   "frmComport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   10035
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.ComboBox cmbxComport 
         Height          =   315
         ItemData        =   "frmComport.frx":0442
         Left            =   120
         List            =   "frmComport.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmComport.frx":0460
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5400
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmComport.frx":0D2A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
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
      Begin VB.Frame frGroup 
         Caption         =   "Buscar Por:"
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1935
         Begin VB.OptionButton optnMar 
            Caption         =   "por &Marca"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton optnCat 
            Caption         =   "por &Categoria"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optnPro 
            Caption         =   "por &Produto."
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico"
         Enabled         =   0   'False
         Height          =   735
         Left            =   1080
         Picture         =   "frmComport.frx":1034
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmComport.frx":1476
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3720
         Width           =   975
      End
      Begin VB.ComboBox cmbxOrderBY 
         Height          =   315
         ItemData        =   "frmComport.frx":1780
         Left            =   120
         List            =   "frmComport.frx":178D
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2640
         Width           =   1935
      End
      Begin VB.ComboBox cmbxCat 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2040
         Width           =   1905
      End
      Begin VB.TextBox txtbxNome_pro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1905
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Comportamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1080
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
      Begin VB.Label lblCat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      ScaleHeight     =   345
      ScaleWidth      =   7665
      TabIndex        =   20
      Top             =   5880
      Width           =   7695
      Begin VB.Label lblQtd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd: 0"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5640
         TabIndex        =   26
         Top             =   0
         Width           =   435
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
         TabIndex        =   25
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
         Left            =   3000
         TabIndex        =   24
         Top             =   0
         Width           =   315
      End
      Begin VB.Label lblItens 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Itens: 0"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4320
         TabIndex        =   23
         Top             =   0
         Width           =   525
      End
      Begin VB.Label lblTotalGerado 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   960
         TabIndex        =   22
         Top             =   0
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
         Left            =   360
         TabIndex        =   21
         Top             =   0
         Width           =   570
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   3360
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
   Begin MSDataGridLib.DataGrid dbgrd 
      Bindings        =   "frmComport.frx":17BB
      Height          =   5655
      Left            =   2280
      TabIndex        =   27
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9975
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
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4394,835
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmComport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmbxCat_Click()
    If optnCat.Value = True Then
        cmbxCat.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
    Else
        cmbxCat.tag = F_BuscaUmCampo("cod_MAR", TAB_MARCAS, "nome_MAR='" & cmbxCat.Text & "'")
    End If
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
            P_ImprX obj, !nome, 20
            P_ImprX obj, !unidade_pro, 120
            P_ImprX obj, FormatVal(!soma), 140, True, True
            P_ImprX obj, FormatVal(!gerado), 160, True, True
            P_ImprX obj, FormatVal(!liquido), 180, True, True
            P_ImprX obj, FormatVal(!er_pro), 195, True, True
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
        'Else'
         '   cmndbtnGraf.Enabled = False
        'End If
    Else
        cmndbtnGraf.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    'dt_Validate.ConnectionString = strConnect
    
    cmbxOrderBY.ListIndex = 0
 '   P_CentraForm Me
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
        strSQL = "select tt.nome_cat as nome,tt.cod_cat as cod,"
    ElseIf optnMar.Value = True Then
        strSQL = "select tm.nome_mar as nome,tm.cod_mar as cod,"
    
    End If
    If cmbxComport.ListIndex = 0 Then
        strSQL = strSQL & " substring(tv.data_ven,1,11) as mes,"
    Else
        strSQL = strSQL & " substring(tv.data_ven,1,7) as mes,"
    End If
    
    
    
    strSQL = strSQL & "tp.*,ti.*,sum(ti.qtd_ite) as soma,"
    
    strSQL = strSQL & " sum(ti.final_ite) as gerado, "
    strSQL = strSQL & " sum((ti.final_ite/ti.qtd_ite-tp.cr_pro)*ti.qtd_ite) as liquido "
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti,"
    strSQL = strSQL & " " & TAB_PRODUTOS & " as tp, "
    strSQL = strSQL & " tab_categorias as tt," & TAB_MARCAS & "  as tm "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tp.cod_mar=tm.cod_mar and tp.cod_Cat=tt.cod_cat "
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
    P_PreencheCateg cmbxCat
    cmbxCat.Visible = True
    txtbxNome_pro.Visible = False
End Sub

Private Sub optnMar_Click()
    lblCat.caption = "&Marca:"
    P_PreencheMarca cmbxCat
    cmbxCat.Visible = True
    txtbxNome_pro.Visible = False
End Sub

Private Sub optnPro_Click()
    lblCat.caption = "&Nome:"
    
    cmbxCat.Visible = False
    txtbxNome_pro.Visible = True
    ' txtbxNome_pro.SetFocus
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




Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
    End If
End Sub
Sub buscaProd()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxNome_pro.Text)
    txtbxNome_pro.Text = p.nome_pro
    txtbxNome_pro.tag = p.cod_pro
End Sub
