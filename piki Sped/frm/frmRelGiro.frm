VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelGiro 
   Caption         =   "Relatório de Giro Periódico"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   13650
   Icon            =   "frmRelGiro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   13650
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   13620
      TabIndex        =   0
      Top             =   0
      Width           =   13650
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4200
         TabIndex        =   30
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   3735
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   595
         Left            =   10920
         Picture         =   "frmRelGiro.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   12240
         Picture         =   "frmRelGiro.frx":0871
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   9600
         Picture         =   "frmRelGiro.frx":0DFF
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   615
         Left            =   8280
         Picture         =   "frmRelGiro.frx":12ED
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox cmbxMarca 
         Height          =   315
         Left            =   6120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1905
      End
      Begin VB.ComboBox cmbxCat 
         Height          =   315
         Left            =   4200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1905
      End
      Begin VB.ComboBox cmbxOrderBY 
         Height          =   315
         ItemData        =   "frmRelGiro.frx":163D
         Left            =   2160
         List            =   "frmRelGiro.frx":1650
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   345
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
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   315
         Left            =   3840
         Picture         =   "frmRelGiro.frx":168F
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton cmdProd 
         Height          =   315
         Left            =   7680
         Picture         =   "frmRelGiro.frx":1A19
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   960
         Width           =   315
      End
      Begin VB.Label lblCofFab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   4800
         TabIndex        =   33
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
         Height          =   195
         Left            =   4200
         TabIndex        =   32
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   195
         Left            =   6120
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   4200
         TabIndex        =   7
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordem:"
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Top             =   120
         Width           =   510
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
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13620
      TabIndex        =   14
      Top             =   6975
      Width           =   13650
      Begin VB.Frame frGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Agrupamento:"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8640
         TabIndex        =   22
         Top             =   0
         Width           =   4815
         Begin VB.OptionButton optnMar 
            Caption         =   "por &Marca"
            Height          =   255
            Left            =   2760
            TabIndex        =   25
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optnCat 
            Caption         =   "por &Categoria"
            Height          =   255
            Left            =   1320
            TabIndex        =   24
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optnPro 
            Caption         =   "por &Produto."
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
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
         TabIndex        =   20
         Top             =   45
         Width           =   390
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
         TabIndex        =   19
         Top             =   45
         Width           =   570
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   30
         Width           =   360
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
         TabIndex        =   16
         Top             =   30
         Width           =   630
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
         TabIndex        =   15
         Top             =   45
         Width           =   585
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmRelGiro.frx":1DA3
      Height          =   5055
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8916
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
Attribute VB_Name = "frmRelGiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bActive As Boolean
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1




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

Private Sub cmdExportar_Click()
    ExportaGridToCSV dtgrd, ado
End Sub

Private Sub cmdProd_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As Produto
    
    p = F_BuscaProdSaida(txtProduto.Text)
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
        lblCofFab.caption = p.codfab_pro
        
    End If
    
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnBuscar_Click()
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
    On Error GoTo trata_Erro
    If ado.RecordCount <= 0 Then
        MsgBox "Não há o quer ser impresso!", vbExclamation
        Exit Sub
    End If
    
    
    P_DtgrdToHtml "Relatório de Giro Periódico (" & txtbxDataIni.Text & " - " & txtbxDAtaFIn.Text & ")  Cliente: " & txtCliente.Text, dtgrd, ado, "", "3,4,5,6"
    
    
    Exit Sub
    
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
trata_Erro:
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

Private Sub cmndbtnLocaliza_Click()
    buscaCli
End Sub
Sub buscaCli()
    Dim c As Clientes.cliente
    
    c = Clientes.F_BuscaCliente(txtCliente.Text)
    If c.cod_cli <> "" Then
        txtCliente.Text = c.nome_Cli
        txtCliente.tag = c.cod_cli
        busca
    End If
    
End Sub

Private Sub dtgrd_DblClick()
    
    If optnPro.Value Then
        frmCadPro.tag = ado!cod
        frmCadPro.Show 1
    End If
End Sub

Private Sub Form_Activate()
    bative = True
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    cmbxOrderBY.ListIndex = 0
    
    P_PreencheCateg cmbxCat
    P_PreencheMarca cmbxMarca
    txtbxDAtaFIn.Text = Date
    txtbxDataIni.Text = Date
    resizeForm Me
    ajustaEstoque
End Sub
Sub soma()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim cTotalGer As Currency
    Dim iCountPro As Single
    bActive = True
    dtgrd.Visible = False
    On Error GoTo trata_Erro
    With ado
    
        Do While Not .EOF
           ' cTotal = cTotal + !Total
            cTotalGer = cTotalGer + !gerado
            cTotal = cTotal + !liquido
            iCountPro = iCountPro + !soma
            .MoveNext
            
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    
    
    lblTotalGerado.caption = FormatVal(cTotalGer)
    lblLiq.caption = FormatVal(cTotal)
    lblItens.caption = iCountPro
    
    bActive = False
    dtgrd.Visible = True
trata_Erro:
End Sub
Sub busca()
    Dim strSQL As String
    
    
    If optnPro.Value = True Then
        strSQL = "select tp.cod_pro as cod,tp.codfab_pro as codfab,tp.nome_pro as nome,nome_mar as marca,"
    ElseIf optnCat.Value = True Then
        strSQL = "select tt.nome_cat as nome,tt.cod_cat as cod,'' as codfab,"
    ElseIf optnMar.Value = True Then
        strSQL = "select tm.nome_mar as nome,tm.cod_mar as cod,'' as codfab,"
    End If
    
    strSQL = strSQL & "sum(ti.qtd_ite) as soma,"
    strSQL = strSQL & "sum(ti.devolvido_ite) as devolvido,"
    strSQL = strSQL & " sum((ti.final_ite/qtd_ite)*(qtd_ite-devolvido_ite)) as gerado, "
    strSQL = strSQL & " sum((ti.final_ite/ti.qtd_ite-tp.cr_pro)*(ti.qtd_ite-devolvido_ite)) as liquido, "
    strSQL = strSQL & " (er_pro-reservado_pro) as Estoque "
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on tp.cod_pro= ti.cod_pro "
    strSQL = strSQL & " left join tab_categorias as tt on tt.cod_cat=tp.cod_cat"
    strSQL = strSQL & " left join " & TAB_MARCAS & " as tm  on tp.cod_mar=tm.cod_mar "
     
    
    strSQL = strSQL & " where tv.situacao_ven=4"
    If txtCliente.tag <> "" Then
        strSQL = strSQL & " and tv.cod_cli='" & txtCliente.tag & "' "
    End If
    
    
        
    
    If cmbxCat.ListIndex > -1 Then
        strSQL = strSQL & " and tp.cod_cat='" & cmbxCat.tag & "' "
    End If
    If cmbxMarca.ListIndex > -1 Then
        strSQL = strSQL & " and tp.cod_mar='" & cmbxMarca.tag & "' "
    End If
    If txtProduto.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pro='" & txtProduto.tag & "' "
    End If
    strSQL = strSQL & " and tv.dataefet_ven >=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven <=" & F_GravaData(txtbxDAtaFIn.Text)
    If optnPro.Value = True Then
        strSQL = strSQL & " group by ti.cod_pro "
    End If
    If optnCat.Value = True Then
        strSQL = strSQL & " group by tp.cod_cat "
    End If
    If optnMar.Value = True Then
        strSQL = strSQL & " group by tp.cod_mar"
    End If
    strSQL = strSQL & " order by "
    If cmbxOrderBY.ListIndex = 0 Then
        strSQL = strSQL & "nome"
    ElseIf cmbxOrderBY.ListIndex = 1 Then
        strSQL = strSQL & "tp.codfab_Pro "
    ElseIf cmbxOrderBY.ListIndex = 2 Then
        strSQL = strSQL & "gerado desc "
    ElseIf cmbxOrderBY.ListIndex = 3 Then
        strSQL = strSQL & "liquido desc"
    Else
        strSQL = strSQL & "soma desc"
    End If
    abreObjetoAdo ado, strSQL, dtgrd
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade dtgrd, Me, Picture1
End Sub

Private Sub optnCat_Click()
    frGroup.caption = "Agrupar por Categoria"
End Sub

Private Sub optnMar_Click()
    frGroup.caption = "Agrupar por Marca"
End Sub

Private Sub optnPro_Click()
    frGroup.caption = "Agrupar por Produto"
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


Private Sub txtProduto_Change()
    txtProduto.tag = ""
End Sub

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdProd_Click
    End If
End Sub
