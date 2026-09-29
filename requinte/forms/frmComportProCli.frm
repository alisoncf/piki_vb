VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmComportProCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comportamento de Produtos por Clientes"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   Icon            =   "frmComportProCli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10140
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   2370
      TabIndex        =   18
      Top             =   0
      Width           =   7695
      Begin VB.Label lblNomePro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   2175
      Begin VB.Frame Frame3 
         Caption         =   "Buscar Por:"
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1935
         Begin VB.OptionButton optnRef 
            Caption         =   "&Referência"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton optnNome 
            Caption         =   "&Nome "
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame Fr2 
         Height          =   1935
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
         Begin VB.TextBox txtbxNome_pro 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1665
         End
         Begin VB.ComboBox cmbxOrderBY 
            Height          =   315
            ItemData        =   "frmComportProCli.frx":014A
            Left            =   120
            List            =   "frmComportProCli.frx":0157
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox cmbxComport 
            Height          =   315
            ItemData        =   "frmComportProCli.frx":0185
            Left            =   120
            List            =   "frmComportProCli.frx":018F
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lblCat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Produto"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Ordenar:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&Comportamento:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1155
         End
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmComportProCli.frx":01A3
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico"
         Enabled         =   0   'False
         Height          =   735
         Left            =   1080
         Picture         =   "frmComportProCli.frx":04AD
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmComportProCli.frx":08EF
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4560
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmComportProCli.frx":0BF9
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         Width           =   765
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmComportProCli.frx":14C3
      Height          =   5535
      Left            =   2370
      OleObjectBlob   =   "frmComportProCli.frx":14D4
      TabIndex        =   16
      Top             =   720
      Width           =   7695
   End
End
Attribute VB_Name = "frmComportProCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean


Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub

Private Sub cmbxOrderBY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        " Até " & txtbxdataFin.Text, obj
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
    Dim b As BlChart
    Dim iCount As Integer
    Dim cSoma As Currency
    Dim max As Integer
    Dim strSQL As String
    With dt.Recordset
        If .RecordCount > 60 Then
            max = 60
        Else
            max = .RecordCount - 1
        End If
        .MoveFirst
        For iCount = 0 To max
            If cmbxComport.ListIndex = 0 Then
                b.val(iCount).caption = WeekdayName(Weekday(!mes)) & "-" & !mes
            Else
                b.val(iCount).caption = !mes & " "
            End If
            If cmbxOrderBY.ListIndex = 0 Then
                b.val(iCount).Value = !gerado
                        
            ElseIf cmbxOrderBY.ListIndex = 1 Then
                b.val(iCount).Value = !liquido
            ElseIf cmbxOrderBY.ListIndex = 2 Then
                b.val(iCount).Value = !soma
            End If
            .MoveNext
        Next
        
        Do While Not .EOF
            cSoma = cSoma + !liquido
            .MoveNext
        Loop
        b.Count = max
        
    
    End With
    If cmbxOrderBY.ListIndex = 2 Then
    
    ElseIf cmbxOrderBY.ListIndex = 3 Then
    
    ElseIf cmbxOrderBY.ListIndex = 4 Then
    
    End If
    'P_Graph b, 6, False   ', LINES, "Gráfico de Giro de Produtos(" & cmbxOrderBy.Text & ")"
    
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
    txtbxdataFin.Text = Date
    txtbxDataIni.Text = Date
End Sub
Sub soma()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim cTotalGer As Currency
    bActive = True
    With dt.Recordset
    
        Do While Not .EOF
           ' cTotal = cTotal + !Total
            cTotalGer = cTotalGer + !gerado
            cTotal = cTotal + !liquido
            .MoveNext
            
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    
    
    
    
    
    
    bActive = False
    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tp.nome_pro as nome,tp.cod_pro as cod,tl.nome_cli, "
    
    If cmbxComport.ListIndex = 0 Then
        strSQL = strSQL & " substring(tv.data_ven,1,11) as mes,"
    Else
        strSQL = strSQL & " substring(tv.data_ven,1,7) as mes,"
    End If
    
    strSQL = strSQL & "tp.*,ti.*,(ti.qtd_ite) as soma,"
    strSQL = strSQL & " (ti.final_ite) as gerado, "
    strSQL = strSQL & " ((ti.final_ite/ti.qtd_ite-tp.cr_pro)*ti.qtd_ite) as liquido "
    strSQL = strSQL & " from tab_vendas as tv, tab_itvendas as ti,"
    strSQL = strSQL & " tab_produtos as tp, "
    strSQL = strSQL & " tab_categorias as tt,tab_marcas as tm, tab_clientes as tl "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tp.cod_mar=tm.cod_mar and tp.cod_Cat=tt.cod_cat "
    strSQL = strSQL & " and ti.num_ven=tv.num_ven and tl.cod_cli = tv.cod_cli "
    strSQL = strSQL & " and tv.situacao_ven = 4 "
    strSQL = strSQL & " and tp.cod_pro='" & txtbxNome_Pro.tag & "' "
    strSQL = strSQL & " and tv.dataefet_ven >=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven <=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " order by mes "
    
    dt.RecordSource = strSQL
    dt.Refresh
End Sub

Private Sub optnPro_Click()
    lblCat.caption = "&Nome:"
    
    
    txtbxNome_Pro.Visible = True
    ' txtbxNome_pro.SetFocus
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




Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
        lblNomePro.caption = txtbxNome_Pro.Text
    End If
End Sub
Sub buscaProd()
    Dim p As Produto
    If optnNome.Value Then
        p = F_BuscaProdSaida(txtbxNome_Pro.Text)
    Else
        p = F_BuscaProdSaida(txtbxNome_Pro.Text, "R")
    End If
    txtbxNome_Pro.Text = p.nome_pro
    txtbxNome_Pro.tag = p.cod_pro
End Sub


