VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelProdFor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Produtos por Fornecedor"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13560
   Icon            =   "frmRelProdFor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   13560
   Begin VB.PictureBox Picture3 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7740
      Left            =   0
      ScaleHeight     =   7710
      ScaleWidth      =   2385
      TabIndex        =   11
      Top             =   735
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "&Detalhes"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdFor.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3960
         Width           =   2175
      End
      Begin VB.ComboBox cmbxFunc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdFor.frx":09B4
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdFor.frx":0D3E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdFor.frx":114B
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6840
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico de Vendas"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdFor.frx":1240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   20
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13530
      TabIndex        =   1
      Top             =   0
      Width           =   13560
      Begin VB.TextBox txtbxFab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         MaxLength       =   40
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         MaxLength       =   40
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmbxSec 
         Height          =   315
         ItemData        =   "frmRelProdFor.frx":138A
         Left            =   120
         List            =   "frmRelProdFor.frx":138C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmbxSub 
         Height          =   315
         ItemData        =   "frmRelProdFor.frx":138E
         Left            =   2880
         List            =   "frmRelProdFor.frx":1390
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmndbtnBuscaProd 
         Height          =   285
         Left            =   10440
         Picture         =   "frmRelProdFor.frx":1392
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante/Marca:"
         Height          =   195
         Left            =   5640
         TabIndex        =   10
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   7560
         TabIndex        =   9
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblSecCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seção:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   510
      End
      Begin VB.Label lblSubCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Subseção:"
         Height          =   195
         Left            =   2880
         TabIndex        =   7
         Top             =   120
         Width           =   765
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
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelProdFor.frx":171C
      Height          =   7575
      Left            =   2520
      OleObjectBlob   =   "frmRelProdFor.frx":172D
      TabIndex        =   0
      Top             =   840
      Width           =   10935
   End
End
Attribute VB_Name = "frmRelProdFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub soma()
    Dim strSQL As String
    Dim cCusto As Currency
    Dim cVen As Currency
    Dim iEst As Integer
    
    With dt.Recordset
    
        Do While Not .EOF
            cCusto = cCusto + (!unit_ite * !qtd_ite)
            iEst = iEst + !qtd_ite
            'cVen = cVen + (!p2_pro * !e2_pro)
            .MoveNext
            
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        
        lblCusto.caption = "Custo: "
        'lblPreco.caption = "Preço: "
        lblQdade.caption = "Quantidade: "
        lblCusto.caption = lblCusto.caption & " " & FormatVal(cCusto)
        'lblPreco.caption = lblPreco.caption & " " & FormatVal(cVen)
        lblQdade.caption = lblQdade.caption & " " & FormatVal(iEst)
    End With

End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tc.*, ti.*, tf.nome_for, tp.*, tu.nome_sub, tp.unidade_pro,ts.nome_sec "
    strSQL = strSQL & " from tab_cmp as tc, tab_itcmp as ti,tab_subsec as tu, tab_secoes as ts, "
    strSQL = strSQL & " tab_produtos as tp, tab_fornecedores as tf  "
    
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and tu.cod_sub=tp.cod_sub "
    strSQL = strSQL & " and tu.cod_sec=ts.cod_sec  "
    strSQL = strSQL & " and tf.cod_for = tc.cod_for "
    
    If txtbxDataIni.Text <> "" And txtbxDataFin.Text <> "" Then
        strSQL = strSQL & " and tc.emissao_not >=" & F_GravaData(txtbxDataIni.Text)
        strSQL = strSQL & " and tc.emissao_not <=" & F_GravaData(txtbxDataFin.Text)
    End If
    If cmbxForn.Text <> "" Then
        strSQL = strSQL & " and tc.cod_for ='" & cmbxForn.tag & "'"
    End If
    
    If cmbxUF.Text <> "" Then
        strSQL = strSQL & " and tf.uf_for ='" & cmbxUF.Text & "'"
    End If
    strSQL = strSQL & " order by tc.emissao_not "
    
    dt.RecordSource = strSQL
    dt.Refresh
    soma
End Sub

Private Sub cmbxForn_Click()
    
    cmbxForn.tag = F_BuscaUmCampo("cod_for", "tab_fornecedores", "nome_for='" & cmbxForn.Text & "'")
    
End Sub

Private Sub cmbxForn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxForn.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxUF.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnImprimir_Click()
    If dt.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que ser impresso!", vbExclamation
        Exit Sub
    End If
    If cmbxForn.tag = "" Then
        MsgBox "Escolha um fornecedor!", vbExclamation
        cmbxForn.SetFocus
        Exit Sub
    End If
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    Dim iPage As Integer
    Dim iCount As Integer
    P_CabRelTimbr "Listagem de Produtos por Fornecedor", obj
    obj.FontName = "arial"
    obj.fontsize = 12
    obj.FontBold = True
    P_ImprX obj, cmbxForn.Text, 10, False
    obj.fontsize = 8
    obj.FontBold = False
    P_ImprLinha obj, "_"
    With dt.Recordset
        .MoveFirst
        P_ImprX obj, "Cód", 0
        P_ImprX obj, "Nome", 15
        P_ImprX obj, "UN", 100
        P_ImprX obj, "Estoque", 115, True, True
        P_ImprX obj, "Custo", 130, True, True
        P_ImprX obj, "Preço", 145, True, True
        P_ImprX obj, "Categoria", 147
        P_ImprX obj, "Marca", 180, False
        P_ImprLinha obj, "_"
        iPage = 1
        Do While Not .EOF
            If iCount = 50 Then
                
                iCount = 0
                P_RodapeVis obj, constNome_sis & "-" & constVersao_Sis, "Pág:" & iPage
                iPage = iPage + 1
                P_ImprX obj, "Cód", 0
                P_ImprX obj, "Nome", 15
                P_ImprX obj, "UN", 100
                P_ImprX obj, "Estoque", 115, True, True
                P_ImprX obj, "Custo", 130, True, True
                P_ImprX obj, "Preço", 145, True, True
                P_ImprX obj, "Categoria", 147
                P_ImprX obj, "Marca", 180, False
                P_ImprLinha obj, "_"
            End If
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome_pro, 15
            P_ImprX obj, !unidade_pro, 100
            P_ImprX obj, FormatVal(!er_pro), 115, True, True
            P_ImprX obj, FormatVal(!cr_pro), 130, True, True
            P_ImprX obj, FormatVal(!pv_pro), 145, True, True
            P_ImprX obj, !nome_cat, 147
            P_ImprX obj, " " & !nome_mar, 179, False
            
            iCount = iCount + 1
            .MoveNext
        Loop
    End With
    P_RodapeVis obj, constNome_sis & "-" & constVersao_Sis, "Pág:" & iPage & "(Fim)"
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dt.Connect = Banco.Connect
    
        
End Sub


Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub txtbxdatafin_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxdatafin_LostFocus()
    P_AjustaData txtbxDataFin
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
Private Sub txtbxFab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFor
    End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_SelectForn(txtbxFab.Text)
    If f.cod_for <> "" Then
        txtbxFab.Text = f.nome_for
        txtbxFab.tag = f.cod_for
        'busca
    End If
    
End Sub
Private Sub cmndbtnBuscaProd_Click()
    Dim p As Produto
    p = F_BuscaProdCabecalho(txtbxExpr.Text)
    If p.cod_pro <> "" Then
        txtbxExpr.Text = p.nome_pro
    
        txtbxExpr.tag = p.cod_pro
        cmndbtnBuscar_Click
    End If
End Sub

