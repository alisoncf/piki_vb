VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelMMD 
   Caption         =   "Relatório de Média de Movimentação Diária de Produto"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   Icon            =   "frmRelMMD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   10050
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   6375
      Left            =   2640
      TabIndex        =   21
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   11245
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton Command2 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelMMD.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ordenar Por:"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   2295
         Begin VB.OptionButton optnRetorno 
            Caption         =   "Retorno"
            Height          =   195
            Left            =   1200
            TabIndex        =   17
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optnGiro 
            Caption         =   "Giro"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optnMedia 
            Caption         =   "Média"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optnProd 
            Caption         =   "Produto"
            Height          =   195
            Left            =   1200
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optnCat 
            Caption         =   "Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optnMarca 
            Caption         =   "Marca"
            Height          =   195
            Left            =   1200
            TabIndex        =   15
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.TextBox txtbxDataIni 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataFin 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelMMD.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5520
         Width           =   2295
      End
      Begin VB.ComboBox cmbxCat 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cmbxMar 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnCalcula 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelMMD.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3840
         Width           =   2295
      End
      Begin VB.ComboBox cmbxForn 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "&Data Inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Data &Final"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Marca"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Fornecedor:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRelMMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmbxForn_Click()
    cmbxForn.tag = F_BuscaUmCampo("cod_for", TAB_FORNECEDORES, "nome_for='" & cmbxForn.Text & "'")
End Sub

Private Sub cmbxForn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyDelete Then
        cmbxForn.ListIndex = -1
    End If
End Sub

Private Sub cmbxMar_Click()
    cmbxMar.tag = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxMar.Text & "'")
End Sub

Private Sub cmbxMar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMar.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnCalcula_Click()
    If txtbxDataIni.Text = "" Then
        MsgBox "Informe a data inicial!", vbExclamation
        txtbxDataIni.SetFocus
        Exit Sub
    End If
    If txtbxDAtaFIn.Text = "" Then
        MsgBox "Informe a data final!", vbExclamation
        txtbxDAtaFIn.SetFocus
        Exit Sub
    End If
    busca
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = 10170
    Me.Height = 6975
    Me.Top = 10
    Me.Left = 10
    P_PreencheCateg cmbxCat
    P_PreencheMarca cmbxMar
    P_PreencheFornecedores cmbxForn
    With Msflxgrd
        .rows = 1
        .FixedCols = 0
        .Cols = 10
        .col = 0: .Text = "Cód:"
        .col = 1: .Text = "Produto:"
        .col = 2: .Text = "MMD:"
        .col = 3: .Text = "Preço:"
        .col = 4: .Text = "Custo:"
        .col = 5: .Text = "Giro:"
        .col = 6: .Text = "Retorno:"
        .col = 7: .Text = "Ret.Liq:"
        .col = 8: .Text = "Marca:"
        .col = 9: .Text = "Categoria:"
        .ColWidth(0) = 1000
        .ColWidth(1) = 3000
        .ColWidth(8) = 2000
        .ColWidth(9) = 2000
    End With
End Sub
Sub busca()
    'On Error GoTo trata_Erro
    Dim strSQL As String
    Dim iDias As Integer
    Dim rPesq As New ADODB.Recordset
    Dim sDias_Trab As String
    
    P_Status "Buscando produtos movimentados no período..."
    
    sDias_Trab = F_BuscaUmCampo(" count(distinct dataefet_ven)", "" & TAB_VENDAS & "", _
                " dataefet_ven >=" & F_PesqData(txtbxDataIni.Text) & " and dataefet_ven <=" & F_PesqData(txtbxDAtaFIn.Text))
    strSQL = " select"
    strSQL = strSQL & " tma.nome_mar,tca.nome_cat, tm.*, ti.cod_pro,"
    strSQL = strSQL & " tm.cod_mar,"
    strSQL = strSQL & " sum(ti.Qtd_ite) as qtd,"
    strSQL = strSQL & " sum(ti.Qtd_ite*ti.unit_ite) as giro,"
    strSQL = strSQL & " sum(ti.Qtd_ite*ti.unit_ite)-sum(ti.Qtd_ite*cr_pro) as retorno,"
    strSQL = strSQL & " Count(tor.dataefet_ven) as dias "
    strSQL = strSQL & " From tab_marcas as tma,tab_categorias as tca, "
    If cmbxForn.tag <> "" Then
        strSQL = strSQL & " tab_itcmp as tic,tab_cmp as tc,  "
    End If
    strSQL = strSQL & " " & TAB_VENDAS & " as tor,"
    strSQL = strSQL & " " & TAB_ITVENDAS & " As ti," & TAB_PRODUTOS & " as tm "
    strSQL = strSQL & " Where ti.cod_pro=tm.cod_pro and tm.cod_mar=tma.cod_mar "
    strSQL = strSQL & " and tm.cod_cat =tca.cod_cat "
    strSQL = strSQL & " and ti.num_ven = tor.num_ven "
    strSQL = strSQL & " and "
    strSQL = strSQL & " tor.dataefet_ven >=" & F_PesqData(txtbxDataIni.Text)
    strSQL = strSQL & " and tor.dataefet_ven <=" & F_PesqData(txtbxDAtaFIn.Text)
    If cmbxForn.tag <> "" Then
        strSQL = strSQL & " and tic.seq_not=tc.seq_not "
        strSQL = strSQL & " and tic.cod_pro=ti.cod_pro "
        strSQL = strSQL & " and tc.cod_for='" & cmbxForn.tag & "'"
    End If
    If cmbxCat.tag <> "" Then
        strSQL = strSQL & " and tm.cod_cat='" & cmbxCat.tag & "' "
    End If
    
    If cmbxMar.tag <> "" Then
        strSQL = strSQL & " and tm.cod_mar='" & cmbxMar.tag & "' "
    End If
'    If sNome <> "" Then
'        strSQL = strSQL & " and tm.nome_pro like '" & sPro & "%' "
'    End If
    strSQL = strSQL & " Group By"
    strSQL = strSQL & " ti.cod_pro "
    If optnMedia.Value Then
        strSQL = strSQL & " order by qtd desc "
    ElseIf optnGiro.Value Then
        strSQL = strSQL & " order by giro desc "
    ElseIf optnMarca.Value Then
        strSQL = strSQL & " order by nome_mar"
    ElseIf optnProd.Value Then
        strSQL = strSQL & " order by nome_pro,cod_pro "
    ElseIf optnRetorno.Value Then
        strSQL = strSQL & " order by retorno desc "
    End If
    rPesq.Open strSQL, conn_client
    With rPesq
        Msflxgrd.rows = 1
        Do While Not .EOF
            strSQL = !cod_pro & Chr(9)
            strSQL = strSQL & !nome_pro & Chr(9)
            strSQL = strSQL & FormatVal(!qtd / CCur(sDias_Trab), "###,###,##0.00") & Chr(9)
            strSQL = strSQL & FormatVal(!pv_pro) & Chr(9)
            strSQL = strSQL & FormatVal(!cr_pro) & Chr(9)
            strSQL = strSQL & FormatVal(!giro) & Chr(9)
            strSQL = strSQL & FormatVal(!Retorno) & Chr(9)
            If !Retorno > 0 Then
                strSQL = strSQL & FormatVal(!Retorno - (!Retorno * F_TaxasEntradas / 100)) & Chr(9)
            Else
                strSQL = strSQL & FormatVal(!Retorno + (!Retorno * F_TaxasEntradas / 100)) & Chr(9)
            End If
            strSQL = strSQL & !nome_mar & Chr(9)
            strSQL = strSQL & !nome_cat
            Msflxgrd.AddItem strSQL
            .MoveNext
        Loop
    End With
    
    ZebraGrid Msflxgrd, 1
    P_Status ""
End Sub

Private Sub Image1_Click()

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
