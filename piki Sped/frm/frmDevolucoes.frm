VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDevolucoes 
   AutoRedraw      =   -1  'True
   Caption         =   "Devoluções de Produtos"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   Icon            =   "frmDevolucoes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbxNum_Ven 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   5505
      TabIndex        =   12
      Top             =   6480
      Width           =   5535
      Begin VB.Label lblValorGerado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   3240
         TabIndex        =   16
         Top             =   37
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Gerado:"
         Height          =   195
         Left            =   2160
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblItens 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   1560
         TabIndex        =   14
         Top             =   37
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Itens Devolvidos:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   855
      Left            =   8880
      Picture         =   "frmDevolucoes.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnConcluir 
      Caption         =   "&Concluir Devolução"
      Height          =   855
      Left            =   7080
      Picture         =   "frmDevolucoes.frx":06D8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1695
   End
   Begin VB.PictureBox pcRec2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   165
      Picture         =   "frmDevolucoes.frx":0822
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   4560
      Width           =   195
   End
   Begin VB.PictureBox pcRec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   165
      Picture         =   "frmDevolucoes.frx":0DAC
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   1320
      Width           =   195
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4895
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Width           =   8415
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Desc:"
         Height          =   195
         Left            =   4320
         TabIndex        =   21
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2520
         TabIndex        =   20
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   525
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4920
         TabIndex        =   9
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblVendedor 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   180
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblDataVen 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   840
         TabIndex        =   5
         Top             =   120
         Width           =   180
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrdDev 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4260
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Número da Venda:"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmDevolucoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub P_buscaItens(sOrd As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rForn As New ADODB.Recordset
    Dim sMsg As String
  '  On Error GoTo trata_erro
    msflxgrd.rows = 1
    
    strSQL = "select tv.*,tp.*,tg.nome_cat,tg.nome_cat,"
    strSQL = strSQL & "nome_mar,tm.cod_mar,ti.*,(ti.unit_ite*ti.qtd_ite) as total "
    strSQL = strSQL & " from  " & TAB_VENDAS & " as tv inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven  "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & "  as tp on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " left join " & tab_categorias & "  as tg on tp.cod_cat=tg.cod_cat "
    strSQL = strSQL & " left join " & TAB_MARCAS & "  as tm on tp.cod_mar=tm.cod_mar "
    
    
    strSQL = strSQL & " where tv.num_ven='" & Me.tag & "' "
    strSQL = strSQL & " and tv.situacao_ven='4' "
    strSQL = strSQL & " order by " & sOrd
    
    'dtProd.RecordSource = strSQL
    'dtProd.Refresh
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If F_BuscaUmCampo("cod_pro", "tab_devolucoes", "cod_pro='" & !cod_pro & _
                "' and num_ven='" & Me.tag & "' and tipo_dev='I'") = "" Then
                sMsg = Chr(9) & FormatCodigo(!cod_pro) & Chr(9)
                sMsg = sMsg & !nome_pro & Chr(9) & FormatVal(!qtd_ite) & Chr(9)
                sMsg = sMsg & FormatVal(!unit_ite) & Chr(9) & FormatVal(!total)
                sMsg = sMsg & Chr(9) & "I" & Chr(9) & "0,00" & Chr(9) & Format(!codfab_pro)
                msflxgrd.AddItem sMsg
            End If
            .MoveNext
            
        Loop
        .Close
    End With
    
    
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub


Private Sub cmndbtnConcluir_Click()
    If msflxgrdDev.rows = 1 Then
        MsgBox "Não há itens na Devolução!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma conclusão da Devolução?", vbYesNo + vbQuestion) = vbYes Then
        
        devolve
            
    End If
End Sub

Sub devolve()
    Dim iCount As Integer
    Dim sTipo As String
    Dim sCod As String
    Dim sQtd As String
    Dim sUnit As String
    Dim sTotal As String
    Dim sDev As String
    With msflxgrdDev
        For iCount = 1 To .rows - 1
            .Row = iCount
            .col = 1
            sCod = F_SoNumero(.Text)
            .col = 3
            sQtd = .Text
            .col = 4
            sUnit = .Text
            .col = 5
            sTotal = .Text
            .col = 7
            sTipo = .Text
            .col = 6
            
            msflxgrd.Row = .Text
            msflxgrd.col = 3
            If Not F_GravaDevolucao(Me.tag, sCod, msflxgrd.Text, CCur(sQtd), Date, CCur(sTotal), sTipo) Then
                MsgBox "Erro na devolução!", vbExclamation
                Exit Sub
            Else
                
            End If
            
        Next
        MsgBox "Devolução concluída!", vbExclamation
        msflxgrd.rows = 1
        msflxgrdDev.rows = 1
        txtbxNum_Ven.Text = ""
        txtbxNum_Ven.SetFocus
    
        P_imprDevolucao Me.tag, Printer
        'P_imprDevolucao Me.Tag, Printer, False
        Dim strSQL As String
        'strSQL = F_GravaCai("D", Date, lblValorGerado.caption, "DEVOLUÇÃO - VENDA: " & Me.Tag, _
         '       Me.Tag, "DIN", strCod_Cai, "0", lblValorGerado.caption, "0", strCod_Cai)
        strSQL = F_GravaCai("D", Date, lblValorGerado.caption, "DEVOLUÇÃO - VENDA: " & Me.tag, _
                Me.tag, "DIN", strCod_Cai, "0", lblValorGerado.caption, strCod_Cai, strCod_Tur, "", "S")
        
        lblValorGerado.caption = "0,00"
        lblItens.caption = "0"
    End With
End Sub

Private Sub cmndbtnConf_Click()
    Dim cQtd As Currency
    Dim sCod As String
    Dim sNome As String
    Dim sTipo As String
    Dim strSQL As String
    If msflxgrd.rows = 1 Then Exit Sub
    If F_InFlxgrd(msflxgrdDev, msflxgrd.Row, 6) >= 0 Then
        MsgBox "Este item já está devolvido!", vbExclamation
        Exit Sub
    End If
    
    
    With frmValoresPReVEn
        .caption = "Itens a devolver"
        msflxgrd.col = 1
         sCod = msflxgrd.Text
        .lblCod_Pro.caption = msflxgrd.Text
        msflxgrd.col = 2
        sNome = msflxgrd.Text
        .lblNome_Pro.caption = msflxgrd.Text
        msflxgrd.col = 3
        cQtd = CCur(msflxgrd.Text)
        .txtbxQtd.Text = msflxgrd.Text
        msflxgrd.col = 4
        .txtbxPreco.Text = msflxgrd.Text
        msflxgrd.col = 6
        sTipo = msflxgrd.Text
        .Show 1
        If .tag <> "" Then
            If CCur(.txtbxQtd.Text) > cQtd Then
                MsgBox "A Quantidade informada foi maior que a permitida!", vbExclamation
                Unload frmValoresPReVEn
                Exit Sub
            End If
            
            
            
            strSQL = Chr(9) & sCod & Chr(9) & sNome & Chr(9) & .txtbxQtd.Text & Chr(9)
            strSQL = strSQL & .txtbxPreco.Text & Chr(9)
            strSQL = strSQL & FormatVal(Round(CCur(.txtbxQtd.Text) * CCur(.txtbxPreco.Text) _
                * (1 - (CCur(lblDesc.tag) / 100)), 2))
            strSQL = strSQL & Chr(9) & msflxgrd.Row & Chr(9) & sTipo
            msflxgrdDev.AddItem strSQL
            msflxgrd.col = 7
            msflxgrd.Text = .txtbxQtd.Text
            msflxgrdDev.Row = msflxgrdDev.rows - 1
            msflxgrdDev.col = 5
            lblValorGerado.caption = FormatVal(CCur(lblValorGerado.caption) + _
            CCur(msflxgrdDev.Text))
        End If
        
        msflxgrdDev_RowColChange
        lblItens.caption = msflxgrdDev.rows - 1
    End With
    
End Sub

Private Sub cmndbtnFechar_Click()
    Dim strSQL As String
    Dim iCount As Integer
    If msflxgrdDev.rows > 1 Then
        If MsgBox("A Devolução não está concluída!" & Chr(13) & _
            "Deseja Realmente Sair?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    With msflxgrdDev
        For iCount = 1 To .rows - 1
            .Row = iCount
            .col = 1
            strSQL = "delete from tab_devforn where num_ven='" & Me.tag & "' "
            strSQL = strSQL & " and cod_pro='" & F_SoNumero(.Text) & "' "
            'conn.execute strSQL
        Next
    End With
    Unload Me
End Sub

Private Sub cmndbtnRem_Click()
    Dim sTipo As String
    Dim strSQL As String
    If msflxgrdDev.rows = 1 Then
        Exit Sub
    End If
    msflxgrdDev.col = 7
    sTipo = msflxgrdDev.Text
    msflxgrdDev.col = 6
    
    msflxgrd.Row = msflxgrdDev.Text
    msflxgrd.col = 7
    msflxgrd.Text = "0,00"
    msflxgrdDev.col = 1
    
    msflxgrdDev.col = 5
            lblValorGerado.caption = FormatVal(CCur(lblValorGerado.caption) - _
            CCur(msflxgrdDev.Text))
    If msflxgrdDev.rows = 2 Then
        msflxgrdDev.rows = 1
    Else
        msflxgrdDev.RemoveItem msflxgrdDev.Row
    End If
    
    lblItens.caption = msflxgrdDev.rows - 1
End Sub

Private Sub Form_Load()
    'dtExt.ConnectionString = strConnect
    'dtPre.ConnectionString = strConnect
    'dtProd.ConnectionString = strConnect
    
    With msflxgrd
        
        .Cols = 9
        .FixedRows = 1
        .FixedCols = 0
        .rows = 1
        .col = 0
        .Text = "..."
        .col = 1
        .Text = "Código"
        .col = 2
        .Text = "Descrição"
        .col = 3
        .Text = "Qtd"
        .col = 4
        .Text = "Unit"
        .col = 5
        .Text = "Total"
        
        .col = 6
        .Text = "Tipo"
        .col = 7
        .Text = "Devolvido"
        .col = 8
        .Text = "Ref.:"
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 3500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 500
        .ColWidth(7) = 1000
        
        
        
    End With
    With msflxgrdDev
        
        .Cols = 8
        .FixedRows = 1
        .FixedCols = 0
        .rows = 1
        .col = 0
        .Text = "..."
        .col = 1
        .Text = "Código"
        .col = 2
        .Text = "Descrição"
        .col = 3
        .Text = "Qtd"
        .col = 4
        .Text = "Unit"
        .col = 5
        .Text = "Gerado"
        
        .col = 6
        .Text = "Item"
        .col = 7
        .Text = "T"
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 3500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 300
        
        
        
    End With
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnConf_Click
    End If
End Sub

Private Sub msflxgrd_RowColChange()
    pcRec.Top = msflxgrd.CellTop + msflxgrd.Top '+ pcRec.Height / 2
    
End Sub

Private Sub msflxgrdDev_KeyDown(KeyCode As Integer, Shift As Integer)
    cmndbtnRem_Click
End Sub

Private Sub msflxgrdDev_RowColChange()
    pcRec2.Top = msflxgrdDev.CellTop + msflxgrdDev.Top '+ pcRec.Height / 2
End Sub

Private Sub txtbxNum_Ven_GotFocus()
    ganhaFoco txtbxNum_Ven
End Sub

Private Sub txtbxNum_Ven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        msflxgrd.rows = 1
        msflxgrdDev.rows = 1
        lblValorGerado.caption = "0,00"
        lblItens.caption = "0"
        If F_Vazio(txtbxNum_Ven.Text) Then
            Exit Sub
        End If
    
    busca
    
    If txtbxNum_Ven.Text = "" Then Exit Sub
    
    P_buscaItens "seq_ite"
        
    If msflxgrd.rows > 1 Then
        msflxgrd_RowColChange
            TrocaTabKey KeyCode
    End If
    End If
End Sub

Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    txtbxNum_Ven.Text = Format(F_SoNumero(txtbxNum_Ven.Text), "000000#")
    Me.tag = txtbxNum_Ven.Text
    strSQL = "select * from " & TAB_VENDAS & " where num_ven='" & txtbxNum_Ven.Text & "'"
    abreObjetoAdo rPesq, strSQL
    
    
        
    
    With rPesq
           If .RecordCount = 0 Then
             MsgBox "Esta venda não existe!", vbExclamation
             txtbxNum_Ven.Text = ""
             txtbxNum_Ven.SetFocus
             Exit Sub
         End If
        If !situacao_ven < "4" Then
            MsgBox "Não é possível haver devoluções nesta venda!" _
                & Chr(13) & "Ela ainda não se encontra finalizada!", vbExclamation
            txtbxNum_Ven.Text = ""
            txtbxNum_Ven.SetFocus
            Exit Sub
        End If
        lblCliente.caption = F_BuscaUmCampo("Nome_cli", "tab_clientes", _
            "cod_cli='" & !cod_cli & "'")
        lblDataVen.caption = !data_ven
        lblValor.caption = FormatVal(!valor_ven)
        lblDesc.caption = Round(!desconto_ven / (!valor_ven), 2) * 100 & " %"
        lblDesc.tag = Round(!desconto_ven / (!valor_ven), 2) * 100
        
        lblVendedor.caption = F_BuscaUmCampo("nome_fun", "tab_funcionarios", _
            "cod_fun='" & !fechavenda_usu & "'")
            .Close
    End With
End Sub


