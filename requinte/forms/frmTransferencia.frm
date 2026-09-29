VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmTransferencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferência"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmTransferencia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9060
   Begin VB.Frame frControle 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Frame frOrd 
         Caption         =   "Ordem:"
         Height          =   1095
         Left            =   120
         TabIndex        =   24
         Top             =   4320
         Width           =   1815
         Begin VB.OptionButton optnDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optnCod 
            Caption         =   "Códi&go"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optnSeq 
            Caption         =   "Seqüência"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame frBotoes 
         Height          =   975
         Left            =   4680
         TabIndex        =   20
         Top             =   120
         Width           =   4215
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   735
            Left            =   2160
            Picture         =   "frmTransferencia.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   180
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnSalvar 
            Caption         =   "&Concluir"
            Height          =   735
            Left            =   120
            Picture         =   "frmTransferencia.frx":058C
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   180
            Width           =   1935
         End
      End
      Begin VB.Frame frOpt 
         Caption         =   "Produtos:"
         Height          =   3255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
         Begin VB.OptionButton optnBar 
            Caption         =   "&B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   255
         End
         Begin VB.OptionButton optnN 
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optnC 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox txtbxCodBarra 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   13
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnEditProd 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   2520
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnExc_Prod 
            Caption         =   "E&xcluir"
            Height          =   495
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton cmndbtnAddProd 
            Caption         =   "&Adicionar"
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtbxQtd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   9
            Text            =   "0,0"
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblUnit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   315
         End
         Begin VB.Label lblExpr 
            AutoSize        =   -1  'True
            Caption         =   "&C&ód. Barra:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Tag             =   "B"
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblMarca 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   2400
            Width           =   45
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Quantidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Tag             =   "B"
            Top             =   840
            Width           =   870
         End
      End
      Begin VB.Data dtProd 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame frCab 
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4455
         Begin VB.TextBox txtbxNome_ter 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   3
            Top             =   360
            Width           =   3975
         End
         Begin VB.CommandButton cmndbtnLocaliza 
            Height          =   285
            Left            =   4080
            Picture         =   "frmTransferencia.frx":06D6
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3765
            TabIndex        =   30
            Top             =   645
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            Height          =   195
            Left            =   2880
            TabIndex        =   29
            Top             =   720
            Width           =   405
         End
         Begin VB.Label lblUF_Ter 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4005
            TabIndex        =   7
            Top             =   120
            Width           =   75
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "&Loja:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   345
         End
         Begin VB.Label lblCnpj 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   420
         End
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "Cancelar Transferência"
         Height          =   735
         Left            =   120
         Picture         =   "frmTransferencia.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5520
         Width           =   1815
      End
      Begin MSDBGrid.DBGrid dbgrdProd 
         Bindings        =   "frmTransferencia.frx":096A
         Height          =   5055
         Left            =   2040
         OleObjectBlob   =   "frmTransferencia.frx":097F
         TabIndex        =   23
         Top             =   1200
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmndbtnAddProd_Click()
    Dim strSQL As String
    Dim sItem As String
    
    If Me.Tag = "" Then
        MsgBox "Não existe uma Transferência gerada!", vbExclamation
        txtbxNome_ter.SetFocus
        Exit Sub
    End If
    If txtbxCodBarra.Tag = "" Then
        MsgBox "Escolha o produto!", vbExclamation
        Exit Sub
    End If
    If F_BuscaUmCampo("cod_pro", "tab_ittransf", "seq_tra='" & Me.Tag & "' and cod_pro='" & _
        txtbxCodBarra.Tag & "'") <> "" Then
        MsgBox "Este item já se encontra lançado!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "insert into tab_ittransf(seq_tra,cod_pro,qtd_itt,unit_itt,item_itt)values("
    strSQL = strSQL & "'" & Me.Tag & "',"
    strSQL = strSQL & "'" & txtbxCodBarra.Tag & "',"
    strSQL = strSQL & F_GravaDec(txtbxQtd.Text) & ","
    strSQL = strSQL & F_GravaDec(lblUnit.caption) & ","
    sItem = F_BuscaUmCampo("max(item_itt)", "tab_ittransf", "seq_tra='" & Me.Tag & "'")
    If sItem = "" Then
        sItem = "1"
    Else
        sItem = Int(sItem + 1)
    End If
    strSQL = strSQL & sItem & ")"
    banco.Execute strSQL
    busca
    txtbxCodBarra.Text = ""
    txtbxQtd.Text = "0,0"
    txtbxCodBarra.SetFocus
End Sub

Private Sub cmndbtnCancelar_Click()
    Dim strSQL As String
    On Error Resume Next
    If Me.Tag = "" Then
        Unload Me
        Exit Sub
    End If
    If MsgBox("Tem certeza que deseja cancelar a Transferência " & Me.Tag & "?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from tab_transf where seq_Tra='" & Me.Tag & "' "
        banco.Execute strSQL
        strSQL = "delete from tab_ittransf where seq_Tra='" & Me.Tag & "' "
        banco.Execute strSQL
    End If
    MsgBox "transferência excluída!", vbExclamation
    If IsLoad("frmEdiTransf") Then
        frmEdiTransf.dtVendas.Refresh
    End If
    Unload Me
End Sub

Private Sub cmndbtnExc_Prod_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    With dtProd.Recordset
    If .RecordCount = 0 Then
        MsgBox "Não há o que ser excluído!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma exclusão do lançamento?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from tab_ittransf where cod_pro='" & !cod_pro & "' "
        strSQL = strSQL & " and seq_Tra='" & Me.Tag & "' "
        banco.Execute strSQL
        busca
    End If
        
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxNome_ter.Tag = "" Then
        MsgBox "Informe a loja!", vbExclamation
        Exit Sub
    End If
    If Me.Tag = "" Then
        
        If MsgBox("Confirma abertura da Transferência?", vbYesNo + vbQuestion) = vbYes Then
        
            'txtbxNome_ter.Text = frmSelectCli.dtBusca.Recordset!nome
            'txtbxNome_ter.Tag = frmSelectCli.dtBusca.Recordset!cod
            If txtbxNome_ter.Tag = "" Then
                Exit Sub
            End If
        
            Me.Tag = F_GeraNovaTransf(txtbxNome_ter.Tag, strCod_Usu)
            If Me.Tag = "" Then
                MsgBox "Transferência inválida!", vbExclamation
                Exit Sub
            End If
            Me.caption = "Transferência nº " & Me.Tag & "-" & txtbxNome_ter.Text
            cmndbtnSalvar.caption = "&Concluir Lançamento"
            
        End If
    Else
        If MsgBox("Confirma conclusão dos lançamentos?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        If txtbxNome_ter.Tag = "" Then
            MsgBox "Escolha A Loja!", vbExclamation
            Exit Sub
        End If
        If dtProd.Recordset.RecordCount = 0 Then
            MsgBox "Não há itens para conclusão!", vbExclamation
            Exit Sub
        End If
        Baixa
        MsgBox "Lançamento de Transferência concluído!", vbInformation
        imprimeTransf Me.Tag
        Unload Me
    End If
End Sub
Sub Baixa()
    Dim strSQL As String
    Dim cEst As Currency
    strSQL = "update tab_transf set situacao_tra='2' where seq_tra='" & Me.Tag & "' "
    banco.Execute strSQL

    With dtProd.Recordset
        Do While Not .EOF
            strSQL = F_BuscaUmCampo("e2_pro", "tab_Composicao", "cod_pro='" & !cod_pro & "'")
            If !qtd_itt > CCur(strSQL) Then
                cEst = "0"
            Else
                cEst = CCur(strSQL) - !qtd_itt
            End If
            strSQL = "update tab_composicao set e2_pro=" & F_GravaDec(cEst)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            
            
            banco.Execute strSQL
            
            .MoveNext
        Loop
        
    End With
End Sub
Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    If Me.Tag = "" Then
        cmndbtnSalvar.caption = "&Abrir Lançamento"
    Else
        cmndbtnSalvar.caption = "&Concluir Lançamento"
    End If
    busca
    bActive = True
End Sub
Sub soma()
    Dim cTotal As Currency
    Dim iPos As Integer
    On Error Resume Next
    With dtProd.Recordset
        iPos = .AbsolutePosition
        Do While Not .EOF
            cTotal = cTotal + !Total
            
            .MoveNext
        Loop
        .MoveFirst
        .Move iPos
    End With
    lblTotal.caption = FormatVal(cTotal)
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select ti.*,ti.unit_itt*ti.qtd_itt as total,tp.nome_pro from tab_produtos as tp, "
    strSQL = strSQL & "tab_ittransf as ti, tab_transf as ta where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & "and ta.seq_tra=ti.seq_tra "
    strSQL = strSQL & " and ti.seq_tra='" & Me.Tag & "' "
    If optnCod.Value = True Then
        strSQL = strSQL & "  order by ti.cod_pro"
    ElseIf optnSeq.Value Then
        strSQL = strSQL & "  order by ti.item_itt"
    Else
        strSQL = strSQL & "  order by tp.nome_pro"
    End If
    
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    soma
End Sub
Private Sub Form_Load()
    P_CentraForm Me
    dtProd.Connect = banco.Connect
End Sub

Private Sub txtbxCodBarra_GotFocus()
    P_SelecTexto txtbxCodBarra
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim p As Produto
    If KeyCode = vbKeyReturn Then
        txtbxCodBarra_LostFocus
        If optnBar.Value Then
            p = F_BuscaProdSaida(txtbxCodBarra.Text, "B")
        ElseIf optnC.Value Then
            p = F_BuscaProdSaida(txtbxCodBarra.Text, "C")
        Else
            p = F_BuscaProdSaida(txtbxCodBarra.Text, "N")
        End If
    End If
    If p.cod_pro <> "" Then
        txtbxCodBarra.Text = p.nome_pro
        txtbxCodBarra.Tag = p.cod_pro
        lblUnit.caption = FormatVal(p.c2_pro)
        txtbxQtd.SetFocus
    
    End If
End Sub

Private Sub txtbxCodBarra_LostFocus()
    P_AjustaTexto txtbxCodBarra
End Sub

Private Sub txtbxNOme_Ter_GotFocus()
    P_SelecTexto txtbxNome_ter
End Sub

Private Sub txtbxNOme_Ter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub
Private Sub cmndbtnLocaliza_Click()
    bActive = True
    frmSelectCli.txtbxNome.Tag = "L"
    frmSelectCli.txtbxNome.Text = txtbxNome_ter.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxNome_ter.Text = frmSelectCli.dtBusca.Recordset!nome
    txtbxNome_ter.Tag = frmSelectCli.dtBusca.Recordset!cod
        
    Unload frmSelectCli
    bActive = False
        
    
    
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd, "###,##0.0"
End Sub


