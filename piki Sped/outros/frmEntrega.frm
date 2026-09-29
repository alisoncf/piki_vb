VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrega"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9180
   Begin VB.Frame Frame1 
      Caption         =   "&Vendas a Entregar"
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.Data dtVendas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   4440
         Top             =   1200
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEntrega.frx":0000
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "frmEntrega.frx":0017
         TabIndex        =   1
         Top             =   240
         Width           =   8895
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   5160
      Width           =   9135
      Begin VB.CommandButton cmndbtnFornadas 
         Caption         =   "&Forn. Pré-venda"
         Height          =   660
         Left            =   5520
         Picture         =   "frmEntrega.frx":10BC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnVisualizar 
         Caption         =   "&Visualizar"
         Height          =   660
         Left            =   3720
         Picture         =   "frmEntrega.frx":1206
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEfetivar 
         Caption         =   "&Entregar"
         Height          =   660
         Left            =   120
         Picture         =   "frmEntrega.frx":1350
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   660
         Left            =   7320
         Picture         =   "frmEntrega.frx":149A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   660
         Left            =   1920
         Picture         =   "frmEntrega.frx":15E4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub busca()
     Dim strSQL As String
    
    
    
    
    
    strSQL = "select tc.nome_cli,tv.* from tab_clientes as tc, "
    strSQL = strSQL & " tab_vendas as tv "
    
    strSQL = strSQL & " where tv.situacao_ven='3' "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    strSQL = strSQL & " order by num_ven"
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
End Sub
    
Private Sub cmndbtnEfetivar_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    If dbgrd.Row = -1 Then
        MsgBox "selecione a Venda!", vbInformation
        dbgrd.SetFocus
        Exit Sub
    End If
    With dtVendas.Recordset
        If MsgBox("Deseja Realmente Entregar a Venda?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
            
        End If
        
        If Not F_senha Then
            MsgBox "Senha Inválida!", vbExclamation
            Exit Sub
        End If
        
        
        If Not ConferePrevenda Then Exit Sub
        If Not ConfereFornadas Then Exit Sub
        P_BaixaVenda !Num_Ven
        MsgBox "Venda Concluída!", vbExclamation
    End With

    
    busca
            
    
End Sub
Function ConfereFornadas() As Boolean
    Dim strSQL As String
    Dim sMsg As String
    Dim bForn As Boolean
    Dim rPesq As Recordset
    strSQL = "select tp.pei_pro, tv.* from tab_prevendas as tv, tab_produtos as tp  "
    strSQL = strSQL & " where tp.cod_pro=tv.cod_pro "
    strSQL = strSQL & " and  num_ven='" & dtVendas.Recordset!Num_Ven & "'"
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        bForn = True
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            If val(!pei_pro) > 0 Then
                strSQL = F_BuscaUmCampo("sum(qtd_ite)", "tab_preforn", _
                    "num_ven='" & dtVendas.Recordset!Num_Ven & "'")
                If strSQL = "" Then
                    strSQL = "0"
                End If
                If CCur(strSQL) < !qtd_ite Then
                    sMsg = sMsg & "No Item (" & .AbsolutePosition + 1 & ")  " & Format(!cod_pro, "00#\.###\-#") & " Restam distribuir " & !qtd_ite - CCur(strSQL) & " Itens " & Chr(13)
                    bForn = False
                End If
                    
                
            End If
            .MoveNext
        Loop
        ConfereFornadas = bForn
        If bForn = False Then
            sMsg = sMsg & Chr(13) & "Selecione o produto descrito e clique em 'Fornadas'"
            MsgBox "Distribuição de Fornadas Incompleta: " & Chr(13) & sMsg, vbCritical
            
        End If
    End With
End Function

Function ConferePrevenda() As Boolean
    Dim strSQL As String
    Dim sMsg As String
    Dim bForn As Boolean
    Dim rPesq As Recordset
    strSQL = "select tp.*,te.estoque_pro,te.nome_pro from tab_prevendas as tp,"
    strSQL = strSQL & "tab_produtos as te where te.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and num_ven='" & dtVendas.Recordset!Num_Ven & "' "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rPesq
        bForn = True
        If .RecordCount = 0 Then
            ConferePrevenda = True
            Exit Function
        End If
        If .RecordCount > 1 Then
            sMsg = "Os Itens Abaixo: " & Chr(13)
        Else
            sMsg = "O Item Abaixo: " & Chr(13)
        End If
        Do While Not .EOF
            
            If !qtd_ite - !estoque_pro > 0 Then
                sMsg = sMsg & Format(!cod_pro, "00#\.###\-#") & "-" & !nome_pro & Chr(13)
                bForn = False
            End If
            .MoveNext
        Loop
        ConferePrevenda = bForn
        If bForn = False Then
            sMsg = sMsg & Chr(13) & "Não Possue(m) estoque suficiente para a entrega!"
            sMsg = sMsg & Chr(13) & "Se este produto já se econtra no estoque, "
            sMsg = sMsg & "Entre em contato com o responsável pela Entrada para que a proceda imediatamente!"
            sMsg = sMsg & Chr(13) & Chr(13) & " A Venda não pode ser concluída!"
            MsgBox "Pré-Venda: " & Chr(13) & Chr(13) & sMsg, vbCritical
            
        End If
    End With
End Function

Private Sub cmndbtnReabrir_Click()
    Dim strSQL As String
    If dbgrd.Row = -1 Then
        MsgBox "selecione a Venda!", vbInformation
        dbgrd.SetFocus
        Exit Sub
    End If
    With dtVendas.Recordset
        If MsgBox("Deseja Realmente Reabrir a Venda?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = strSQL & " update tab_vendas set situacao_ven='1'"
            strSQL = strSQL & " where num_ven='" & !Num_Ven & "' "
            banco.Execute strSQL
            busca
        End If
    End With
End Sub

Private Sub cmndbtnFornadas_Click()
    frmProdutoEntregaForn.Tag = dtVendas.Recordset!Num_Ven
    frmProdutoEntregaForn.Show 1
End Sub

Private Sub cmndbtnImpr_Click()
    P_Status "Gerando impressão... "
    P_ImprimeVenda frmVisualiza.Pic, dtVendas.Recordset!Num_Ven, False
    P_Status ""
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
        Unload frmVisualiza
        P_Status "Imprimindo..."
        P_ImprimeVenda Printer, dtVendas.Recordset!Num_Ven, True
        P_Status ""
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVisualizar_Click()
        
    P_Status "Montando venda..."
    With dtVendas.Recordset
        frmVenda.Tag = !Num_Ven
        'fr.lblAbert.caption = !abertura_ven
        frmVenda.lblNome_Cli.caption = !nome_cli
        frmVenda.lblNome_Cli.Tag = !cod_cli
        frmVenda.lblHora.caption = !data_ven
        frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
        frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", "tab_Funcionarios", _
            "cod_fun='" & !cod_usu & "'")
        frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
        frmVenda.lblValor.caption = FormatVal(!valor_ven)
        frmVenda.caption = "Venda " & !Num_Ven & "-" & !nome_cli
        frmVenda.cmndbtnAddExr.Enabled = False
        frmVenda.cmndbtnAddPre.Enabled = False
        frmVenda.cmndbtnAddProd.Enabled = False
        frmVenda.cmndbtnCancel.Enabled = False
        frmVenda.cmndbtnEdit_Pre.Enabled = False
        frmVenda.cmndbtnEditEx.Enabled = False
        frmVenda.cmndbtnExcEx.Enabled = False
        frmVenda.cmndbtnFornada.Enabled = False
        frmVenda.cmndbtnSalvar.Enabled = False
        frmVenda.cmndbtnExcl_Pre.Enabled = False
        frmVenda.cmndbtnExc_Prod.Enabled = False
        frmVenda.cmndbtnEditProd.Enabled = False
        frmVenda.frControle.Tag = "VISUALIZAÇÃO"
        frmVenda.Show 1
    End With
    P_Status ""

End Sub

Private Sub dtVendas_Reposition()
    If dtVendas.Recordset.RecordCount = 0 Then
        'cmndbtnCancelar.Enabled = False
        cmndbtnEfetivar.Enabled = False
        cmndbtnImpr.Enabled = False
        'cmndbtnReabrir.Enabled = False
    Else
        'cmndbtnCancelar.Enabled = True
        
        If F_Fornada Then
            cmndbtnFornadas.Enabled = True
        Else
            cmndbtnFornadas.Enabled = False
        End If
        cmndbtnEfetivar.Enabled = True
        cmndbtnImpr.Enabled = True
        'cmndbtnReabrir.Enabled = True
    End If
End Sub
Function F_Fornada() As Boolean
    Dim strSQL As String
    Dim bForn As Boolean
    Dim rPesq As Recordset
    strSQL = "select tp.cod_pro from tab_produtos as tp,"
    strSQL = strSQL & " tab_prevendas as tv "
    strSQL = strSQL & " where tv.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tp.pei_pro <> '' "
    strSQL = strSQL & " and tv.num_ven='" & dtVendas.Recordset!Num_Ven & "' "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            F_Fornada = True
        Else
            F_Fornada = False
        End If
        .Close
    End With
End Function
Private Sub Form_Activate()
    busca
    Me.caption = "Vendas a Entregar  - Usuário: " & strNome_Usu
End Sub

Private Sub Form_Load()
    dtVendas.Connect = banco.Connect
    P_CentraForm Me
End Sub

Private Sub Timer1_Timer()
    RefreshDataPosition dtVendas
End Sub


