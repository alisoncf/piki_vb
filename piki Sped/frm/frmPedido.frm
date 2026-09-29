VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido de Produto"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frControle 
      BackColor       =   &H8000000A&
      Height          =   7095
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7575
      Begin VB.Data dtProd 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   5160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame frmCli 
         Height          =   1335
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   7335
         Begin VB.TextBox txtbxEndereco 
            Height          =   285
            Left            =   3480
            TabIndex        =   6
            Top             =   960
            Width           =   3615
         End
         Begin VB.TextBox txtbxNome_cli 
            Height          =   285
            Left            =   3480
            TabIndex        =   4
            Top             =   600
            Width           =   3615
         End
         Begin VB.ComboBox cmbxLoja 
            Height          =   315
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&End.:"
            Height          =   195
            Left            =   3030
            TabIndex        =   5
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Cliente:"
            Height          =   195
            Left            =   2880
            TabIndex        =   3
            Top             =   600
            Width           =   525
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   450
            TabIndex        =   24
            Top             =   480
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Itens:"
            Height          =   195
            Left            =   465
            TabIndex        =   21
            Top             =   720
            Width           =   390
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Loja:"
            Height          =   195
            Left            =   3060
            TabIndex        =   1
            Top             =   240
            Width           =   345
         End
         Begin VB.Label lblValor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   960
            TabIndex        =   25
            Tag             =   "!"
            Top             =   480
            Width           =   390
         End
         Begin VB.Label lblPedido 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "001"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   960
            TabIndex        =   23
            Tag             =   "!"
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pedido:"
            Height          =   195
            Left            =   315
            TabIndex        =   22
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblItens 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   960
            TabIndex        =   20
            Tag             =   "!"
            Top             =   720
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblNome_Ven 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do Vendedor"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   960
            TabIndex        =   18
            Tag             =   "!"
            Top             =   960
            Width           =   1380
         End
      End
      Begin MSDBGrid.DBGrid dbgrdProd 
         Bindings        =   "frmPedido.frx":0000
         Height          =   3255
         Left            =   120
         OleObjectBlob   =   "frmPedido.frx":0015
         TabIndex        =   12
         Top             =   2640
         Width           =   7335
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   120
         TabIndex        =   13
         Top             =   5880
         Width           =   7335
         Begin VB.CommandButton cmndbtnSalvar 
            Caption         =   "&Concluir"
            Height          =   735
            Left            =   120
            Picture         =   "frmPedido.frx":0BCE
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnCancel 
            Caption         =   "Ca&ncelar"
            Height          =   735
            Left            =   2520
            Picture         =   "frmPedido.frx":1158
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   735
            Left            =   4920
            Picture         =   "frmPedido.frx":12A2
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   7335
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   315
            Picture         =   "frmPedido.frx":13EC
            ScaleHeight     =   270
            ScaleWidth      =   255
            TabIndex        =   29
            Top             =   412
            Width           =   255
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2040
            Picture         =   "frmPedido.frx":1536
            ScaleHeight     =   270
            ScaleWidth      =   255
            TabIndex        =   28
            Top             =   412
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   3840
            Picture         =   "frmPedido.frx":1680
            ScaleHeight     =   270
            ScaleWidth      =   255
            TabIndex        =   27
            Top             =   412
            Width           =   255
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   5640
            Picture         =   "frmPedido.frx":17CA
            ScaleHeight     =   270
            ScaleWidth      =   255
            TabIndex        =   26
            Top             =   412
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnExc_Prod 
            Caption         =   "E&xcluir"
            Enabled         =   0   'False
            Height          =   615
            Left            =   3720
            TabIndex        =   10
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnEditProd 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   615
            Left            =   1920
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnAddProd 
            Caption         =   "&Adicionar "
            Height          =   615
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnFornada 
            Caption         =   "&Fornada"
            Height          =   615
            Left            =   5520
            TabIndex        =   11
            Top             =   240
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmbxLoja_Click()
    P_AtualizaVenda
End Sub

Private Sub cmbxLoja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnAddProd_Click()
    Dim p As Produto
    If cmbxLoja.Text = "" Then
        MsgBox "Informe primeiro a Loja!", vbExclamation
        cmbxLoja.SetFocus
        Exit Sub
    End If
    
    p = F_BuscaProd(, , False)
    
    
    
    
        If p.qtd_pro = 0 Then
            
            
            MsgBox "Não foi informado uma quantidade válida!", vbExclamation
            
            cmndbtnAddProd.SetFocus
        Else
            If F_BuscaUmCampo("Cod_Pro", "tab_itped", "cod_pro='" & p.cod_pro & "' and num_ped ='" & Me.Tag & "'") <> "" Then
                MsgBox "Este item já se encontra neste pedido!", vbExclamation
                Exit Sub
            End If
            P_InsereProd p.cod_pro, p.preco, p.qtd_pro
             If val(p.pei) > 0 Then
            On Error Resume Next
            With frmPedidoFornada
                .Tag = p.cod_pro
                .frControle.caption = p.nome_pro
                 
                .lblQtd_Ite.caption = FormatVal(p.qtd_pro)
                .lblVenda.caption = Me.Tag
            
            .Show 1
        End With
        
    
            
                P_buscaItens "nome_pro"
            
            
            
        End If
        P_AtualizaVenda
    End If
    cmndbtnAddProd.SetFocus
    
End Sub
Sub P_AtualizaVenda()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iPos As Integer
   On Error GoTo trata_erro
    P_Status "Atualizando Valores..."
    
    'iPos = dtProd.Recordset.AbsolutePosition
    dtProd.Refresh
    With dtProd.Recordset
        
        
        Do While Not .EOF
            cTotal = cTotal + !Total
            .MoveNext
        Loop
        lblValor.caption = FormatVal(CStr(cTotal))
        lblItens.caption = dtProd.Recordset.RecordCount
    End With
        
    
    lblValor.caption = FormatVal(CStr(cTotal))
    
    dtProd.Refresh
    dtProd.Recordset.Move iPos
    
    'lblItens.caption = dtProd.Recordset.RecordCount + dtExt.Recordset.RecordCount + _
    dtPre.Recordset.RecordCount
    
    strSQL = "update tab_pedidos set "
    'valor_ped =" & F_GravaDec(cTotal) & ",
    strSQL = strSQL & " cod_loj='" & Left(cmbxLoja.Text, 3) & "',  "
    strSQL = strSQL & " cod_usu='" & strCod_Usu & "', "
    strSQL = strSQL & " nome_cli='" & txtbxNome_Cli.Text & "', "
    strSQL = strSQL & " endereco_cli='" & txtbxEndereco.Text & "' "
    strSQL = strSQL & " where num_ped='" & Me.Tag & "' "
    banco.Execute strSQL
    P_Status ""
Exit Sub
trata_erro:
    If Err.Number <> 91 Then
        MsgBox Err.Description, vbCritical
    End If
    
    
    P_Status ""
    
End Sub
Sub P_buscaItens(sOrd As String)
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "select tv.*,tp.nome_pro,tp.pei_pro,tp.tipo_pro,tp.preco_pro,tp.custo_pro,tg.nome_cat,tg.nome_cat,"
    strSQL = strSQL & "te.nome_mar,te.cod_mar,ti.*,(tp.preco_pro*ti.qtd_ped) as total "
    strSQL = strSQL & " from tab_produtos as tp,tab_pedidos as tv, tab_categorias as tg,"
    strSQL = strSQL & " tab_marcas as te, tab_Itped as ti "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro and tp.cod_cat=tg.cod_cat "
    strSQL = strSQL & " and tp.cod_mar=te.cod_mar "
    strSQL = strSQL & " and ti.num_ped=tv.num_ped and tv.num_ped='" & Me.Tag & "' "
    strSQL = strSQL & " order by " & sOrd
    dtProd.RecordSource = strSQL
    dtProd.Refresh
            
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub


Sub P_InsereProd(sCod_Pro As String, cUnit As Currency, cQtd As Currency)
    Dim strSQL As String
    Dim iIndex As Integer
    
    strSQL = "insert into tab_itped(num_ped,cod_pro,qtd_ped)values("
     strSQL = strSQL & "'" & Me.Tag & "', "
    strSQL = strSQL & "'" & sCod_Pro & "',"
    
    strSQL = strSQL & F_GravaDec(CStr(cQtd)) & ")"
    
    banco.Execute strSQL
    'RESERVANDO O ESTOQUE
    strSQL = "update tab_produtos set reservado_pro=reservado_pro + " & F_GravaDec(CStr(cQtd))
    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
    banco.Execute strSQL
    strSQL = "update tab_estoque set reservado_pro=reservado_pro + " & F_GravaDec(CStr(cQtd))
    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
    banco.Execute strSQL
End Sub
Private Sub cmndbtnCancel_Click()
    Dim strSQL  As String
    If MsgBox("Tem certeza que deseja cancelar este Pedido?", vbYesNo + vbCritical) = vbYes Then
        With dtProd.Recordset
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                    
                excluir False
                .MoveNext
            Loop
        End With
        strSQL = "Update tab_pedidos set situacao_ped = 0 where num_ped='" & Me.Tag & "' "
        banco.Execute strSQL
        
        strSQL = "delete from tab_itped where num_ped='" & Me.Tag & "' "
        banco.Execute strSQL
        strSQL = "delete from tab_pedforn where num_ped='" & Me.Tag & "' "
        banco.Execute strSQL
        P_ImprimePedidoCancelado Printer, Me.Tag, True
        
        If IsLoad("frmEdiPed") Then
            RefreshDataPosition frmEdiPed.dtVendas
        End If
        Unload Me
        
    End If
End Sub


Private Sub cmndbtnEdit_Pre_Click()
    
    Dim strSQL As String
    Dim cUlt_Qtd As Currency
    With dtPre.Recordset
    
        frmValoresPReVEn.lblCod_Pro.caption = !cod_pro
        frmValoresPReVEn.lblNome_Pro.caption = !nome_pro
        
        frmValoresPReVEn.txtbxPreco.Text = FormatVal(!unit_ite)
        frmValoresPReVEn.txtbxQtd.Text = FormatVal(!qtd_ite)
        frmValoresPReVEn.Tag = Me.Tag
        cUlt_Qtd = !qtd_ite
        On Error Resume Next
        frmValoresPReVEn.Show 1
        If CCur(frmValoresPReVEn.txtbxQtd.Text) <= 0 Then
            MsgBox "Informe um valor decente!", vbExclamation
            Exit Sub
        End If
        strSQL = "Update tab_prevendas set qtd_ite=" & _
        F_GravaDec(frmValoresExt.txtbxQtd.Text)
        
        
        strSQL = strSQL & " where num_ven='" & Me.Tag & "' "
        strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
        banco.Execute strSQL
        'ajustando estoque reservado
        
        Unload frmValoresPReVEn
    End With
    dtPre.Refresh
    
    P_AtualizaVenda

End Sub

Private Sub cmndbtnEditProd_Click()
        Dim iPos As Integer
    Dim strSQL As String
    Dim cUlt_Qtd As Currency
    With dtProd.Recordset
        iPos = .AbsolutePosition
        frmValoresProd.lblCod_Pro.caption = !cod_pro
        frmValoresProd.lblNome_Pro.caption = !nome_pro
        frmValoresProd.txtbxCusto.Text = FormatVal(!custo_pro)
        frmValoresProd.txtbxPreco.Text = FormatVal(!preco_pro)
        frmValoresProd.txtbxQtd.Text = FormatVal(!qtd_ped)
        frmValoresProd.Tag = Me.Tag
        cUlt_Qtd = !qtd_ped
        On Error Resume Next
        frmValoresProd.Show 1
        If CCur(frmValoresProd.txtbxQtd.Text) <= 0 Then
            MsgBox "Informe um valor decente!", vbExclamation
            Exit Sub
        End If
        strSQL = "Update tab_itped set qtd_ite=" & _
        F_GravaDec(frmValoresProd.txtbxQtd.Text)
        strSQL = strSQL & ", unit_ite=" & _
        F_GravaDec(frmValoresProd.txtbxPreco.Text)
        strSQL = strSQL & " where num_ped='" & Me.Tag & "' "
        strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
        banco.Execute strSQL
        'ajustando estoque reservado
        strSQL = "Update tab_produtos set reservado_pro=reservado_pro-" & F_GravaDec(CStr(cUlt_Qtd))
        strSQL = strSQL & "+" & F_GravaDec(frmValoresProd.txtbxQtd.Text)
        strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
        banco.Execute strSQL
        Unload frmValoresProd
    End With
    dtProd.Refresh
    dtProd.Recordset.Move iPos
    P_AtualizaVenda

End Sub

Private Sub cmndbtnEditEx_Click()
    Dim iPos As Integer
    Dim strSQL As String
    Dim cUlt_Qtd As Currency
    With dtPre.Recordset
        iPos = .AbsolutePosition
        frmValoresExt.lblCod_Pro.caption = !cod_pro
        frmValoresExt.lblNome_Pro.caption = !nome_pro
        
        frmValoresExt.txtbxPreco.Text = FormatVal(!unit_ite)
        frmValoresExt.txtbxQtd.Text = FormatVal(!qtd_ite)
        frmValoresExt.Tag = Me.Tag
        cUlt_Qtd = !qtd_ite
        On Error Resume Next
        frmValoresExt.Show 1
        If CCur(frmValoresExt.txtbxQtd.Text) <= 0 Then
            MsgBox "Informe um valor decente!", vbExclamation
            Exit Sub
        End If
        strSQL = "Update tab_venped set qtd_ite=" & _
        F_GravaDec(frmValoresExt.txtbxQtd.Text)
        
        
        strSQL = strSQL & " where num_ven='" & Me.Tag & "' "
        strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
        banco.Execute strSQL
        'ajustando estoque reservado
        
        Unload frmValoresExt
    End With
    dtExt.Refresh
    
    P_AtualizaVenda
End Sub
Sub excluir(bQuestion As Boolean)
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As Recordset
    With dtProd.Recordset
        If bQuestion = True Then
            If MsgBox("Confirma Exclusão do Produto " & !nome_pro & "?", _
            vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        
        iPos = .AbsolutePosition
        strSQL = "delete from tab_itped where cod_pro='" & !cod_pro & "' "
        strSQL = strSQL & " and num_ped='" & Me.Tag & "' "
        banco.Execute strSQL
        'ajustando estoque reservado
        strSQL = "select * from tab_pedforn where num_ped='" & Me.Tag & "' "
        
        Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
            Do While Not .EOF
                 strSQL = "Delete from tab_pedforn "
                strSQL = strSQL & " where num_ped='" & Me.Tag & "' "
                
                strSQL = strSQL & " and cod_pro = '" & !cod_pro & "' "
                strSQL = strSQL & " and fornada_ped = '" & !fornada_ped & "' "
                banco.Execute strSQL
                strSQL = "UPDATE tab_fornadas set reservado_pro= reservado_pro- "
                strSQL = strSQL & F_GravaDec(!qtd_ped)
                strSQL = strSQL & " where "
                'strSQL = strSQL & "  cod_loj = '" & strCod_Loj & "' and  "
                strSQL = strSQL & " cod_pro = '" & !cod_pro & "' "
                strSQL = strSQL & " and fornada_pro = '" & !fornada_ped & "' "
                banco.Execute strSQL
                .MoveNext
            Loop
        End With
        strSQL = "update tab_produtos set reservado_pro=reservado_pro - " & F_GravaDec(CStr(!qtd_ped))
        strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
        banco.Execute strSQL
        strSQL = "update tab_estoque set reservado_pro=reservado_pro - " & F_GravaDec(CStr(!qtd_ped))
        strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
        strSQL = strSQL & " and cod_loj='" & strCod_Loj & "' "
        banco.Execute strSQL
        dtProd.Refresh
        If iPos = .RecordCount - 1 Then
            If dtProd.Recordset.RecordCount > 0 Then
                dtProd.Recordset.MoveLast
            End If
        Else
            dtProd.Recordset.Move iPos
        End If
        P_AtualizaVenda
    End With

End Sub
Private Sub cmndbtnExc_Prod_Click()
    excluir True
End Sub

Private Sub cmndbtnExc_serv_Click()
    Dim iPos As Integer
    Dim strSQL As String
    With dtServ.Recordset
        If MsgBox("Confirma Exclusão do Serviço " & !Descricao_Ser & "?", _
        vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
        iPos = .AbsolutePosition
        strSQL = "delete from tab_venserv where cod_ser='" & !cod_ser & "' "
        strSQL = strSQL & " and num_ven='" & Me.Tag & "' "
        banco.Execute strSQL
        strSQL = "delete from tab_funserv where "
        strSQL = strSQL & "  item_ser='" & dtServ.Recordset!item_ser & "' "
        banco.Execute strSQL
        dtServ.Refresh
        If iPos = .RecordCount - 1 Then
            If dtServ.Recordset.RecordCount > 0 Then
                dtServ.Recordset.MoveLast
            End If
        Else
            dtServ.Recordset.Move iPos
        End If
        P_AtualizaVenda
    End With
End Sub

Private Sub cmndbtnExcEx_Click()
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As Recordset
    With dtExt.Recordset
        If MsgBox("Confirma Exclusão do Produto " & !nome_pro & "?", _
        vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        End If
        iPos = .AbsolutePosition
        strSQL = "delete from tab_venped where cod_pro='" & !cod_pro & "' "
        strSQL = strSQL & " and num_ven='" & Me.Tag & "' "
        banco.Execute strSQL
        'ajustando estoque reservado
        
        
        dtExt.Refresh
        
        P_AtualizaVenda
    End With

End Sub

Private Sub cmndbtnExcl_Pre_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    With dtPre.Recordset
        If MsgBox("Confirma Exclusão do Produto " & !nome_pro & "?", _
        vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        End If
        
        strSQL = "delete from tab_prevendas where cod_pro='" & !cod_pro & "' "
        strSQL = strSQL & " and num_ven='" & Me.Tag & "' "
        banco.Execute strSQL
        'ajustando estoque reservado
        
        
        dtPre.Refresh
        
        P_AtualizaVenda
    End With

End Sub

Private Sub cmndbtnFornada_Click()
    If dbgrdProd.Row = -1 Then
        MsgBox "Selecione um Produto!", vbExclamation
        dbgrdProd.SetFocus
        Exit Sub
    End If
    If dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não há produtos!", vbExclamation
        Exit Sub
    End If
    If val(dtProd.Recordset!pei_pro) = 0 Then
        MsgBox "Este produto não possui fornadas!", vbExclamation
        Exit Sub
    End If
    With frmPedidoFornada
        
        
        .Tag = dtProd.Recordset!cod_pro
        .frControle.caption = dtProd.Recordset!nome_pro
        .lblQtd_Ite.caption = FormatVal(dtProd.Recordset!qtd_ped)
        .lblVenda.caption = Me.Tag
        
        .Show 1
    End With
End Sub

Private Sub cmndbtnSair_Click()
    If IsLoad(frmEdiPed.Name) Then
        frmEdiPed.dtVendas.Refresh
    End If
    Unload Me
End Sub


Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não é possível fechar este pedido, ele não possui itens!", vbExclamation
        Exit Sub
    End If
    
    
    If MsgBox("Deseja Concluir o Pedido? ", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
        
    End If
    On Error Resume Next
    If Not ConfereFornadas Then
        Exit Sub
    End If
             
    strSQL = "update tab_pedidos set situacao_ped='2' where num_ped = '" & Me.Tag & "'"
    banco.Execute strSQL
              
    P_ImprimePedido frmVisualiza, Me.Tag, False
    'frmVisualiza.Refresh
    On Error Resume Next
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
        
        P_ImprimePedido Printer, Me.Tag, True
    End If
    Unload frmVisualiza
    If IsLoad("frmEdiPed") Then
        frmEdiPed.dtVendas.Refresh
    End If
    Unload Me
        
    MsgBox "Pedido Concluído!", vbExclamation
        
    
End Sub

Private Sub dbgrdComiss_GotFocus()
    dbgrdComiss.MarqueeStyle = 3
End Sub

Private Sub dbgrdComiss_LostFocus()
    dbgrdComiss.MarqueeStyle = 0
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command6_Click()

End Sub

Private Sub dbgrdProd_DblClick()
    cmndbtnEditProd_Click
End Sub

Private Sub dbgrdProd_GotFocus()
    dbgrdProd.MarqueeStyle = 3
End Sub

Private Sub dbgrdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEditProd_Click
    End If
End Sub

Private Sub dbgrdProd_LostFocus()
    dbgrdProd.MarqueeStyle = 0
End Sub

Private Sub dbgrdServ_GotFocus()
    dbgrdServ.MarqueeStyle = 3
End Sub

Private Sub dbgrdServ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
    End If
End Sub

Private Sub dbgrdServ_LostFocus()
    dbgrdServ.MarqueeStyle = 0
End Sub

Private Sub dtExt_Reposition()
    If dtExt.Recordset.RecordCount > 0 Then
        cmndbtnExcEx.Enabled = True
        cmndbtnEditEx.Enabled = True
    Else
        cmndbtnExcEx.Enabled = False
        cmndbtnEditEx.Enabled = False
    End If

End Sub
Private Sub dtPre_Reposition()
    If dtPre.Recordset.RecordCount > 0 Then
        cmndbtnEdit_Pre.Enabled = True
        cmndbtnExcl_Pre.Enabled = True
    Else
        cmndbtnEdit_Pre.Enabled = False
        cmndbtnExcl_Pre.Enabled = False
    End If

End Sub
Private Sub dtProd_Reposition()
    If dtProd.Recordset.RecordCount > 0 Then
        cmndbtnExc_Prod.Enabled = True
        cmndbtnEditProd.Enabled = True
    Else
        cmndbtnExc_Prod.Enabled = False
        cmndbtnEditProd.Enabled = False
    End If
End Sub

Private Sub dtServ_Reposition()
    If dtServ.Recordset.RecordCount > 0 And Not dtServ.Recordset.EOF Then
        cmndbtnExc_serv.Enabled = True
        Dim strSQL As String
        strSQL = "select tvf.*, tf.nome_fun from tab_funcionarios as tf,  "
        strSQL = strSQL & "tab_funserv as tvf "
        strSQL = strSQL & " where tvf.cod_fun=tf.cod_fun and tvf.num_ven='" & Me.Tag & "' "
        strSQL = strSQL & " and item_ser='" & dtServ.Recordset.Fields(7).Value & "' "
        dtComiss.RecordSource = strSQL
        dtComiss.Refresh
    Else
        cmndbtnExc_serv.Enabled = False
        
    End If
End Sub

Private Sub Form_Activate()
    
        
    P_Status "Buscando Itens..."
    P_buscaItens "Nome_Pro"
    
    P_Status ""
    P_AtualizaVenda
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA Then
        cmndbtnAddProd_Click
    End If
End Sub

Private Sub Form_Load()
   ' Me.Width = 11115
   ' Me.Height = 6750
    P_CentraForm Me
    P_PreencheLojas cmbxLoja, False
    dtProd.Connect = banco.Connect
'
     MudaAparencia Me
  
End Sub

Private Sub Form_Resize()
'    If Me.Width > 11115 Then
'        Me.Width = 11115
'    End If
'    If Me.Height > 6750 Then
'        Me.Height = 6750
'    End If
    P_CentraObejct frControle, Me
End Sub

Private Sub optnCod_Click()
    P_buscaItens "cod_pro"
End Sub

Private Sub optnDesc_Click()
    P_buscaItens "nome_pro"
End Sub

Private Sub optnEnd_Click()
    P_buscaItens "endereco_pro"
End Sub

Private Sub optnSeq_Click()
    P_buscaItens "seq_ite"
End Sub
Function ConfereFornadas() As Boolean
    Dim strSQL As String
    Dim sMsg As String
    Dim bForn As Boolean
    With dtProd.Recordset
        bForn = True
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            If val(!pei_pro) > 0 Then
                strSQL = F_BuscaUmCampo("sum(qtd_ped)", "tab_pedforn", _
                    "num_ped='" & Me.Tag & "'")
                If strSQL = "" Then
                    strSQL = "0"
                End If
                If CCur(strSQL) < !qtd_ped Then
                    sMsg = sMsg & "No Item (" & .AbsolutePosition + 1 & ")  " & Format(!cod_pro, "00#\.###\-#") & " Restam distribuir " & !qtd_ped - CCur(strSQL) & " Itens " & Chr(13)
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


Private Sub SSTab1_DblClick()

End Sub


Private Sub txtbxEndereco_GotFocus()
    P_SelecTexto txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxEndereco_LostFocus()
    P_AjustaTexto txtbxEndereco
    P_AtualizaVenda
End Sub

Private Sub txtbxNome_Cli_GotFocus()
    P_SelecTexto txtbxNome_Cli
End Sub

Private Sub txtbxNome_cli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Cli_LostFocus()
    P_AjustaTexto txtbxNome_Cli
    P_AtualizaVenda
End Sub
