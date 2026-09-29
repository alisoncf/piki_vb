VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDevolucao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devolução"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11340
   Icon            =   "frmDevolucao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frControle 
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   10935
      Begin VB.Frame frmCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   6120
         TabIndex        =   24
         Top             =   120
         Width           =   4695
         Begin VB.Label lblValor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   675
            Left            =   1920
            TabIndex        =   30
            Tag             =   "!"
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Gerado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   585
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produtos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   28
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblProd 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   1080
            TabIndex        =   27
            Tag             =   "!"
            Top             =   0
            Width           =   135
         End
         Begin VB.Label lblItens 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2280
            TabIndex        =   26
            Tag             =   "!"
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Itens:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   1680
            TabIndex        =   25
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   5895
         Begin VB.TextBox txtbxProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   4
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox txtbxQtd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            TabIndex        =   3
            Text            =   "1"
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmndbtnLocaliza 
            Height          =   375
            Left            =   120
            Picture         =   "frmDevolucao.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Produto"
            Height          =   195
            Left            =   600
            TabIndex        =   7
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Qtd:"
            Height          =   195
            Left            =   5040
            TabIndex        =   6
            Top             =   120
            Width           =   300
         End
         Begin VB.Label lblNome_pro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   720
            TabIndex        =   5
            Top             =   720
            Width           =   45
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   8400
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
         Begin VB.CommandButton cmndbtnSalvar 
            Caption         =   "&Concluir"
            Height          =   855
            Left            =   120
            Picture         =   "frmDevolucao.frx":0294
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnCancel 
            Caption         =   "Ca&ncelar"
            Height          =   855
            Left            =   120
            Picture         =   "frmDevolucao.frx":081E
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   855
            Left            =   120
            Picture         =   "frmDevolucao.frx":0968
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2160
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "&Itens"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   10695
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1230
            Left            =   8280
            ScaleHeight     =   1230
            ScaleWidth      =   2190
            TabIndex        =   23
            Top             =   4560
            Width           =   2190
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               Height          =   1170
               Left            =   -240
               Picture         =   "frmDevolucao.frx":0AB2
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2730
            End
         End
         Begin VB.Data dtProd 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   1  'UseODBC
            Exclusive       =   0   'False
            Height          =   495
            Left            =   5040
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1080
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSDBGrid.DBGrid dbgrdProd 
            Bindings        =   "frmDevolucao.frx":7F11
            Height          =   5535
            Left            =   120
            OleObjectBlob   =   "frmDevolucao.frx":7F26
            TabIndex        =   9
            Top             =   360
            Width           =   8055
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F2 - Buscar Produto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   7080
         Width           =   1410
      End
      Begin VB.Label lblMsg1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Edita o Item"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2400
         TabIndex        =   18
         Top             =   7080
         Width           =   1155
      End
      Begin VB.Label lblMsg2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Exclui o Item"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   17
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Navega no Grid"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5760
         TabIndex        =   16
         Top             =   7080
         Width           =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F7 - Conclui Devolução"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7800
         TabIndex        =   15
         Top             =   7080
         Width           =   1665
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atendente:"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   8085
      Width           =   780
   End
   Begin VB.Label lblNome_Fun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Atendente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   21
      Top             =   8055
      Width           =   2070
   End
   Begin VB.Label lblnum_ven 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5280
      TabIndex        =   20
      Top             =   8055
      Width           =   1215
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controle:"
      Height          =   195
      Left            =   4560
      TabIndex        =   19
      Top             =   8085
      Width           =   630
   End
End
Attribute VB_Name = "frmDevolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean


Sub P_AtualizaVenda()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iPos As Integer
    Dim cItens As Integer
    P_Status "Atualizando Valores..."
    
    iPos = dtProd.Recordset.AbsolutePosition
    dtProd.Refresh
    With dtProd.Recordset
        
        
        Do While Not .EOF
            cTotal = cTotal + !Total
            cItens = cItens + !qtd_ite
            .MoveNext
        Loop
        lblValor.caption = FormatVal(CStr(cTotal))
        lblItens.caption = dtProd.Recordset.RecordCount
    End With
        
    
    lblValor.caption = FormatVal(CStr(cTotal))
    
    dtProd.Refresh
    dtProd.Recordset.Move iPos
    
    lblProd.caption = dtProd.Recordset.RecordCount
    lblItens.caption = cItens

    strSQL = "update tab_devolucoes set valor_Dev =" & F_GravaDec(cTotal)
    'strSQL = strSQL & ", endereco_ven='" & txtbxEndereco_ven.Text & "' "
    strSQL = strSQL & " where num_Dev='" & Me.Tag & "' "
    banco.Execute strSQL
    P_Status ""
End Sub
Sub P_buscaItens()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "select tv.*,tp.nome_pro,tq.p2_pro,tq.c2_pro,tg.nome_cat,tg.nome_cat,"
    strSQL = strSQL & "te.nome_mar,te.cod_mar,ti.*,(ti.unit_ite*ti.qtd_ite) as total "
    strSQL = strSQL & " from tab_produtos as tp,tab_devolucoes as tv, tab_categorias as tg,"
    strSQL = strSQL & " tab_marcas as te, tab_Itdev as ti,tab_composicao as tq  "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro and tp.cod_cat=tg.cod_cat "
    strSQL = strSQL & " and tq.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tp.cod_mar=te.cod_mar "
    strSQL = strSQL & " and ti.num_Dev=tv.num_Dev and tv.num_dev='" & Me.Tag & "' "
    strSQL = strSQL & " order by seq_ite "
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    
            
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub


Sub P_InsereProd(sCod_Pro As String, cUnit As Currency, cQtd As Currency)
    Dim strSQL As String
    Dim iIndex As Integer
    strSQL = F_BuscaUmCampo("max(seq_ite)", "tab_itvendas", "num_ven='" & Me.Tag & "'")
    If strSQL = "" Then
        iIndex = 1
    Else
        iIndex = val(strSQL) + 1
    End If
    strSQL = "insert into tab_itdev(num_Dev,cod_pro,unit_ite,qtd_ite,seq_ite)values("
    
    strSQL = strSQL & "'" & Me.Tag & "', "
    strSQL = strSQL & "'" & sCod_Pro & "',"
    strSQL = strSQL & F_GravaDec(CStr(cUnit)) & ","
    strSQL = strSQL & F_GravaDec(CStr(cQtd)) & ","
    strSQL = strSQL & iIndex & ") "
    banco.Execute strSQL
    
    
        
End Sub
Sub P_AumentaQTD(sCod_Pro As String, cQtd As Currency)
    Dim strSQL As String
    Dim iIndex As Integer
    
    strSQL = "Update tab_itDev set "
    
    strSQL = strSQL & "qtd_ite = qtd_ite+"
    strSQL = strSQL & F_GravaDec(CStr(cQtd)) & " "
    strSQL = strSQL & " where num_Dev='" & Me.Tag & "'  "
    strSQL = strSQL & " and cod_pro ='" & sCod_Pro & "' "
    banco.Execute strSQL
    
    
    
    
    
End Sub

Sub P_InsereComiss(cod_fun As String, Num_Ven As String, cod_ser As String, _
                    valor_ser As Currency, valorcomiss_ser As Currency, _
                    percent_ser As Currency, unit_ser As Currency, item_ser As Integer)
    Dim strSQL As String
    strSQL = "insert into tab_funserv(cod_fun,num_ven,valor_ser,valorcomiss_ser,"
    strSQL = strSQL & "percent_ser,item_ser)Values("
    strSQL = strSQL & "'" & cod_fun & "',"
    strSQL = strSQL & "'" & Num_Ven & "',"
    
    strSQL = strSQL & F_GravaDec(valor_ser) & ","
    strSQL = strSQL & F_GravaDec(valorcomiss_ser) & ","
    strSQL = strSQL & F_GravaDec(percent_ser) & ","
    
    strSQL = strSQL & item_ser & ") "
    banco.Execute strSQL
    
End Sub

Private Sub cmndbtnCancel_Click()
    Dim strSQL  As String
    If MsgBox("Tem certeza que deseja cancelar esta Devolucao?", vbYesNo + vbCritical) = vbYes Then
        
        
        strSQL = "delete from tab_devolucoes where num_Dev='" & Me.Tag & "' "
        banco.Execute strSQL
        strSQL = "delete from tab_itDev where num_Dev='" & Me.Tag & "' "
        banco.Execute strSQL
        
    
        If IsLoad("frmEdiDevolucoes") Then
            RefreshDataPosition frmEdiVendas.dtVendas
        End If
        Unload Me
        
    End If
End Sub




Sub excluir(bQuestion As Boolean)
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As Recordset
    With dtProd.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If bQuestion = True Then
            If MsgBox("Confirma Exclusão do Produto " & !nome_pro & "?", _
            vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        iPos = .AbsolutePosition
        strSQL = "delete from tab_itdev where cod_pro='" & !cod_pro & "' "
        strSQL = strSQL & " and num_Dev='" & Me.Tag & "' "
        banco.Execute strSQL
        'ajustando estoque reservado

        
        dtProd.Refresh
        If iPos = .RecordCount - 1 Then
            If dtProd.Recordset.RecordCount > 0 Then
                dtProd.Recordset.MoveLast
            End If
        Else
            dtProd.Recordset.Move iPos
        End If
        P_AtualizaVenda
        If dtProd.Recordset.RecordCount = 0 Then
            txtbxQtd.Enabled = False
            txtbxProduto.Enabled = True
            txtbxProduto.SetFocus
        End If
    End With
End Sub
Private Sub cmndbtnLocaliza_Click()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxProduto.Text, cmndbtnLocaliza.Tag)
    If p.cod_pro <> "" Then
        txtbxProduto.Text = p.codbarra_pro
        lblNome_pro.caption = p.nome_pro
        txtbxProduto.SetFocus
        SendKeys Chr(13)
    End If
    cmndbtnLocaliza.Tag = ""
End Sub

Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    If IsLoad(frmEdiVendas.Name) Then
        frmEdiVendas.dtVendas.Refresh
    End If
    If dtProd.Recordset.RecordCount = 0 Then
        strSQL = "delete from tab_vendas where num_ven='" & Me.Tag & "' "
        banco.Execute strSQL
        strSQL = "delete from tab_itvendas where num_ven='" & Me.Tag & "' "
        banco.Execute strSQL
        
    End If
    Unload Me
End Sub


Private Sub cmndbtnSalvar_Click()

    Dim strSQL As String
    If dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não há itens na sua venda!", vbExclamation
        Exit Sub
    End If
    
    
    If MsgBox("Deseja realmente concluir a Devolução?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = "update tab_devolucoes set situacao_dev='2' where num_dev='" & Me.Tag & "'"
        banco.Execute strSQL
        With dtProd.Recordset
            .MoveFirst
            Do While Not .EOF
                strSQL = "update tab_composicao set e2_pro=e2_pro+" & F_GravaDec(!qtd_ite)
                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                .MoveNext
            Loop
        End With
        strSQL = F_GravaCai("D", Date, lblValor.caption, "DEVOLUÇÃO " & _
            lblnum_ven.caption & "-" & Date, lblnum_ven.caption, "DIN", strCod_Cai, 0, 0, strCod_Cai)
    
        strSQL = F_GeraNovaDevolucao("00000", strCod_Usu)
        lblnum_ven.caption = strSQL
        Me.Tag = strSQL
        P_buscaItens
        P_AtualizaVenda
        txtbxProduto.Text = ""
        txtbxProduto.SetFocus
        dtProd.Refresh
    End If
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

Private Sub dbgrdProd_GotFocus()
    
    dbgrdProd.MarqueeStyle = 3
    
    
End Sub

Private Sub dbgrdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyDelete Then
        excluir False
    ElseIf KeyCode = vbKeyP Or KeyCode = vbKeyReturn Then
        txtbxProduto_GotFocus
        txtbxProduto.SetFocus
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
    If frControle.Tag <> "" Then
        Exit Sub
    End If
    
    If dtExt.Recordset.RecordCount > 0 Then
        cmndbtnExcEx.Enabled = True
        cmndbtnEditEx.Enabled = True
    Else
        cmndbtnExcEx.Enabled = False
        cmndbtnEditEx.Enabled = False
    End If

End Sub
Private Sub dtPre_Reposition()
    If frControle.Tag <> "" Then
        Exit Sub
    End If
    If dtPre.Recordset.RecordCount > 0 Then
        cmndbtnEdit_Pre.Enabled = True
        cmndbtnExcl_Pre.Enabled = True
    Else
        cmndbtnEdit_Pre.Enabled = False
        cmndbtnExcl_Pre.Enabled = False
    End If

End Sub
Private Sub dtProd_Reposition()
    If frControle.Tag <> "" Then
        Exit Sub
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
    lblnum_ven.caption = Me.Tag
    lblNome_Fun.caption = strCod_Usu & "-" & strNome_Usu
    If bActive = True Then Exit Sub
    P_Status "Buscando Itens..."
    txtbxProduto.SetFocus
    P_buscaItens
    P_AtualizaVenda
    P_Status ""
    bActive = True
    If strCod_Cai = "" Then
        MsgBox "Não existe caixa para este computador, o caixa ficará para o '01' ", vbExclamation
        strCod_Cai = "01"
    End If
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyF5 Then
        
        dbgrdProd.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
        excluir True
    ElseIf KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF2 Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub Form_Load()
   ' Me.Width = 11115
   ' Me.Height = 6750
    bActive = False
    'P_CentraForm Me
    dtProd.Connect = banco.Connect
    'dtExt.Connect = banco.Connect
    'dtPre.Connect = banco.Connect
     MudaAparencia Me
'    sstb.Tab = 0
    P_CentraObejct frControle, Me, False
End Sub

Private Sub Form_Resize()
'    If Me.Width > 11115 Then
'        Me.Width = 11115
'    End If
'    If Me.Height > 6750 Then
'        Me.Height = 6750
'    End If
    P_CentraObejct frControle, Me, False
End Sub

Sub editaItem()
    If dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que ser Editado!", vbExclamation
        Exit Sub
    End If
    With dtProd.Recordset
        txtbxProduto.Text = F_BuscaUmCampo("codbarra_pro", "tab_produtos", _
            "cod_pro='" & !cod_pro & "'")
        lblNome_pro.caption = F_BuscaUmCampo("nome_pro", "tab_produtos", "cod_pro='" & !cod_pro & "'")
        txtbxProduto.Enabled = False
        txtbxQtd.Text = !qtd_ite
        txtbxQtd.Tag = !qtd_ite
        txtbxQtd.Enabled = True
        txtbxQtd.SetFocus
        
    End With
End Sub

Private Sub optnCod_Click()
    txtbxProduto.SetFocus
End Sub

Private Sub optnDesc_Click()
    txtbxProduto.SetFocus
End Sub

Private Sub optnEnd_Click()
    
End Sub

Private Sub optnSeq_Click()
    P_buscaItens
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
                strSQL = F_BuscaUmCampo("sum(qtd_ite)", "tab_venforn", _
                    "num_ven='" & Me.Tag & "'")
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


Private Sub SSTab1_DblClick()

End Sub

Private Sub sstb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If sstb.Tab = 0 Then
            cmndbtnAddProd.SetFocus
        ElseIf sstb.Tab = 1 Then
            cmndbtnAddExr.SetFocus
        ElseIf sstb.Tab = 2 Then
            cmndbtnAddPre.SetFocus
        
        End If
    End If
End Sub

Private Sub txtbxEndereco_ven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    SendKeys Chr(9)
        End If
End Sub


Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iPos As Integer
    Dim p As Produto
    Dim rPesq As Recordset
    If KeyCode = vbKeyReturn Then
        If txtbxProduto.Text <> "" Then
            insere
        End If
    ElseIf KeyCode = vbKeyTab Then
        txtbxQtd.Enabled = True
    ElseIf KeyCode = vbKeyDown Then
        dbgrdProd.SetFocus
    ElseIf KeyCode = vbKeyF6 Then
        cmndbtnLocaliza_Click
    ElseIf KeyCode = vbKeyF11 Then
        
        frmReemissao.Show 1
    End If
    
End Sub
Sub insere()
    Dim iPos As Integer
    
        strSQL = "select tc.nome_Cat, tm.nome_mar,"
        strSQL = strSQL & "tp.*,tq.c2_pro,tq.e2_pro,tq.p2_pro "
        strSQL = strSQL & " from tab_produtos as tp,tab_categorias as tc, "
        strSQL = strSQL & " tab_marcas as tm,tab_composicao as tq  "
        strSQL = strSQL & " where codbarra_PRO = '"
        strSQL = strSQL & txtbxProduto.Text & "' "
        
        strSQL = strSQL & " and tq.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and tc.cod_cat=tp.cod_Cat and tm.cod_mar=tp.cod_mar "
        strSQL = strSQL & " order by tp.nome_pro"
        Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        If rPesq.RecordCount = 0 Then
            MsgBox "Produto não encontrado!", vbExclamation
            txtbxProduto_GotFocus
            Exit Sub
        End If
        If rPesq.RecordCount = 1 Then
            With rPesq
                lblNome_pro.caption = !nome_pro
                If F_BuscaUmCampo("cod_pro", "tab_itdev", "cod_pro='" & _
                    !cod_pro & "' and num_dev='" & Me.Tag & "'") = "" Then
                    
                    
                    P_InsereProd !cod_pro, !p2_pro, CCur(txtbxQtd.Text)
                    P_buscaItens
                    dtProd.Recordset.Move dtProd.Recordset.RecordCount - 1
                    txtbxProduto_GotFocus
                    txtbxProduto.SetFocus
                    
                Else
                    P_AumentaQTD !cod_pro, CCur(txtbxQtd.Text)
                    iPos = dtProd.Recordset.AbsolutePosition
                    P_buscaItens
                    dtProd.Recordset.Move iPos
                    txtbxProduto_GotFocus
                    txtbxProduto.SetFocus
                    
                End If
                    
                    .Close
            End With
            
            P_AtualizaVenda
        
        ElseIf rPesq.RecordCount > 1 Then
            
            Dim p As Produto
            p = F_BuscaProdSaida(txtbxProduto.Text, "B")
            If p.cod_pro <> "" Then
                txtbxProduto.Text = p.codbarra_pro
                lblNome_pro.caption = p.nome_pro
                
                P_InsereProd p.cod_pro, p.p2_pro, CCur(txtbxQtd.Text)
                P_buscaItens
                dtProd.Recordset.Move dtProd.Recordset.RecordCount - 1
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
            End If
            
        End If
        
End Sub
Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iPos As Integer
    On Error GoTo trata_erro
    If KeyCode = vbKeyReturn Then
        If CCur(txtbxQtd.Text) = "0" Then
            MsgBox "Informe uma quantidade válida!", vbExclamation
            txtbxQtd.SetFocus
        End If
        P_AumentaQTD dtProd.Recordset!cod_pro, CCur(txtbxQtd.Text) - dtProd.Recordset!qtd_ite
        iPos = dtProd.Recordset.AbsolutePosition
        P_buscaItens
        dtProd.Recordset.Move iPos
        txtbxQtd.Text = "1"
        txtbxQtd.Enabled = False
        txtbxProduto.Text = ""
        txtbxProduto.Enabled = True
        
        P_AtualizaVenda
        txtbxProduto.SetFocus
    End If
Exit Sub
trata_erro:
    MsgBox Err.Description, vbcre
End Sub

Private Sub txtbxQtd_LostFocus()
    txtbxQtd.Text = "1"
    txtbxQtd.Enabled = False
    txtbxProduto.Text = ""
    txtbxProduto.Enabled = True
    txtbxProduto.SetFocus
    
    
    
End Sub


