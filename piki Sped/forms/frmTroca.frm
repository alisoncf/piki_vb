VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTroca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troca de Produtos"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   Icon            =   "frmTroca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   3120
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   9360
      Picture         =   "frmTroca.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Concluir"
      Height          =   615
      Left            =   8040
      Picture         =   "frmTroca.frx":06F5
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtbxVendedor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   32
      Text            =   "Escolha o Vendedor"
      Top             =   6255
      Width           =   3015
   End
   Begin VB.CommandButton cmndbtnBuscaFun 
      Height          =   315
      Left            =   3960
      Picture         =   "frmTroca.frx":0B9D
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6240
      Width           =   255
   End
   Begin VB.CommandButton cmndbtnDesconto 
      Caption         =   "De&sconto:"
      Height          =   255
      Left            =   9600
      TabIndex        =   30
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtbxDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9600
      TabIndex        =   29
      Text            =   "0,00"
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1440
      ScaleHeight     =   705
      ScaleWidth      =   5025
      TabIndex        =   23
      Top             =   120
      Width           =   5055
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc:"
         Height          =   195
         Left            =   2400
         TabIndex        =   28
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblVendedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   690
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2400
         TabIndex        =   26
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblDataVen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   390
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6600
      ScaleHeight     =   705
      ScaleWidth      =   2865
      TabIndex        =   16
      Top             =   120
      Width           =   2895
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Abater"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblAbater 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Acrescer:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblAcresc 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final Previsto:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   990
      End
      Begin VB.Label lblFinalPrev 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   480
         Width           =   390
      End
   End
   Begin VB.TextBox txtbxQtd_alt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   3240
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid Msflxgrd 
      Height          =   4095
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7223
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Nº da Venda:"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      Begin VB.TextBox txtbxNum_Ven 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   10455
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7920
         TabIndex        =   12
         Text            =   "0,00"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6480
         TabIndex        =   9
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnAdd 
         Caption         =   "&Adicionar"
         Height          =   375
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   5640
         TabIndex        =   7
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   315
         Left            =   4920
         Picture         =   "frmTroca.frx":0C96
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7920
         TabIndex        =   13
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6480
         TabIndex        =   8
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblNome_pro 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   11
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5640
         TabIndex        =   6
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   555
      End
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vendedor:"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   33
      Top             =   6255
      Width           =   840
   End
End
Attribute VB_Name = "frmTroca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnAdd_Click()
    Dim strSQL As String
    Dim iCount As Integer
    If CCur(txtbxTotal.Text) = 0 Then
        MsgBox "Não foi informado um valor válido!", vbCritical
        txtbxQtd.SetFocus
        Exit Sub
    End If
    If F_InFlxgrd(msflxgrd, FormatCodigo(txtbxProduto.tag), 1) > -1 Then
        MsgBox "Este produto já existe"
        Exit Sub
    End If
    If cmndbtnAdd.tag = "" Then
        If MsgBox("Confirma inclusão do Produto " & txtbxProduto.Text & "?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "" & Chr(9) & FormatCodigo(txtbxProduto.tag) & Chr(9)
            strSQL = strSQL & txtbxProduto.Text & Chr(9) & txtbxQtd.Text & Chr(9)
            strSQL = strSQL & txtbxUnit.Text & Chr(9) & txtbxTotal.Text & Chr(9)
            strSQL = strSQL & txtbxTotal.Text & Chr(9) & txtbxQtd.Text & Chr(9) & "S"
            strSQL = strSQL & Chr(9) & txtbxProduto.tag
            msflxgrd.AddItem strSQL
            msflxgrd.row = msflxgrd.rows - 1
            For iCount = 0 To msflxgrd.Cols - 1
                
                msflxgrd.col = iCount
                msflxgrd.CellBackColor = vbInfoBackground
                msflxgrd.CellForeColor = vbBlue
            Next
            lblAcresc.caption = FormatVal(CCur(lblAcresc.caption) + CCur(txtbxTotal.Text))
            lblFinalPrev.caption = FormatVal(CCur(lblAcresc.caption) - CCur(lblAbater.caption))
            lblFinalPrev.tag = lblFinalPrev.caption
        End If
    Else
    
    End If
End Sub

Private Sub cmndbtnBuscaFun_Click()
    Dim f As funcionario.funcionario
    f = F_SelectVendedor(txtbxVendedor.Text)
    If f.cod_fun <> "" Then
        txtbxVendedor.Text = f.guerra_fun
        txtbxVendedor.tag = f.cod_fun
    End If
    
End Sub

Private Sub cmndbtnLocaliza_Click()
    Dim p As Produto
    
    p = F_BuscaProdSaida(txtbxProduto.Text)
    txtbxProduto.Text = p.nome_pro
    txtbxProduto.tag = p.cod_pro
    
    txtbxUnit.Text = p.pv_pro
    
    If p.pq_pro > 0 Then
        If MsgBox("Este produto está em queima. Deseja lançar valor promocional?", vbYesNo + vbQuestion) = vbYes Then
            txtbxUnit.Text = p.pq_pro
        End If
        
    End If
    
    
    txtbxQtd.SetFocus
End Sub
Sub P_InsereProd(sCod_Pro As String, cUnit As Currency, cQtd As Currency)
    Dim strSQL As String
    Dim iIndex As Integer
    strSQL = F_BuscaUmCampo("max(seq_ite)", "tab_itvendas", "num_ven='" & Me.tag & "'")
    If strSQL = "" Then
        iIndex = 1
    Else
        iIndex = val(strSQL) + 1
    End If
    strSQL = "insert into tab_itvendas(num_ven,cod_pro,unit_ite,qtd_ite,seq_ite)values("
    
    strSQL = strSQL & "'" & Me.tag & "', "
    strSQL = strSQL & "'" & sCod_Pro & "',"
    strSQL = strSQL & F_GravaDec(CStr(cUnit)) & ","
    strSQL = strSQL & F_GravaDec(CStr(cQtd)) & ","
    strSQL = strSQL & iIndex & ") "
    conn.Execute strSQL
    'RESERVANDO O ESTOQUE
    strSQL = "update tab_composicao set er_pro=er_pro - " & F_GravaDec(CStr(cQtd))
    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
    conn.Execute strSQL
        
End Sub



Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If F_VerificaAltera = False And F_VerificaNovoProduto = False Then
        MsgBox "Não há alterações na troca!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma conclusão da troca?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    If txtbxVendedor.tag = "" Then
        MsgBox "Informe o vendedor responsável pela troca!", vbExclamation
        txtbxVendedor.SetFocus
        Exit Sub
    End If
    conclui
End Sub
Sub conclui()
    Dim strSQL As String
    Dim sSeq As String
    Dim iCount As Integer
    Dim cQtd_Ant As Currency
    Dim cQtd_Atual As Currency
    Dim sPro As String
    Dim sCod_Pro As String
    Dim cFinal As Currency
    Dim sSeq_ite As String
    Dim sUnit_Ite As String
    Dim cAux As Currency
    Dim cIndice As Currency
    sSeq = MaxKeySQL("tab_trocas", "seq_tro", "", "T00000#")
    
    If CCur(lblFinalPrev.caption) < 0 Then
        
        
            If MsgBox("Deseja realmente concluir uma troca com valor menor que a venda?", vbYesNo + vbCritical) = vbNo Then
                Exit Sub
            End If
        
    End If
    
    If CCur(lblFinalPrev.caption) > 0 Then
            frmTipoVenda.operacao = "Troca"
            frmTipoVenda.lblNome_Cli.caption = lblCliente.caption
            frmTipoVenda.lblNome_Cli.tag = lblCliente.tag
            frmTipoVenda.tag = txtbxNum_Ven.Text
            frmTipoVenda.lblnum_ven.caption = txtbxNum_Ven.Text
            frmTipoVenda.lblValor_Ven.caption = lblFinalPrev.caption
            frmTipoVenda.lblValorFinal.caption = lblFinalPrev.caption
            frmTipoVenda.lblVendedor.caption = strNome_Usu
            frmTipoVenda.txtbxDesconto.Text = FormatVal(txtbxDesconto.Text)
            frmTipoVenda.Show 1
            If frmTipoVenda.tag = "" Then
                MsgBox "Os Valores não foram informados corretamente!", vbExclamation
                
                Me.tag = ""
                Unload frmTipoVenda
                
                Exit Sub
            End If
            Unload frmTipoVenda
    ElseIf CCur(lblFinalPrev.caption) < 0 Then
        strSQL = F_GravaCai("D", Date, FormatVal(-1 * CCur(lblFinalPrev.caption)), "TROCA REF. VENDA " & txtbxNum_Ven.Text & "EM " & Date, txtbxNum_Ven.Text, "DIN", "", 0, 0, strCod_Cai, strCod_Tur, "000", "S")
        
    End If
    
    
    strSQL = "insert into tab_trocas("
    strSQL = strSQL & "Num_Ven, data_tro, cod_usu, seq_tro, abat_tro,acres_tro,desconto_tro)values("
    strSQL = strSQL & "'" & txtbxNum_Ven.Text & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & txtbxVendedor.tag & "', "
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaDec(lblAbater.caption) & ", "
    strSQL = strSQL & F_GravaDec(lblAcresc.caption) & ", "
    strSQL = strSQL & F_GravaDec(txtbxDesconto.Text) & ") "
    conn.Execute strSQL
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 0: sPro = .Text
            .col = 1: sCod_Pro = F_SoNumero(.Text)
            .col = 3: cQtd_Atual = CCur(.Text)
            .col = 4: sUnit_Ite = (.Text)
            .col = 7: cQtd_Ant = CCur(.Text)
            .col = 6: cFinal = CCur(.Text)
            If sPro <> "" Then
                If cQtd_Ant <> cQtd_Atual Then
                    strSQL = "insert into tab_ittrocas(seq_tro,cod_pro,qtdant_itt,qtdatual_itt,unit_itt)values("
                    strSQL = strSQL & "'" & sSeq & "',"
                    strSQL = strSQL & "'" & sCod_Pro & "',"
                    strSQL = strSQL & F_GravaDec(cQtd_Ant) & ",  "
                    strSQL = strSQL & F_GravaDec(cQtd_Atual) & ",  "
                    strSQL = strSQL & F_GravaDec(cFinal) & ")  "
                    conn.Execute strSQL
                    strSQL = "update tab_itvendas set "
                    strSQL = strSQL & "qtd_ite=" & F_GravaDec(cQtd_Atual) & ","
                    If cQtd_Ant > cQtd_Atual Then
                        strSQL = strSQL & "final_ite=" & F_GravaDec(cFinal / cQtd_Ant * cQtd_Atual)
                    Else
                        cIndice = (cQtd_Ant * cFinal) / CCur(lblFinal.caption)
                        cAux = (cFinal / cQtd_Ant * cQtd_Atual)
                        strSQL = strSQL & "final_ite=" & F_GravaDec((CCur(lblFinalPrev.caption) * cIndice) + (CCur(lblFinal.caption) * cIndice))
                    End If
                    strSQL = strSQL & " where num_ven='" & txtbxNum_Ven.Text & "' "
                    strSQL = strSQL & " and seq_ite=" & sPro
                    conn.Execute strSQL
                Else
                    
                
                End If
                If cQtd_Ant > cQtd_Atual Then
                    strSQL = "update " & TAB_PRODUTOS & " set er_pro=er_pro+" & F_GravaDec(cQtd_Ant - cQtd_Atual)
                    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
                    conn.Execute strSQL
                ElseIf cQtd_Ant < cQtd_Atual Then
                    strSQL = "update " & TAB_PRODUTOS & " set er_pro=er_pro-" & F_GravaDec(cQtd_Atual - cQtd_Ant)
                    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
                    conn.Execute strSQL
                
                End If

                
            Else
                'inserir produto
                sSeq_ite = MaxKeySQL("tab_itvendas", "seq_ite", "num_ven='" & txtbxNum_Ven.Text & "'", "")
                strSQL = "insert into tab_itvendas(num_ven,cod_pro,unit_ite,qtd_ite,seq_ite,desc_ite,final_ite)values( "
                strSQL = strSQL & "'" & txtbxNum_Ven.Text & "',"
                .col = 1
                strSQL = strSQL & "'" & F_SoNumero(.Text) & "',"
                .col = 4
                strSQL = strSQL & F_GravaDec(.Text) & ","
                .col = 3
                strSQL = strSQL & F_GravaDec(.Text) & ","
                strSQL = strSQL & sSeq_ite & ","
                .col = 2
                strSQL = strSQL & "'" & .Text & "',"
                .col = 6
                strSQL = strSQL & F_GravaDec(.Text) & ")"
                conn.Execute strSQL
                .col = 3
                strSQL = "update " & TAB_PRODUTOS & " set er_pro=er_pro-" & F_GravaDec(.Text)
                .col = 1
                strSQL = strSQL & " where cod_pro='" & F_SoNumero(.Text) & "' "
                conn.Execute strSQL
            End If
        Next
        MsgBox "Troca Concluída!", vbInformation
        imprimeTroca frmVisualiza.Pic
        frmVisualiza.Show 1
        If frmVisualiza.tag <> "" Then
            imprimeTroca Printer
            Printer.EndDoc
        End If
        Unload frmVisualiza
        txtbxNum_Ven.Text = ""
        msflxgrd.rows = 1
        Unload Me
    End With
    
    
End Sub
    
Function F_VerificaAltera() As Boolean
    Dim iCount As Integer
    Dim cQtd_Atual As Currency
    Dim cQtd_Anterior As Currency
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 3: cQtd_Atual = CCur(.Text)
            .col = 7: cQtd_Anterior = CCur(.Text)
            If cQtd_Anterior <> cQtd_Atual Then
                        
                F_VerificaAltera = True
                Exit Function
            End If
        Next
        F_VerificaAltera = False
    End With
End Function
Function F_VerificaNovoProduto() As Boolean
    Dim iCount As Integer
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 0
            If .Text = "" Then
                        
                F_VerificaNovoProduto = True
                Exit Function
            End If
        Next
        F_VerificaNovoProduto = False
    End With
End Function

Private Sub cmndbtnDesconto_Click()
    Dim sUsu As String
    
    If CCur(lblFinalPrev.caption) <= 0 Then
        MsgBox "Operação inválida!", vbCritical
        Exit Sub
    End If
    
    sUsu = F_SenhaCodBar
    If sUsu <> "" Then
        txtbxDesconto.Enabled = True
        cmndbtnDesconto.tag = sUsu ' deixar 100% na troca
        'cmndbtnDesconto.tag = "100"
        txtbxDesconto.SetFocus
    End If
    
End Sub
Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    Dim cDesc As Currency
    On Error GoTo trata_erro
    Dim descontoDado As Currency
    'P_ValorLostFocus txtbxDesconto
    cDesc = CCur(F_BuscaUmCampo("desconto_usu", TAB_DEFUSU, _
        "cod_usu='" & cmndbtnDesconto.tag & "'"))
    descontoDado = CCur(txtbxDesconto.Text) / CCur(lblFinalPrev.caption) * 100
    
    If CCur(txtbxDesconto.Text) > CCur(lblFinalPrev.caption) * cDesc / 100 Then
        If MsgBox("Seu desconto previsto é de " & cDesc & "% e está dando um desconto de " & descontoDado & "%. Deseja continuar?", vbYesNo + vbExclamation) = vbNo Then
            txtbxDesconto.SetFocus
            Exit Sub
        End If
    End If
    lblFinalPrev.caption = FormatVal(lblFinalPrev.tag) - CCur(txtbxDesconto.Text)
    txtbxDesconto.Enabled = False
Exit Sub
trata_erro:
    P_MostraErro
    txtbxDesconto.Text = "0"
End Sub
Private Sub Form_Click()
    txtbxQtd_alt.Visible = False
End Sub
Sub imprimeTroca(obj As Object)
    On Error Resume Next
    Dim strSQL As String
    Dim sSeq As String
    Dim iCount As Integer
    Dim cQtd_Ant As Currency
    Dim cQtd_Atual As Currency
    Dim sPro As String
    Dim sCod_Pro As String
    Dim cFinal As Currency
    Dim sSeq_ite As String
    Dim sUnit_Ite As String
    P_CabRelTimbr "TROCA REFERENTE À VENDA " & txtbxNum_Ven.Text & " EM " & Date, obj
    P_ImprX obj, "Venda: " & txtbxNum_Ven.Text, 0
    P_ImprX obj, "" & lblCliente.caption, 30, False
    P_ImprX obj, "" & lblDataVen.caption, 0
    P_ImprX obj, "" & lblVendedor.caption, 45
    P_ImprX obj, "" & lblValor.caption, 100
    P_ImprX obj, "" & lblDesc.caption, 160, False
    P_ImprLinha obj
    P_ImprX obj, "Inicial: " & lblInicial.caption, 0
    P_ImprX obj, "À Vista: " & lblAvIsta.caption, 30
    P_ImprX obj, "A Prazo: " & lblAPrazo.caption, 60
    P_ImprX obj, "Financeira: " & lblFinan.caption, 90
    
    
    P_ImprX obj, "Desconto: R$ " & lblDesconto.caption, 130
    P_ImprX obj, "Final: " & lblFinal.caption, 170, False
    P_ImprLinha obj, "="
    With msflxgrd
        obj.FontBold = True
        P_ImprX obj, "...", 0
        P_ImprX obj, "Cód", 5
        P_ImprX obj, "Descrição", 24
        P_ImprX obj, "Qtd", 120, True, True
        P_ImprX obj, "Unit", 135, True, True
        P_ImprX obj, "Total", 150, True, True
        P_ImprX obj, "Final", 165, True, True
        P_ImprX obj, "Inicial", 180, True, True
        P_ImprX obj, "Ref:", 182, False
        obj.FontBold = False
        P_ImprLinha obj
        For iCount = 1 To msflxgrd.rows - 1
            .row = iCount
            .col = 0: sPro = .Text
            .col = 3: cQtd_Atual = CCur(.Text)
            .col = 7: cQtd_Ant = CCur(.Text)
            If cQtd_Atual <> cQtd_Ant Then
                obj.FontBold = True
                obj.ForeColor = vbBlue
            ElseIf sPro = "" Then
                obj.FontBold = True
                obj.ForeColor = vbRed
            Else
                obj.FontBold = False
                obj.ForeColor = vbBlack
            End If
            .col = 0: P_ImprX obj, .Text, 0
            .col = 1: P_ImprX obj, .Text, 5
            .col = 2: P_ImprX obj, .Text, 24
            .col = 3: P_ImprX obj, .Text, 120, True, True
            .col = 4: P_ImprX obj, .Text, 135, True, True
            .col = 5: P_ImprX obj, .Text, 150, True, True
            .col = 6: P_ImprX obj, .Text, 165, True, True
            .col = 7: P_ImprX obj, .Text, 180, True, True
            .col = 9: P_ImprX obj, .Text, 182, False
            obj.FontBold = False
            obj.ForeColor = vbBlack
        Next
        P_ImprLinha obj
        P_ImprX obj, "A Abater: " & lblAbater.caption, 10
        P_ImprX obj, "A Acrescer: " & lblAcresc.caption, 60
        P_ImprX obj, "Desconto: " & txtbxDesconto.Text, 110
        P_ImprX obj, "Final Previsto: " & lblFinalPrev.caption, 160, False
        P_ImprLinha obj, "_"
    End With
End Sub
Private Sub Form_Load()
    adc.ConnectionString = conn.ConnectionString

    With msflxgrd
        .FixedCols = 0
        .FixedRows = 1
        .Cols = 10
        .rows = 1
        .col = 0:  .Text = "..."
        .col = 1: .Text = "Cód"
        .col = 2: .Text = "Descrição"
        .col = 3: .Text = "Qtd"
        .col = 4: .Text = "Unit"
        .col = 5: .Text = "Total "
        .col = 6: .Text = "Final "
        .col = 7: .Text = "Qtd Ini."
        .col = 8: .Text = "*"
        .col = 9: .Text = "Ref."
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 400
        .ColWidth(9) = 2000
        
    End With
    
End Sub

Private Sub msflxgrd_Click()
    txtbxQtd_alt.Visible = False
        
End Sub

Private Sub msflxgrd_DblClick()
    If msflxgrd.col = 3 Then
        msflxgrd_KeyDown 13, 0
    End If
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim cQtd As Currency
    Dim cValor As Currency
    Dim cTotal As Currency
    If KeyCode = vbKeyReturn Then
        msflxgrd.col = 3
        txtbxQtd_alt.Left = msflxgrd.ColPos(3) + msflxgrd.Left
        txtbxQtd_alt.Top = msflxgrd.Top + (msflxgrd.RowHeight(msflxgrd.row) * msflxgrd.row)
        txtbxQtd_alt.Height = msflxgrd.RowHeight(msflxgrd.row)
        txtbxQtd_alt.width = msflxgrd.ColWidth(msflxgrd.col)
        txtbxQtd_alt.Text = msflxgrd.Text
        txtbxQtd_alt.Visible = True
        txtbxQtd_alt.SetFocus
    ElseIf KeyCode = vbKeyDelete Then
        msflxgrd.col = 0
        If msflxgrd.Text = "" Then
            If MsgBox("Deseja realmente excluir o item?", vbYesNo + vbCritical) = vbYes Then
                msflxgrd.col = 3
                cQtd = CCur(msflxgrd.Text)
                msflxgrd.col = 4
                cValor = CCur(msflxgrd.Text)
                cTotal = cValor * cQtd
                
                
                
                
                lblAcresc.caption = FormatVal(CCur(lblAcresc.caption) - CCur(cTotal))
                
                lblFinalPrev.caption = FormatVal(CCur(lblAcresc.caption) - CCur(lblAbater.caption))
                lblFinalPrev.tag = lblFinalPrev.caption
                
                msflxgrd.RemoveItem msflxgrd.row
            End If
        End If
    End If
End Sub

Private Sub txtbxNum_Ven_GotFocus()
    P_SelecTexto txtbxNum_Ven
End Sub

Private Sub txtbxNum_Ven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub
Sub BuscaVenda()
    'lblValorGerado.caption = "0,00"
    'lblItens.caption = "0"
    If F_Vazio(txtbxNum_Ven.Text) Then
        Exit Sub
    End If
    txtbxNum_Ven.Text = Format(F_SoNumero(txtbxNum_Ven.Text), "000000#")
    Me.tag = txtbxNum_Ven.Text
    busca
    If txtbxNum_Ven.Text = "" Then Exit Sub
    P_buscaitens "seq_ite"
        
End Sub
Sub P_buscaitens(sOrd As String)
    strSQL = Vendas.F_BuscaItens(txtbxNum_Ven.Text)
    adc.RecordSource = strSQL
    adc.Refresh
    With adc.Recordset
        msflxgrd.rows = 1
        If .RecordCount = 0 Then Exit Sub
        
        Do While Not .EOF
            strSQL = !seq_ite & Chr(9) & (!cod_pro) & Chr(9)
            strSQL = strSQL & !desc_ite & Chr(9) & (!qtd_ite) & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite) & Chr(9) & FormatVal(!total) & Chr(9)
            strSQL = strSQL & FormatVal(!final_ite) & Chr(9) & !qtd_ite & Chr(9) & "N"
            strSQL = strSQL & Chr(9) & Format(!ref_pro)
            msflxgrd.AddItem strSQL
            .MoveNext
        Loop
        .MoveFirst
    End With
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error Resume Next
    strSQL = "select tv.* "
    strSQL = strSQL & " from tab_vendas "
    strSQL = strSQL & " as tv where "
    strSQL = strSQL & " tv.num_ven='" & txtbxNum_Ven.Text & "'"
    strSQL = strSQL & " and situacao_ven='4'"
    
    rPesq.Open strSQL, conn
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Esta venda não existe!", vbExclamation
            txtbxNum_Ven.Text = ""
            txtbxNum_Ven.SetFocus
            Exit Sub
        End If
        If !situacao_ven < "3" Then
            MsgBox "Não é possível haver devoluções nesta venda!" _
                & Chr(13) & "Ela ainda não se encontra finalizada!", vbExclamation
            txtbxNum_Ven.Text = ""
            txtbxNum_Ven.SetFocus
            Exit Sub
        End If
        lblCliente.caption = "Cliente: " & F_BuscaUmCampo("Nome_cli", TAB_CLIENTES, _
            "cod_cli='" & !cod_cli & "'")
        lblCliente.tag = !cod_cli
        lblDataVen.caption = "Data: " & !data_ven
        lblValor.caption = "Valor: " & FormatVal(!valor_ven)
        lblDesc.caption = "Desc: " & Round(!desconto_ven / !valor_ven, 2) * 100 & " % "
        lblDesc.tag = Round(!desconto_ven / !valor_ven, 2) * 100
        
        lblVendedor.caption = "Vendedor: " & F_BuscaUmCampo("guerra_fun", TAB_FUNCIONARIOS, _
            "cod_fun='" & !abrevenda_usu & "'")
        
        lblAbater.caption = "0,00"
        lblAcresc.caption = "0,00"
        lblFinalPrev.caption = "0,00"
        lblFinalPrev.tag = "0,00"
    End With
End Sub


Private Sub txtbxNum_Ven_LostFocus()
    BuscaVenda
End Sub

Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub txtbxQtd_alt_GotFocus()
    P_SelecTexto txtbxQtd_alt
End Sub

Private Sub txtbxQtd_alt_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cFinal As Currency
    Dim iQtd As Integer
    
    On Error GoTo trata_erro
    If KeyCode = vbKeyReturn Then
        txtbxQtd_alt_LostFocus
        If msflxgrd.Text = txtbxQtd_alt.Text Then
            txtbxQtd_alt.Visible = False
            Exit Sub
        End If
        
        If MsgBox("Confirma alteração da quantidade de " & msflxgrd.Text & _
            " para " & txtbxQtd_alt.Text, vbYesNo + vbCritical) = vbYes Then
            If txtbxQtd_alt.Text = 0 Then
                msflxgrd.col = 6
                cFinal = CCur(msflxgrd.Text)
            Else
                msflxgrd.col = 3
                iQtd = msflxgrd.Text
                If iQtd = 0 Then
                    iQtd = 1
                End If
                msflxgrd.col = 6
                cFinal = Round(CCur(msflxgrd.Text / iQtd) * (iQtd - CCur(txtbxQtd_alt.Text)), 2)
            End If
            msflxgrd.col = 3
            If cFinal = 0 Then
'                cFinal = ""
            End If
            If CCur(txtbxQtd_alt.Text) > CCur(msflxgrd.Text) Then
                
'                lblAcresc.caption = FormatVal(CCur(lblAcresc.caption) + _
'                    ((CCur(txtbxQtd_alt.Text) - CCur(msflxgrd.Text)) * cFinal))
                 lblAcresc.caption = FormatVal(CCur(lblAcresc.caption) + _
                    cFinal)
                
            Else
'                lblAcresc.caption = FormatVal(CCur(lblAcresc.caption) + _
'                    ((CCur(msflxgrd.Text) - CCur(txtbxQtd_alt.Text)) * cFinal))
                lblAbater.caption = FormatVal(CCur(lblAbater.caption) + _
                     cFinal)
            End If
            lblFinalPrev.caption = FormatVal(CCur(lblAcresc.caption) - CCur(lblAbater.caption))
            lblFinalPrev.tag = lblFinalPrev.caption
            msflxgrd.Text = txtbxQtd_alt.Text
            msflxgrd.CellBackColor = vbInfoBackground
            msflxgrd.CellForeColor = vbBlue
            txtbxQtd_alt.Visible = False
            
        End If
    End If ' parei aki
    If KeyCode = vbKeyEscape Then
        txtbxQtd_alt.Visible = False
        msflxgrd.SetFocus
    End If
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub txtbxQtd_alt_LostFocus()
    P_ValorLostFocus txtbxQtd_alt, "###,##0"
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd, "###,##0"
    txtbxTotal.Text = FormatVal(CCur(txtbxQtd.Text) * CCur(txtbxUnit.Text))
End Sub

Private Sub txtbxVendedor_Change()
    txtbxVendedor.tag = ""
End Sub

Private Sub txtbxVendedor_GotFocus()
    P_SelecTexto txtbxVendedor
End Sub

Private Sub txtbxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaFun_Click
    End If
End Sub
