VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOrcamento 
   Caption         =   "Orçamento de Venda"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12285
   Icon            =   "frmOrcamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   12285
   Begin VB.TextBox txtbxGrid 
      Height          =   285
      Left            =   840
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox pcProd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   12255
      TabIndex        =   10
      Top             =   6540
      Width           =   12285
      Begin VB.PictureBox pcTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H0088C9EC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   9120
         ScaleHeight     =   1215
         ScaleWidth      =   4095
         TabIndex        =   17
         Top             =   0
         Width           =   4095
         Begin VB.Label lblValor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   540
            Left            =   1080
            TabIndex        =   23
            Tag             =   "!"
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produtos:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   840
         End
         Begin VB.Label lblProd 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   1080
            TabIndex        =   20
            Tag             =   "!"
            Top             =   0
            Width           =   165
         End
         Begin VB.Label lblCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            TabIndex        =   19
            Top             =   3000
            Width           =   1965
         End
         Begin VB.Label lblCodBarr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2040
            TabIndex        =   18
            Top             =   3000
            Width           =   2325
         End
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   11
         Text            =   "0,00"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   300
      End
      Begin VB.Label lblUltItem 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   15
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Preço Unitário"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5160
         TabIndex        =   14
         Top             =   120
         Width           =   1005
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   12255
      TabIndex        =   0
      Top             =   7515
      Width           =   12285
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   855
         Left            =   8400
         Picture         =   "frmOrcamento.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   855
         Left            =   10320
         Picture         =   "frmOrcamento.frx":0924
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   855
         Left            =   6360
         Picture         =   "frmOrcamento.frx":0A8F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxCondicoes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   5415
      End
      Begin VB.TextBox txtbxCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Tag             =   "0000"
         Text            =   "Venda a Consumidor"
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtbxValidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4560
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Condições:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Válido Até:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         Top             =   0
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   6015
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10610
      _Version        =   393216
      BackColor       =   -2147483636
      BackColorBkg    =   -2147483624
      GridColor       =   16777215
      GridColorFixed  =   12632256
      AllowUserResizing=   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "select tv.*,tp.*,tg.nome_cat,tg.nome_cat,"
    strSQL = strSQL & "te.nome_mar,te.cod_mar,ti.*,(ti.unit_ite*ti.qtd_ite) as total "
    strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp,tab_orc as tv, tab_categorias as tg,"
    strSQL = strSQL & TAB_MARCAS & " as te, tab_Itorc as ti   "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro and tp.cod_cat=tg.cod_cat "
    
    strSQL = strSQL & " and tp.cod_mar=te.cod_mar "
    strSQL = strSQL & " and ti.num_orc=tv.num_orc and tv.num_orc='" & Me.tag & "' "
    strSQL = strSQL & " order by seq_ite "
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    
            
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub insere()
    Dim iPos As Integer
    Dim i As Integer
    Dim p As Produto
    Dim strSQL As String
            
            
            
            
            p = F_BuscaProdSaida(txtbxProduto.Text)
            If p.cod_pro <> "" Then
                
                txtbxProduto.Text = p.nome_pro
                
                
                
                txtbxProduto.tag = p.cod_pro
                If p.pv_pro = 0 Then
                    MsgBox "Este produto encontra-se com o preço (p2) de venda zerado!", vbCritical
                    Exit Sub
                End If
                
                If strContrpre_usu = "S" Then
                    txtbxUnit.Enabled = True
                    txtbxUnit.Text = p.pv_pro
                    txtbxUnit.SetFocus
                    'cmndbtnInserir.Visible = True
                    Exit Sub
                End If
                If p.cod_pro <> "" Then
                    strSQL = Msflxgrd.rows & Chr(9) & p.nome_pro & Chr(9) & txtbxQtd.Text & Chr(9)
                    strSQL = strSQL & FormatVal(p.pv_pro) & Chr(9)
                    strSQL = strSQL & FormatVal(p.pv_pro * CCur(txtbxQtd.Text)) & Chr(9) & p.cod_pro
                    Msflxgrd.AddItem strSQL
                    
                    lblValor.caption = FormatVal(CCur(lblValor.caption) + p.pv_pro * CCur(txtbxQtd.Text))
                End If
                
                                          
                txtbxQtd.Text = "1"
                    
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
                    
                
        End If
        
        
    ZebraGrid Msflxgrd, 2, &HD3E7EB
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = True
    lblProd.caption = Msflxgrd.rows
End Sub

Private Sub cmndbtnImprimir_Click()
    'ImprOrc Me.tag
End Sub

Private Sub cmndbtnSair_Click()
    If cmndbtnSalvar.Enabled = True Then
        If MsgBox("Você fez alterações no orçamento, deseja Salvar?", vbYesNo + vbQuestion) = vbYes Then
            salva
        End If
    End If
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Msflxgrd.rows = 1 Then
        MsgBox "Não há o que ser salvo!", vbCritical
        Exit Sub
    End If

        salva

    
        
    
    
    
End Sub

Sub salva()
    Dim strSQL As String
    Dim iCount As Integer
    If F_Vazio(txtbxValidade.Text) Then
        MsgBox "Informe a validade corretamente!", vbExclamation
        txtbxValidade.SetFocus
        Exit Sub
        
    End If
    If Me.tag = "" Then
        strSQL = F_GeraNovoOrc(txtbxCliente.tag, txtbxCliente.Text, txtbxValidade.Text, lblValor.caption, "0", txtbxCondicoes.Text)
        Me.tag = strSQL
    Else
        strSQL = "update tab_orc set valor_orc=" & F_GravaDec(lblValor) & ","
        strSQL = strSQL & "cod_cli='" & txtbxCliente.tag & "',"
        strSQL = strSQL & "nomecli_orc='" & txtbxCliente.Text & "' "
        strSQL = strSQL & " where num_orc='" & Me.tag & "' "
        conn.Execute strSQL
        strSQL = Me.tag
    End If
    If strSQL <> "" Then
        P_InsereProd
    End If
End Sub
Sub P_InsereProd()
    Dim strSQL As String
    Dim iCount As Integer
    strSQL = "delete from tab_itorc where num_orc='" & Me.tag & "' "
    conn.Execute strSQL
    With Msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            strSQL = "insert into tab_itorc(num_orc,cod_pro,unit_ite,qtd_ite,seq_ite)values("
        
            strSQL = strSQL & "'" & Me.tag & "', "
            .col = 5
            strSQL = strSQL & "'" & .Text & "',"
            .col = 3
            strSQL = strSQL & F_GravaDec(.Text) & ","
            .col = 2
            strSQL = strSQL & F_GravaDec(.Text) & ","
            strSQL = strSQL & iCount & ")"
            conn.Execute strSQL
        Next
    End With

    'RESERVANDO O ESTOQUE
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = False
End Sub

Private Sub lblCat_Click()

End Sub

Private Sub Form_Resize()
    posicionaADireita cmndbtnSair, pcBot
    posicionaADireita pcTotal, pcProd
    PosicionaGrade Msflxgrd, Me, pcProd
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub txtbxCliente_GotFocus()
    ganhaFoco txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    
           
    If KeyCode = vbKeyReturn Then
        frmSelectCli.txtbxNome.tag = "C"
        frmSelectCli.txtbxNome.Text = txtbxCliente.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxCliente.Text = frmSelectCli.adoR!nome_Cli
    txtbxCliente.tag = frmSelectCli.adoR!cod_cli
    
    
    
'

If txtbxCliente.tag = "" Then
        txtbxCliente.tag = "0000"
        txtbxCliente.Text = "Venda a Consumidor"
    End If
    
    Unload frmSelectCli
    
    End If
End Sub
Sub buscaCli()
    
End Sub
Private Sub txtbxCliente_LostFocus()
    AjustaTexto txtbxCliente
End Sub

Private Sub txtbxCondicoes_GotFocus()
    ganhaFoco txtbxCondicoes
End Sub

Private Sub txtbxCondicoes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtbxCondicoes_LostFocus()
    AjustaTexto txtbxCondicoes
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_erro
    FechaObjeto rPesq
    strSQL = "select * from tab_orc where num_orc='" & Me.tag & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        txtbxCliente.Text = !nomecli_orc
        txtbxCliente.tag = !cod_cli
        txtbxValidade.Text = !validade_orc
        txtbxCondicoes.Text = !obs_orc
        .Close
    End With
    strSQL = "select tp.*,ti.* from tab_itorc as ti," & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and nuM_orc='" & Me.tag & "' "
    strSQL = strSQL & " order by seq_ite "
    
    rPesq.Open strSQL, conn_client
    With rPesq
        lblValor.caption = "0"
        Do While Not .EOF
            strSQL = Msflxgrd.rows & Chr(9) & !nome_pro & Chr(9) & !qtd_ite & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite) & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite * !qtd_ite) & Chr(9) & !cod_pro
            Msflxgrd.AddItem strSQL
                    
            lblValor.caption = FormatVal(CCur(lblValor.caption) + !unit_ite * !qtd_ite)
            .MoveNext
        Loop
        .Close
    End With
    ZebraGrid Msflxgrd, 1
    cmndbtnSalvar.Enabled = True
    cmndbtnImprimir.Enabled = True
    Exit Sub
    
trata_erro:
    MostraErro
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        editaItem
    End If
End Sub

Private Sub Form_Load()
    resizeForm Me
    bActive = False
    Me.Top = 10: Me.Left = 10
    With Msflxgrd
        .FixedCols = 0
        .rows = 1
        .Cols = 6
        .col = 0: .Text = "..."
        .col = 1: .Text = "Descrição"
        .col = 2: .Text = "Qtd."
        .col = 3: .Text = "Custo(R$)"
        .col = 4: .Text = "Total(R$)"
        .col = 5: .Text = "Código"
        .ColWidth(0) = 300
        .ColWidth(1) = 3500
        .ColWidth(2) = 600
        .ColWidth(3) = 700
        .ColWidth(4) = 800
        .ColWidth(5) = 2000
        
    End With
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyF4 Then
        exclui
    End If
    If KeyCode = vbKeyReturn Then
        With Msflxgrd
            If .col = 2 Or .col = 3 Then
                txtbxGrid.Text = .Text
                txtbxGrid.Left = .CellLeft + Msflxgrd.Left
                txtbxGrid.Top = .CellTop + Msflxgrd.Top
                txtbxGrid.Height = .CellHeight
                txtbxGrid.Width = .CellWidth
                txtbxGrid.Visible = True
                txtbxGrid.SetFocus
                
            End If
        End With
    End If
End Sub
Sub exclui()
    With Msflxgrd
        If .rows = 2 Then
            .col = 4
            lblValor.caption = FormatVal(CCur(lblValor.caption) - CCur(.Text))
            Msflxgrd.rows = 1
        ElseIf .rows > 2 Then
            
            .col = 4
            lblValor.caption = FormatVal(CCur(lblValor.caption) - CCur(.Text))
            Msflxgrd.RemoveItem .Row
        Else
            MsgBox "Não há o que excluir!", vbCritical
        End If
    End With
    If Msflxgrd.rows > 1 Then
        cmndbtnImprimir.Enabled = True
    End If
    cmndbtnSalvar.Enabled = True
    ZebraGrid Msflxgrd, 2, &HD3E7EB
    lblProd.caption = Msflxgrd.rows
End Sub



Private Sub txtbxGrid_GotFocus()
    ganhaFoco txtbxGrid
End Sub

Private Sub txtbxGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        PerdeFoco txtbxGrid
        If CCur(txtbxGrid.Text) <= 0 Then
            MsgBox "Valor incorreto!", vbCritical
            txtbxGrid.SetFocus
            Exit Sub
        End If
        Msflxgrd.Text = txtbxGrid.Text
        txtbxGrid.Visible = False
        SomaLinha
        SomaGrid
    End If
End Sub

Private Sub txtbxGrid_LostFocus()
    PerdeFoco txtbxGrid
    txtbxGrid.Visible = False
End Sub
Sub SomaGrid()
    Dim iCount As Integer
    Dim cTotal As Currency
    With Msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .col = 4
            cTotal = cTotal + CCur(.Text)
            
        Next
        lblValor.caption = FormatVal(cTotal)
        lblProd.caption = .rows - 1
    End With
End Sub
Sub SomaLinha()
    Dim cQtd As Currency
    Dim cUnit As Currency
    Msflxgrd.col = 2
    cQtd = CCur(Msflxgrd.Text)
    Msflxgrd.col = 3
    cUnit = CCur(Msflxgrd.Text)
    Msflxgrd.col = 4
    Msflxgrd.Text = FormatVal(cUnit * cQtd)
    
End Sub
Private Sub txtbxProduto_Change()
    txtbxProduto.tag = ""
End Sub

Private Sub txtbxProduto_GotFocus()
    ganhaFoco txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If F_Vazio(txtbxProduto.Text) Then
            MsgBox "Informe o produto!", vbExclamation
            txtbxProduto.SetFocus
            Exit Sub
        End If
        insere
    End If
End Sub

Private Sub txtbxQtd_GotFocus()
    
    ganhaFoco txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo trata_erro
    If KeyCode = vbKeyReturn Then
        If CCur(txtbxQtd.Text) = 0 Then
            MsgBox "O Valor Informado não está Correto!", vbCritical
            txtbxQtd.SetFocus
            Exit Sub
        End If
        txtbxQtd.Enabled = False
        txtbxProduto.SetFocus
    End If
Exit Sub
trata_erro:
    MsgBox Err.Description, vbcre
End Sub

Private Sub txtbxQtd_LostFocus()
    On Error Resume Next
    If CCur(txtbxQtd.Text) > 999 Then
        If MsgBox("A quantidade digitada é muito grande." & _
            "Deseja Continuar?", vbYesNo + vbCritical + vbDefaultButton2, "Atenção!!!") = vbNo Then
            txtbxQtd.Text = "1"
            Exit Sub
        End If
    End If
    txtbxQtd.Enabled = False
    txtbxProduto.Text = ""
    txtbxProduto.Enabled = True
    txtbxProduto.SetFocus
    
    
    
End Sub
Sub editaItem()
    txtbxQtd.Enabled = True
    txtbxQtd.Text = "1"
    txtbxQtd.SetFocus
End Sub

Private Sub txtbxUnit_GotFocus()
    ganhaFoco txtbxUnit
End Sub

Private Sub txtbxUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        insereManual
        
    End If
End Sub
Sub insereManual()
    Dim strSQL As String
    If txtbxProduto.tag = "" Then
        MsgBox "Informe o produto!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    If CCur(txtbxUnit.Text) <= 0 Then
        MsgBox "O valor unitário está zerado!", vbExclamation
        Exit Sub
    End If
    If CCur(txtbxQtd.Text) <= 0 Then
        MsgBox "Quantidade está zerada!", vbCritical
        Exit Sub
    End If
    
    strSQL = Msflxgrd.rows & Chr(9) & txtbxProduto.Text & Chr(9) & txtbxQtd.Text & Chr(9)
    strSQL = strSQL & FormatVal(txtbxUnit.Text) & Chr(9)
    strSQL = strSQL & FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text)) & Chr(9) & txtbxProduto.tag
    Msflxgrd.AddItem strSQL
                    
    lblValor.caption = FormatVal(CCur(lblValor.caption) + CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text))
    ZebraGrid Msflxgrd, 2, &HD3E7EB
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = True
    txtbxProduto.Text = ""
    txtbxQtd.Text = "1"
    txtbxProduto.SetFocus
    
End Sub
Private Sub txtbxUnit_LostFocus()
    PerdeFoco txtbxUnit
End Sub

Private Sub txtbxValidade_GotFocus()
    ganhaFoco txtbxValidade
End Sub

Private Sub txtbxValidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValidade_LostFocus()
    P_AjustaData txtbxValidade
End Sub

