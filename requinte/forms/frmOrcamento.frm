VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOrcamento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Orçamento de Venda"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtbxCondicoes 
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   6360
      TabIndex        =   34
      Top             =   5880
      Width           =   3975
   End
   Begin VB.TextBox txtbxValidade 
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   6360
      TabIndex        =   33
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtbxCliente 
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   6360
      TabIndex        =   32
      Tag             =   "0000"
      Text            =   "Venda a Consumidor"
      Top             =   4680
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EB&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   6360
      ScaleHeight     =   4305
      ScaleWidth      =   4065
      TabIndex        =   11
      Top             =   120
      Width           =   4095
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         Height          =   195
         Left            =   15
         TabIndex        =   30
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. Barra:"
         Height          =   195
         Left            =   2040
         TabIndex        =   29
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód:"
         Height          =   195
         Left            =   10
         TabIndex        =   28
         Top             =   2760
         Width           =   330
      End
      Begin VB.Label lblValor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   42
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   945
         Left            =   1680
         TabIndex        =   27
         Tag             =   "!"
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   630
         Left            =   0
         TabIndex        =   26
         Top             =   600
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10
         TabIndex        =   25
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblProd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Tag             =   "!"
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência:"
         Height          =   195
         Left            =   2040
         TabIndex        =   23
         Top             =   3480
         Width           =   825
      End
      Begin VB.Label lblRef 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2040
         TabIndex        =   22
         Top             =   3720
         Width           =   1965
      End
      Begin VB.Label lblPreco 
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
         TabIndex        =   21
         Top             =   3720
         Width           =   1725
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         Height          =   195
         Left            =   2040
         TabIndex        =   17
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblMarca 
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
         TabIndex        =   16
         Top             =   2280
         Width           =   1845
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         Height          =   195
         Left            =   10
         TabIndex        =   15
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblProduto 
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
         TabIndex        =   14
         Top             =   1680
         Width           =   4845
      End
      Begin VB.Label lblCapPro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   10
         TabIndex        =   13
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label lblCat 
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
         Top             =   2280
         Width           =   2325
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
         TabIndex        =   20
         Top             =   3000
         Width           =   2325
      End
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6360
      Picture         =   "frmOrcamento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   9120
      Picture         =   "frmOrcamento.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   7800
      Picture         =   "frmOrcamento.frx":0F8C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtbxGrid 
         Height          =   285
         Left            =   960
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   6015
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   10610
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorBkg    =   -2147483624
         GridColor       =   16777215
         GridColorFixed  =   12632256
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   0
         Top             =   6120
         Width           =   6015
         Begin VB.TextBox txtbxUnit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5160
            TabIndex        =   6
            Text            =   "0,00"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtbxQtd 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Text            =   "1"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtbxProduto 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   720
            TabIndex        =   4
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Unit:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5160
            TabIndex        =   5
            Top             =   120
            Width           =   330
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
            TabIndex        =   3
            Top             =   120
            Width           =   555
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
            TabIndex        =   1
            Top             =   120
            Width           =   300
         End
      End
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Condições:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6360
      TabIndex        =   37
      Top             =   5640
      Width           =   795
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Válido Até:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6360
      TabIndex        =   36
      Top             =   5040
      Width           =   765
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Cliente:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6360
      TabIndex        =   35
      Top             =   4440
      Width           =   525
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
    strSQL = strSQL & " from tab_produtos as tp,tab_orc as tv, tab_categorias as tg,"
    strSQL = strSQL & " tab_marcas as te, tab_Itorc as ti   "
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
    Dim p As ProdutoSaida
    Dim strSQL As String
            
            
            
            
            p = F_BuscaProdSaida(txtbxProduto.Text, "B")
            If p.cod_pro <> "" Then
                
                txtbxProduto.Text = p.nome_pro
                txtbxProduto.tag = p.cod_pro
                lblCat.caption = p.sub.nome_sub
                lblCod.caption = p.cod_pro
                'lblCodBarr.caption = p.codbarra_pro
                'lblMarca.caption = p.nome_mar
                lblProduto.caption = p.nome_pro
                lblPreco.caption = p.pv_pro
                lblRef.caption = p.ref_pro
                
                
                lblProduto.caption = p.nome_pro
                
                If p.pv_pro = 0 Then
                    MsgBox "Este produto encontra-se com o preço de venda zerado!", vbCritical
                    Exit Sub
                End If
                lblProduto.caption = p.nome_pro
                If strContrpre_usu = "S" Then
                    txtbxUnit.Enabled = True
                    txtbxUnit.Text = p.pv_pro
                    txtbxUnit.SetFocus
                    'cmndbtnInserir.Visible = True
                    Exit Sub
                End If
                If p.cod_pro <> "" Then
                    strSQL = msflxgrd.rows & Chr(9) & p.nome_pro & Chr(9) & txtbxQtd.Text & Chr(9)
                    strSQL = strSQL & FormatVal(p.pv_pro) & Chr(9)
                    strSQL = strSQL & FormatVal(p.pv_pro * CCur(txtbxQtd.Text)) & Chr(9) & p.cod_pro
                    msflxgrd.AddItem strSQL
                    
                    lblValor.caption = FormatVal(CCur(lblValor.caption) + p.pv_pro * CCur(txtbxQtd.Text))
                End If
                
                                          
                txtbxQtd.Text = "1"
                    
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
                    
                
        End If
        
        
    ZebraGrid msflxgrd, 2, &HD3E7EB
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = True
    lblProd.caption = msflxgrd.rows
End Sub

Private Sub cmndbtnImprimir_Click()
    ImprOrc Me.tag
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
    If msflxgrd.rows = 1 Then
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
        Banco.Execute strSQL
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
    Banco.Execute strSQL
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            strSQL = "insert into tab_itorc(num_orc,cod_pro,unit_ite,qtd_ite,seq_ite)values("
        
            strSQL = strSQL & "'" & Me.tag & "', "
            .Col = 5
            strSQL = strSQL & "'" & .Text & "',"
            .Col = 3
            strSQL = strSQL & F_GravaDec(.Text) & ","
            .Col = 2
            strSQL = strSQL & F_GravaDec(.Text) & ","
            strSQL = strSQL & iCount & ")"
            Banco.Execute strSQL
        Next
    End With

    'RESERVANDO O ESTOQUE
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = False
End Sub

Private Sub txtbxCliente_GotFocus()
    P_SelecTexto txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    
           
    If KeyCode = vbKeyReturn Then
        frmSelectCli.txtbxNome.tag = "C"
        frmSelectCli.txtbxNome.Text = txtbxCliente.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxCliente.Text = frmSelectCli.dtBusca.Recordset!Nome
    txtbxCliente.tag = frmSelectCli.dtBusca.Recordset!cod
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
    P_AjustaTexto txtbxCliente
End Sub

Private Sub txtbxCondicoes_GotFocus()
    P_SelecTexto txtbxCondicoes
End Sub

Private Sub txtbxCondicoes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxCondicoes_LostFocus()
    P_AjustaTexto txtbxCondicoes
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        Carrega
    End If
    bActive = True
End Sub
Sub Carrega()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_erro
    strSQL = "select * from tab_orc where num_orc='" & Me.tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        txtbxCliente.Text = !nomecli_orc
        txtbxCliente.tag = !cod_cli
        txtbxValidade.Text = !validade_orc
        txtbxCondicoes.Text = !obs_orc
    End With
    strSQL = "select tp.*,ti.* from tab_itorc as ti,tab_produtos as tp "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and nuM_orc='" & Me.tag & "' "
    strSQL = strSQL & " order by seq_ite "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        lblValor.caption = "0"
        Do While Not .EOF
            strSQL = msflxgrd.rows & Chr(9) & !nome_pro & Chr(9) & !qtd_ite & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite) & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite * !qtd_ite) & Chr(9) & !cod_pro
            msflxgrd.AddItem strSQL
                    
            lblValor.caption = FormatVal(CCur(lblValor.caption) + !unit_ite * !qtd_ite)
            .MoveNext
        Loop
        .Close
    End With
    ZebraGrid msflxgrd, 1
    cmndbtnSalvar.Enabled = True
    cmndbtnImprimir.Enabled = True
    Exit Sub
    
trata_erro:
    P_MostraErro
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        editaItem
    End If
End Sub

Private Sub Form_Load()
    bActive = False
    Me.Top = 10: Me.Left = 10
    With msflxgrd
        .FixedCols = 0
        .rows = 1
        .Cols = 6
        .Col = 0: .Text = "..."
        .Col = 1: .Text = "Descrição"
        .Col = 2: .Text = "Qtd."
        .Col = 3: .Text = "Custo(R$)"
        .Col = 4: .Text = "Total(R$)"
        .Col = 5: .Text = "Código"
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
        With msflxgrd
            If .Col = 2 Or .Col = 3 Then
                txtbxGrid.Text = .Text
                txtbxGrid.Left = .CellLeft + msflxgrd.Left
                txtbxGrid.Top = .CellTop + msflxgrd.Top
                txtbxGrid.Height = .CellHeight
                txtbxGrid.Width = .CellWidth
                txtbxGrid.Visible = True
                txtbxGrid.SetFocus
                
            End If
        End With
    End If
End Sub
Sub exclui()
    With msflxgrd
        If .rows = 2 Then
            .Col = 4
            lblValor.caption = FormatVal(CCur(lblValor.caption) - CCur(.Text))
            msflxgrd.rows = 1
        ElseIf .rows > 2 Then
            
            .Col = 4
            lblValor.caption = FormatVal(CCur(lblValor.caption) - CCur(.Text))
            msflxgrd.RemoveItem .Row
        Else
            MsgBox "Não há o que excluir!", vbCritical
        End If
    End With
    If msflxgrd.rows > 1 Then
        cmndbtnImprimir.Enabled = True
    End If
    cmndbtnSalvar.Enabled = True
    ZebraGrid msflxgrd, 2, &HD3E7EB
    lblProd.caption = msflxgrd.rows
End Sub



Private Sub txtbxGrid_GotFocus()
    P_SelecTexto txtbxGrid
End Sub

Private Sub txtbxGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        P_ValorLostFocus txtbxGrid
        If CCur(txtbxGrid.Text) <= 0 Then
            MsgBox "Valor incorreto!", vbCritical
            txtbxGrid.SetFocus
            Exit Sub
        End If
        msflxgrd.Text = txtbxGrid.Text
        txtbxGrid.Visible = False
        SomaLinha
        SomaGrid
    End If
End Sub

Private Sub txtbxGrid_LostFocus()
    P_ValorLostFocus txtbxGrid
    txtbxGrid.Visible = False
End Sub
Sub SomaGrid()
    Dim iCount As Integer
    Dim cTotal As Currency
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 4
            cTotal = cTotal + CCur(.Text)
            
        Next
        lblValor.caption = FormatVal(cTotal)
        lblProd.caption = .rows - 1
    End With
End Sub
Sub SomaLinha()
    Dim cQtd As Currency
    Dim cUnit As Currency
    msflxgrd.Col = 2
    cQtd = CCur(msflxgrd.Text)
    msflxgrd.Col = 3
    cUnit = CCur(msflxgrd.Text)
    msflxgrd.Col = 4
    msflxgrd.Text = FormatVal(cUnit * cQtd)
    
End Sub
Private Sub txtbxProduto_Change()
    txtbxProduto.tag = ""
End Sub

Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
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
    
    P_SelecTexto txtbxQtd
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
    P_SelecTexto txtbxUnit
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
    
    strSQL = msflxgrd.rows & Chr(9) & txtbxProduto.Text & Chr(9) & txtbxQtd.Text & Chr(9)
    strSQL = strSQL & FormatVal(txtbxUnit.Text) & Chr(9)
    strSQL = strSQL & FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text)) & Chr(9) & txtbxProduto.tag
    msflxgrd.AddItem strSQL
                    
    lblValor.caption = FormatVal(CCur(lblValor.caption) + CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text))
    ZebraGrid msflxgrd, 2, &HD3E7EB
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = True
    txtbxProduto.Text = ""
    txtbxQtd.Text = "1"
    txtbxProduto.SetFocus
    
End Sub
Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit
End Sub

Private Sub txtbxValidade_GotFocus()
    P_SelecTexto txtbxValidade
End Sub

Private Sub txtbxValidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValidade_LostFocus()
    P_AjustaData txtbxValidade
End Sub

