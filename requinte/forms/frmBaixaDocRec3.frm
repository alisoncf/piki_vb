VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBaixaDocReq 
   Caption         =   "Baixa de Parcela"
   ClientHeight    =   8715
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12945
   Icon            =   "frmBaixaDocRec3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   7440
      ScaleHeight     =   555
      ScaleWidth      =   3795
      TabIndex        =   43
      Top             =   8160
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ComboBox cmbxConta 
         BackColor       =   &H00008000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmBaixaDocRec3.frx":038A
         Left            =   120
         List            =   "frmBaixaDocRec3.frx":0394
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Tag             =   "!"
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton optnFin 
         Caption         =   "Financeiro:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   960
         TabIndex        =   45
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optnCaixa 
         Caption         =   "Caixa"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CheckBox chkImpr 
      Caption         =   "Imprimir Recibo"
      Height          =   195
      Left            =   2880
      TabIndex        =   42
      Top             =   8280
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtTotalDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4440
      TabIndex        =   41
      Text            =   "0,00"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox txtbxSeq 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4200
      TabIndex        =   25
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Terceiro:"
      Height          =   495
      Left            =   5400
      TabIndex        =   34
      Top             =   0
      Width           =   3975
      Begin VB.OptionButton optnFun 
         Caption         =   "Funcionário"
         Height          =   195
         Left            =   2400
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optnRep 
         Caption         =   "Representante"
         Height          =   195
         Left            =   960
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optnCli 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox txtbxMsf 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7200
      TabIndex        =   33
      Text            =   "0,00"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnDistr 
      Height          =   285
      Left            =   11040
      Picture         =   "frmBaixaDocRec3.frx":039C
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtbxDistr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      TabIndex        =   31
      Text            =   "0,00"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtbxJuros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0005C9EB&
      Height          =   285
      Left            =   3360
      TabIndex        =   27
      Text            =   "0,00"
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtbxDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0005C9EB&
      Height          =   285
      Left            =   2280
      TabIndex        =   26
      Text            =   "0,00"
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtbxCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   240
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid msflxParcelas 
      Height          =   2775
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   4895
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtbxValor_Par 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6960
      TabIndex        =   15
      Text            =   "0,00"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtbxReman 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Text            =   "0,00"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox txtbxTotalBaixa 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   10
      Text            =   "0,00"
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Cancela&r"
      Height          =   735
      Left            =   11400
      Picture         =   "frmBaixaDocRec3.frx":0726
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar Baixa"
      Height          =   735
      Left            =   9960
      Picture         =   "frmBaixaDocRec3.frx":0891
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escolha a Forma de Pagamento:"
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   12615
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   525
         Left            =   11160
         Picture         =   "frmBaixaDocRec3.frx":0FFB
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnLancar 
         Caption         =   "&Lançar"
         Height          =   525
         Left            =   9720
         Picture         =   "frmBaixaDocRec3.frx":1585
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtbxValorBaixa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox cmbxEspecie 
         Height          =   315
         ItemData        =   "frmBaixaDocRec3.frx":1B0F
         Left            =   120
         List            =   "frmBaixaDocRec3.frx":1B22
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2760
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   3836
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         FillStyle       =   1
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label lblTur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         Height          =   195
         Left            =   4560
         TabIndex        =   13
         Top             =   2160
         Width           =   180
      End
      Begin VB.Label lblCaixa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         Height          =   195
         Left            =   4200
         TabIndex        =   12
         Top             =   2160
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor a Baixar:"
         Height          =   195
         Left            =   2280
         TabIndex        =   3
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&spécie:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   615
      End
   End
   Begin VB.CommandButton cmndbtnBuscaCli 
      Height          =   285
      Left            =   3840
      Picture         =   "frmBaixaDocRec3.frx":1B84
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Desc.(R$):"
      Height          =   195
      Left            =   4440
      TabIndex        =   40
      Top             =   7440
      Width           =   1170
   End
   Begin VB.Label lblStatusCli 
      AutoSize        =   -1  'True
      Caption         =   "Liberado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   39
      Top             =   8227
      Width           =   1065
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Situação Cliente:"
      Height          =   195
      Left            =   120
      TabIndex        =   38
      Top             =   8280
      Width           =   1200
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seq.Doc."
      Height          =   195
      Left            =   4200
      TabIndex        =   24
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Baixar:"
      Height          =   195
      Left            =   9720
      TabIndex        =   30
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Juros(R$):"
      Height          =   195
      Left            =   3360
      TabIndex        =   29
      Top             =   7440
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto(R$):"
      Height          =   195
      Left            =   2280
      TabIndex        =   28
      Top             =   7440
      Width           =   1035
   End
   Begin VB.Label lblCli 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   525
   End
   Begin VB.Label lblCPF 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "000000"
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
      Left            =   3480
      TabIndex        =   20
      Tag             =   "!"
      Top             =   0
      Width           =   645
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Escolha as Parcelas que Deseja Baixar:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   12600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total (sem juros)"
      Height          =   195
      Left            =   6960
      TabIndex        =   17
      Top             =   7440
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Baixar:"
      Height          =   195
      Left            =   5640
      TabIndex        =   16
      Top             =   7440
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total da Baixa:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   1065
   End
   Begin VB.Menu mnPar 
      Caption         =   "&Parcela"
      Begin VB.Menu mnSelec 
         Caption         =   "&Marcar/ Desmarcar "
      End
      Begin VB.Menu mnEdit 
         Caption         =   "&Editar"
      End
      Begin VB.Menu mnMudarValor 
         Caption         =   "&Mudar Valor a Baixar"
      End
   End
   Begin VB.Menu mnOpcPag 
      Caption         =   "&Opção de Pagamento"
      Begin VB.Menu mnLanc 
         Caption         =   "&Lançar"
      End
      Begin VB.Menu mnExcl 
         Caption         =   "&Excluir"
      End
   End
   Begin VB.Menu mnAção 
      Caption         =   "&Ação"
      Begin VB.Menu mnConfBaixa 
         Caption         =   "&Confirmar Baixa"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "frmBaixaDocReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim che As cheque

Public bUsaJuro As Boolean


Private Sub chkVale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxConta_Click()
    If Left(cmbxEspecie.Text, 1) = "1" Then
        cmbxConta.tag = F_BuscaUmCampo("cod_con", "tab_contas", "nome_con='" & cmbxConta.Text & "'")
    
    Else
    
        cmbxConta.tag = F_BuscaUmCampo("cod_fin", "tab_financeiras", "nome_fin='" & cmbxConta.Text & "'")
    End If
End Sub

Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxEspecie_Click()
    If cmbxEspecie.ListIndex = 1 Then
        txtbxValorBaixa.Text = "0,00"
        cmbxConta.Enabled = False
        txtbxValorBaixa.Enabled = False
        
    ElseIf cmbxEspecie.ListIndex = 2 Then
        P_PreencheFinanceiras cmbxConta, "C"
        txtbxValorBaixa.Enabled = True
        cmbxConta.Enabled = True
    ElseIf cmbxEspecie.ListIndex = 3 Then
        P_PreencheFinanceiras cmbxConta, "D"
        txtbxValorBaixa.Enabled = True
        cmbxConta.Enabled = True
    Else
        P_PreencheContas cmbxConta, 3
        txtbxValorBaixa.Enabled = True
        cmbxConta.Enabled = False
        
'        If optnFun.Value Then
'            chkVale.Visible = True
'        Else
'            chkVale.Visible = False
'        End If
    End If
End Sub

Private Sub cmbxEspecie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Sub buscaCLi()
    Dim c As cliente
    If txtbxCliente.tag <> "" Then
        c = F_BuscaCliente(txtbxCliente.tag)
        lblStatusCli.tag = c.status_cli
        If c.status_cli = "2" Then
            MsgBox "Lembre-se. O cliente está negativado!", vbExclamation
            lblStatusCli.caption = "Negativado"
            lblStatusCli.ForeColor = vbRed
        Else
            lblStatusCli.caption = "Liberado"
            lblStatusCli.ForeColor = vbGreen
        End If
        
    Else
        c = F_BuscaCliente(txtbxCliente.Text)
    End If
    
    If c.cod_cli <> "" Then
        
        txtbxCliente.Text = c.nome_Cli
        txtbxCliente.tag = c.cod_cli
        lblCPF.caption = c.cnpj_Cli
        lblCPF.tag = "C"
        busca
    End If

End Sub
Sub buscaFun()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectFuncionario(txtbxCliente.Text)
    If f.cod_fun <> "" Then
        txtbxCliente.Text = f.nome_fun
        txtbxCliente.tag = f.cod_fun
        lblCPF.caption = f.cpf_fun
        lblCPF.tag = "U"
        busca
    End If
End Sub

Sub buscaRep()
    Dim c As representante.representante
    c = F_BuscaRep(txtbxCliente.Text)
    
    
    If c.cod_Rep <> "" Then
        
        txtbxCliente.Text = c.nome_rep
        txtbxCliente.tag = c.cod_Rep
        lblCPF.caption = c.cpf_rep
        lblCPF.tag = "R"
        busca
    End If
End Sub
Public Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    AjustaGrid
    If lblCPF.tag = "!" Then
        lblCPF.tag = "C"
    End If
    strSQL = F_SelectDocsReceber(lblCPF.tag, "", "", "", "", "", "", "", "", txtbxSeq.Text, "", "", txtbxCliente.tag, 1)
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        Do While Not .EOF
            strSQL = .AbsolutePosition + 1 & Chr(9)
            strSQL = strSQL & !seq_doc & Chr(9) & !seq_itd & Chr(9)
            strSQL = strSQL & !venc_itd & Chr(9) & FormatVal(!reman_itd) & Chr(9) & "0,00" & Chr(9)
            strSQL = strSQL & FormatVal(!valor_itd) & Chr(9) & !num_doc & Chr(9)
            strSQL = strSQL & FormatVal(!valor_doc) & Chr(9) & !data_doc & Chr(9)
            strSQL = strSQL & !parcela_itd & "/" & !parcelas_doc & Chr(9) & !dias & Chr(9)
            strSQL = strSQL & FormatVal(!valor_juro) & Chr(9) & FormatVal(!sugerido) & Chr(9)
            strSQL = strSQL & !nome_cob & Chr(9) & !cod_cob & Chr(9) & !historico_doc
            msflxParcelas.AddItem strSQL
            
            
            
            msflxParcelas.row = msflxParcelas.rows - 1
            msflxParcelas.Col = 14
            msflxParcelas.CellBackColor = vbButtonShadow
            msflxParcelas.CellForeColor = vbRed
            msflxParcelas.CellFontBold = True
            
            
            
            .MoveNext
        Loop
        .Close
    End With
    somaABaixa
    ZebraGrid msflxParcelas, 1
    
    
End Sub
Sub AjustaGrid()
    With msflxParcelas
        .Cols = 17
        .rows = 1
        .Col = 0: .Text = "..."
        .Col = 1: .Text = "Seq.Doc."
        .Col = 2: .Text = "Seq.Itd."
        .Col = 3: .Text = "Vencimento."
        .Col = 4: .Text = "Restante."
        .Col = 5: .Text = "Baixar"
        .Col = 6: .Text = "Inicial."
        .Col = 7: .Text = "Nº Doc."
        .Col = 8: .Text = "Valor Doc."
        .Col = 9: .Text = "Data Doc."
        .Col = 10: .Text = "Parcela."
        .Col = 11: .Text = "Dias."
        .Col = 12: .Text = "Acréscimo."
        .Col = 13: .Text = "Vlr Sug."
        
        .Col = 14: .Text = "Cob."
        .Col = 15: .Text = "Conta"
        .Col = 16: .Text = "Histórico."
        .ColWidth(15) = 600
        .ColWidth(16) = 3500
        .ColWidth(10) = 600
        .ColWidth(11) = 600
        .ColWidth(12) = 600
        .ColWidth(0) = 500
        .ColWidth(1) = 0
        .ColWidth(2) = 0
    End With
End Sub

Private Sub cmndbtnBuscaCli_Click()
    If optnCli.Value Then
        buscaCLi
    ElseIf optnRep.Value Then
        buscaRep
    ElseIf optnFun.Value Then
        buscaFun
    End If
End Sub

Private Sub cmndbtnConfirma_Click()
    Dim sSEQ As String
    Dim sDocs As String
    Dim sSeqBai As String
    Dim sSeqFin As String
    Dim sSeqCai As String
    Dim iCount As Integer
    Dim c As cheque
    Dim sChe As String
    Dim sEsp As String
    Dim cValor As Currency
    Dim cJuros As Currency
    Dim sEspBai As String
    Dim sVal As String
    Dim sDoc As String
    Dim sAge As String
    Dim sCPF As String
    Dim sCCO As String
    Dim sCodPla As String
    Dim sBaixado As String
    Dim sCob As String
    Dim sSeqItd As String
    Dim sTipoTer As String
    On Error GoTo trata_Erro
    Dim cBaixaItd As New Collection
    Set cBaixaItd = Nothing
    Dim cDesc As Currency
    Dim sDesc As String
    If strBaixadoc_usu <> "S" Then
        MsgBox "Não tem permissão pra baixar documento!", vbExclamation
        Exit Sub
    End If
    
    If msflxParcelas.rows = 1 Then
        MsgBox "Não foi informado nenhuma parcela para ser baixada!", vbExclamation
        Exit Sub
    End If
    
    If msflxgrd.rows = 1 Then
        MsgBox "Não foi informado nenhum valor para ser baixado!", vbExclamation
        Exit Sub
    Else
        
    End If
    
    If CCur(txtbxJuros.Text) < 0 Then
        MsgBox "Juros inválidos!", vbExclamation
        txtbxJuros.SetFocus
        Exit Sub
    End If
    
    If CCur(txtbxTotalBaixa.Text) <> CCur(txtbxReman.Text) + CCur(txtbxJuros.Text) - CCur(txtbxDesconto.Text) Then
        MsgBox "Você só pode baixar o valor completo!", vbExclamation
        Exit Sub
    End If
    If CCur(txtbxDesconto.Text) > 0 Then
        cDesc = CCur(txtbxDesconto.Text) / CCur(txtbxDistr.Text) * 100
        If crcDescontoBaixa_usu < cDesc Then
            MsgBox "Você está tentando dar um desconto de " & cDesc & "% quando o máximo possível é " & crcDescontoBaixa_usu & "%", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    If MsgBox("Confirma Baixa dos Documento?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
        
    End If
    
    
    
    
    sSeqBai = financeiro.F_BuscaChaveBaixa
    
    
    With msflxParcelas
        For iCount = 1 To .rows - 1
            .row = iCount
            .Col = 5
            If .Text > 0 Then
                .Col = 15
                sCob = .Text
                .Col = 2
                sSEQ = .Text
                .Col = 5
                sVal = .Text
                .Col = 7
                sDoc = .Text
                sDocs = sDocs & .Text & "-"
                .Col = 2
                If Left(.Text, 1) = "1" Then
                
                Else
                
                End If
                strSQL = F_BuscaUmCampo("reman_itd", "tab_itdoc", "  seq_itd='" & sSEQ & "'")
                If strSQL = "" Then strSQL = "0"
                
                
                strSQL = "Update tab_itdoc set reman_itd=reman_itd-" & F_GravaDec(sVal)
                strSQL = strSQL & " where seq_itd='" & sSEQ & "'"
                Conn.Execute strSQL
                P_ParcelaBaixadaSPC sSEQ
            
                strSQL = F_GravaBaixaItd(sSeqBai, sSEQ, sVal, sEsp, "", sCob, Date, txtTotalDesconto.Text, strCod_Cai)
                If strSQL = "" Then
                    MsgBox "Erro na gravação do lançamento do caixa. O processo irá continuar mas o lançamento não estará disponível!", vbExclamation
                    
                End If
            End If
            
            
            
        Next
    End With
    If optnCli.Value = True Then
        sTipoTer = "C"
    ElseIf optnRep = True Then
        sTipoTer = "R"
    Else
        sTipoTer = "U"
    End If
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .Col = 2
            sEsp = Left(.Text, 1)
            .Col = 0: c.Venc_che = .Text
            .Col = 1: c.Valor_che = .Text
            .Col = 3: c.num_che = .Text
            .Col = 4: c.cod_bco = .Text
            .Col = 5: c.age_che = .Text
            .Col = 6: c.cco_che = .Text
            .Col = 7: c.nome_che = .Text
            
            .Col = 9: sDesc = .Text
            If sEsp = "1" Then
                
                    If optnFin.Value Then
                        sSEQ = F_GravaFinanceiro("R", CStr(Date), CStr(c.Valor_che), _
                        "Baixa docs " & sDocs & txtbxCliente.Text, _
                        cmbxConta.tag, "S", "000")
                        sEspBai = "RDI" 'recebimento em dinheiro financeiro
'                        cBaixaFin.Add sSeq
                    Else
                        sSEQ = F_GravaCai("R", CStr(Date), CStr(c.Valor_che), _
                        "Baixa docs " & sDocs & txtbxCliente.Text, "", _
                            "RDI", strCod_Cai, CStr(c.Valor_che), 0, strCod_Cai, strCod_Tur, "000", "S", txtTotalDesconto.Text)
                            sEspBai = "RDC" 'recebimento dinheiro caixa
                            
                        If chkVale.Value = 1 And sTipoTer = "U" Then
                            If F_GravaValesCai(sSEQ, txtbxCliente.tag) Then
                                MsgBox "Lançamento em Dinheiro lançado como Vale", vbInformation
                                P_ImprimeValeCai sSEQ
                            End If
                        End If
                            
                            
'                        cBaixaFin.Add sSeq
                    End If
                
                
                 'sSeq = F_GravaFinanceiro("R", Date, CStr(c.Valor_che), "HISTORICO", "DIN", _
                        frmBaixaPar.cmbxConta.Tag, adc.Recordset!seq_itd, "1", "")
                        
                        If F_GravaBaixaItdFin(sSeqBai, sSEQ, sEspBai, CStr(c.Valor_che), sDesc) = "" Then
                            MsgBox "Erro na gravação da baixa!", vbExclamation
                        End If
                        
                'P_Class c.Valor_che, "R", sSEQ, "F", adc.Recordset!historico_doc
            ElseIf sEsp = "2" Then
                sSEQ = F_GravaDoc("BAIXA DOCS " & sDocs & txtbxCliente.Text, _
                     Date, CStr(c.Valor_che), Date, _
                    txtbxCliente.tag, "R", "", "1", "F", sTipoTer, "000")
                    'If F_GravaBaixaItd(adc.Recordset!seq_itd, CStr(c.Valor_che), "DUP", sSeq) = "" Then
                    '        MsgBox "Erro na gravação da baixa!", vbExclamation
                    '    End If
                    
                If "R" = "R" Then
                    
                    sChe = F_GravaCh(sSEQ, c.num_che, c.cod_bco, c.age_che, c.cco_che, _
                            c.nome_che, CCur(c.Valor_che), c.Venc_che, c.cod_bco, "R", _
                            "", "", "1", "", "")
                            
                            
                    If optnCaixa.Value = True Then
                        sSeqCai = F_GravaCai("C", Date, CStr(c.Valor_che), "Baixa docs " & sDocs & txtbxCliente.Text _
                    , "", "RCH", strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, "000", "S", txtbxDesconto.Text)
                    End If
                    
                    If F_GravaBaixaItdFin(sSeqBai, sChe, "CHP", CStr(c.Valor_che)) = "" Then
                            MsgBox "Erro na gravação da baixa!", vbExclamation
                        End If
                Else
                    sChe = F_GravaChePag(sSEQ, c.num_che, c.cod_bco, _
                            CCur(c.Valor_che), c.Venc_che, adc.Recordset!codigo, adc.Recordset!parcela_itd)
                End If
            Else
                sSEQ = MaxKeySQL("tab_venfin", "num_ven", "num_ven like 'B%'", "B00000#")
                
                
                
                If F_GravaVenFin(sSEQ, c.cco_che, c.Valor_che, Date, "", "1", "B", "") Then
                End If
                
                If sEsp = "3" Then
                    sEspBai = "RCC"
                Else
                    sEspBai = "RCD"
                End If
                
                
                    sSeqCai = F_GravaCai("C", Date, CStr(c.Valor_che), "Baixa docs " & sDocs & txtbxCliente.Text _
                    & "-" & Date, sSEQ, _
                     sEspBai, strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, "000", "S", txtTotalDesconto.Text)
                
                
                If F_GravaBaixaItdFin(sSeqBai, sSEQ, "FIN", CStr(c.Valor_che)) = "" Then
                            MsgBox "Erro na gravação da baixa!", vbExclamation
                        End If
            End If
            
            cValor = cValor + c.Valor_che
        Next
        
        
            'strSQL = "update tab_itdoc set reman_itd = reman_itd-" & F_GravaDec(cValor - CCur(frmBaixaPar.txtbxJuros.Text) + CCur(frmBaixaPar.txtbxDesconto.Text))
            'strSQL = strSQL & ",juros_itd=" & F_GravaDec(frmBaixaPar.txtbxJuros.Text)
            'strSQL = strSQL & ",Desconto_itd=" & F_GravaDec(frmBaixaPar.txtbxDesconto.Text)
            'strSQL = strSQL & ", baixa_itd=" & F_GravaData(Date)
            'strSQL = strSQL & " where seq_itd='" & adc.Recordset!seq_itd & "' "
            'conn.execute strSQL
        
         'retornar carnês
            
        With msflxParcelas
            For iCount = 1 To .rows - 1
                .row = iCount
                .Col = 5
                If CCur(.Text) > 0 Then
                    .Col = 2
                    sSeqItd = .Text
                    .Col = 1
                    
                    strSQL = "update tab_itdoc set cod_cob='000' where seq_doc='" & .Text & "' "
                
                    
                    'strSQL = strSQL & " and reman_itd<valor_itd "
                
                    Conn.Execute strSQL
                
                End If
                
                
                
            Next
        End With
'        MsgBox "Documento Baixado!", vbInformation
        
        
        
        
        
'        If sCob <> "000" Or strCod_Cai <> "04" Then
'
'            P_ImprimeReciboBai sSeqBai, Printer
'            Printer.EndDoc
'            P_ImprimeReciboBai sSeqBai, Printer
'            Printer.EndDoc
            
            
        '    MsgBox "O recibo foi impresso na impressora " & Printer.deviceName, vbInformation
'        Else
        
        If chkImpr.Value = 1 Then
        
            ImprReciboBai sSeqBai
        End If
'        End If
        If TotalDebitoCliente(txtbxCliente.tag, Left(lblCli.caption, 1)) = 0 Then
            MsgBox "Este cliente já quitou todos os seus débitos! Repasse essa informação ao vendedor responsável!", vbExclamation
            
            
            
            
        End If
        
        
                
                
                
        
        
        Unload Me
    End With

Exit Sub
trata_Erro:
    P_MostraErro
    Unload Me
End Sub

Private Sub cmndbtnDistr_Click()
    Dim cRest As Currency
    Dim cVal As Currency
    If MsgBox("Confirma Distribuição do Valor Informado?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    With msflxParcelas
        For iCount = 1 To .rows - 1
            .row = iCount
            .Col = 5
            .Text = "0"
            
            
        Next
        cRest = txtbxDistr.Text
        If txtbxDistr.Text > 0 Then
            For iCount = 1 To .rows - 1
                .row = iCount
                If cRest > 0 Then
                    
                    .Col = 4
                    cVal = CCur(.Text)
                    If cRest >= cVal Then
                        .Col = 5
                        .Text = cVal
                        cRest = cRest - cVal
                    Else
                        .Col = 5
                        .Text = cRest
                        cRest = 0
                    End If
                    
                    
                    
                    
                    
                End If
            Next
        
        End If
        somaABaixa
    End With
    
    
    
    
End Sub

Private Sub cmndbtnExcl_Click()
      Dim cValor As Currency
    If msflxgrd.rows = 1 Then
            MsgBox "Não há o que ser deletado!", vbExclamation
        Else
            If MsgBox("Confirma exclusão da linha?", vbYesNo + vbCritical) = vbYes Then
                msflxgrd.Col = 1
                cValor = CCur(msflxgrd.Text)
                If msflxgrd.rows = 2 Then
                    msflxgrd.rows = 1
                Else
                    msflxgrd.RemoveItem msflxgrd.row
                End If
                txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) - cValor)
            End If
            
        End If
End Sub

Private Sub cmndbtnLancar_Click()
    
    Dim strSQL As String
    On Error Resume Next
    If CCur(txtbxValorBaixa.Text) > F_Restante Then   'cmbxEspecie.Text = "1 - Dinheiro" And + CCur(txtbxTotalBaixa.Text)
        MsgBox "O Valor a ser lançado deve ser menor ou igual ao remanescente!", vbExclamation
          
        'txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    
    If optnCaixa.Value Then
        If Not F_IsCaixa Then
            MsgBox "Este computador não é um caixa, por tanto não pode lançar no caixa por aqui!", vbExclamation
            Exit Sub
        End If
    End If
    If cmbxEspecie.Text = "" Then
        MsgBox "Informe o tipo de pagamento!", vbExclamation
        cmbxEspecie.SetFocus
        Exit Sub
    End If
    If CCur(txtbxValorBaixa.Text) = "0" And Left(cmbxEspecie.Text, 1) = "1" Then
        MsgBox "Informe o valor que deseja lançar para a baixa e clique em lançar!", vbExclamation
        txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    If optnFin.Value And cmbxConta.tag = "" Then
        MsgBox "Informe a Conta!", vbExclamation
        cmbxConta.SetFocus
    End If
    If cmbxEspecie.ListIndex = 0 Then
        If cmbxConta.Text = "" And cmbxEspecie.Text = "1 - Dinheiro" And optnFin.Value Then
                MsgBox "Escolha a Conta!", vbExclamation
                cmbxConta.SetFocus
                Exit Sub
            End If
        strSQL = Date & Chr(9) & txtbxValorBaixa.Text & Chr(9) & cmbxEspecie.Text
        strSQL = strSQL & Chr(9) & Chr(9) & Chr(9) & Chr(9) & cmbxConta.tag & Chr(9) & Chr(9) & Chr(9) & txtbxDesconto.Text
        msflxgrd.AddItem strSQL
        txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) + CCur(txtbxValorBaixa.Text))
        
    
    
    
    
    ElseIf cmbxEspecie.ListIndex = 1 Then
        Dim c As cheque
        If "R" = "R" Then
'            If cmbxConta.Text = "" And cmbxEspecie.Text = "1 - Dinheiro" Then
'                MsgBox "Escolha a Conta!", vbInformation
'                cmbxConta.SetFocus
'                Exit Sub
'            End If
            If che.nome_che = "" Then che.nome_che = txtbxCliente.Text
            If che.cpf_che = "" Then che.cpf_che = lblCPF.caption
            c = F_DadosCheque(True, che.Venc_che, che.nome_che, che.num_che, che.cco_che, che.age_che, che.cod_bco, CStr(che.Valor_che), che.cpf_che)
            If c.Valor_che = 0 Then
                Unload frmDadosChe
                Exit Sub
            End If
            Do While c.Valor_che > F_Restante
                MsgBox "O Valor a ser lançado deve ser menor ou igual ao remanescente!", vbExclamation
                c = F_DadosCheque(False, Date, txtbxCliente.Text, c.num_che, c.cco_che, c.age_che, c.cod_bco, "0,00")
        
            
            Loop
            strSQL = c.Venc_che & Chr(9) & c.Valor_che & Chr(9)
            strSQL = strSQL & cmbxEspecie.Text & Chr(9) & c.num_che & Chr(9)
            strSQL = strSQL & c.cod_bco & Chr(9) & c.age_che & Chr(9)
            strSQL = strSQL & c.cco_che & Chr(9) & c.nome_che & Chr(9)
            strSQL = strSQL & c.cpf_che & Chr(9) & txtbxDesconto.Text
            msflxgrd.AddItem strSQL
            che = c
            txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) + c.Valor_che)
        Else
            Dim cp As chePag
            cp = F_DadosChePag(Date, , , F_Restante)
            If cp.Valor_che = 0 Then
                Unload frmDadosChe
                Exit Sub
            End If
            If cp.Valor_che > F_Restante Then
                MsgBox "O Valor está maior que o restante!", vbExclamation
                Unload frmDadosChe
                Exit Sub
            End If
            strSQL = cp.Venc_che & Chr(9) & cp.Valor_che & Chr(9)
            strSQL = strSQL & cmbxEspecie.Text & Chr(9) & cp.num_che & Chr(9)
            strSQL = strSQL & cp.cod_con & Chr(9) & "" & Chr(9)
            strSQL = strSQL & cp.cod_con & Chr(9) & F_BuscaUmCampo("Nome_con", "tab_Contas", "cod_con='" & cp.num_che & "'")
            strSQL = strSQL & Chr(9) & txtbxDesconto.Text
            msflxgrd.AddItem strSQL
            txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) + cp.Valor_che)
        End If
    Else
            If cmbxConta.Text = "" And Left(cmbxEspecie.Text, 3) = "3" And optnFin.Value Then
                MsgBox "Escolha o Cartão!", vbExclamation
                cmbxConta.SetFocus
                Exit Sub
            End If
            strSQL = Date & Chr(9) & txtbxValorBaixa.Text & Chr(9) & cmbxEspecie.Text
            strSQL = strSQL & Chr(9) & Chr(9) & Chr(9) & Chr(9) & cmbxConta.tag
            msflxgrd.AddItem strSQL
            txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) + CCur(txtbxValorBaixa.Text))
    End If
    soma
    
    
    cmndbtnConfirma_Click
End Sub

Private Sub cmndbtnSair_Click()
    Me.tag = ""
    Me.Hide
End Sub


Sub somaABaixa()
    Dim iCount As Integer
    Dim cTotal As Currency
    Dim iRow As Integer
    Dim cJuros As Currency
    Dim cBai As Currency
    Dim cRem As Currency
    Dim cIni As Currency
    With msflxParcelas
    
        iRow = .row
        
        For iCount = 1 To .rows - 1
            .row = iCount
            .Col = 4: cRem = CCur(.Text)
            cIni = cIni + CCur(.Text)
            .Col = 5: cBai = CCur(.Text)
            .CellFontBold = True
            If cBai > 0 Then
                
                ZebraLinha msflxParcelas, .row, vbYellow
            Else
                ZebraLinha msflxParcelas, .row, vbWhite
            End If
            .Col = 5:
            .row = iCount
            cTotal = cTotal + CCur(.Text)
            
            .Col = 12
            
            cJuros = cJuros + (CCur(.Text) * cBai / cRem)
            
        Next
        
        
        
        
        .row = iRow
    End With
    
    txtbxDistr.Text = FormatVal(cTotal)
    txtbxReman.Text = FormatVal(cTotal)
    If bUsaJuro Then
        txtbxJuros.Text = FormatVal(cJuros)
    End If
    txtbxValor_Par.Text = FormatVal(cIni)
    
    
    
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        If txtbxCliente.tag <> "" Then
            cmndbtnBuscaCli_Click
        End If
    
        
    
    
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmbxEspecie.ListIndex = 0
        On Error Resume Next
        txtbxValorBaixa.SetFocus
    ElseIf KeyCode = vbKeyF5 Then
        cmndbtnBuscaCli_Click
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me, 0
    P_PreencheContas cmbxConta, 3
    msflxgrd.rows = 1
    msflxgrd.Cols = 10
    msflxgrd.Col = 0
    msflxgrd.Text = "Vencimento"
    msflxgrd.Col = 1
    msflxgrd.Text = "Valor"
    msflxgrd.Col = 2
    msflxgrd.Text = "Espécie"
    msflxgrd.Col = 3
    msflxgrd.Text = "Cheque Nº"
    msflxgrd.Col = 4
    msflxgrd.Text = "Banco"
    msflxgrd.Col = 5
    msflxgrd.Text = "Agencia"
    msflxgrd.Col = 6
    msflxgrd.Text = "Conta C."
    
    msflxgrd.Col = 7
    msflxgrd.Text = "Titular"
    msflxgrd.Col = 8
    msflxgrd.Text = "CPF"
    msflxgrd.Col = 9
    msflxgrd.Text = "Desconto"
    
    msflxgrd.ColAlignment(2) = 1
    
    msflxgrd.ColWidth(0) = 1000
    msflxgrd.ColWidth(1) = 1000
    msflxParcelas.FixedCols = 0
    msflxParcelas.rows = 1
    msflxParcelas.Cols = 3
    msflxParcelas.Col = 0:
    msflxParcelas.Text = "Seq."
    msflxParcelas.Col = 1:
    msflxParcelas.Text = "Valor"
    msflxParcelas.Col = 2:
    msflxParcelas.Text = "Juros"
    
    cmbxEspecie.ListIndex = 0
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnAção
    End If
End Sub

Private Sub mnConfBaixa_Click()
    cmndbtnConfirma_Click
End Sub

Private Sub mnEdit_Click()

    
    
    
        msflxParcelas.Col = 2
        frmEdiLan.tag = msflxParcelas.Text
        frmEdiLan.Show 1
        busca
End Sub

Private Sub mnExcl_Click()
    cmndbtnExcl_Click
End Sub

Private Sub mnLanc_Click()
    cmndbtnLancar_Click
End Sub

Private Sub mnMudarValor_Click()
    On Error GoTo trata_Erro
    msflxParcelas.Col = 5
    msflxParcelas_KeyDown 13, 0
    Exit Sub
trata_Erro:
    
End Sub

Private Sub mnSair_Click()
    cmndbtnSair_Click
End Sub

Private Sub mnSelec_Click()
    msflxParcelas_DblClick
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyDelete Then
        cmndbtnExcl_Click
    End If
    
End Sub

Private Sub msflxgrd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpcPag
    End If
End Sub

Private Sub msflxParcelas_DblClick()
    Dim cRem As Currency
    On Error GoTo trata_Erro
    With msflxParcelas
        .Col = 0
        
        If .CellBackColor = vbYellow Then
            .Col = 5: .Text = "0,00"
            ZebraLinha msflxParcelas, .row, vbWhite
            
            
        Else
            .Col = 4
            cRem = CCur(.Text)
            .Col = 5
            .Text = FormatVal(cRem)
            ZebraLinha msflxParcelas, .row, vbYellow
        End If
        somaABaixa
        txtbxValorBaixa_GotFocus
    End With
    
    Exit Sub
trata_Erro:
    
End Sub

Private Sub msflxParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If msflxParcelas.Col = 5 Then
            editaCelula
        Else
            msflxParcelas_DblClick
        End If
    End If
End Sub
Sub editaCelula()
    With msflxParcelas
        .Col = 4
        txtbxMsf.Text = .Text
        .Col = 5
        txtbxMsf.Left = .CellLeft + .Left
        txtbxMsf.Top = .CellTop + .Top
        txtbxMsf.width = .CellWidth
        txtbxMsf.height = .CellHeight
        txtbxMsf.Visible = True
        
        txtbxMsf.SetFocus
    End With
End Sub

Private Sub msflxParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnPar
    End If
End Sub

Private Sub optnCaixa_Click()
    cmbxConta.Enabled = False
    Dim sStatus As String
    If strCod_Cai = "" Or strCod_Tur = "" Then
        MsgBox "Operação impossível!", vbExclamation
        optnFin.Value = True
        Exit Sub
    End If
    sStatus = F_BuscaUmCampo("status_Cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
    If sStatus = "F" Then
        MsgBox "Não é possível efetuar esse recebimento pois o caixa se encontra fechado!", vbExclamation
        optnFin.Value = True
        Exit Sub
    End If
    
    
    
    
    
End Sub

Private Sub optnCli_Click()
    lblCli.caption = "&Cliente:"
End Sub

Private Sub optnFin_Click()
    cmbxConta.Enabled = True
    cmbxConta.SetFocus
End Sub

Private Sub optnRep_Click()
    lblCli.caption = "&Representante:"
End Sub

Private Sub txtbxCliente_GotFocus()
    P_SelecTexto txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaCli_Click
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
    
    
    P_ValorLostFocus txtbxDesconto
    
    If CCur(txtbxDesconto.Text) > 0 Then
        If Not F_senha Then
            txtbxDesconto.Text = "0,00"
        End If
    End If
    
    
End Sub

Private Sub txtbxDistr_GotFocus()
    P_SelecTexto txtbxDistr
End Sub

Private Sub txtbxDistr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDistr_LostFocus()
    P_ValorLostFocus txtbxDistr
End Sub

Private Sub txtbxJuros_GotFocus()
    P_SelecTexto txtbxJuros
End Sub

Private Sub txtbxJuros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxJuros_LostFocus()
    P_ValorLostFocus txtbxJuros
End Sub

Private Sub txtbxMsf_GotFocus()
    P_SelecTexto txtbxMsf
End Sub

Private Sub txtbxMsf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxMsf_LostFocus()
    P_ValorLostFocus txtbxMsf
    msflxParcelas.Text = txtbxMsf.Text
    txtbxMsf.Visible = False
    somaABaixa
    msflxParcelas.SetFocus
End Sub

Private Sub txtbxSeq_GotFocus()
    P_SelecTexto txtbxSeq
End Sub

Private Sub txtbxSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxSeq_LostFocus()
    txtbxSeq.Text = Format(txtbxSeq.Text, "000000#")
End Sub

Private Sub txtbxValorBaixa_GotFocus()
    txtbxValorBaixa.Text = F_Restante  'CCur(txtbxReman.Text) - CCur(txtbxTotalBaixa.Text)
    P_SelecTexto txtbxValorBaixa
        
End Sub
Function F_Restante() As Currency
    F_Restante = CCur(txtbxReman.Text) - CCur(txtbxTotalBaixa.Text) + CCur(txtbxJuros.Text) - CCur(txtbxDesconto.Text)
End Function
Sub soma()
    Dim cDesc
    Dim i As Integer
    With msflxgrd
        For i = 1 To .rows - 1
            .Col = 9
            .row = i
            If .Text <> "" Then
                cDesc = cDesc + CCur(.Text)
            End If
        Next
        
        
        
    End With
    txtTotalDesconto.Text = FormatVal(cDesc)
    
End Sub
Private Sub txtbxValorBaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLancar_Click
    End If
End Sub

Private Sub txtbxValorBaixa_LostFocus()
    P_ValorLostFocus txtbxValorBaixa
    If CCur(txtbxValorBaixa.Text) > F_Restante Then
        MsgBox "O Valor não pode ser maior que o restante!", vbExclamation
        txtbxValorBaixa_GotFocus
    '        txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    
End Sub
