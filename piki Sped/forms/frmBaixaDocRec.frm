VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBaixaDocReq 
   Caption         =   "Baixa de Parcela"
   ClientHeight    =   8925
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12855
   Icon            =   "frmBaixaDocRec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotalDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5760
      TabIndex        =   45
      Text            =   "0,00"
      Top             =   8400
      Width           =   975
   End
   Begin VB.TextBox txtbxSeq 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4200
      TabIndex        =   28
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Terceiro:"
      Height          =   495
      Left            =   5400
      TabIndex        =   37
      Top             =   0
      Width           =   3975
      Begin VB.OptionButton optnFun 
         Caption         =   "Funcionário"
         Height          =   195
         Left            =   2400
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optnRep 
         Caption         =   "Representante"
         Height          =   195
         Left            =   960
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optnCli 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   38
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
      TabIndex        =   36
      Text            =   "0,00"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnDistr 
      Height          =   285
      Left            =   11040
      Picture         =   "frmBaixaDocRec.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtbxDistr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      TabIndex        =   34
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
      TabIndex        =   30
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
      TabIndex        =   29
      Text            =   "0,00"
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtbxCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid msflxParcelas 
      Height          =   2775
      Left            =   120
      TabIndex        =   21
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
      TabIndex        =   18
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
      Left            =   5760
      TabIndex        =   17
      Text            =   "0,00"
      Top             =   7680
      Width           =   975
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
      TabIndex        =   13
      Text            =   "0,00"
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Cancela&r"
      Height          =   735
      Left            =   10080
      Picture         =   "frmBaixaDocRec.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar Baixa"
      Height          =   735
      Left            =   8640
      Picture         =   "frmBaixaDocRec.frx":087F
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Escolha a Forma de Pagamento:"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   12495
      Begin VB.CheckBox chkVale 
         Caption         =   "Vale"
         Height          =   375
         Left            =   7080
         TabIndex        =   41
         Top             =   2640
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optnCaixa 
         Caption         =   "Caixa"
         Height          =   195
         Left            =   2280
         TabIndex        =   6
         Top             =   2520
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optnFin 
         Caption         =   "Financeiro:"
         Height          =   195
         Left            =   3120
         TabIndex        =   5
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ComboBox cmbxConta 
         BackColor       =   &H00008000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmBaixaDocRec.frx":0FE9
         Left            =   2280
         List            =   "frmBaixaDocRec.frx":0FF3
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "!"
         Top             =   2760
         Width           =   3255
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   525
         Left            =   9960
         Picture         =   "frmBaixaDocRec.frx":0FFB
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnLancar 
         Caption         =   "&Lançar"
         Height          =   525
         Left            =   8520
         Picture         =   "frmBaixaDocRec.frx":1585
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtbxValorBaixa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox cmbxEspecie 
         Height          =   315
         ItemData        =   "frmBaixaDocRec.frx":1B0F
         Left            =   120
         List            =   "frmBaixaDocRec.frx":1B22
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2760
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   2175
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
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
         TabIndex        =   16
         Top             =   2160
         Width           =   180
      End
      Begin VB.Label lblCaixa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         Height          =   195
         Left            =   4200
         TabIndex        =   15
         Top             =   2160
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor a Baixar:"
         Height          =   195
         Left            =   5640
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
      Picture         =   "frmBaixaDocRec.frx":1B7F
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Desconto(R$):"
      Height          =   195
      Left            =   5760
      TabIndex        =   44
      Top             =   8160
      Width           =   1440
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
      TabIndex        =   43
      Top             =   8160
      Width           =   1065
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Situação Cliente:"
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   8280
      Width           =   1200
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seq.Doc."
      Height          =   195
      Left            =   4200
      TabIndex        =   27
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Baixar:"
      Height          =   195
      Left            =   9720
      TabIndex        =   33
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Juros(R$):"
      Height          =   195
      Left            =   3360
      TabIndex        =   32
      Top             =   7440
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto(R$):"
      Height          =   195
      Left            =   2280
      TabIndex        =   31
      Top             =   7440
      Width           =   1035
   End
   Begin VB.Label lblCli 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   120
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   600
      Width           =   12600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total (sem juros)"
      Height          =   195
      Left            =   6960
      TabIndex        =   20
      Top             =   7440
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Baixar:"
      Height          =   195
      Left            =   5760
      TabIndex        =   19
      Top             =   7440
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total da Baixa:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
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
Dim bActive As Boolean
Private Sub chkVale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxConta_Click()
    If Left(cmbxEspecie.text, 1) = "1" Or Left(cmbxEspecie.text, 1) = "5" Then
        cmbxConta.tag = F_BuscaUmCampo("cod_con", "tab_contas", "nome_con='" & cmbxConta.text & "'")
    
    Else
    
        cmbxConta.tag = F_BuscaUmCampo("cod_fin", "tab_financeiras", "nome_fin='" & cmbxConta.text & "'")
    End If
End Sub

Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxEspecie_Click()
    If cmbxEspecie.ListIndex = 1 Then
        txtbxValorBaixa.text = "0,00"
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
        
        If optnFun.Value Then
            chkVale.Visible = True
        Else
            chkVale.Visible = False
        End If
    End If
End Sub

Private Sub cmbxEspecie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Sub buscaCli()
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
        c = F_BuscaCliente(txtbxCliente.text)
    End If
    
    If c.cod_cli <> "" Then
        
        txtbxCliente.text = c.nome_Cli
        txtbxCliente.tag = c.cod_cli
        lblCpf.caption = c.cnpj_Cli
        lblCpf.tag = "C"
        busca
    End If

End Sub
Sub buscaFun()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectFuncionario(txtbxCliente.text)
    If f.cod_fun <> "" Then
        txtbxCliente.text = f.nome_fun
        txtbxCliente.tag = f.cod_fun
        lblCpf.caption = f.cpf_fun
        lblCpf.tag = "U"
        busca
    End If
End Sub

Sub buscaRep()
    Dim c As representante.representante
    c = F_BuscaRep(txtbxCliente.text)
    
    
    If c.cod_Rep <> "" Then
        
        txtbxCliente.text = c.nome_rep
        txtbxCliente.tag = c.cod_Rep
        lblCpf.caption = c.cpf_rep
        lblCpf.tag = "R"
        busca
    End If
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    ajustaGrid
    strSQL = F_SelectDocsReceber(lblCpf.tag, "", "", "", "", "", "", "", "", txtbxSeq.text, "", "", txtbxCliente.tag, 1)
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
            strSQL = strSQL & !nome_cob & Chr(9) & !cod_cob & Chr(9) & !historico_doc & Chr(9) & !boleto_itd
            msflxParcelas.AddItem strSQL
            
            
            
            msflxParcelas.row = msflxParcelas.rows - 1
            msflxParcelas.col = 14
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
Sub ajustaGrid()
    With msflxParcelas
        .Cols = 18
        .rows = 1
        .col = 0: .text = "..."
        .col = 1: .text = "Seq.Doc."
        .col = 2: .text = "Seq.Itd."
        .col = 3: .text = "Vencimento."
        .col = 4: .text = "Restante."
        .col = 5: .text = "Baixar"
        .col = 6: .text = "Inicial."
        .col = 7: .text = "Nº Doc."
        .col = 8: .text = "Valor Doc."
        .col = 9: .text = "Data Doc."
        .col = 10: .text = "Parcela."
        .col = 11: .text = "Dias."
        .col = 12: .text = "Acréscimo."
        .col = 13: .text = "Vlr Sug."
        
        .col = 14: .text = "Cob."
        .col = 15: .text = "Conta"
        .col = 16: .text = "Histórico."
        .col = 17: .text = "Boleto"
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
        buscaCli
    ElseIf optnRep.Value Then
        buscaRep
    ElseIf optnFun.Value Then
        buscaFun
    End If
End Sub

Private Sub cmndbtnConfirma_Click()
    Dim iParc As Integer
    Dim sSeq As String
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
'    On Error GoTo trata_Erro
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
    
    If CCur(txtbxTotalBaixa.text) <> CCur(txtbxReman.text) + CCur(txtbxJuros.text) - CCur(txtbxDesconto.text) Then
        MsgBox "Você só pode baixar o valor completo!", vbExclamation
        Exit Sub
    End If
    If CCur(txtbxDesconto.text) > 0 Then
        cDesc = CCur(txtbxDesconto.text) / CCur(txtbxDistr.text) * 100
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
            .col = 5
            If .text > 0 Then
                .col = 15
                sCob = .text
                .col = 2
                sSeq = .text
                .col = 5
                sVal = .text
                .col = 7
                sDoc = .text
                sDocs = sDocs & .text & "-"
                .col = 2
                If Left(.text, 1) = "1" Then
                
                Else
                
                End If
                strSQL = F_BuscaUmCampo("reman_itd", " " & TAB_ITDOC & " ", "  seq_itd='" & sSeq & "'")
                If strSQL = "" Then strSQL = "0"
                
                
                strSQL = "Update  " & TAB_ITDOC & "  set reman_itd=reman_itd-" & F_GravaDec(sVal)
                strSQL = strSQL & " where seq_itd='" & sSeq & "'"
                conn.Execute strSQL
                'P_ParcelaBaixadaSPC sSeq
            
                strSQL = F_GravaBaixaItd(sSeqBai, sSeq, sVal, sEsp, sSeqFin, sCob, Date, txtTotalDesconto.text, strCod_Cai)
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
            .col = 2
            sEsp = Left(.text, 1)
            .col = 0: c.Venc_che = .text
            .col = 1: c.Valor_che = .text
            .col = 3: c.num_che = .text
            .col = 4: c.cod_bco = .text
            .col = 5: c.age_che = .text
            .col = 6: c.cco_che = .text
            .col = 7: c.nome_che = .text
            
            .col = 9: sDesc = .text
            If sEsp = "1" Then
                
                    If optnFin.Value Then
                        sSeq = F_GravaFinanceiro("R", CStr(Date), CStr(c.Valor_che), _
                        "Baixa docs " & sDocs & txtbxCliente.text, _
                        cmbxConta.tag, "S", strPlaRecebDupl)
                        If sSeq = "" Then
                            MsgBox "Erro na gravação do movimento financeiro!", vbExclamation
                            'Exit Sub
                        End If
                        
                        
                        
                        sEspBai = "RDI" 'recebimento em dinheiro financeiro
'                        cBaixaFin.Add sSeq
                    Else
                        sSeq = F_GravaCai("R", CStr(Date), CStr(c.Valor_che), _
                        "Baixa docs " & sDocs & txtbxCliente.text, "", _
                            "RDI", strCod_Cai, CStr(c.Valor_che), 0, strCod_Cai, strCod_Tur, "000", "S", txtTotalDesconto.text)
                            sEspBai = "RDC" 'recebimento dinheiro caixa
                            
                        If chkVale.Value = 1 And sTipoTer = "U" Then
                            If F_GravaValesCai(sSeq, txtbxCliente.tag) Then
                                MsgBox "Lançamento em Dinheiro lançado como Vale", vbInformation
                                P_ImprimeValeCai sSeq
                            End If
                        End If
                            
                            
'                        cBaixaFin.Add sSeq
                    End If
                
                
                 'sSeq = F_GravaFinanceiro( "R", Date, CStr(c.Valor_che), "HISTORICO", "DIN", _
                        )
                        
                        If F_GravaBaixaItdFin(sSeqBai, sSeq, sEspBai, CStr(c.Valor_che), sDesc) = "" Then
                            MsgBox "Erro na gravação da baixa!", vbExclamation
                        End If
                        
                'P_Class c.Valor_che, "R", sSEQ, "F", adc.Recordset!historico_doc
            ElseIf sEsp = "2" Then
                sSeq = F_GravaDoc("BAIXA DOCS " & sDocs & txtbxCliente.text, _
                     Date, CStr(c.Valor_che), Date, _
                    txtbxCliente.tag, "R", "", "1", "F", sTipoTer, "000")
                    'If F_GravaBaixaItd(adc.Recordset!seq_itd, CStr(c.Valor_che), "DUP", sSeq) = "" Then
                    '        MsgBox "Erro na gravação da baixa!", vbExclamation
                    '    End If
                    
                If "R" = "R" Then
                    
                    sChe = F_GravaCh(sSeq, c.num_che, c.cod_bco, c.age_che, c.cco_che, _
                            c.nome_che, CCur(c.Valor_che), c.Venc_che, c.cod_bco, "R", _
                            "", "", "1", "", "")
                            
                            
                    If optnCaixa.Value = True Then
                        sSeqCai = F_GravaCai("C", Date, CStr(c.Valor_che), "Baixa docs " & sDocs & txtbxCliente.text _
                    , "", "RCH", strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, "000", "S", txtbxDesconto.text)
                    End If
                    
                    If F_GravaBaixaItdFin(sSeqBai, sChe, "CHP", CStr(c.Valor_che)) = "" Then
                            MsgBox "Erro na gravação da baixa!", vbExclamation
                        End If
                Else
                    sChe = F_GravaChePag(sSeq, c.num_che, c.cod_bco, _
                            CCur(c.Valor_che), c.Venc_che, adc.Recordset!codigo, adc.Recordset!parcela_itd)
                End If
            
            
            ElseIf sEsp = "5" Then
                
                    If optnFin.Value Then
                        sSeq = F_GravaFinanceiro("R", CStr(Date), CStr(c.Valor_che), _
                        "Baixa docs " & sDocs & txtbxCliente.text, _
                        cmbxConta.tag, "S", strPlaRecebDupl)
                        If sSeq = "" Then
                            MsgBox "Erro na gravação do movimento financeiro!", vbExclamation
                            'Exit Sub
                        End If
                        
                        
                        
                        sEspBai = "PIX" 'recebimento em dinheiro financeiro
'                        cBaixaFin.Add sSeq
                    Else
                        sSeq = F_GravaCai("R", CStr(Date), CStr(c.Valor_che), _
                        "Baixa docs " & sDocs & txtbxCliente.text, "", _
                            "RDI", strCod_Cai, CStr(c.Valor_che), 0, strCod_Cai, strCod_Tur, "000", "S", txtTotalDesconto.text)
                            sEspBai = "PIX" 'recebimento dinheiro caixa
                            
                        If chkVale.Value = 1 And sTipoTer = "U" Then
                            If F_GravaValesCai(sSeq, txtbxCliente.tag) Then
                                MsgBox "Lançamento em Dinheiro lançado como Vale", vbInformation
                                P_ImprimeValeCai sSeq
                            End If
                        End If
                            
                            
'                        cBaixaFin.Add sSeq
                    End If
                
                
                 'sSeq = F_GravaFinanceiro( "R", Date, CStr(c.Valor_che), "HISTORICO", "DIN", _
                        )
                        
                        If F_GravaBaixaItdFin(sSeqBai, sSeq, sEspBai, CStr(c.Valor_che), sDesc) = "" Then
                            MsgBox "Erro na gravação da baixa!", vbExclamation
                        End If
                        
            
            Else
                
                
                
                For iParc = 1 To c.num_che
                    sSeq = MaxKeySQL(TAB_VENFIN, "num_ven", "num_ven like 'B%'", "B00000#")
                    If F_GravaVenFin(sSeq, c.cco_che, c.Valor_che / c.num_che, Date, c.nome_che, "1", "B", CStr(iParc), DateAdd("D", IIf(sEsp = "3", 30 * iParc, 1), Date)) Then
                    End If
                Next
                If sEsp = "3" Then
                    sEspBai = "RCC"
                Else
                    sEspBai = "RCD"
                End If
                
                
                    sSeqCai = F_GravaCai("C", Date, CStr(c.Valor_che), "Baixa docs " & sDocs & txtbxCliente.text _
                    & "-" & Date, sSeq, _
                     sEspBai, strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, "000", "S", txtTotalDesconto.text)
                
                
                If F_GravaBaixaItdFin(sSeqBai, sSeq, "FIN", CStr(c.Valor_che)) = "" Then
                            MsgBox "Erro na gravação da baixa!", vbExclamation
                        End If
            End If
            
            cValor = cValor + c.Valor_che
        Next
        
        
            'strSQL = "update  " & tab_itdoc & "  set reman_itd = reman_itd-" & F_GravaDec(cValor - CCur(frmBaixaPar.txtbxJuros.Text) + CCur(frmBaixaPar.txtbxDesconto.Text))
            'strSQL = strSQL & ",juros_itd=" & F_GravaDec(frmBaixaPar.txtbxJuros.Text)
            'strSQL = strSQL & ",Desconto_itd=" & F_GravaDec(frmBaixaPar.txtbxDesconto.Text5)
            'strSQL = strSQL & ", baixa_itd=" & F_GravaData(Date)
            'strSQL = strSQL & " where seq_itd='" & adc.Recordset!seq_itd & "' "
            'conn.execute strSQL
        
         'retornar carnês
            
        With msflxParcelas
            For iCount = 1 To .rows - 1
                .row = iCount
                .col = 5
                If CCur(.text) > 0 Then
                    .col = 2
                    sSeqItd = .text
                    .col = 1
                    
                    strSQL = "update  " & TAB_ITDOC & "  set cod_cob='000' where seq_doc='" & .text & "' "
                
                    
                    'strSQL = strSQL & " and reman_itd<valor_itd "
                
                    conn.Execute strSQL
                
                End If
                
                
                
            Next
        End With
        MsgBox "Documento Baixado!", vbInformation
        
        ImprReciboBai sSeqBai
        
        
        
                
        
        
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
            .col = 5
            .text = "0"
            
            
        Next
        cRest = txtbxDistr.text
        If txtbxDistr.text > 0 Then
            For iCount = 1 To .rows - 1
                .row = iCount
                If cRest > 0 Then
                    
                    .col = 4
                    cVal = CCur(.text)
                    If cRest >= cVal Then
                        .col = 5
                        .text = cVal
                        cRest = cRest - cVal
                    Else
                        .col = 5
                        .text = cRest
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
                msflxgrd.col = 1
                cValor = CCur(msflxgrd.text)
                If msflxgrd.rows = 2 Then
                    msflxgrd.rows = 1
                Else
                    msflxgrd.RemoveItem msflxgrd.row
                End If
                txtbxTotalBaixa.text = FormatVal(CCur(txtbxTotalBaixa.text) - cValor)
            End If
            
        End If
End Sub

Private Sub cmndbtnLancar_Click()
    
    Dim strSQL As String
    On Error Resume Next
    If CCur(txtbxValorBaixa.text) > F_Restante Then   'cmbxEspecie.Text = "1 - Dinheiro" And + CCur(txtbxTotalBaixa.Text)
        MsgBox "O Valor a ser lançado deve ser menor ou igual ao remanescente!", vbExclamation
          
        'txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    
    If optnCaixa.Value Then
        If strCod_Cai = "" Or strCod_Cai = "00" Or strCod_Cai = "99" Then
            MsgBox "Este computador não é um caixa, por tanto não pode lançar no caixa por aqui!", vbExclamation
            Exit Sub
        End If
    End If
    If cmbxEspecie.text = "" Then
        MsgBox "Informe o tipo de pagamento!", vbExclamation
        cmbxEspecie.SetFocus
        Exit Sub
    End If
    If CCur(txtbxValorBaixa.text) = "0" And Left(cmbxEspecie.text, 1) = "1" Then
        MsgBox "Informe o valor que deseja lançar para a baixa e clique em lançar!", vbExclamation
        txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    If optnFin.Value And cmbxConta.tag = "" Then
        MsgBox "Informe a Conta!", vbExclamation
        cmbxConta.SetFocus
    End If
    If cmbxEspecie.ListIndex = 0 Then
        If cmbxConta.text = "" And cmbxEspecie.text = "1 - Dinheiro" And optnFin.Value Then
                MsgBox "Escolha a Conta!", vbExclamation
                cmbxConta.SetFocus
                Exit Sub
            End If
        strSQL = Date & Chr(9) & txtbxValorBaixa.text & Chr(9) & cmbxEspecie.text
        strSQL = strSQL & Chr(9) & Chr(9) & Chr(9) & Chr(9) & cmbxConta.tag & Chr(9) & Chr(9) & Chr(9) & txtbxDesconto.text
        msflxgrd.AddItem strSQL
        txtbxTotalBaixa.text = FormatVal(CCur(txtbxTotalBaixa.text) + CCur(txtbxValorBaixa.text))
        
    
    
    
    
    ElseIf cmbxEspecie.ListIndex = 1 Then
        Dim c As cheque
        If "R" = "R" Then
'            If cmbxConta.Text = "" And cmbxEspecie.Text = "1 - Dinheiro" Then
'                MsgBox "Escolha a Conta!", vbInformation
'                cmbxConta.SetFocus
'                Exit Sub
'            End If
            If che.nome_che = "" Then che.nome_che = txtbxCliente.text
            If che.cpf_che = "" Then che.cpf_che = lblCpf.caption
            c = F_DadosCheque(True, che.Venc_che, che.nome_che, che.num_che, che.cco_che, che.age_che, che.cod_bco, CStr(che.Valor_che), che.cpf_che)
            If c.Valor_che = 0 Then
                Unload frmDadosChe
                Exit Sub
            End If
            Do While c.Valor_che > F_Restante
                MsgBox "O Valor a ser lançado deve ser menor ou igual ao remanescente!", vbExclamation
                c = F_DadosCheque(False, Date, txtbxCliente.text, c.num_che, c.cco_che, c.age_che, c.cod_bco, "0,00")
        
            
            Loop
            strSQL = c.Venc_che & Chr(9) & c.Valor_che & Chr(9)
            strSQL = strSQL & cmbxEspecie.text & Chr(9) & c.num_che & Chr(9)
            strSQL = strSQL & c.cod_bco & Chr(9) & c.age_che & Chr(9)
            strSQL = strSQL & c.cco_che & Chr(9) & c.nome_che & Chr(9)
            strSQL = strSQL & c.cpf_che & Chr(9) & txtbxDesconto.text
            msflxgrd.AddItem strSQL
            che = c
            txtbxTotalBaixa.text = FormatVal(CCur(txtbxTotalBaixa.text) + c.Valor_che)
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
            strSQL = strSQL & cmbxEspecie.text & Chr(9) & cp.num_che & Chr(9)
            strSQL = strSQL & cp.cod_con & Chr(9) & "" & Chr(9)
            strSQL = strSQL & cp.cod_con & Chr(9) & F_BuscaUmCampo("Nome_con", "tab_Contas", "cod_con='" & cp.num_che & "'")
            strSQL = strSQL & Chr(9) & txtbxDesconto.text
            msflxgrd.AddItem strSQL
            txtbxTotalBaixa.text = FormatVal(CCur(txtbxTotalBaixa.text) + cp.Valor_che)
        End If
    Else
            Dim p As pagamentocartao
            If Left(cmbxEspecie.text, 1) <> "5" Then
                p = F_DadosCartao(, , , , txtbxValorBaixa.text)
                
                If cmbxConta.text = "" And Left(cmbxEspecie.text, 3) = "3" And optnFin.Value Then
                    MsgBox "Escolha o Cartão!", vbExclamation
                    cmbxConta.SetFocus
                    Exit Sub
                End If
            End If
            strSQL = Date & Chr(9) & txtbxValorBaixa.text & Chr(9) & cmbxEspecie.text
            strSQL = strSQL & Chr(9) & p.parcelas & Chr(9) & p.cod_fin & Chr(9) & Chr(9) & p.cod_fin & Chr(9) & p.autorizacao
            msflxgrd.AddItem strSQL
            txtbxTotalBaixa.text = FormatVal(CCur(txtbxTotalBaixa.text) + CCur(txtbxValorBaixa.text))
    End If
    soma
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
            .col = 4: cRem = CCur(.text)
            cIni = cIni + CCur(.text)
            .col = 5: cBai = CCur(.text)
            .CellFontBold = True
            If cBai > 0 Then
                
                ZebraLinha msflxParcelas, .row, vbYellow
            Else
                ZebraLinha msflxParcelas, .row, vbWhite
            End If
            .col = 5:
            .row = iCount
            cTotal = cTotal + CCur(.text)
            
            .col = 12
            
            cJuros = cJuros + (CCur(.text) * cBai / cRem)
            
        Next
        
        
        
        
        .row = iRow
    End With
    
    txtbxDistr.text = FormatVal(cTotal)
    txtbxReman.text = FormatVal(cTotal)
    txtbxJuros.text = FormatVal(cJuros)
    txtbxValor_Par.text = FormatVal(cIni)
    
    
    
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    
    If Me.tag <> "" Then
        If txtbxCliente.tag <> "" Then
            cmndbtnBuscaCli_Click
        End If
    
    End If
    bActive = True
End Sub

Private Sub Form_Load()
    bActive = False
    P_CentraForm Me, 0
    P_PreencheContas cmbxConta, 3
    msflxgrd.rows = 1
    msflxgrd.Cols = 10
    msflxgrd.col = 0
    msflxgrd.text = "Vencimento"
    msflxgrd.col = 1
    msflxgrd.text = "Valor"
    msflxgrd.col = 2
    msflxgrd.text = "Espécie"
    msflxgrd.col = 3
    msflxgrd.text = "Cheque Nº"
    msflxgrd.col = 4
    msflxgrd.text = "Banco"
    msflxgrd.col = 5
    msflxgrd.text = "Agencia"
    msflxgrd.col = 6
    msflxgrd.text = "Conta C."
    
    msflxgrd.col = 7
    msflxgrd.text = "Titular"
    msflxgrd.col = 8
    msflxgrd.text = "CPF"
    msflxgrd.col = 9
    msflxgrd.text = "Desconto"
    
    msflxgrd.ColAlignment(2) = 1
    
    msflxgrd.ColWidth(0) = 1000
    msflxgrd.ColWidth(1) = 1000
    msflxParcelas.FixedCols = 0
    msflxParcelas.rows = 1
    msflxParcelas.Cols = 3
    msflxParcelas.col = 0:
    msflxParcelas.text = "Seq."
    msflxParcelas.col = 1:
    msflxParcelas.text = "Valor"
    msflxParcelas.col = 2:
    msflxParcelas.text = "Juros"
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

    
    
    
        msflxParcelas.col = 2
        frmEdiLan.tag = msflxParcelas.text
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
    msflxParcelas.col = 5
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
        .col = 0
        
        If .CellBackColor = vbYellow Then
            .col = 5: .text = "0,00"
            ZebraLinha msflxParcelas, .row, vbWhite
            
            
        Else
            .col = 4
            cRem = CCur(.text)
            .col = 5
            .text = FormatVal(cRem)
            ZebraLinha msflxParcelas, .row, vbYellow
        End If
        somaABaixa
    End With
    
    Exit Sub
trata_Erro:
    
End Sub

Private Sub msflxParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If msflxParcelas.col = 5 Then
            editaCelula
        Else
            msflxParcelas_DblClick
        End If
    End If
End Sub
Sub editaCelula()
    With msflxParcelas
        .col = 4
        txtbxMsf.text = .text
        .col = 5
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
    If strCod_Cai = "" Then
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
    msflxParcelas.text = txtbxMsf.text
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
    txtbxSeq.text = Format(txtbxSeq.text, "000000#")
End Sub

Private Sub txtbxValorBaixa_GotFocus()
    txtbxValorBaixa.text = F_Restante  'CCur(txtbxReman.Text) - CCur(txtbxTotalBaixa.Text)
    P_SelecTexto txtbxValorBaixa
        
End Sub
Function F_Restante() As Currency
    F_Restante = CCur(txtbxReman.text) - CCur(txtbxTotalBaixa.text) + CCur(txtbxJuros.text) - CCur(txtbxDesconto.text)
End Function
Sub soma()
    Dim cDesc
    Dim i As Integer
    With msflxgrd
        For i = 1 To .rows - 1
            .col = 9
            .row = i
            If .text <> "" Then
                cDesc = cDesc + CCur(.text)
            End If
        Next
        
        
        
    End With
    txtTotalDesconto.text = FormatVal(cDesc)
    
End Sub
Private Sub txtbxValorBaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValorBaixa_LostFocus()
    P_ValorLostFocus txtbxValorBaixa
    If CCur(txtbxValorBaixa.text) > F_Restante Then
        MsgBox "O Valor não pode ser maior que o restante!", vbExclamation
        txtbxValorBaixa_GotFocus
    '        txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    
End Sub
