VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBaixaDocReq2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa de Parcela"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   Icon            =   "frmBaixaDocRec2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msflxParcelas 
      Height          =   4935
      Left            =   120
      TabIndex        =   27
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8705
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
      Left            =   3120
      TabIndex        =   24
      Text            =   "0,00"
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtbxReman 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2040
      TabIndex        =   23
      Text            =   "0,00"
      Top             =   5760
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
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Cancela&r"
      Height          =   735
      Left            =   9960
      Picture         =   "frmBaixaDocRec2.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar Baixa"
      Height          =   735
      Left            =   8520
      Picture         =   "frmBaixaDocRec2.frx":04F5
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.OptionButton optnCaixa 
         Caption         =   "Caixa"
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   3960
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optnFin 
         Caption         =   "Financeiro:"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   4200
         Width           =   1095
      End
      Begin VB.ComboBox cmbxConta 
         BackColor       =   &H00008000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmBaixaDocRec2.frx":0C5F
         Left            =   3720
         List            =   "frmBaixaDocRec2.frx":0C69
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "!"
         Top             =   4050
         Width           =   3255
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   525
         Left            =   7080
         Picture         =   "frmBaixaDocRec2.frx":0C71
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0005C9EB&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtbxJuros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0005C9EB&
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnLancar 
         Caption         =   "&Lançar"
         Height          =   525
         Left            =   5760
         Picture         =   "frmBaixaDocRec2.frx":11FB
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtbxValorBaixa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.ComboBox cmbxEspecie 
         Height          =   315
         ItemData        =   "frmBaixaDocRec2.frx":1785
         Left            =   120
         List            =   "frmBaixaDocRec2.frx":1795
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   4050
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   3375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5953
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
         Left            =   4440
         TabIndex        =   20
         Top             =   2280
         Width           =   180
      End
      Begin VB.Label lblCaixa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         Height          =   195
         Left            =   4200
         TabIndex        =   19
         Top             =   2280
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto(R$):"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   4440
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Juros(R$):"
         Height          =   195
         Left            =   1200
         TabIndex        =   16
         Top             =   4440
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor a Baixar:"
         Height          =   195
         Left            =   7080
         TabIndex        =   3
         Top             =   3840
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&spécie:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   3840
         Width           =   615
      End
   End
   Begin VB.Label lblCPF 
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
      Left            =   120
      TabIndex        =   29
      Tag             =   "!"
      Top             =   5160
      Width           =   645
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas:"
      Height          =   195
      Left            =   1200
      TabIndex        =   28
      Top             =   0
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicial:"
      Height          =   195
      Left            =   3120
      TabIndex        =   26
      Top             =   5520
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remanesc.:"
      Height          =   195
      Left            =   2040
      TabIndex        =   25
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblNome_Ter 
      AutoSize        =   -1  'True
      Caption         =   "Terceiro:"
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
      Left            =   3000
      TabIndex        =   22
      Tag             =   "!"
      Top             =   5160
      Width           =   780
   End
   Begin VB.Label lblParc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parcela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Tag             =   "!"
      Top             =   6120
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total da Baixa:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   1065
   End
End
Attribute VB_Name = "frmBaixaDocReq2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim che As cheque
Private Sub cmbxConta_Click()
    If Left(cmbxEspecie.Text, 1) = "1" Then
        cmbxConta.tag = F_BuscaUmCampo("cod_con", "tab_contas", "nome_con='" & cmbxConta.Text & "'")
    
    Else
    
        cmbxConta.tag = F_BuscaUmCampo("cod_fin", "tab_financeiras", "nome_fin='" & cmbxConta.Text & "'")
    End If
End Sub

Private Sub cmbxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        cmbxConta.Enabled = True
    End If
End Sub

Private Sub cmbxEspecie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnConfirma_Click()
    Dim sSeq As String
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
    On Error GoTo trata_Erro
    Dim cBaixaItd As New Collection
    Set cBaixaItd = Nothing
    If msflxParcelas.rows = 1 Then
        MsgBox "Não foi informado nenhuma parcela para ser baixada!", vbExclamation
    End If
    
    If msflxgrd.rows = 1 Then
        MsgBox "Não foi informado nenhum valor para ser baixado!", vbExclamation
    Else
        
    End If
    
    If CCur(txtbxTotalBaixa.Text) <> CCur(txtbxReman.Text) + CCur(txtbxJuros.Text) - CCur(txtbxDesconto.Text) Then
        MsgBox "Você só pode baixar o valor completo!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma Baixa dos Documento?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
        
    End If
    sSeqBai = financeiro.F_BuscaChaveBaixa
    
    
    With msflxParcelas
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 0
            sSeq = .Text
            .Col = 1
            sVal = .Text
            .Col = 3
            sDoc = .Text
            
            strSQL = "Update tab_itdoc set reman_itd=reman_itd-" & F_GravaDec(sVal)
            strSQL = strSQL & " where seq_itd='" & sSeq & "'"
            Banco.Execute strSQL
            
            strSQL = F_GravaBaixaItd(sSeqBai, sSeq, sVal, "", "", "")
            
            P_VerificaBaixaComissProv sSeq, F_GravaDec(sVal)
            
            
        Next
    End With
    
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 2
            sEsp = Left(.Text, 1)
            .Col = 0: c.Venc_che = .Text
            .Col = 1: c.Valor_che = .Text
            .Col = 3: c.num_che = .Text
            .Col = 4: c.cod_bco = .Text
            .Col = 5: c.age_che = .Text
            .Col = 6: c.cco_che = .Text
            .Col = 7: c.nome_che = .Text
            If sEsp = "1" Then
                
                    If optnFin.Value Then
                        sSeq = F_GravaFinanceiro("R", CStr(Date), CStr(c.Valor_che), _
                        "Baixa docs " & lblNome_Ter.caption, _
                        cmbxConta.tag, "S", "000")
                        sEspBai = "RDI" 'recebimento em dinheiro financeiro
'                        cBaixaFin.Add sSeq
                    Else
                        sSeq = F_GravaCai("R", CStr(Date), CStr(c.Valor_che), _
                        "Baixa docs " & lblNome_Ter.caption, "", _
                            "RDI", strCod_Cai, CStr(c.Valor_che), 0, strCod_Cai, strCod_Tur, "000", "S")
                            sEspBai = "RDC" 'recebimento dinheiro caixa
'                        cBaixaFin.Add sSeq
                    End If
                
                
                 'sSeq = F_GravaFinanceiro("R", Date, CStr(c.Valor_che), "HISTORICO", "DIN", _
                        frmBaixaPar.cmbxConta.Tag, dtNot.Recordset!seq_itd, "1", "")
                        
                        If F_GravaBaixaItdFin(sSeqBai, sSeq, sEspBai, CStr(c.Valor_che)) = "" Then
                            MsgBox "Erro na gravação da baixa!", vbExclamation
                        End If
                        
                'P_Class c.Valor_che, "R", sSEQ, "F", dtNot.Recordset!historico_doc
            ElseIf sEsp = "2" Then
                sSeq = F_GravaDoc("BAIXA DOCS" & lblNome_Ter.caption, _
                     Date, CStr(c.Valor_che), Date, _
                    lblNome_Ter.tag, "R", "", "1", "F", "C", "000")
                    'If F_GravaBaixaItd(dtNot.Recordset!seq_itd, CStr(c.Valor_che), "DUP", sSeq) = "" Then
                    '        MsgBox "Erro na gravação da baixa!", vbExclamation
                    '    End If
                    
                If "R" = "R" Then
                    
                    sChe = F_GravaCh(sSeq, c.num_che, c.cod_bco, c.age_che, c.cco_che, _
                            c.nome_che, CCur(c.Valor_che), c.Venc_che, c.cod_bco, "R", _
                            "", "", "1", "", "")
                            
                            
                    If optnCaixa.Value = True Then
                        sSeqCai = F_GravaCai("C", Date, CStr(c.Valor_che), "Baixa docs " & lblNome_Ter.caption _
                    , "", "RCH", strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, "000", "S")
                    End If
                    
                    If F_GravaBaixaItdFin(sSeqBai, sChe, "CHP", CStr(c.Valor_che)) = "" Then
                            MsgBox "Erro na gravação da baixa!", vbExclamation
                        End If
                Else
                    sChe = F_GravaChePag(sSeq, c.num_che, c.cod_bco, _
                            CCur(c.Valor_che), c.Venc_che, dtNot.Recordset!Codigo, dtNot.Recordset!parcela_itd)
                End If
            Else
                sSeq = MaxKeySQL("tab_venfin", "num_ven", "num_ven like 'B%'", "B00000#")
                
                
                
                If F_GravaVenFin(sSeq, c.cco_che, c.Valor_che, Date, "", "1", "B") Then
                End If
                
                If sEsp = "3" Then
                    sEspBai = "RCC"
                Else
                    sEspBai = "RCD"
                End If
                
                If optnCaixa.Value Then
                    sSeqCai = F_GravaCai("C", Date, CStr(c.Valor_che), "Baixa doc " & dtNot.Recordset!num_doc _
                    & "-" & dtNot.Recordset!nome, sSeq, _
                     sEspBai, strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, dtNot.Recordset!cod_pla, "S")
                End If
                
                If F_GravaBaixaItdFin(sSeqBai, sSeq, "FIN", CStr(c.Valor_che)) = "" Then
                            MsgBox "Erro na gravação da baixa!", vbExclamation
                        End If
            End If
            
            cValor = cValor + c.Valor_che
        Next
        
        
            'strSQL = "update tab_itdoc set reman_itd = reman_itd-" & F_GravaDec(cValor - CCur(frmBaixaPar.txtbxJuros.Text) + CCur(frmBaixaPar.txtbxDesconto.Text))
            'strSQL = strSQL & ",juros_itd=" & F_GravaDec(frmBaixaPar.txtbxJuros.Text)
            'strSQL = strSQL & ",Desconto_itd=" & F_GravaDec(frmBaixaPar.txtbxDesconto.Text)
            'strSQL = strSQL & ", baixa_itd=" & F_GravaData(Date)
            'strSQL = strSQL & " where seq_itd='" & dtNot.Recordset!seq_itd & "' "
            'Banco.Execute strSQL
        
         'retornar carnês
            
        With msflxParcelas
            For iCount = 1 To .rows - 1
                .Col = 6
                .Row = iCount
                strSQL = "update tab_itdoc set cod_cob='000' where seq_doc='" & .Text & "' "
                strSQL = strSQL & " and reman_itd>0 "
                Banco.Execute strSQL
                
                
            Next
        End With
        MsgBox "Documento Baixado!", vbInformation
        
        ImprReciboBai sSeqBai
        Unload Me
    End With

Exit Sub
trata_Erro:
    P_MostraErro
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
                    msflxgrd.RemoveItem msflxgrd.Row
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
    If CCur(txtbxValorBaixa.Text) = "0" Then
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
        strSQL = strSQL & Chr(9) & Chr(9) & Chr(9) & Chr(9) & cmbxConta.tag
        msflxgrd.AddItem strSQL
        txtbxTotalBaixa.Text = FormatVal(CCur(txtbxTotalBaixa.Text) + CCur(txtbxValorBaixa.Text))
        
    
    
    
    
    ElseIf cmbxEspecie.ListIndex = 1 Then
        Dim c As cheque
        If cmndbtnCOnfirma.tag = "R" Then
'            If cmbxConta.Text = "" And cmbxEspecie.Text = "1 - Dinheiro" Then
'                MsgBox "Escolha a Conta!", vbInformation
'                cmbxConta.SetFocus
'                Exit Sub
'            End If
            If che.nome_che = "" Then che.nome_che = lblNome_Ter.caption
            If che.cpf_che = "" Then che.cpf_che = lblCPF.caption
            c = F_DadosCheque(True, che.Venc_che, che.nome_che, che.num_che, che.cco_che, che.age_che, che.cod_bco, CStr(che.Valor_che), che.cpf_che)
            If c.Valor_che = 0 Then
                Unload frmDadosChe
                Exit Sub
            End If
            Do While c.Valor_che > F_Restante
                MsgBox "O Valor a ser lançado deve ser menor ou igual ao remanescente!", vbExclamation
                c = F_DadosCheque(False, Date, lblNome_Ter.caption, c.num_che, c.cco_che, c.age_che, c.cod_bco, "0,00")
        
            
            Loop
            strSQL = c.Venc_che & Chr(9) & c.Valor_che & Chr(9)
            strSQL = strSQL & cmbxEspecie.Text & Chr(9) & c.num_che & Chr(9)
            strSQL = strSQL & c.cod_bco & Chr(9) & c.age_che & Chr(9)
            strSQL = strSQL & c.cco_che & Chr(9) & c.nome_che & Chr(9)
            strSQL = strSQL & c.cpf_che
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
End Sub

Private Sub cmndbtnSair_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub Form_Load()
    P_CentraForm Me, 0
    P_PreencheContas cmbxConta, 3
    msflxgrd.rows = 1
    msflxgrd.Cols = 9
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
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyDelete Then
        cmndbtnExcl_Click
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

Private Sub optnFin_Click()
    cmbxConta.Enabled = True
    cmbxConta.SetFocus
End Sub

Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    P_ValorLostFocus txtbxDesconto
End Sub

Private Sub txtbxJuros_GotFocus()
    P_SelecTexto txtbxJuros
End Sub

Private Sub txtbxJuros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxJuros_LostFocus()
    P_ValorLostFocus txtbxJuros
End Sub

Private Sub txtbxValorBaixa_GotFocus()
    txtbxValorBaixa.Text = F_Restante  'CCur(txtbxReman.Text) - CCur(txtbxTotalBaixa.Text)
    P_SelecTexto txtbxValorBaixa
        
End Sub
Function F_Restante() As Currency
    F_Restante = CCur(txtbxReman.Text) - CCur(txtbxTotalBaixa.Text) + CCur(txtbxJuros.Text) - CCur(txtbxDesconto.Text)
End Function
Sub soma()
    
    
End Sub
Private Sub txtbxValorBaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValorBaixa_LostFocus()
    P_ValorLostFocus txtbxValorBaixa
    If CCur(txtbxValorBaixa.Text) > F_Restante Then
        MsgBox "O Valor não pode ser maior que o restante!", vbExclamation
        txtbxValorBaixa_GotFocus
        txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    
End Sub
