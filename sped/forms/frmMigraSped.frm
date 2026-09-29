VERSION 5.00
Begin VB.Form frmMigraSped 
   Caption         =   "Migrar SPED"
   ClientHeight    =   2550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPart 
      Caption         =   "Participantes"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdMigraNfe 
      Caption         =   "NF-e"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cupons"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtAte 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdCup 
         Caption         =   "Cupons Fiscais"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtDe 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Até"
         Height          =   195
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmMigraSped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCup_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rItens As New ADODB.Recordset
    Dim cod_sit As String
    Dim id As String
    Dim cst As String
    
    
    
    
    strSQL = "select * from cupons where 1>0 "
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and data_cup>=" & F_GravaDataTempo(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and data_cup<=" & F_GravaDataTempo(txtAte.Text & " 23:59")
    End If
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    
    With rPesq
        Do While Not .EOF
            strSQL = "select tim.perc_imp, ti.*,tp.unidade_pro,nfe_cst from tab_itcup as ti, " & TAB_PRODUTOS & " as tp,"
            strSQL = strSQL & "tab_impostos as tim where tim.cod_imp=tp.cod_imp"
            strSQL = strSQL & " and tp.cod_pro=ti.cod_pro and seq_cup=" & !seq_cup
            
            If !cancelado_cup = "S" Then
                cod_sit = "01" 'cancelado
            Else
                cod_sit = "00" 'normal
            End If
            
            id = SpedCupom.gravasped_cupom(!num_ven, !cod_cai, Format(!data_cup, "dd/mm/yyyy"), Format(!data_cup, "hh:mm:ss"), "2D", _
                cod_sit, CStr(!num_cup), CStr(!valor_cup), "0", "0", "", "")
            If id = "" Then
                MsgBox "Erro na gravação do spedcupom"
            End If
            abreObjetoAdo rItens, strSQL
            With rItens
                Do While Not .EOF
                    
                    If IsNull(!nfe_cst) Then
                        cst = "00"
                    Else
                        cst = !nfe_cst
                    End If
                    SpedCupom.gravasped_cupom_itens id, !cod_pro, !qtd_ite, "0", !unidade_pro, !unit_ite * !qtd_ite, cst, _
                    "5102", !perc_imp, !unit_ite * !qtd_ite * 0.00065, !unit_ite * !qtd_ite * 0.003, !unit_ite * !qtd_ite * !perc_imp / 100, _
                    CStr(!desconto_ite)
                                         
                        
                    .MoveNext
                Loop
            End With
            Me.caption = .AbsolutePosition & "/" & .RecordCount
            Me.Refresh
                    
            .MoveNext
        Loop
    End With
    
End Sub

Private Sub cmdMigraNfe_Click()
    Dim strSQL As String
    Dim ado As New ADODB.Recordset
    Dim sPart As String
    Dim r2 As New ADODB.Recordset
    Dim sId As String
    strSQL = "delete from sped_nf where ind_oper=1"
    conn.Execute strSQL
    strSQL = "select tc.cpf_cli,ts.* from tab_saidasnf as ts, " & TAB_CLIENTES & " as tc"
    strSQL = strSQL & " where tc.cod_cli=ts.cod_cli order by seq_sai"
    
    abreObjetoAdo ado, strSQL
    
    
    
    
    With ado
        Do While Not .EOF
            sId = BuscaParticipanteCodigo(ado!cpf_cli)
            strSQL = SpedNF.GravaSpedNf("1", sPart, sPart, "55", "00", !serie_sai, !nota_sai, "", !emissao_sai, _
            !emissao_sai, CStr(!valor_sai), "0", "0", "0", !valp_sai, "9", CStr(!frete_sai), CStr(!seguro_sai), _
            CStr(!despesas_sai), CStr(!bc_sai), "0", "0", CStr(!vicms_sai), CStr(!ipi_sai), CStr(!pis_sai), _
            CStr(!cofins_sai), "0", "0", "4", "", "")
            
            strSQL = "select tp.*,ti.* from " & TAB_PRODUTOS & " as tp, tab_itsainf as ti where tp.cod_pro=ti.cod_pro"
            strSQL = strSQL & " and ti.seq_sai='" & !Seq_sai & "'"
            abreObjetoAdo r2, strSQL
            With r2
                Do While Not .EOF
                    If Not SpedNF.InsereItemSpedNF(sId, .AbsolutePosition, !cod_pro, !cod_pro, !nome_pro, !qtd_its, _
                    !unidade_pro, CStr(!qtd_its), !unidade_pro, CStr(!qtd_its * !unit_its), "0", !vDesconto, "0", Format(!nfe_cst), Format(!CFOP), _
                    "", CStr(!bc_icms), CStr(!aliq_icms), CStr(!vICMS), "0", "0", "0", "0", "99", "", "0", "0", "0", "49", _
                    "0", "0", "0", "0", "0", "", "0", "0", "0", "0", "0", "", !codbarra_pro, CStr(!qtd_its * !qtd_emb), "", "0", "0", "N") Then
                        MsgBox "erro na inserção!", vbExclamation
                    End If
                    .MoveNext
                Loop
            End With
            
            
            
            .MoveNext
        Loop
    End With
    
    
End Sub

Private Sub cmdPart_Click()
    Dim strSQL As String
    Dim ado As New ADODB.Recordset
    Dim sPart As String
    
    strSQL = "select tc.* from tab_saidasnf as ts, " & TAB_CLIENTES & " as tc "
    strSQL = strSQL & "  where tc.cod_cli=ts.cod_cli order by cod_cli "
    
    abreObjetoAdo ado, strSQL
    
    
    
    With ado
        Do While Not .EOF
            sPart = BuscaParticipanteCodigo(!cpf_cli)
            If sPart = "" Then
                If !cod_cli <> "00000" Then
                    If GravaParticipante(!nome_Cli, "1058", IIf(Len(F_SoNumero(!cpf_cli)) > 11, !cpf_cli, ""), IIf(Len(F_SoNumero(!cpf_cli)) > 11, "", !cpf_cli), Format(!ie_Cli), !cod_cid, _
                        "", Format(!endereco_cli), Format(!NUMERO_cli), Format(!COMPLEMENTO_cli), Format(!bairro_cli), !cod_cli, "", Format(!cep_cli), !uf_cli) = False Then
                         MsgBox "erro"
                    End If
                End If
            Else
            End If
            .MoveNext
        Loop
    End With
    
    
    strSQL = "select tc.* from " & TAB_FORNECEDORES & " as tc"
    
    abreObjetoAdo ado, strSQL
    
    
    
    With ado
        Do While Not .EOF
            sPart = BuscaParticipanteCodigo(!cnpj_for)
            If sPart = "" Then
                'If GravaParticipante(!nome_for, "1058", !cnpj_for, "", !ie_for, _
                '      BuscaCidadePorNome(!cidade_For, !uf_for), "", !endereco_for, "", "", !bairro_for, "", !cod_for, !cep_for) Then
                '
                'End If
            Else
            End If
            .MoveNext
        Loop
        MsgBox "ok"
    End With
    
End Sub

Private Sub Form_Load()
    txtDe.Text = F_PrimDiaMes
    txtAte.Text = F_UltDiaMes
End Sub
