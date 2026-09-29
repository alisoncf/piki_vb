VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSped 
   Caption         =   "SPED"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFim 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   14
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtIni 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eventos SPED"
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtCaminho 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "C:\"
      Top             =   5160
      Width           =   3855
   End
   Begin VB.PictureBox pcBlocos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3705
      ScaleWidth      =   7785
      TabIndex        =   4
      Top             =   840
      Width           =   7815
      Begin MSComctlLib.TreeView tree 
         Height          =   2895
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5106
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.CheckBox chkZero 
         Caption         =   "Bloco 0"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox txtMes 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fim:"
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   285
   End
   Begin VB.Label lblIni 
      AutoSize        =   -1  'True
      Caption         =   "Início"
      Height          =   195
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salvar em:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   750
   End
   Begin VB.Label Label2 
      Caption         =   "Mês:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Ano:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmSped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGerar_Click()
    geraSped
End Sub
Sub geraSped()
    Dim s As String
    Dim ado As New ADODB.Recordset
    Dim rPesq As New ADODB.Recordset
    Dim r As Reg_0
    Dim sCodMun As String
    Dim sCodVer As String
    Dim sEmailEmp As String
    Dim strSQL As String
    Dim linhas0990 As String
    Dim rItens As New ADODB.Recordset
    Dim r1 As New ADODB.Recordset
    Dim r2 As New ADODB.Recordset
    Dim r3 As New ADODB.Recordset
    Dim r4 As New ADODB.Recordset
    
    Dim linhasC As Integer
    sCodMun = ""
    sCodVer = "007"
    linhas0990 = 0
    AbreArquivoExterno txtCaminho.Text, 1
            
    r.CNPJ = F_SoNumero(strCNPJ_Emp)
    r.cod_fin = Remessa_Arquivo_Original
    r.COD_MUN = sCodMun
    r.cod_ver = sCodVer
    r.CPF = ""
    r.DT_FIN = txtFim.Text
    r.DT_INI = txtIni.Text
    r.IE = strIE_Emp
    r.im = ""
    r.Iind_atv = outros_atv
    r.ind_perfil = "A"
    r.nome = strNome_Emp
    r.suframa = ""
    r.uf = strUF_Emp
    
    s = Registro_0000(r) 'nível 0
    ImprimeEmArquivoExterno s, 1
    linhas0990 = linhas0990 + 1
    
    s = Registro_0001("0")
    ImprimeEmArquivoExterno s, 1
    linhas0990 = linhas0990 + 1
    strSQL = "select * from tab_parametros"
    abreObjetoAdo ado, strSQL
    With ado
        If ado.RecordCount > 0 Then
            s = Registro_0005(!fantasia_emp, !cep_emp, !endereco_emp, Format(!numero_emp), Format(!complemento_emp), !bairro_emp, _
            !fone_emp, !fax_emp, Format(!email_emp))
            linhas0990 = linhas0990 + 1
        End If
    End With
    ImprimeEmArquivoExterno s, 1
    linhas0990 = linhas0990 + 1
    's = Registro_0015 ( ) verificar se é substituto tributário
    
    strSQL = "select * from contabilista"
    
    abreObjetoAdo ado, strSQL
    With ado
        If .RecordCount > 0 Then
            s = Registro_0100(!nome, !CPF, !crc, Format(!CNPJ), !cep, !endereco, !Num, Format(!complemento), !bairro, Format(!fone), Format(!fax), Format(!email))
            ImprimeEmArquivoExterno s, 1
            linhas0990 = linhas0990 + 1
        Else
            MsgBox "Não foram encontrados registros de contabilistas!", vbExclamation
            
        End If
        FechaObjeto ado
    End With
    
    strSQL = "select tf.* from sped_participante as tf, sped_nf as tc where tf.codigo=tc.cod_part "
    strSQL = strSQL & " and dt_e_s>=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and dt_e_s<=" & F_GravaData(txtFim.Text)
    'strSQL = strSQL & " and ind_oper=0"
    
    
    abreObjetoAdo ado, strSQL
    With ado
        Do While Not .EOF
            If .RecordCount > 0 Then
            
                
                s = Registro_0150(!codigo, !nome, "1058", !CNPJ, "", !CPF, !COD_MUN, !suframa, !endereco, !numero, !complemento, !bairro)
                ImprimeEmArquivoExterno s, 1
                linhas0990 = linhas0990 + 1
                
                strSQL = "select * from sped_alteracao_participante where data_alt>=" & F_GravaData(txtIni.Text)
                strSQL = strSQL & " and data_alt<=" & F_GravaData(txtFim.Text)
                strSQL = strSQL & " and codigo=" & !codigo
                strSQL = strSQL & " order by codigo"
                abreObjetoAdo r1, strSQL
                
                With r1
                    Do While Not .EOF
                        s = Registro_0175(!data_alt, !numero_campo, !conteudo_anterior)
                        ImprimeEmArquivoExterno s, 1
                        linhas0990 = linhas0990 + 1
                        .MoveNext
                    Loop
                    FechaObjeto r1
                End With
                
                
            End If
            .MoveNext
        Loop
        FechaObjeto ado
    End With
    
    
    strSQL = "select  * from sped_unidades order by unidade"
    abreObjetoAdo ado, strSQL
    With ado
        Do While Not .EOF
            s = Registro_0190(!unidade, !Descricao)
            ImprimeEmArquivoExterno s, 1
            linhas0990 = linhas0990 + 1
            .MoveNext
        Loop
        FechaObjeto ado
    End With
    
    strSQL = "(select tp.*,imp.perc_imp from tab_produtos as tp, sped_nf as nf, sped_nf_itens as i,"
    strSQL = strSQL & " tab_impostos as imp where nf.id=id_nf "
    strSQL = strSQL & " and tp.cod_imp=imp.cod_imp "
    strSQL = strSQL & " and tp.cod_pro=i.cod_pro "
    strSQL = strSQL & " and nf.dt_doc>=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and nf.dt_doc<=" & F_GravaData(txtFim.Text)
    strSQL = strSQL & " order by tp.cod_pro) union  "
    strSQL = "(select tp.*,imp.perc_imp from tab_produtos as tp, sped_cupom as nf, sped_cupom_itens as i,"
    strSQL = strSQL & " tab_impostos as imp where nf.id=id_cupom "
    strSQL = strSQL & " and tp.cod_imp=imp.cod_imp "
    strSQL = strSQL & " and tp.cod_pro=i.cod_item "
    strSQL = strSQL & " and nf.data>=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and nf.data<=" & F_GravaData(txtFim.Text)
    strSQL = strSQL & " order by tp.cod_pro) "
    abreObjetoAdo ado, strSQL
    With ado
        Do While Not .EOF
            s = Registro_0200(!cod_pro, !nome_pro, !codbarra_pro, "", !unidade_pro, mercadoria_para_revenda, Format(!ncm_pro), _
                "", Format(!cod_gen), "", !perc_imp)
            ImprimeEmArquivoExterno s, 1
            linhas0990 = linhas0990 + 1
            .MoveNext
        Loop
        FechaObjeto ado
    End With
    
    
    strSQL = "select * from sped_alteracao_produto where fim>=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and fim<=" & F_GravaData(txtFim.Text) & " order by codigo"
    abreObjetoAdo ado, strSQL
    
    With ado
        Do While Not .EOF
            s = Registro_0205(!descricao_anterior, !inicio, !fim, !codigo_anterior)
            ImprimeEmArquivoExterno s, 1
            linhas0990 = linhas0990 + 1
            .MoveNext
        Loop
        
    End With
    
    
    
    
    strSQL = "select * from sped_ativo_imobilizado order by codigo"
    abreObjetoAdo ado, strSQL
    '206 combustiveis
    '220 fatores de conversão de unidades
    
    With ado
        Do While Not .EOF
            s = Registro_0300(!codigo, !tipo, !Descricao, !codigo_principal, !COD_CTA, !numero_parcelas)
            ImprimeEmArquivoExterno s, 1
            linhas0990 = linhas0990 + 1
            If !tipo = "1" Then
                s = Registro_0305(!cod_ccus, !funcao, !vida_util)
                ImprimeEmArquivoExterno s, 1
                linhas0990 = linhas0990 + 1
            End If
            
            .MoveNext
        Loop
        FechaObjeto ado
    End With
    
    
    '0400 não tem.. registros de operação que não são CFOP (controlados pela própria empresa)
    
    
    strSQL = "select * from sped_informacao_complementar order by codigo"
    abreObjetoAdo ado, strSQL
    '206 combustiveis
    '220 fatores de conversão de unidades
    
    With ado
        Do While Not .EOF
            s = Registro_0450(!codigo, !Descricao)
            ImprimeEmArquivoExterno s, 1
            
            linhas0990 = linhas0990 + 1
            
            .MoveNext
        Loop
        FechaObjeto ado
    End With
    
    
    strSQL = "select * from sped_pc_contabil order by dt_alt"
    abreObjetoAdo ado, strSQL
    '206 combustiveis
    '220 fatores de conversão de unidades
    
    With ado
        If .RecordCount > 0 Then
            Do While Not .EOF
                s = Registro_0500(!dt_alt, !cod_nat_cc, !ind_cta, !nivel, !COD_CTA, !nome_cta)
                ImprimeEmArquivoExterno s, 1
                linhas0990 = linhas0990 + 1
                .MoveNext
            Loop
            .Close
            
        Else
            'MsgBox "Não foram encontrados registros de fornecedores com compras no período!", vbExclamation
            
        End If
        FechaObjeto ado
    End With
    
    
    strSQL = "select * from sped_ccus order by dt_alt"
    abreObjetoAdo ado, strSQL
    
    
    With ado
        If .RecordCount > 0 Then
            Do While Not .EOF
                s = Registro_0600(!dt_alt, !codigo, !nome)
                ImprimeEmArquivoExterno s, 1
                linhas0990 = linhas0990 + 1
                .MoveNext
            Loop
            .Close
            
        Else
            'MsgBox "Não foram encontrados registros de fornecedores com compras no período!", vbExclamation
            
        End If
        FechaObjeto ado
    End With
    linhas0990 = linhas0990 + 1
    s = Registro_0990(linhas0990)
    ImprimeEmArquivoExterno s, 1
    
    s = Registro_C001(1)
    ImprimeEmArquivoExterno s, 1
    linhasC = 1
    
    
    strSQL = "select * from sped_nf where dt_e_s>=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and dt_e_s<=" & F_GravaData(txtFim.Text)
    'strSQL = strSQL & " and ind_emit=0 "
    strSQL = strSQL & " and status=4"
    abreObjetoAdo ado, strSQL
    
    
    With ado
        If .RecordCount > 0 Then
            Do While Not .EOF
                s = Registro_C100(!IND_OPER, !ind_emit, !COD_PART, !COD_MOD, !COD_SIT, !SER, _
                !NUM_DOC, !chv_nfe, CStr(!DT_DOC), CStr(!DT_E_S), CStr(!VL_DOC), !IND_PGTO, CStr(!VL_DESC), CStr(!vl_abat_nt), _
                CStr(!VL_MERC), !IND_FRT, CStr(!vl_frt), CStr(!vl_seg), CStr(!vl_out_da), CStr(!VL_BC_ICMS), CStr(!VL_ICMS), CStr(!VL_BC_ICMS_ST), CStr(!VL_ICMS), _
                CStr(!VL_IPI), CStr(!VL_PIS), CStr(!VL_COFINS), CStr(!vl_pis_st), CStr(!vl_cofins_st))
                
                
                
                
                
                ImprimeEmArquivoExterno s, 1
                
                strSQL = "select b.* from sped_nf_ic as a, sped_informacao_complementar as b"
                strSQL = strSQL & " where a.codigo=b.codigo and id_nf=" & !id
                abreObjetoAdo rItens, strSQL
                With rItens
                    Do While Not .EOF
                        
                       s = Registro_C110(!codigo, !texto)
                       ImprimeEmArquivoExterno s, 1
                       linhasC = linhasC + 1
                       
                        
                        .MoveNext
                    
                    Loop
                    .Close
                End With
                
                
                
                strSQL = "select * from sped_nf_itens where id_nf=" & !id
                abreObjetoAdo rItens, strSQL
                With rItens
                    Do While Not .EOF
                        
                       s = Registro_C170(!NUM_ITEM, !COD_ITEM, !descr_compl, !QTD, !UNID, CStr(!VL_ITEM), CStr(!VL_DESC), _
                       !Ind_mov, !CST_ICMS, !CFOP, !cod_nat, CStr(!VL_BC_ICMS), CStr(!ALIQ_ICMS), CStr(!VL_ICMS), _
                       CStr(!VL_BC_ICMS_ST), CStr(!ALIQ_ST), CStr(!VL_ICMS_ST), !IND_APUR, !CST_IPI, !cod_enq, _
                       !VL_BC_IPI, !aliq_ipi, !VL_IPI, !cst_pis, !vl_bc_pis, !quant_bc_pis, !aliq_pis, !VL_PIS, _
                       !cst_cofins, !vl_bc_cofins, !aliq_cofins, !quant_bc_cofins, !aliq_cofins_r, !VL_COFINS, !COD_CTA)
                       ImprimeEmArquivoExterno s, 1
                       linhasC = linhasC + 1
                       
                        
                        .MoveNext
                    
                    Loop
                    .Close
                End With
                
                
                linhasC = linhasC + 1
                .MoveNext
            Loop
            .Close
            
        Else
            'MsgBox "Não foram encontrados registros de fornecedores com compras no período!", vbExclamation
            
        End If
        FechaObjeto ado
    End With
    
    
    
    'c171 armazenamento de combustíveis. Não usa
    'REGISTRO C172: OPERAÇÕES COM ISSQN (CÓDIGO 01) não usa
    'REGISTRO C173: OPERAÇÕES COM MEDICAMENTOS (CÓDIGO 01 e 55).
    'REGISTRO C174: OPERAÇÕES COM ARMAS DE FOGO (CÓDIGO 01).
    'REGISTRO C175: OPERAÇÕES COM VEÍCULOS NOVOS (CÓDIGO 01 e 55).
    'REGISTRO C176: RESSARCIMENTO DE ICMS EM OPERAÇÕES COM SUBSTITUIÇÃO TRIBUTÁRIA (CÓDIGO 01, 55). 'ver com o contador
    'REGISTRO C177: OPERAÇÕES COM PRODUTOS SUJEITOS A SELO DE CONTROLE IPI.
    'REGISTRO C178: OPERAÇÕES COM PRODUTOS SUJEITOS À TRIBUTAÇÀO DE IPI POR UNIDADE OU QUANTIDADE DE PRODUTO.
    'REGISTRO C179: INFORMAÇÕES COMPLEMENTARES ST (CÓDIGO 01).
    'REGISTRO C190: REGISTRO ANALÍTICO DO DOCUMENTO (CÓDIGO 01, 1B, 04 ,55 e 65).
    'REGISTRO C195: OBSERVAÇOES DO LANÇAMENTO FISCAL (CÓDIGO 01, 1B E 55)
    'REGISTRO C197: OUTRAS OBRIGAÇÕES TRIBUTÁRIAS, AJUSTES E INFORMAÇÕES DE VALORES PROVENIENTES DE DOCUMENTO FISCAL.
    
    'REGISTRO C300: RESUMO DIÁRIO DAS NOTAS FISCAIS DE VENDA A CONSUMIDOR (CÓDIGO 02)
        
    strSQL = "select * from sped_nf_c300 where dt_doc>=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and dt_doc<=" & F_GravaData(txtFim.Text)
    abreObjetoAdo rItens, strSQL
    With rItens
        Do While Not .EOF
            s = Registro_C300(!COD_MOD, !SER, !Sub, !NUM_DOC_INI, !NUM_DOC_INI, !DT_DOC, CStr(!VL_DOC), CStr(!VL_IPI), CStr(!VL_COFINS), !COD_CTA)
            ImprimeEmArquivoExterno s, 1
            linhasC = linhasC + 1
            .MoveNext
        Loop
        .Close
    End With
    
    
    
    'REGISTRO C310: DOCUMENTOS CANCELADOS DE NOTAS FISCAIS DE VENDA A CONSUMIDOR (CÓDIGO 02).
    'fazer
    
    'REGISTRO C320: REGISTRO ANALÍTICO DO RESUMO DIÁRIO DAS NOTAS FISCAIS DE VENDA A CONSUMIDOR (CÓDIGO 02).
    'fazer
    
    'REGISTRO C321: ITENS DO RESUMO DIÁRIO DOS DOCUMENTOS (CÓDIGO 02).
    'fazer
    'REGISTRO C350: NOTA FISCAL DE VENDA A CONSUMIDOR (CÓDIGO 02)
    'fazer
    'REGISTRO C370: ITENS DO DOCUMENTO (CÓDIGO 02)
    'fazer
    'REGISTRO C390: REGISTRO ANALÍTICO DAS NOTAS FISCAIS DE VENDA A CONSUMIDOR (CÓDIGO 02)
    'fazer
    'REGISTRO C400: EQUIPAMENTO ECF (CÓDIGO 02, 2D e 60).
    
    
    strSQL = "select * from sped_ecf where  "
    strSQL = strSQL & " inicio<=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and (fim>=" & F_GravaData(txtFim.Text) & " or fim is null)"
    abreObjetoAdo rItens, strSQL
    With rItens
        Do While Not .EOF
            s = Registro_C400(!COD_MOD, !ECF_MOD, !ECF_FAB, !ECF_CX)
            ImprimeEmArquivoExterno s, 1
            linhasC = linhasC + 1
            
            'REGISTRO C420: REGISTRO DOS TOTALIZADORES PARCIAIS DA REDUÇÃO Z (COD 02, 2D e 60).
            
            'REGISTRO C425: RESUMO DE ITENS DO MOVIMENTO DIÁRIO (CÓDIGO 02 e 2D).
            
            strSQL = "select * from sped_redz where caixa='" & !ECF_CX & "'"
            strSQL = strSQL & " and dt_doc>=" & F_GravaData(txtIni.Text)
            strSQL = strSQL & " and dt_doc<=" & F_GravaData(txtFim.Text)
            strSQL = strSQL & " order by dt_doc "
            
            
            abreObjetoAdo r2, strSQL
            With r2
                Do While Not .EOF
                    s = Registro_C405(!DT_DOC, !CRO, !CRZ, !NUM_COO_FIN, !GT_FIN, !VL_BRT)
                    ImprimeEmArquivoExterno s, 1
                    linhasC = linhasC + 1
                    
                    s = Registro_C410(!VL_PIS, !VL_COFINS)
                    ImprimeEmArquivoExterno s, 1
                    linhasC = linhasC + 1
                    
                    
                    
                    strSQL = "select r.*,ti.* from sped_redz_tot as r, tab_impostos as ti "
                    strSQL = strSQL & " where r.cod_imp=ti.cod_imp "
                    strSQL = strSQL & " and r.id_red=" & !id
                    strSQL = strSQL & " order by id "
                    
                    abreObjetoAdo r3, strSQL
                    With r3
                        Do While Not .EOF
                            s = Registro_C420(!codigo, !valor, "", "", "")
                            ImprimeEmArquivoExterno s, 1
                            linhasC = linhasC + 1
                            
                            strSQL = "select p.unidade_pro,sum(qtd) as qtd,sum(vl_item) as vl_item,sum(vl_pis) as vl_pis,"
                            strSQL = strSQL & " sum(vl_cofins) as vl_cofins "
                            strSQL = strSQL & " from sped_cupom as c, sped_cupom_itens as i,"
                            strSQL = strSQL & " tab_produtos as p "
                            strSQL = strSQL & " where c.id=i.id_cupom "
                            strSQL = strSQL & " and i.cod_item=p.cod_pro "
                            strSQL = strSQL & " and data=" & F_GravaData(!DT_DOC)
                            strSQL = strSQL & " and tp.cod_imp='" & !cod_imp & "'"
                            strSQL = strSQL & " caixa='" & rItens!caixa
                            strSQL = strSQL & " group by cod_pro order by cod_pro "
                            abreObjetoAdo r4, strSQL
                            With r4
                                Do While Not .EOF
                                    s = Registro_C425(!codigo, !QTD, !unidade_pro, CStr(!VL_ITEM), CStr(!VL_PIS), CStr(!VL_COFINS))
                                    ImprimeEmArquivoExterno s, 1
                                    linhasC = linhasC + 1
                                    .MoveNext
                                Loop
                                .Close
                            End With
                            
                            
                            
                            
                            .MoveNext
                        Loop
                    End With
                    
                    
                    
                    
                
                
                    strSQL = "select c.* from sped_cupom as c where data=" & F_GravaData(r2!DT_DOC)
                    strSQL = strSQL & " and c.caixa='" & r2!caixa & "' "
                    strSQL = strSQL & " order by num_doc "
                    
                    abreObjetoAdo r3, strSQL
                    With r3
                        Do While Not .EOF
                            s = Registro_C460(!COD_MOD, !COD_SIT, !NUM_DOC, !data, CStr(!VL_DOC), CStr(!VL_PIS), CStr(!VL_COFINS), !cpf_adq, !nom_adq)
                            ImprimeEmArquivoExterno s, 1
                            linhasC = linhasC + 1
                            strSQL = "select * from sped_cupom_itens where id=" & !id
                            abreObjetoAdo r4, strSQL
                            
                            With r4
                                Do While Not .EOF
                                    s = Registro_C470(!COD_ITEM, CStr(!QTD), CStr(!QTD_CANC), !UNID, CStr(!VL_ITEM), !CST_ICMS, !CFOP, CStr(!ALIQ_ICMS), _
                                        CStr(!VL_PIS), CStr(!VL_COFINS))
                                    ImprimeEmArquivoExterno s, 1
                                    linhasC = linhasC + 1
                                    .MoveNext
                                Loop
                                .Close
                            End With
                        
                            .MoveNext
                        Loop
                        .Close
                    End With
                    
                    
                    strSQL = "select sum(qtd) as qtd, sum(vl_item) as total_item, sum(i.vl_pis) as total_pis,"
                    strSQL = strSQL & " sum(i.vl_cofins) as vl_cofins from sped_cupom as c, sped_cupom_itens as i"
                    strSQL = strSQL & " where c.id=i.id "
                    strSQL = strSQL & " and data=" & F_GravaData(r2!DT_DOC)
                    strSQL = strSQL & " and c.caixa='" & r2!caixa & "' "
                    strSQL = strSQL & " group by i.cst_icms, i.cfop, i.aliq_icms"
                    strSQL = strSQL & " order by i.cst_icms, i.cfop, i.aliq_icms"
                    
                    
                    abreObjetoAdo r3, strSQL
                    With r3
                        Do While Not .EOF
                            s = Registro_C490(!CST_ICMS, !CFOP, CStr(!ALIQ_ICMS), !VL_ITEM, !VL_ITEM, !VL_ICMS, !UNID, !COD_OBS)
                            ImprimeEmArquivoExterno s, 1
                            linhasC = linhasC + 1
                            
                            
                            .MoveNext
                        Loop
                    End With
                
                
                .MoveNext
                Loop
                    
                
            'r2'c405
            End With
            
            
            
            
            
            
            .MoveNext
        Loop
        strSQL = "select sum(qtd) as qtd_total, sum(qtd_canc) as qtd_canc, sum(qtd_canc*vl_item) as total_canc,"
        strSQL = strSQL & " sum(i.vl_desconto) as total_desconto,sum(qtd_canc*vl_item) as total_item, "
        strSQL = strSQL & " sum(i.vl_pis) as total_pis,sum(i.vl_icms) as total_icms,"
        strSQL = strSQL & " sum(i.vl_cofins) as total_cofins from sped_cupom as c, sped_cupom_itens as i"
        strSQL = strSQL & " where c.id=i.id "
        strSQL = strSQL & " and data>=" & F_GravaData(txtIni.Text)
        strSQL = strSQL & " and data<=" & F_GravaData(txtFim.Text)
        
        strSQL = strSQL & " group by i.cod_item, i.aliq_icms"
        strSQL = strSQL & " order by i.cod_item, i.aliq_icms"
        
        
        
        abreObjetoAdo r3, strSQL
        With r3
            Do While Not .EOF
                s = Registro_C495(!ALIQ_ICMS, !COD_ITEM, CStr(!QTD_total), CStr(!QTD_CANC), !UNID, CStr(!total_item), _
                     CStr(!total_desconto), CStr(!total_canc), "0", CStr(!total_item), _
                     CStr(!total_icms), IIf(!st_pro = "I", CStr(!total_item), "0"), IIf(!st_pro = "R", CStr(!total_item), "0"), IIf(!st_pro = "I", CStr(!total_item), "0"))
                
                ImprimeEmArquivoExterno s, 1
                linhasC = linhasC + 1
                
                
                .MoveNext
            Loop
        End With

        .Close
    End With
    
    
    strSQL = "select * from sped_nf_energia where dt_doc>=" & F_GravaData(txtIni.Text) & " and dt_doc<=" & F_GravaData(txtFim.Text)
    strSQL = strSQL & " order by dt_doc "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            s = Registro_C500(!ind_opr, !ind_emit, !COD_PART, !COD_MOD, !COD_SIT, !SER, !SUBSER, !COD_CONS, _
                !NUM_DOC, !DT_DOC, !DT_E_S, CStr(!VL_DOC), CStr(!VL_DESC), CStr(!VL_FORN), CStr(!vl_serv_net), _
                CStr(!vl_ter), CStr(!VL_DA), CStr(!VL_BC_ICMS), CStr(!VL_ICMS), CStr(!VL_BC_ICMS_ST), _
                CStr(!VL_ICMS_ST), !COD_INF, CStr(!VL_PIS), CStr(!VL_COFINS), !TP_LIGACAO, !cod_gru)
            ImprimeEmArquivoExterno s, 1
            linhasC = linhasC + 1
            
            .MoveNext
        Loop
    End With
    
    'Energia Eletrica = CFOP 1253 - 060 Sub/Trib

    'Telefone = CFOP 1303 - 090 Outras

    'Agua = CFOP 1556 - 040 Isento
    
    
    
    
    
               
    linhasC = linhasC + 1
    s = Registro_C990(CStr(linhasC))
    ImprimeEmArquivoExterno s, 1
       
    
        
    'bloco do
    
    s = Registro_D001("1") 'sem dados informado no registro D
    ImprimeEmArquivoExterno s, 1
    s = Registro_D990("2")
    ImprimeEmArquivoExterno s, 1
    
     
    
    
    
    
    strSQL = "select * from sped_nf where dt_e_s>=" & F_GravaData(txtIni.Text)
    strSQL = strSQL & " and dt_e_s<=" & F_GravaData(txtFim.Text)
    strSQL = strSQL & " and ind_emit=1 "
    strSQL = strSQL & " and status=4"
    abreObjetoAdo ado, strSQL
    
    
    With ado
        If .RecordCount > 0 Then
            Do While Not .EOF
                s = Registro_C100(saida, emissao_propria, !COD_PART, !COD_MOD, !COD_SIT, !SER, _
                !NUM_DOC, !chv_nfe, !DT_DOC, !DT_E_S, !VL_DOC, !IND_PGTO, !VL_DESC, !vl_abat_nt, _
                !VL_MERC, !IND_FRT, !vl_frt, !vl_seg, !vl_out_da, !VL_BC_ICMS, !VL_ICMS, !VL_BC_ICMS_ST, !VL_ICMS, _
                !VL_IPI, !VL_PIS, !VL_COFINS, !vl_pis_st, !vl_cofins_st)
                ImprimeEmArquivoExterno s, 1
                linhas0990 = linhas0990 + 1
                .MoveNext
            Loop
            .Close
            
        Else
            'MsgBox "Não foram encontrados registros de fornecedores com compras no período!", vbExclamation
            
        End If
        FechaObjeto ado
    End With
    
    
    
    
    
    
    FechaArquivoExterno 1
    Shell "explorer " & txtCaminho.Text
     
    
End Sub
Private Sub Form_Load()
    txtAno.Text = Year(Date)
    txtMes.Text = Month(Date)
    txtCaminho.Text = App.Path & "\sped.txt"
    
End Sub

Private Sub txtAno_GotFocus()
    P_SelecTexto txtAno
End Sub

Private Sub txtAno_KeyDown(KeyCode As Integer, shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtMes_LostFocus()
    txtMes.Text = Format(Month("01/" & txtMes.Text & "/" & Year(Date)), "0#")
    buscadata
End Sub
Sub buscadata()
    On Error GoTo trata_erro
    txtIni.Text = F_PrimDiaMes("01/" & txtMes.Text & "/" & txtAno.Text)
    txtFim.Text = F_UltDiaMes("01/" & txtMes.Text & "/" & txtAno.Text)
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Private Sub txtMes_GotFocus()
    P_SelecTexto txtMes
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtAno_LostFocus()
    txtAno.Text = Year(CDate("01/01/" & txtAno.Text))
End Sub
