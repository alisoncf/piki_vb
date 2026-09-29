Attribute VB_Name = "Sped"
Public CrLf As String
Public Enum cod_fin
     Remessa_Arquivo_Original = 0
     Remessa_Arquivo_Substituto = 1
     
End Enum
Public Enum TP_ASSINANTE
    ComercialIndustrial = 1
    poderPublic = 2
    ResidencialPEssoaFisica = 3
    Publico = 4
    SemiPublico = 5
    outros = 6
    
End Enum
Public Enum IND_APUR
    menal = 0
    decendial = 1
End Enum
Public Enum ind_tit
    duplicada = 0
    cheque = 1
    promissoria = 2
    recibo = 3
    outros = 99
End Enum
Public Enum ind_atv
     industrial_ou_equiparado = 0
     outros = 1
     
End Enum
Public Enum IND_OPER
    Entrada = 0
    saida = 1
End Enum
Public Enum IND_EMIT
    emissao_propria = 0
    terceiros = 1
End Enum
Public Enum IND_PGTO
    avista = 0
    aprazo = 1
    outros = 2
End Enum
Public Enum IND_CARGA
   RODOVIARIO = 0
   FERROVIARIO = 1
   RODO_FERRO = 2
   AQUAVIARIO = 3
   DUTO_VIARIO = 4
   AEREO = 5
   outros = 9
   
End Enum
Public Enum IND_FRT
    PorContaDoEmitente = 0
    
    PorContaDoDestinatarioRemetente = 1
    PorContaDeTerceiros = 2
    SemCobrancaDefrete = 3
End Enum


Public Enum tipo_item
     mercadoria_para_revenda = 0
     materia_prima = 1
     embalagem = 2
     produto_em_processo = 3
     produto_acabado = 4
     subproduto = 5
     produto_intermediario = 6
     material_de_uso_e_consumo = 7
     ativo_imobilizado = 8
     Servicos = 9
     outros_insumos = 10
     outras = 99
     
    
End Enum

Public Type Reg_0
    cod_ver As String
    cod_fin As cod_fin
    DT_INI As String
    DT_FIN As String
    nome As String
    CNPJ As String
    CPF As String
    UF As String
    IE As String
    cod_mun As String
    im As String
    suframa As String
    ind_perfil As String 'a b ou c
    ind_atv As ind_atv
    
    
    
    
End Type
Function GravaDataSped(Data As String) As String
    
    GravaDataSped = Format(Data, "ddmmyyyy")
End Function
Function GravaValorSped(valor As String, Optional casasdecimais As Integer = 2)
    zeros = String(casasdecimais, "0")

    GravaValorSped = Format(valor, "##0." & zeros)
End Function
Function GravaNumeroSped(valor As String, tamanho As Integer) As String

    GravaNumeroSped = EsticaTexto(valor, tamanho, 2, "0")
End Function

Function Registro_0000(r As Reg_0) As String
    Dim s As String
    s = "|0000|"
    s = s & EsticaTexto(r.cod_ver, 3) & "|"
    s = s & EsticaTexto(CStr(r.cod_fin), 3) & "|"
    s = s & EsticaTexto(r.cod_ver, 1) & "|"
    s = s & GravaDataSintegra(r.DT_INI) & "|"
    s = s & GravaDataSintegra(r.DT_FIN) & "|"
    s = s & EsticaTexto(r.nome, 100) & "|"
    s = s & EsticaTexto(r.CNPJ, 14) & "|"
    s = s & EsticaTexto(r.CPF, 11) & "|"
    s = s & EsticaTexto(r.UF, 2) & "|"
    s = s & EsticaTexto(r.IE, 4) & "|"
    s = s & EsticaTexto(r.cod_mun, 7) & "|"
    s = s & EsticaTexto(r.im, 10) & "|"
    s = s & EsticaTexto(r.suframa, 9) & "|"
    s = s & EsticaTexto(r.ind_perfil, 1) & "|" 'A b ou c
    s = s & EsticaTexto(CStr(r.ind_atv), 1) & "|" & CrLf
    Registro_0000 = s
End Function
Function Registro_0001(Optional Ind_mov As String = "0") As String
    Dim s As String
    s = "|0001|"
    s = s & EsticaTexto(Ind_mov, 1) & "|"
    s = s & CrLf
    Registro_0001 = s
End Function


Function Registro_0005(fantasia As String, cep As String, ende As String, Num As String, _
    compl As String, bairro As String, fone As String, fax As String, email As String) As String
    
    Dim s As String
    s = "|0005|"
    s = s & EsticaTexto(fantasia, 60) & "|"
    s = s & EsticaTexto(cep, 8) & "|"
    s = s & EsticaTexto(ende, 60) & "|"
    s = s & EsticaTexto(Num, 10) & "|"
    s = s & EsticaTexto(compl, 60) & "|"
    s = s & EsticaTexto(bairro, 60) & "|"
    s = s & EsticaTexto(fone, 11) & "|"
    s = s & EsticaTexto(fax, 11) & "|"
    s = s & EsticaTexto(email, 60) & "|" & CrLf
    
    Registro_0005 = s
    
End Function

Function Registro_0100(nome As String, CPF As String, crc As String, CNPJ As String, cep As String, ende As String, Num As String, _
    compl As String, bairro As String, fone As String, fax As String, email As String) As String
    
    Dim s As String
    s = "|0100|"
    s = s & EsticaTexto(nome, 60) & "|"
    s = s & EsticaTexto(CPF, 11) & "|"
    s = s & EsticaTexto(crc, 15) & "|"
    s = s & EsticaTexto(CNPJ, 14) & "|"
    s = s & EsticaTexto(cep, 8) & "|"
    s = s & EsticaTexto(ende, 60) & "|"
    s = s & EsticaTexto(Num, 10) & "|"
    s = s & EsticaTexto(compl, 60) & "|"
    s = s & EsticaTexto(bairro, 60) & "|"
    s = s & EsticaTexto(fone, 11) & "|"
    s = s & EsticaTexto(fax, 11) & "|"
    s = s & EsticaTexto(email, 60) & "|" & CrLf
    
    Registro_0100 = s
    
End Function
Function Registro_0015(uf_st As String, ie_st As String) As String
    Dim s As String
    s = "|0015|" & EsticaTexto(uf_st, 2) & "|" & EsticaTexto(ie_st, 10) & "|" & CrLf
    
    Registro_0015 = s
End Function

Function Registro_0150(COD_PART As String, nome As String, cod_pais As String, CNPJ As String, _
    CPF As String, IE As String, cod_mun As String, suframa As String, ende As String, Num As String, _
    compl As String, bairro As String) As String
Dim s As String
    s = "|0150|"
    s = s & EsticaTexto(COD_PART, 60) & "|"
    s = s & EsticaTexto(nome, 100) & "|"
    s = s & EsticaTexto(cod_pais, 5) & "|"
    s = s & EsticaTexto(CNPJ, 14) & "|"
    s = s & EsticaTexto(IE, 11) & "|"
    s = s & EsticaTexto(suframa, 9) & "|"
    s = s & EsticaTexto(ende, 60) & "|"
    s = s & EsticaTexto(Num, 10) & "|"
    s = s & EsticaTexto(compl, 60) & "|"
    s = s & EsticaTexto(bairro, 60) & "|" & CrLf
    
    
    
    
    Registro_0150 = s
End Function
Function Registro_0175(dt_alt As String, nr_campo As String, cont_ant As String) As String
Dim s As String
    s = "|0175|"
    s = s & GravaDataSped(dt_alt) & "|"
    s = s & EsticaTexto(nr_campo, 2) & "|"
    s = s & EsticaTexto(cont_ant, 100) & "|"
    s = s & CrLf
    
    
    
    
    Registro_0175 = s
End Function

Function Registro_0190(UNID As String, descr As String) As String
Dim s As String
    s = "|0175|"
    
    s = s & EsticaTexto(UNID, 4) & "|"
    s = s & EsticaTexto(descr, 100) & "|"
    s = s & CrLf
    
    
    
    
    Registro_0190 = s
End Function

Function Registro_0200(COD_ITEM As String, descr_item As String, cod_barra As String, _
    cod_ant_item As String, unid_inv As String, tipo_it As tipo_item, _
    cod_ncm As String, ex_ipi As String, cod_gen As String, cod_lst As String, ALIQ_ICMS As String) As String
Dim s As String
    s = "|0200|"
    
    s = s & EsticaTexto(COD_ITEM, 60) & "|"
    s = s & EsticaTexto(descr_item, 100) & "|"
    s = s & cod_barra & "|"
    s = s & EsticaTexto(cod_ant_item, 60) & "|"
    s = s & EsticaTexto(unid_inv, 6) & "|"
    s = s & EsticaTexto(tipo_it, 2) & "|"
    s = s & EsticaTexto(cod_ncm, 8) & "|"
    s = s & EsticaTexto(ex_ipi, 3) & "|"
    s = s & EsticaTexto(cod_gen, 2) & "|"
    s = s & EsticaTexto(cod_lst, 4) & "|"
    s = s & EsticaTexto(GravaValorSped(ALIQ_ICMS), 6) & "|"
    s = s & CrLf
    
    
    
    
    Registro_0200 = s
End Function


Function Registro_0205(descr_anterior_item As String, DT_INI As String, DT_FIN As String, cod_ant_item As String) As String
Dim s As String
    s = "|0205|"
    s = s & descr_anterior_item & "|"
    s = s & GravaDataSped(DT_INI) & "|"
    s = s & GravaDataSped(DT_FIN) & "|"
    s = s & EsticaTexto(cod_ant_item, 60) & "|"
    s = s & CrLf
    
    
    
    
    Registro_0205 = s
End Function

Function Registro_0300(cod_ind_bem As String, ident_merc As String, descr_item As String, cod_prnc As String, _
    COD_CTA As String, nr_parc As String) As String
Dim s As String
    s = "|0300|"
    s = s & EsticaTexto(cod_ind_bem, 60) & "|"
    s = s & EsticaTexto(ident_merc, 1) & "|"  '1 mercadoria 2 componente
    s = s & EsticaTexto(cod_prnc, 1) & "|"   'caso seja componente, o código do produto principal
    s = s & EsticaTexto(COD_CTA, 60) & "|"   'conta analítica de contabilização
    s = s & EsticaTexto(nr_parc, 3) & "|" 'parcelas a serem apropriadas
    s = s & CrLf
    
    
    
    
    Registro_0300 = s
End Function

Function Registro_0990(qtd_linhas As String) As String
    Dim s As String
    s = "|0300|" & qtd_linhas & "|"
    Registro_0990 = s
End Function


Function Registro_C001(Ind_mov As String) As String
    Dim s As String
    s = "|C001|" & Ind_mov & "|" '0 bloco sem dados 1 bloco com dados informados
    Registro_C001 = s
    
End Function

Function Registro_C100(Ind_Op As IND_OPER, ind_em As IND_EMIT, COD_PART As String, cod_modelo_nota As String, _
    cod_sit_nota As String, SERIE As String, NUM_DOC As String, chv_nfe As String, DT_DOC As String, _
        DT_ENTRADA_SAIDA As String, valor_doc As String, ind_pg As IND_PGTO, valor_desconto As String, _
        vlr_abat_not As String, valor_total_merc As String, Ind_frete As IND_FRT, _
        valor_frete As String, valor_seguro As String, valor_out_da As String, vlr_bcicms As String, _
        vlr_icms As String, vlr_bc_icms_st As String, vlr_icms_st As String, _
        vlr_ipi As String, vlr_pis As String, vlr_cofins As String, vlr_pis_st As String, _
        vlr_cofins_st As String) As String
    Dim s As String
    s = "|C100|"
    s = s & Ind_Op & "|"
    s = s & ind_em & "|"
    s = s & EsticaTexto(COD_PART, 60) & "|"
    s = s & EsticaTexto(cod_modelo_nota, 2) & "|"
    s = s & EsticaTexto(cod_sit_nota, 2) & "|"
    s = s & EsticaTexto(SERIE, 3) & "|"
    s = s & EsticaTexto(NUM_DOC, 9) & "|"
    s = s & EsticaTexto(chv_nfe, 44) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & GravaDataSped(DT_ENTRADA_SAIDA) & "|"
    s = s & GravaValorSped(valor_doc) & "|"
    s = s & ind_pg & "|"
    s = s & GravaValorSped(valor_desconto) & "|"
    s = s & GravaValorSped(vlr_abat_not) & "|"
    s = s & GravaValorSped(valor_total_merc) & "|"
    s = s & Ind_frete & "|"
    s = s & GravaValorSped(valor_frete) & "|"
    s = s & GravaValorSped(valor_seguro) & "|"
    s = s & GravaValorSped(valor_out_da) & "|"
    s = s & GravaValorSped(vlr_bcicms) & "|"
    s = s & GravaValorSped(vlr_icms) & "|"
    s = s & GravaValorSped(vlr_bc_icms_st) & "|"
    s = s & GravaValorSped(vlr_icms_st) & "|"
    s = s & GravaValorSped(vlr_ipi) & "|"
    s = s & GravaValorSped(vlr_pis) & "|"
    s = s & GravaValorSped(vlr_cofins) & "|"
    s = s & GravaValorSped(vlr_pis_st) & "|"
    s = s & GravaValorSped(vlr_cofins_st) & "|"
    
    Registro_C100 = s
    
End Function
Function Registro_C105(oper As String, UF As String) As String
    'oper 0 combustivel e lubrificantes 1 veiculos leansing
    Dim s As String
    s = "|C105|"
    s = s & oper & "|" & UF & "|" & CrLf
    
    
    Registro_C105 = s
        
    
End Function



Function Registro_C110(COD_INF As String, Texto As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    s = "|C110|"
    s = s & EsticaTexto(COD_INF, 6) & Texto & "|" & CrLf
    
    
    Registro_C110 = s
        
    
End Function

Function Registro_C111(NUM_PROC As String, IND_PROC As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    'ind_proc - 0 sefaz, 1 justiça federal, 2 justiça estadual 3 SECEX/SRF 9 outros
    
    
    s = "|C110|"
    s = s & EsticaTexto(NUM_PROC, 15) & IND_PROC & "|" & CrLf
    
    
    Registro_C111 = s
        
    
End Function
Function Registro_C112(cod_DA As String, UF As String, NUM_DA As String, cod_aut As String, DT_VCTO As String, dt_pgto As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    'cod_da= 0 doc arrecadação estadual  1 GNRE
    s = "|C110|"
    s = s & EsticaTexto(cod_DA, 1) & "|"
    s = s & EsticaTexto(UF, 2) & "|"
    s = s & NUM_DA & "|"
    s = s & GravaDataSped(DT_VCTO) & "|"
    s = s & GravaDataSped(dt_pgto) & "|"
    s = s & CrLf
    
    
    Registro_C112 = s
        
    
End Function

Function Registro_C113(IND_OPER As String, IND_EMIT, COD_PART As String, COD_MOD As String, SER As String, sub_s As String, NUM_DOC As String, DT_DOC As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    'ind_oper= 0 entrada 1 saida
    'ind_emit 0 = emissao propria 1 terceiros
    s = "|C113|"
    s = s & EsticaTexto(IND_OPER, 1) & "|"
    s = s & EsticaTexto(CStr(IND_EMIT), 1) & "|"
    s = s & EsticaTexto(COD_PART, 60) & "|"
    s = s & EsticaTexto(COD_MOD, 2) & "|"
    s = s & EsticaTexto(SER, 4) & "|"
    s = s & EsticaTexto(sub_s, 3) & "|"
    s = s & EsticaTexto(NUM_DOC, 9) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    
    s = s & CrLf
    
    
    Registro_C113 = s
        
    
End Function

Function Registro_C114(COD_MOD As String, ECF_FAB As String, ECF_CX As String, NUM_DOC As String, DT_DOC As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    'ind_oper= 0 entrada 1 saida
    'ind_emit 0 = emissao propria 1 terceiros
    s = "|C114|"
    s = s & EsticaTexto(COD_MOD, 2) & "|"
    s = s & EsticaTexto(ECF_FAB, 21) & "|"
    s = s & EsticaTexto(ECF_CX, 3, 2, "0") & "|"
    s = s & EsticaTexto(NUM_DOC, 6, 2, "0") & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    
    s = s & CrLf
    
    
    Registro_C114 = s
        
    
End Function
Function Registro_C115(IND_CARGA As String, cnpj_col As String, ie_col As String, cpf_col As String, cod_mun_col As String, _
    cnpj_entg As String, ie_entg As String, cpf_entg As String, cod_mun_entg As String) As String
    'ind_carga 0:rodo 1:ferroviário 2: rodo-ferro 3:aquav 4: dutovi 5: aéreo 9:outros
    Dim s As String
    'ind_oper= 0 entrada 1 saida
    'ind_emit 0 = emissao propria 1 terceiros
    s = "|C115|"
    s = s & EsticaTexto(IND_CARGA, 1) & "|"
    s = s & EsticaTexto(cnpj_col, 14) & "|"
    s = s & EsticaTexto(ie_col, 14) & "|"
    s = s & EsticaTexto(cpf_col, 11) & "|"
    s = s & EsticaTexto(cod_mun_col, 7) & "|"
    s = s & EsticaTexto(cnpj_entg, 14) & "|"
    s = s & EsticaTexto(ie_entg, 14) & "|"
    s = s & EsticaTexto(cpf_entg, 11) & "|"
    s = s & EsticaTexto(cod_mun_entg, 7) & "|"
    s = s & CrLf
    
    
    Registro_C115 = s
        
    
End Function


Function Registro_C120(cod_doc_imp As String, num_doc_imp As String, pis_imp As String, cofins_imp As String, num_acdraw As String) As String
    
    Dim s As String
    
    'cod_doc_imp 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C120|"
    s = s & EsticaTexto(cod_doc_imp, 1) & "|"
    s = s & EsticaTexto(num_doc_imp, 12) & "|"
    s = s & GravaValorSped(pis_imp) & "|"
    s = s & GravaValorSped(cofins_imp) & "|"
    s = s & EsticaTexto(num_acdraw, 20) & "|"
    s = s & CrLf
    
    
    Registro_C120 = s
        
    
End Function

Function Registro_C140(IND_EMIT As IND_EMIT, ind_tit As ind_tit, desc_tit As String, num_tit As String, qtd_parc As String, vl_tit As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C140|"
    s = s & EsticaTexto(CStr(IND_EMIT), 1) & "|"
    s = s & EsticaTexto(CStr(ind_tit), 2) & "|"
    s = s & desc_tit & "|"
    s = s & num_tit & "|"
    s = s & EsticaTexto(qtd_parc, 2) & "|"
    s = s & GravaValorSped(vl_tit) & "|"
    
    s = s & CrLf
    
    
    Registro_C140 = s
        
    
End Function

Function Registro_C160(COD_PART As String, veic_id As String, qtd_vol As String, peso_brt As String, peso_liq As String, uf_id As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C160|"
    s = s & EsticaTexto(COD_PART, 60) & "|"
    s = s & EsticaTexto(veic_id, 7) & "|"
    s = s & GravaValorSped(qtd_vol) & "|"
    s = s & GravaValorSped(peso_brt) & "|"
    s = s & GravaValorSped(peso_liq) & "|"
    s = s & EsticaTexto(uf_id, 2) & "|"
    s = s & CrLf
    
    
    Registro_C160 = s
        
    
End Function
Function Registro_C170(num_item As String, COD_ITEM As String, desc_compl As String, QTD As String, UNID As String, _
    vl_ITEM As String, vl_DESC As String, Ind_mov As String, CST_ICMS As String, CFOP As String, cod_nat As String, _
    VL_BC_ICMS As String, ALIQ_ICMS As String, VL_ICMS As String, VL_BC_ICMS_ST As String, _
    aliq_st As String, VL_ICMS_ST As String, IND_APUR As String, CST_IPI As String, cod_enq As String, _
    VL_BC_IPI As String, aliq_ipi As String, VL_IPI As String, cst_pis As String, vl_bc_pis As String, quant_bc_pis As String, _
    aliq_pis As String, VL_PIS As String, cst_cofins As String, vl_bc_cofins As String, aliq_cofins As String, _
    quant_bc_cofins As String, aliq_confins As String, VL_COFINS As String, COD_CTA As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C170|"
    s = s & EsticaTexto(num_item, 3, 2, "0") & "|"
    s = s & EsticaTexto(COD_ITEM, 60) & "|"
    s = s & desc_compl & "|"
    s = s & GravaValorSped(QTD) & "|"
    s = s & EsticaTexto(UNID, 6) & "|"
    s = s & GravaValorSped(vl_ITEM) & "|"
    s = s & GravaValorSped(vl_DESC) & "|"
    s = s & EsticaTexto(Ind_mov, 1) & "|" ' movimentação fisica item produto :  0: sim 1: não
    s = s & EsticaTexto(CST_ICMS, 3) & "|"
    s = s & EsticaTexto(CFOP, 4) & "|"
    s = s & EsticaTexto(cod_nat, 10) & "|"
    
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & GravaValorSped(aliq_st) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    
    s = s & EsticaTexto(IND_APUR, 1) & "|"  '0 mensal 1 decenial
    s = s & EsticaTexto(CST_IPI, 2) & "|"
    s = s & EsticaTexto(cod_enq, 3) & "|"
    
    s = s & GravaValorSped(VL_BC_IPI) & "|"
    s = s & GravaValorSped(aliq_ipi) & "|"
    s = s & GravaValorSped(VL_IPI) & "|"
    
    s = s & EsticaTexto(cst_pis, 2) & "|"
    s = s & GravaValorSped(vl_bc_pis) & "|"
    s = s & GravaValorSped(aliq_pis) & "|"
    s = s & GravaValorSped(quant_bc_pis) & "|"
    s = s & GravaValorSped(aliq_pis) & "|"
    s = s & GravaValorSped(VL_PIS) & "|"
    
    s = s & EsticaTexto(cst_cofins, 2) & "|"
    s = s & GravaValorSped(vl_bc_cofins) & "|"
    s = s & GravaValorSped(aliq_cofins) & "|"
    s = s & GravaValorSped(quant_bc_cofins) & "|"
    s = s & GravaValorSped(aliq_cofins) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    
    
    s = s & EsticaTexto(COD_CTA, 2) & "|"
    s = s & CrLf
    
    
    Registro_C170 = s
        
    
End Function


Function Registro_C176(COD_MOD_ULT_E As String, NUM_DOC_ULT_E As String, SER_ULT_E As String, _
    DT_ULT_E As String, COD_PART_ULT_E As String, QUANT_ULT_E As String, VL_UNIT_ULT_E As String, _
    VL_UNIT_BC_ST As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C176|"
    s = s & EsticaTexto(COD_MOD_ULT_E, 2) & "|"
    s = s & EsticaTexto(NUM_DOC_ULT_E, 9) & "|"
    s = s & EsticaTexto(SER_ULT_E, 3) & "|"
    s = s & GravaDataSped(DT_ULT_E) & "|"
    s = s & EsticaTexto(COD_PART_ULT_E, 60) & "|"
    s = s & GravaValorSped(QUANT_ULT_E) & "|"
    s = s & GravaValorSped(VL_UNIT_ULT_E) & "|"
    s = s & GravaValorSped(VL_UNIT_BC_ST) & "|"
    
    s = s & CrLf
    
    
    Registro_C176 = s
        
    
End Function



Function Registro_C190(CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, _
    VL_OPR As String, VL_BC_ICMS As String, VL_ICMS As String, VL_BC_ICMS_ST As String, _
    VL_ICMS_ST As String, VL_RED_BC As String, VL_IPI As String, COD_OBS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C190|"
    s = s & EsticaTexto(CST_ICMS, 3) & "|"
    s = s & EsticaTexto(CFOP, 4) & "|"
    s = s & gravaAliquotaSPED(ALIQ_ICMS, 2) & "|"
    s = s & GravaValorSped(VL_OPR) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_RED_BC) & "|"
    s = s & GravaValorSped(VL_IPI) & "|"
    
    s = s & EsticaTexto(COD_OBS, 6) & "|"
    
    s = s & CrLf
    
    
    Registro_C190 = s
        
    
End Function

Function Registro_C400(COD_MOD As String, ECF_MOD As String, ECF_FAB As String, _
    ECF_CX As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C400|"
    s = s & EsticaTexto(COD_MOD, 2) & "|"
    s = s & EsticaTexto(ECF_MOD, 20) & "|"
    s = s & EsticaTexto(ECF_FAB, 21) & "|"
    s = s & EsticaTexto(ECF_CX, 3) & "|"
    
    
    s = s & CrLf
    
    
    Registro_C400 = s
        
    
End Function

Function Registro_C405(DT_DOC As String, CRO As String, CRZ As String, _
    NUM_COO_FIN As String, GT_FIN As String, VL_BRT As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C400|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & EsticaTexto(CRO, 3, 2, "0") & "|"
    s = s & EsticaTexto(CRZ, 3, 2, "0") & "|"
    s = s & EsticaTexto(NUM_COO_FIN, 6, 2, "0") & "|"
    s = s & GravaValorSped(GT_FIN) & "|"
    s = s & GravaValorSped(VL_BRT) & "|"
    
    
    s = s & CrLf
    
    
    Registro_C405 = s
        
    
End Function

Function Registro_C420(COD_TOT_PAR As String, VLR_ACUM_TOT As String, NR_TOT As String, _
    CRZ As String, DESCR_NR_TOT As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C420|"
    s = s & EsticaTexto(COD_TOT_PAR, 7) & "|"
    s = s & GravaValorSped(VLR_ACUM_TOT) & "|"
    s = s & EsticaTexto(NR_TOT, 2, 2, "0") & "|"
    s = s & EsticaTexto(CRZ, 3, 2, "0") & "|"
    s = s & DESCR_NR_TOT & "|"
    
    
    
    s = s & CrLf
    
    
    Registro_C420 = s
        
    
End Function
Function Registro_C425(COD_ITEM As String, QTD As String, UNID As String, _
    vl_ITEM As String, VL_PIS As String, VL_COFINS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C425|"
    s = s & EsticaTexto(COD_ITEM, 60) & "|"
    's = s & EsticaTexto(NR_TOT, 2, 2, "0") & "|"
    s = s & GravaNumeroSped(QTD, 3) & "|"
    s = s & EsticaTexto(UNID, 6) & "|"
    s = s & GravaValorSped(vl_ITEM) & "|"
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    
    
    s = s & CrLf
    
    
    Registro_C425 = s
        
    
End Function

Function Registro_C460(COD_MOD As String, COD_SIT As String, NUM_DOC As String, _
    DT_DOC As String, vl_DOC As String, VL_PIS As String, VL_COFINS As String, CPFCNPJ As String, _
    NOME_ADQ As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C425|"
    s = s & EsticaTexto(COD_MOD, 2) & "|"
    s = s & EsticaTexto(COD_SIT, 2, 2, "0") & "|"
    s = s & EsticaTexto(NUM_DOC, 2, 2, "0") & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & GravaValorSped(vl_DOC) & "|"
    
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    s = s & EsticaTexto(CPFCNPJ, 14) & "|"
    s = s & EsticaTexto(NOME_ADQ, 60) & "|"
    s = s & CrLf
    
    
    Registro_C460 = s
        
    
End Function
Function Registro_C470(COD_ITEM As String, QTD As String, QTD_CANC As String, _
    UNID As String, vl_ITEM As String, CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, _
    VL_PIS As String, VL_COFINS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C470|"
    s = s & EsticaTexto(COD_ITEM, 60) & "|"
    s = s & GravaNumeroSped(QTD, 3) & "|"
    s = s & GravaNumeroSped(QTD_CANC, 3) & "|"
    s = s & EsticaTexto(UNID, 6) & "|"
    s = s & GravaValorSped(vl_ITEM) & "|"
    s = s & GravaValorSped(CST_ICMS) & "|"
    s = s & EsticaTexto(CFOP, 4) & "|"
    s = s & GravaValorSped(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    
    
    s = s & CrLf
    
    
    Registro_C470 = s
        
    
End Function

Function Registro_C490(CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, VL_OPR As String, _
    VL_BC_ICMS As String, VL_ICMS As String, _
    UNID As String, COD_OBS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C490|"
    s = s & GravaNumeroSped(CST_ICMS, 3) & "|"
    s = s & GravaNumeroSped(CFOP, 4) & "|"
    's = s & GravaNumeroSped(QTD_CANC, 3) & "|"
    s = s & GravaValorSped(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_OPR) & "|"
    
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & EsticaTexto(COD_OBS, 6) & "|"
    
    
    s = s & CrLf
    
    
    Registro_C490 = s
        
    
End Function

Function Registro_C990(qtd_linhas As String) As String
    Dim s As String
    s = "|C990|" & qtd_linhas & "|"
    Registro_C990 = s
End Function



Function Registro_D001(Ind_mov As String) As String
    Dim s As String
    s = "|D001|" & Ind_mov & "|" '0 bloco sem dados 1 bloco com dados informados
    Registro_D001 = s
End Function


Function Registro_D100(IND_OPERA As IND_OPER, IND_EMIT As IND_EMIT, COD_PART As String, COD_MOD As String, _
    COD_SIT As String, SER As String, SUBSER As String, NUM_DOC As String, CHV_CTE As String, DT_DOC As String, _
    DT_AQUISICAO_PRESTACAO As String, TP_CTe As String, CHV_CTE_REF As String, vl_DOC As String, _
    vl_DESC As String, IND_FRE As IND_FRT, vl_SERV As String, VL_BC_ICMS As String, _
    VL_ICMS As String, VL_NT As String, COD_INF As String, COD_CTA As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|D100|"
    s = s & EsticaTexto(CStr(IND_OPERA), 1) & "|"
    s = s & EsticaTexto(CStr(IND_EMIT), 1) & "|"
    s = s & EsticaTexto(COD_PART, 60) & "|"
    s = s & EsticaTexto(COD_MOD, 2) & "|"
    s = s & GravaNumeroSped(COD_SIT, 2) & "|"
    s = s & EsticaTexto(SER, 4) & "|"
    s = s & EsticaTexto(SUBSER, 3) & "|"
    s = s & EsticaTexto(NUM_DOC, 3) & "|"
    s = s & EsticaTexto(CHV_CTE, 44) & "|"
    s = s & GravaDataSintegra(DT_DOC) & "|"
    s = s & GravaDataSintegra(DT_AQUISICAO_PRESTACAO) & "|"
    s = s & EsticaTexto(TP_CTe, 1) & "|"
    s = s & EsticaTexto(CHV_CTE_REF, 44) & "|"
    s = s & GravaValorSped(vl_DOC) & "|"
    s = s & GravaValorSped(vl_DESC) & "|"
    s = s & EsticaTexto(CStr(IND_FRE), 1) & "|"
    s = s & GravaValorSped(vl_SERV) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_NT) & "|"
       
    s = s & EsticaTexto(COD_INF, 6) & "|"
    s = s & COD_CTA & "|"
    
    s = s & CrLf
    
    
    Registro_D100 = s
        
    
End Function


Function Registro_D110(NUM_ITEN As String, COD_ITEM As String, vl_SERV As String, VL_OUT As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    s = "|D110|"
    s = s & EsticaTexto(NUM_ITEN, 4) & "|"
    s = s & EsticaTexto(COD_ITEM, 60) & "|"
    s = s & GravaValorSintegra(vl_SERV) & "|"
    s = s & GravaValorSintegra(VL_OUT) & "|"
    
    s = s & CrLf
    Registro_D110 = s
        
    
End Function




Function Registro_D120(COD_MUN_ORIG As String, COD_MUN_DEST As String, placa As String, UF_PLACA As String) As String
    Dim s As String
    
    s = "|D120|"
    s = s & EsticaTexto(COD_MUN_ORIG, 7) & "|"
    s = s & EsticaTexto(COD_MUN_DEST, 7) & "|"
    s = s & EsticaTexto(placa, 7) & "|"
    s = s & EsticaTexto(UF_PLACA, 2) & "|"
    s = s & CrLf
    Registro_D120 = s
        
    
End Function

Function Registro_D130(COD_PART As String, COD_REDESPACHADO As String, IND_FRT_RED As String, _
COD_MUN_ORIG As String, COD_MUN_DEST As String, placa As String, VL_LIQ_FRE As String, _
VL_SEC_CAT As String, VL_DESP As String, VL_PEDAGIOS As String, VL_OUTROS As String, _
VL_TOTAL_FRT As String, UF_VEICULO As String) As String
    Dim s As String
    
    s = "|D130|"
    s = s & EsticaTexto(COD_PART, 60) & "|"
    s = s & EsticaTexto(COD_REDESPACHADO, 60) & "|"
    s = s & EsticaTexto(IND_FRT_RED, 1) & "|"  '0 sem redespacho 1- emiente 2-destinatario 9-outros
    s = s & EsticaTexto(COD_MUN_ORIG, 7) & "|"
    s = s & EsticaTexto(COD_MUN_DEST, 7) & "|"
    s = s & EsticaTexto(placa, 7) & "|"
    s = s & GravaValorSped(VL_LIQ_FRE, 2) & "|"
    s = s & GravaValorSped(VL_SEC_CAT, 2) & "|"
    s = s & GravaValorSped(VL_DESP, 2) & "|"
    s = s & GravaValorSped(VL_PEDAGIOS, 2) & "|"
    s = s & GravaValorSped(VL_OUTROS, 2) & "|"
    s = s & GravaValorSped(VL_TOTAL_FRT, 2) & "|"
    s = s & EsticaTexto(UF_VEICULO, 2) & "|"
    s = s & CrLf
    Registro_D130 = s
        
    
End Function

Function Registro_D160(DESPACHO As String, CNPJ_REMETENTE As String, IE_REMET As String, _
COD_MUN_ORIGEN As String, CNPJ_DESTINO As String, IE_DESTINO As String, COD_MUN_DEST As String) As String
    Dim s As String
    
    s = "|D160|"
    s = s & DESPACHO & "|"
    s = s & EsticaTexto(CNPJ_REMETENTE, 14) & "|"
    s = s & EsticaTexto(IE_REMET, 14) & "|"
    s = s & EsticaTexto(COD_MUN_ORIGEN, 7) & "|"
    s = s & EsticaTexto(CNPJ_DESTINO, 14) & "|"
    s = s & EsticaTexto(IE_DESTINO, 14) & "|"
    s = s & EsticaTexto(COD_MUN_DEST, 7) & "|"
    s = s & CrLf
    Registro_D160 = s
        
    
End Function

Function Registro_D161(IND_CARG As IND_CARGA, CNPJ_COLETA As String, IE_COLETA As String, _
COD_MUN_COLETA As String, CNPJ_ENTREGA As String, IE_ENTREGA As String, COD_MUN_ENTREGA As String) As String
    Dim s As String
    
    s = "|D161|"
    s = s & IND_CARG & "|"
    s = s & EsticaTexto(CNPJ_COLETA, 14) & "|"
    s = s & EsticaTexto(IE_COLETA, 14) & "|"
    s = s & EsticaTexto(COD_MUN_COLETA, 7) & "|"
    s = s & EsticaTexto(CNPJ_ENTREGA, 14) & "|"
    s = s & EsticaTexto(IE_ENTREGA, 14) & "|"
    s = s & EsticaTexto(COD_MUN_ENTREGA, 7) & "|"
    s = s & CrLf
    Registro_D161 = s
        
    
End Function


Function Registro_D500(IND_OPERACAO As IND_OPER, IND_EMITEnte As IND_EMIT, COD_PART As String, _
COD_MOD As String, COD_SIT As String, SERIE As String, SUBSER As String, NUM_DOC As String, DT_DOC As String, _
DT_ENTRADA_SAIDA As String, vl_DOC As String, vl_DESC As String, vl_SERV As String, VL_SERV_NT As String, _
VL_TERC As String, VL_DA As String, VL_BC_ICMS As String, VL_ICMS As String, COD_INF As String, VL_PIS As String, _
VL_COFINS As String, COD_CTA As String, TIPO_ASSINANTE As TP_ASSINANTE) As String
    Dim s As String
    
    s = "|D500|"
    s = s & IND_OPERACAO & "|"
    s = s & IND_EMITEnte & "|"
    s = s & EsticaTexto(COD_PART, 60) & "|"
    s = s & GravaNumeroSped(COD_MOD, 2) & "|"
    s = s & GravaNumeroSped(COD_SIT, 2) & "|"
    s = s & EsticaTexto(SERIE, 4) & "|"
    s = s & EsticaTexto(SUBSER, 3) & "|"
    s = s & GravaNumeroSped(NUM_DOC, 9) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & GravaDataSped(DT_ENTRADA_SAIDA) & "|"
    
    s = s & GravaValorSped(vl_DOC, 2) & "|"
    s = s & GravaValorSped(vl_DESC, 2) & "|"
    s = s & GravaValorSped(vl_SERV, 2) & "|"
    s = s & GravaValorSped(VL_SERV_NT, 2) & "|"
    s = s & GravaValorSped(VL_TERC, 2) & "|"
    s = s & GravaValorSped(VL_DA, 2) & "|"
    s = s & GravaValorSped(VL_BC_ICMS, 2) & "|"
    s = s & GravaValorSped(VL_ICMS, 2) & "|"
    s = s & EsticaTexto(COD_INF, 6) & "|"
    s = s & GravaValorSped(VL_PIS, 2) & "|"
    s = s & GravaValorSped(VL_COFINS, 2) & "|"
    s = s & COD_CTA & "|"
    
    s = s & EsticaTexto(CStr(TIPO_ASSINANTE), 1) & "|"
    s = s & CrLf
    Registro_D500 = s
        
    
End Function
Function Registro_D590(CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, _
VL_OPR As String, VL_BC_ICMS As String, VL_ICMS As String, VL_BC_ICMS_UF As String, _
VL_ICMS_UF As String, VL_RED_BC As String, COD_OBS As String) As String
    Dim s As String
    
    s = "|D590|"
    s = s & GravaNumeroSped(CST_ICMS, 3) & "|"
    s = s & GravaNumeroSped(CFOP, 4) & "|"
    s = s & GravaValorSped(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_OPR) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_UF) & "|"
    s = s & GravaValorSped(VL_ICMS_UF) & "|"
    s = s & GravaValorSped(VL_RED_BC) & "|"
    s = s & EsticaTexto(COD_OBS, 6) & "|"
    s = s & CrLf
    Registro_D590 = s
        
    
End Function

Function Registro_D610(COD_CLASS As String, COD_ITEM As String, QTD As String, _
UNID As String, vl_ITEM As String, vl_DESC As String, CST_ICMS As String, _
CFOP As String, ALIQ_ICMS As String, VL_BC_ICMS As String, VL_ICMS As String, _
VL_BC_ICMS_UF As String, VL_ICMS_UF As String, VL_RED_BC As String, _
VL_PIS As String, VL_COFINS As String, COD_CTA As String) As String
    Dim s As String
    
    s = "|D610|"
    s = s & EsticaTexto(COD_CLASS, 4) & "|"
    s = s & EsticaTexto(COD_ITEM, 60) & "|"
    s = s & GravaNumeroSped(QTD, 3) & "|"
    s = s & EsticaTexto(UNID, 6) & "|"
    s = s & GravaNumeroSped(CFOP, 4) & "|"
    s = s & GravaValorSped(vl_ITEM) & "|"
    s = s & GravaValorSped(vl_DESC) & "|"
    s = s & EsticaTexto(CST_ICMS, 3) & "|"
    s = s & EsticaTexto(CFOP, 4) & "|"
    s = s & GravaValorSped(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_UF) & "|"
    s = s & GravaValorSped(VL_ICMS_UF) & "|"
    s = s & GravaValorSped(VL_RED_BC) & "|"
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    s = s & COD_CTA & "|"
    s = s & CrLf
    Registro_D610 = s
        
    
End Function



Function Registro_C690(CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, _
    VL_OPR As String, VL_BC_ICMS As String, VL_ICMS As String, VL_BC_ICMS_ST As String, _
    VL_ICMS_ST As String, VL_RED_BC As String, COD_OBS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C690|"
    s = s & EsticaTexto(CST_ICMS, 3) & "|"
    s = s & EsticaTexto(CFOP, 4) & "|"
    s = s & gravaAliquotaSPED(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_OPR) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_RED_BC) & "|"
    's = s & GravaValorSped(VL_IPI) & "|"
    s = s & EsticaTexto(COD_OBS, 6) & "|"
    s = s & CrLf
    
    
    Registro_C690 = s
        
    
End Function

Function Registro_D696(CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, _
    VL_OPR As String, VL_BC_ICMS As String, VL_ICMS As String, VL_BC_ICMS_UF As String, _
    VL_ICMS_UF As String, VL_RED_BC As String, COD_OBS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C696|"
    s = s & EsticaTexto(CST_ICMS, 3) & "|"
    s = s & EsticaTexto(CFOP, 4) & "|"
    s = s & gravaAliquotaSPED(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_OPR) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_UF) & "|"
    s = s & GravaValorSped(VL_ICMS_UF) & "|"
    s = s & GravaValorSped(VL_RED_BC) & "|"
    s = s & EsticaTexto(COD_OBS, 6) & "|"
    
    s = s & CrLf
    
    
    Registro_D696 = s
        
    
End Function

Function Registro_D990(qtd_linhas As String) As String
    Dim s As String
    s = "|D990|" & qtd_linhas & "|"
    s = s & CrLf
    Registro_D990 = s
End Function

Function Registro_E001(Ind_mov As String) As String
    Dim s As String
    s = "|E001|" & Ind_mov & "|" '0 bloco sem dados 1 bloco com dados informados
    s = s & CrLf
    Registro_E001 = s
End Function
Function Registro_E100(DT_INI As String, DT_FIN As String) As String
    Dim s As String
    s = "|E100|"
    s = s & GravaDataSped(DT_INI) & "|"
    s = s & GravaDataSped(DT_FIN) & "|"
    s = s & CrLf
    Registro_E100 = s
End Function

Function Registro_E110(VL_TOT_DEBITOS As String, VL_AJ_DEBITOS As String, VL_TOT_AJ_DEBITOS As String, _
    VL_ESTORNOS_CRED As String, VL_TOT_CREDITOS As String, VL_AJ_CREDITOS As String, _
    VL_TOT_AJ_CREDITOS As String, VL_ESTORNOS_DEB As String, VL_SLD_CREDOR_ANT As String, _
    VL_SLD_APURADO As String, VL_TOT_DED As String, VL_ICMS_RECOLHER As String, _
    VL_SLD_CREDOR_TRANSPORTAR As String, DEB_ESP As String) As String
    Dim s As String
    s = "|E110|"
    s = s & GravaValorSped(VL_TOT_DEBITOS) & "|"
    s = s & GravaValorSped(VL_AJ_DEBITOS) & "|"
    s = s & GravaValorSped(VL_TOT_AJ_DEBITOS) & "|"
    s = s & GravaValorSped(VL_ESTORNOS_CRED) & "|"
    s = s & GravaValorSped(VL_TOT_CREDITOS) & "|"
    s = s & GravaValorSped(VL_AJ_CREDITOS) & "|"
    s = s & GravaValorSped(VL_SLD_CREDOR_ANT) & "|"
    s = s & GravaValorSped(VL_SLD_APURADO) & "|"
    s = s & GravaValorSped(VL_TOT_DED) & "|"
    s = s & GravaValorSped(VL_ICMS_RECOLHER) & "|"
    s = s & GravaValorSped(VL_SLD_CREDOR_TRANSPORTAR) & "|"
    s = s & GravaValorSped(DEB_ESP) & "|"
    
    s = s & CrLf
    Registro_E110 = s
End Function



Function Registro_E111(COD_AJ_APUR As String, DESCR_COMPL_AJ As String, vl_AJ_APUR As String) As String
    Dim s As String
    s = "|E111|"
    s = s & EsticaTexto(COD_AJ_APUR, 8) & "|"
    s = s & DESCR_COMPL_AJ & "|"
    s = s & GravaValorSped(vl_AJ_APUR) & "|"
    s = s & CrLf
    Registro_E111 = s
End Function

Function Registro_E112(NUM_DA As String, NUM_PROC As String, IND_PROC As String) As String
    Dim s As String
    s = "|E112|"
    s = s & NUM_DA & "|"
    s = s & EsticaTexto(NUM_PROC, 15) & "|"
    s = s & EsticaTexto(IND_PROC, 1) & "|" '0 sefaz 1 federal 2 estado 9 outros
    s = s & CrLf
    Registro_E112 = s
End Function

Function Registro_E113(COD_PART As String, COD_MOD As String, SERIE As String, SUBSER As String, _
NUM_DOC As String, DT_DOC As String, COD_ITEM As String, VL_AJUSTE_ITEM As String) As String
    Dim s As String
    s = "|E113|"
    s = s & EsticaTexto(COD_PART, 60) & "|"
    s = s & EsticaTexto(COD_MOD, 2) & "|"
    s = s & EsticaTexto(SERIE, 4) & "|"
    s = s & EsticaTexto(SUBSER, 3) & "|"
    s = s & EsticaTexto(NUM_DOC, 9) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & EsticaTexto(COD_ITEM, 60) & "|"
    
    s = s & GravaValorSped(VL_AJUSTE_ITEM) & "|"
    s = s & CrLf
    Registro_E113 = s
End Function
Function Registro_E115(COD_INF_ADIC As String, VL_INF_ADIC As String, DESC_COMPL_AJ As String) As String
    Dim s As String
    s = "|E115|"
    s = s & EsticaTexto(COD_INF_ADIC, 8) & "|"
    
    s = s & DESC_COMPL_AJ & "|"
    
    s = s & CrLf
    Registro_E115 = s
End Function



Function Registro_E116(COD_OR As String, VL_OR As String, DT_VCTO As String, _
COD_REC As String, NUM_PROC As String, IND_PROC As String, PROC As String, txt_COMPL As String, _
MES_REF As String) As String
    Dim s As String
    s = "|E116|"
    s = s & EsticaTexto(COD_OR, 4) & "|"
    s = s & GravaValorSped(VL_OR) & "|"
    s = s & GravaDataSped(DT_VCTO) & "|"
    s = s & COD_REC & "|"
    s = s & EsticaTexto(NUM_PROC, 15) & "|"
    s = s & EsticaTexto(IND_PROC, 1) & "|" '0 sefaz 1 federal 2 estado 9 outros
    s = s & txt_COMPL & "|"
    s = s & MES_REF & "|" 'MMAAAA
    s = s & CrLf
    Registro_E116 = s
End Function



Function Registro_E200(UF As String, DT_INI As String, DT_FIN As String, cod_ant_item As String) As String
Dim s As String
    s = "|E200|"
    s = s & UF & "|"
    s = s & GravaDataSped(DT_INI) & "|"
    s = s & GravaDataSped(DT_FIN) & "|"
    
    s = s & CrLf
    
    
    
    
    Registro_E200 = s
End Function


Function Registro_E210(Ind_mov As String, VL_SD_CRED_ANT_ST As String, VL_DEVOL_ST As String, _
    VL_RESSARC_ST As String, VL_OUT_CRED_ST As String, VL_AJ_CREDITOS_ST As String, VL_RETENCAO_ST As String, _
    VL_OUT_DEB_ST As String, VL_AJ_DEB_ST As String, _
    VL_SALDO_DEV_ANT_ST As String, VL_DEDUCOES_ST As String, VL_ICMS_RECOL_ST As String, _
    VL_SD_CRED_ST_TRANSPORTAR As String, DEB_ESP_ST As String) As String
    Dim s As String
    s = "|E210|"
    s = s & EsticaTexto(Ind_mov, 1) & "|" '0 sem operação st 1: com operação st
    s = s & GravaValorSped(VL_SD_CRED_ANT_ST) & "|"
    s = s & GravaValorSped(VL_DEVOL_ST) & "|"
    s = s & GravaValorSped(VL_RESSARC_ST) & "|"
    s = s & GravaValorSped(VL_OUT_CRED_ST) & "|"
    s = s & GravaValorSped(VL_AJ_CREDITOS_ST) & "|"
    s = s & GravaValorSped(VL_RETENCAO_ST) & "|"
    s = s & GravaValorSped(VL_OUT_DEB_ST) & "|"
    s = s & GravaValorSped(VL_SALDO_DEV_ANT_ST) & "|"
    s = s & GravaValorSped(VL_DEDUCOES_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_RECOL_ST) & "|"
    s = s & GravaValorSped(VL_SD_CRED_ST_TRANSPORTAR) & "|"
    s = s & GravaValorSped(DEB_ESP_ST) & "|"
    
    s = s & CrLf
    Registro_E210 = s
End Function


Function Registro_E220(COD_AJ_APUR As String, DESCR_COMPL_AJ As String, vl_AJ_APUR As String) As String
Dim s As String
    s = "|E220|"
    s = s & EsticaTexto(COD_AJ_APUR, 8) & "|" 'tabela 5.1.1
    s = s & DESCR_COMPL_AJ & "|"
    
    s = s & GravaValorSped(vl_AJ_APUR) & "|"
    
    s = s & CrLf
    
    
    
    
    Registro_E220 = s
End Function
Function Registro_E230(NUM_DA As String, NUM_PROC As String, IND_PROC As String, RESUMO_PROCESSO As String, _
txt_COMPL As String) As String
    Dim s As String
    s = "|E230|"
    s = s & NUM_DA & "|"
    s = s & EsticaTexto(NUM_PROC, 15) & "|"
    s = s & EsticaTexto(IND_PROC, 1) & "|" '0 sefaz 1 federal 2 estadual 9 outros
    s = s & RESUMO_PROCESSO & "|"
    s = s & txt_COMPL & "|"
    s = s & CrLf
    Registro_E230 = s
End Function

Function Registro_E240(COD_PART As String, COD_MOD As String, SERIE As String, SUBSER As String, _
NUM_DOC As String, DT_DOC As String, COD_ITEM As String, VL_AJUSTE_ITEM As String) As String
    Dim s As String
    s = "|E240|"
    s = s & EsticaTexto(COD_PART, 60) & "|"
    s = s & EsticaTexto(COD_MOD, 2) & "|"
    s = s & EsticaTexto(SERIE, 4) & "|"
    s = s & EsticaTexto(SUBSER, 3) & "|"
    s = s & EsticaTexto(NUM_DOC, 9) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & EsticaTexto(COD_ITEM, 60) & "|"
    
    s = s & GravaValorSped(VL_AJUSTE_ITEM) & "|"
    s = s & CrLf
    Registro_E240 = s
End Function


Function Registro_E250(COD_OR As String, VL_OR As String, DT_VCTO As String, _
COD_REC As String, NUM_PROC As String, IND_PROC As String, PROC As String, txt_COMPL As String, _
MES_REF As String) As String
    Dim s As String
    s = "|E116|"
    s = s & EsticaTexto(COD_OR, 4) & "|"
    s = s & GravaValorSped(VL_OR) & "|"
    s = s & GravaDataSped(DT_VCTO) & "|"
    s = s & COD_REC & "|"
    s = s & EsticaTexto(NUM_PROC, 15) & "|"
    s = s & EsticaTexto(IND_PROC, 1) & "|" '0 sefaz 1 federal 2 estado 9 outros
    s = s & txt_COMPL & "|"
    s = s & MES_REF & "|" 'MMAAAA
    s = s & CrLf
    Registro_E250 = s
End Function


Function Registro_E500(IND_APURA As IND_APUR, DT_INI As String, DT_FIN As String) As String
    Dim s As String
    s = "|E500|"
    s = s & IND_APURA & "|"
    s = s & GravaDataSped(DT_INI) & "|"
    s = s & GravaDataSped(DT_FIN) & "|"
    s = s & CrLf
    Registro_E500 = s
End Function

Function Registro_E510(CFOP As String, CST_IPI As String, VL_CONTABIL_IPI As String, _
VL_BC_IPI As String, VL_IPI As String) As String
    Dim s As String
    s = "|E510|"
    s = s & EsticaTexto(CFOP, 4) & "|" 'SINIEF 07/01.
    s = s & EsticaTexto(CST_IPI, 2) & "|"  'tabela 4.3.2
    
    s = s & GravaValorSped(VL_BC_IPI) & "|"
    s = s & GravaValorSped(VL_CONTABIL_IPI) & "|"
    s = s & GravaValorSped(VL_IPI) & "|"
    
    s = s & CrLf
    Registro_E510 = s
End Function
Function Registro_E520(vl_SD_ANT_IPI As String, VL_DEB_IPI As String, VL_CRED_IPI As String, _
VL_OD_IPI As String, VL_OC_IPI As String, VL_SaldoCredor_IPI As String, _
VL_Saldodevedor As String) As String
    Dim s As String
    s = "|E520|"
    s = s & GravaValorSped(vl_SD_ANT_IPI) & "|"
    
    s = s & GravaValorSped(VL_DEB_IPI) & "|"
    s = s & GravaValorSped(VL_CRED_IPI) & "|"
    s = s & GravaValorSped(VL_OD_IPI) & "|"
    s = s & GravaValorSped(VL_OC_IPI) & "|"
    s = s & GravaValorSped(VL_SaldoCredor_IPI) & "|"
    s = s & GravaValorSped(VL_Saldodevedor) & "|"
    s = s & CrLf
    Registro_E520 = s
End Function

Function Registro_E530(IND_AJ As String, VL_AJ As String, COD_AJ As String, IND_DOC As String, _
NUM_DC As String, DESCR_AJ As String) As String
    Dim s As String
    s = "|E530|"
    s = s & EsticaTexto(IND_AJ, 1) & "|"  '0 deb 1 cred
    s = s & GravaValorSped(VL_AJ) & "|"
    s = s & EsticaTexto(COD_AJ, 3) & "|"  'tabwla 4.5.4
    s = s & EsticaTexto(IND_DOC, 1) & "|"  '0 proc judicial 1:proc adm 2:per/dcomp 9:outros
    s = s & NUM_DC & "|"
    s = s & DESCR_AJ & "|"
    
    
    s = s & CrLf
    Registro_E530 = s


End Function

Function Registro_E990(qtd_linhas As String) As String
    Dim s As String
    s = "|E990|" & qtd_linhas & "|"
    s = s & CrLf
    Registro_E990 = s
End Function

Function Registro_H001(Ind_mov As String) As String
    Dim s As String
    s = "|H001|" & Ind_mov & "|" '0 bloco sem dados 1 bloco com dados informados
    s = s & CrLf
    Registro_H001 = s
End Function

Function Registro_E005(dt_INV As String, vl_INV As String, MOT_INV As String) As String
    Dim s As String
    s = "|E005|"
    s = s & GravaDataSped(dt_INV) & "|"
    s = s & GravaValorSped(vl_INV) & "|"
    s = s & GravaNumeroSped(MOT_INV, 2) & "|"
    s = s & CrLf
    Registro_E005 = s
    
    '01 final periodo
    '02 – Na mudança de forma de tributação da mercadoria (ICMS);
    '03 – Na solicitação da baixa cadastral, paralisação temporária e outras situações;
    '04 – Na alteração de regime de pagamento – condição do contribuinte;
    '05 – Por determinação dos fiscos.
End Function



Function Registro_H010(COD_ITEM As String, UNID As String, QTD As String, VL_UNIT As String, _
    vl_ITEM As String, IND_PROP As String, COD_PART As String, txtCompl As String, COD_CTA As String) As String
    Dim s As String
    s = "|H010|"
    s = s & EsticaTexto(COD_ITEM, 60) & "|"
    s = s & EsticaTexto(UNID, 6) & "|"
    s = s & GravaValorSped(QTD, 3) & "|"
    s = s & txtCompl & "|"
    s = s & COD_CTA & "|"
    s = s & CrLf
    Registro_H010 = s
End Function


Function Registro_H020(CST_ICMS As String, BC_ICMS As String, VL_ICMS As String) As String
    Dim s As String
    s = "|H020|"
    s = s & EsticaTexto(CST_ICMS, 3) & "|"
    s = s & GravaValorSped(BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    
    s = s & CrLf
    Registro_H020 = s
    
    
End Function

Function Registro_H990(qtd_linhas As String) As String
    Dim s As String
    s = "|H990|" & qtd_linhas & "|"
    s = s & CrLf
    Registro_H990 = s
End Function

Function Registro_1001(Ind_mov As String) As String
    Dim s As String
    s = "|1001|" & Ind_mov & "|" '0 bloco sem dados 1 bloco com dados informados
    s = s & CrLf
    Registro_1001 = s
End Function

Function Registro_1010(IND_EXP As String, IND_CCRF As String, IND_COMB As String, IND_USINA As String, _
    IND_VA As String, IND_EE As String, IND_CART As String, IND_FORM As String, IND_AER As String) As String
    Dim s As String
    s = "|1010|"
    s = s & IND_EXP & "|"
    s = s & IND_CCRF & "|"
    s = s & IND_COMB & "|"
    s = s & IND_USINA & "|"
    s = s & IND_VA & "|"
    s = s & IND_EE & "|"
    s = s & IND_CART & "|"
    s = s & IND_FORM & "|"
    s = s & IND_AER & "|"
    s = s & CrLf
    Registro_1010 = s
End Function

Function Registro_1600(COD_PART As String, TOT_CREDITO As String, TOT_DEBITO As String) As String
    Dim s As String
    s = "|1600|"
    s = s & EsticaTexto(COD_PART, 60) & "|"
    s = s & GravaValorSped(TOT_CREDITO) & "|"
    s = s & GravaValorSped(TOT_DEBITO) & "|"
    s = s & CrLf
    Registro_1600 = s
End Function

Function Registro_9001(Ind_mov As String) As String
    Dim s As String
    s = "|9001|" & Ind_mov & "|"
    s = s & CrLf
    Registro_9001 = s
End Function

Function Registro_9900(REG_BLC As String, qtd_REG_BLC As String) As String
    Dim s As String
    s = "|9900|"
    s = s & EsticaTexto(REG_BLC, 4) & "|"
    s = s & qtd_REG_BLC & "|"
    s = s & CrLf
    Registro_9900 = s
End Function



Function Registro_9990(qtd_linhas As String) As String
    Dim s As String
    s = "|9990|" & qtd_linhas & "|"
    s = s & CrLf
    Registro_9990 = s
End Function

Function Registro_9999(qtd_linhas As String) As String
    Dim s As String
    s = "|9999|" & qtd_linhas & "|" 'linhas do arquivo digital
    s = s & CrLf
    Registro_9999 = s
End Function

