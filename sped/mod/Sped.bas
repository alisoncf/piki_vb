Attribute VB_Name = "Sped"

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
     outros_atv = 1
     
End Enum
Public Enum IND_OPER
    Entrada = 0
    saida = 1
End Enum
Public Enum ind_emit
    emissao_propria = 0
    terceiros = 1
End Enum
Public Enum IND_PGTO
    avista = 0
    aprazo = 1
    outros_pag = 2
End Enum
Public Enum IND_CARGA
   RODOVIARIO = 0
   FERROVIARIO = 1
   RODO_FERRO = 2
   AQUAVIARIO = 3
   DUTO_VIARIO = 4
   AEREO = 5
   outros_carga = 9
   
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
    uf As String
    IE As String
    COD_MUN As String
    im As String
    suframa As String
    ind_perfil As String 'a b ou c
    Iind_atv As ind_atv
    
    
    
    
End Type
Function GravaDataSped(data As String) As String
    
    GravaDataSped = Format(data, "ddmmyyyy")
End Function
Function GravaValorSped(valor As String, Optional CasasDecimais As Integer = 2)
    zeros = String(CasasDecimais, "0")

    GravaValorSped = Format(valor, "##0." & zeros)
End Function
Function GravaNumeroSped(valor As String, tamanho As Integer) As String

    GravaNumeroSped = EsticaTextoSped(valor, tamanho, 2, "0")
End Function
Function gravaAliquotaSPED(Aliquota As String, Optional tamanho As Integer = 2) As String
    If Aliquota = "FF" Or Aliquota = "II" Then
        Aliquota = "17"
    End If
    Aliquota = Format(Aliquota, "##0.00")
    gravaAliquotaSPED = EsticaTextoSped(Aliquota, tamanho, 2, "0", True)
End Function


Function Registro_0000(r As Reg_0) As String
    Dim s As String
    s = "|0000|"
    s = s & EsticaTextoSped(r.cod_ver, 3) & "|"
    s = s & EsticaTextoSped(CStr(r.cod_fin), 3) & "|"
    s = s & EsticaTextoSped(r.cod_ver, 1) & "|"
    s = s & GravaDataSped(r.DT_INI) & "|"
    s = s & GravaDataSped(r.DT_FIN) & "|"
    s = s & EsticaTextoSped(r.nome, 100) & "|"
    s = s & EsticaTextoSped(r.CNPJ, 14) & "|"
    s = s & EsticaTextoSped(r.CPF, 11) & "|"
    s = s & EsticaTextoSped(r.uf, 2) & "|"
    s = s & EsticaTextoSped(r.IE, 4) & "|"
    s = s & EsticaTextoSped(r.COD_MUN, 7) & "|"
    s = s & EsticaTextoSped(r.im, 10) & "|"
    s = s & EsticaTextoSped(r.suframa, 9) & "|"
    s = s & EsticaTextoSped(r.ind_perfil, 1) & "|" 'A b ou c
    s = s & EsticaTextoSped(CStr(r.Iind_atv), 1) & "|" & CrLf
    Registro_0000 = s
End Function
Function Registro_0001(Optional Ind_mov As String = "0") As String
    Dim s As String
    s = "|0001|"
    s = s & EsticaTextoSped(Ind_mov, 1) & "|"
    s = s & CrLf
    Registro_0001 = s
End Function


Function Registro_0005(fantasia As String, cep As String, ende As String, Num As String, _
    compl As String, bairro As String, fone As String, fax As String, email As String) As String
    
    Dim s As String
    s = "|0005|"
    s = s & EsticaTextoSped(fantasia, 60) & "|"
    s = s & EsticaTextoSped(cep, 8) & "|"
    s = s & EsticaTextoSped(ende, 60) & "|"
    s = s & EsticaTextoSped(Num, 10) & "|"
    s = s & EsticaTextoSped(compl, 60) & "|"
    s = s & EsticaTextoSped(bairro, 60) & "|"
    s = s & EsticaTextoSped(fone, 11) & "|"
    s = s & EsticaTextoSped(fax, 11) & "|"
    s = s & EsticaTextoSped(email, 60) & "|" & CrLf
    
    Registro_0005 = s
    
End Function

Function Registro_0100(nome As String, CPF As String, crc As String, CNPJ As String, cep As String, ende As String, Num As String, _
    compl As String, bairro As String, fone As String, fax As String, email As String) As String
    
    Dim s As String
    s = "|0100|"
    s = s & EsticaTextoSped(nome, 60) & "|"
    s = s & EsticaTextoSped(CPF, 11) & "|"
    s = s & EsticaTextoSped(crc, 15) & "|"
    s = s & EsticaTextoSped(CNPJ, 14) & "|"
    s = s & EsticaTextoSped(cep, 8) & "|"
    s = s & EsticaTextoSped(ende, 60) & "|"
    s = s & EsticaTextoSped(Num, 10) & "|"
    s = s & EsticaTextoSped(compl, 60) & "|"
    s = s & EsticaTextoSped(bairro, 60) & "|"
    s = s & EsticaTextoSped(fone, 11) & "|"
    s = s & EsticaTextoSped(fax, 11) & "|"
    s = s & EsticaTextoSped(email, 60) & "|" & CrLf
    
    Registro_0100 = s
    
End Function
Function Registro_0015(uf_st As String, ie_st As String) As String
    Dim s As String
    s = "|0015|" & EsticaTextoSped(uf_st, 2) & "|" & EsticaTextoSped(ie_st, 10) & "|" & CrLf
    
    Registro_0015 = s
End Function

Function Registro_0150(COD_PART As String, nome As String, cod_pais As String, CNPJ As String, _
    CPF As String, IE As String, COD_MUN As String, suframa As String, ende As String, Num As String, _
    compl As String, bairro As String) As String
Dim s As String
    s = "|0150|"
    s = s & EsticaTextoSped(COD_PART, 60) & "|"
    s = s & EsticaTextoSped(nome, 100) & "|"
    s = s & EsticaTextoSped(cod_pais, 5) & "|"
    s = s & EsticaTextoSped(CNPJ, 14) & "|"
    s = s & EsticaTextoSped(IE, 11) & "|"
    s = s & EsticaTextoSped(suframa, 9) & "|"
    s = s & EsticaTextoSped(ende, 60) & "|"
    s = s & EsticaTextoSped(Num, 10) & "|"
    s = s & EsticaTextoSped(compl, 60) & "|"
    s = s & EsticaTextoSped(bairro, 60) & "|" & CrLf
    
    
    
    
    Registro_0150 = s
End Function
Function Registro_0175(dt_alt As String, nr_campo As String, cont_ant As String) As String
Dim s As String
    s = "|0175|"
    s = s & GravaDataSped(dt_alt) & "|"
    s = s & EsticaTextoSped(nr_campo, 2) & "|"
    s = s & EsticaTextoSped(cont_ant, 100) & "|"
    s = s & CrLf
    
    
    
    
    Registro_0175 = s
End Function

Function Registro_0190(UNID As String, descr As String) As String
Dim s As String
    s = "|0175|"
    
    s = s & EsticaTextoSped(UNID, 4) & "|"
    s = s & EsticaTextoSped(descr, 100) & "|"
    s = s & CrLf
    
    
    
    
    Registro_0190 = s
End Function

Function Registro_0200(COD_ITEM As String, descr_item As String, cod_barra As String, _
    cod_ant_item As String, unid_inv As String, tipo_it As tipo_item, _
    cod_ncm As String, ex_ipi As String, cod_gen As String, cod_lst As String, ALIQ_ICMS As String) As String
Dim s As String
    s = "|0200|"
    
    s = s & EsticaTextoSped(COD_ITEM, 60) & "|"
    s = s & EsticaTextoSped(descr_item, 100) & "|"
    s = s & cod_barra & "|"
    s = s & EsticaTextoSped(cod_ant_item, 60) & "|"
    s = s & EsticaTextoSped(unid_inv, 6) & "|"
    s = s & EsticaTextoSped(CStr(tipo_it), 2) & "|"
    s = s & EsticaTextoSped(cod_ncm, 8) & "|"
    s = s & EsticaTextoSped(ex_ipi, 3) & "|"
    s = s & EsticaTextoSped(cod_gen, 2) & "|"
    s = s & EsticaTextoSped(cod_lst, 4) & "|"
    s = s & EsticaTextoSped(GravaValorSped(ALIQ_ICMS), 6) & "|"
    s = s & CrLf
    
    
    
    
    Registro_0200 = s
End Function


Function Registro_0205(descr_anterior_item As String, DT_INI As String, DT_FIN As String, cod_ant_item As String) As String
Dim s As String
    s = "|0205|"
    s = s & descr_anterior_item & "|"
    s = s & GravaDataSped(DT_INI) & "|"
    s = s & GravaDataSped(DT_FIN) & "|"
    s = s & EsticaTextoSped(cod_ant_item, 60) & "|"
    s = s & CrLf
    
    
    
    
    Registro_0205 = s
End Function

Function Registro_0300(cod_ind_bem As String, ident_merc As String, descr_item As String, cod_prnc As String, _
    COD_CTA As String, nr_parc As String) As String
Dim s As String
    s = "|0300|"
    s = s & EsticaTextoSped(cod_ind_bem, 60) & "|"
    s = s & EsticaTextoSped(ident_merc, 1) & "|"  '1 mercadoria 2 componente
    s = s & EsticaTextoSped(cod_prnc, 1) & "|"   'caso seja componente, o código do produto principal
    s = s & EsticaTextoSped(COD_CTA, 60) & "|"   'conta analítica de contabilização
    s = s & EsticaTextoSped(nr_parc, 3) & "|" 'parcelas a serem apropriadas
    s = s & CrLf
    
    
    
    
    Registro_0300 = s
End Function
Function Registro_0305(cod_ccus As String, func As String, vida_util As String) As String
Dim s As String
    s = "|0305|"
    s = s & EsticaTextoSped(cod_ccus, 60) & "|"
    s = s & func & "|"
    s = s & GravaNumeroSped(vida_util, 3) & "|"    'caso seja componente, o código do produto principal
    
    s = s & CrLf
    
    
    
    
    Registro_0305 = s
End Function
Function Registro_0400(cod_nat As String, descr As String) As String
Dim s As String
    s = "|0400|"
    s = s & EsticaTextoSped(cod_nat, 60) & "|"
    s = s & descr & "|"
    
    
    s = s & CrLf
    
    
    
    
    Registro_0400 = s
End Function
Function Registro_0450(codigo As String, Descricao As String) As String
Dim s As String
    s = "|0450|"
    s = s & EsticaTextoSped(codigo, 6) & "|"
    s = s & Descricao & "|"
    
    
    s = s & CrLf
    
    
    
    
    Registro_0450 = s
End Function
Function Registro_0460(codigo As String, Descricao As String) As String
Dim s As String
    s = "|0460|"
    s = s & EsticaTextoSped(codigo, 6) & "|"
    s = s & Descricao & "|"
    
    
    s = s & CrLf
    
    
    
    
    Registro_0460 = s
End Function
Function Registro_0500(dt_alt As String, cod_nat_cc As String, ind_cta As String, nivel As String, _
    COD_CTA As String, nome_cta As String) As String
Dim s As String
    s = "|0500|"
    s = s & GravaDataSped(dt_alt) & "|"
    s = s & GravaNumeroSped(cod_nat_cc, 2) & "|"
    s = s & GravaNumeroSped(ind_cta, 1) & "|"
    s = s & GravaNumeroSped(nivel, 5) & "|"
    s = s & EsticaTextoSped(COD_CTA, 60) & "|"
    s = s & EsticaTextoSped(nome_cta, 60) & "|"
    
    
    s = s & CrLf
    
    
    
    
    Registro_0500 = s
End Function
Function Registro_0600(dt_alt As String, cod_ccus As String, ccus As String) As String
Dim s As String
    s = "|0600|"
    s = s & GravaDataSped(dt_alt) & "|"
    s = s & EsticaTextoSped(cod_ccus, 60) & "|"
    
    
    s = s & EsticaTextoSped(ccus, 60) & "|"
    
    
    s = s & CrLf
    
    
    
    
    Registro_0600 = s
End Function

Function Registro_0990(qtd_linhas As String) As String
    Dim s As String
    s = "|0990|" & qtd_linhas & "|"
    Registro_0990 = s
End Function


Function Registro_C001(Ind_mov As String) As String
    Dim s As String
    s = "|C001|" & Ind_mov & "|" '0 bloco sem dados 1 bloco com dados informados
    Registro_C001 = s
    
End Function

Function Registro_C100(Ind_Op As String, ind_em As String, COD_PART As String, cod_modelo_nota As String, _
    cod_sit_nota As String, SERIE As String, NUM_DOC As String, chv_nfe As String, DT_DOC As String, _
        DT_ENTRADA_SAIDA As String, Valor_Doc As String, ind_pg As IND_PGTO, valor_desconto As String, _
        vlr_abat_not As String, valor_total_merc As String, Ind_frete As IND_FRT, _
        valor_frete As String, valor_seguro As String, valor_out_da As String, vlr_bcicms As String, _
        vlr_icms As String, vlr_bc_icms_st As String, vlr_icms_st As String, _
        vlr_ipi As String, vlr_pis As String, vlr_cofins As String, vlr_pis_st As String, _
        vlr_cofins_st As String) As String
    Dim s As String
    s = "|C100|"
    s = s & Ind_Op & "|"
    s = s & ind_em & "|"
    s = s & EsticaTextoSped(COD_PART, 60) & "|"
    s = s & EsticaTextoSped(cod_modelo_nota, 2) & "|"
    s = s & EsticaTextoSped(cod_sit_nota, 2) & "|"
    s = s & EsticaTextoSped(SERIE, 3) & "|"
    s = s & EsticaTextoSped(NUM_DOC, 9) & "|"
    s = s & EsticaTextoSped(chv_nfe, 44) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & GravaDataSped(DT_ENTRADA_SAIDA) & "|"
    s = s & GravaValorSped(Valor_Doc) & "|"
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
Function Registro_C105(oper As String, uf As String) As String
    'oper 0 combustivel e lubrificantes 1 veiculos leansing
    Dim s As String
    s = "|C105|"
    s = s & oper & "|" & uf & "|" & CrLf
    
    
    Registro_C105 = s
        
    
End Function



Function Registro_C110(COD_INF As String, texto As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    s = "|C110|"
    s = s & EsticaTextoSped(COD_INF, 6) & texto & "|" & CrLf
    
    
    Registro_C110 = s
        
    
End Function

Function Registro_C111(NUM_PROC As String, IND_PROC As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    'ind_proc - 0 sefaz, 1 justiça federal, 2 justiça estadual 3 SECEX/SRF 9 outros
    
    
    s = "|C110|"
    s = s & EsticaTextoSped(NUM_PROC, 15) & IND_PROC & "|" & CrLf
    
    
    Registro_C111 = s
        
    
End Function
Function Registro_C112(cod_DA As String, uf As String, NUM_DA As String, cod_aut As String, DT_VCTO As String, dt_pgto As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    'cod_da= 0 doc arrecadação estadual  1 GNRE
    s = "|C110|"
    s = s & EsticaTextoSped(cod_DA, 1) & "|"
    s = s & EsticaTextoSped(uf, 2) & "|"
    s = s & NUM_DA & "|"
    s = s & GravaDataSped(DT_VCTO) & "|"
    s = s & GravaDataSped(dt_pgto) & "|"
    s = s & CrLf
    
    
    Registro_C112 = s
        
    
End Function

Function Registro_C113(IND_OPER As String, ind_emit, COD_PART As String, COD_MOD As String, SER As String, sub_s As String, NUM_DOC As String, DT_DOC As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    'ind_oper= 0 entrada 1 saida
    'ind_emit 0 = emissao propria 1 terceiros
    s = "|C113|"
    s = s & EsticaTextoSped(IND_OPER, 1) & "|"
    s = s & EsticaTextoSped(CStr(ind_emit), 1) & "|"
    s = s & EsticaTextoSped(COD_PART, 60) & "|"
    s = s & EsticaTextoSped(COD_MOD, 2) & "|"
    s = s & EsticaTextoSped(SER, 4) & "|"
    s = s & EsticaTextoSped(sub_s, 3) & "|"
    s = s & EsticaTextoSped(NUM_DOC, 9) & "|"
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
    s = s & EsticaTextoSped(COD_MOD, 2) & "|"
    s = s & EsticaTextoSped(ECF_FAB, 21) & "|"
    s = s & EsticaTextoSped(ECF_CX, 3, 2, "0") & "|"
    s = s & EsticaTextoSped(NUM_DOC, 6, 2, "0") & "|"
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
    s = s & EsticaTextoSped(IND_CARGA, 1) & "|"
    s = s & EsticaTextoSped(cnpj_col, 14) & "|"
    s = s & EsticaTextoSped(ie_col, 14) & "|"
    s = s & EsticaTextoSped(cpf_col, 11) & "|"
    s = s & EsticaTextoSped(cod_mun_col, 7) & "|"
    s = s & EsticaTextoSped(cnpj_entg, 14) & "|"
    s = s & EsticaTextoSped(ie_entg, 14) & "|"
    s = s & EsticaTextoSped(cpf_entg, 11) & "|"
    s = s & EsticaTextoSped(cod_mun_entg, 7) & "|"
    s = s & CrLf
    
    
    Registro_C115 = s
        
    
End Function


Function Registro_C120(cod_doc_imp As String, num_doc_imp As String, pis_imp As String, cofins_imp As String, num_acdraw As String) As String
    
    Dim s As String
    
    'cod_doc_imp 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C120|"
    s = s & EsticaTextoSped(cod_doc_imp, 1) & "|"
    s = s & EsticaTextoSped(num_doc_imp, 12) & "|"
    s = s & GravaValorSped(pis_imp) & "|"
    s = s & GravaValorSped(cofins_imp) & "|"
    s = s & EsticaTextoSped(num_acdraw, 20) & "|"
    s = s & CrLf
    
    
    Registro_C120 = s
        
    
End Function

Function Registro_C140(ind_emit As ind_emit, ind_tit As ind_tit, desc_tit As String, num_tit As String, qtd_parc As String, vl_tit As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C140|"
    s = s & EsticaTextoSped(CStr(ind_emit), 1) & "|"
    s = s & EsticaTextoSped(CStr(ind_tit), 2) & "|"
    s = s & desc_tit & "|"
    s = s & num_tit & "|"
    s = s & EsticaTextoSped(qtd_parc, 2) & "|"
    s = s & GravaValorSped(vl_tit) & "|"
    
    s = s & CrLf
    
    
    Registro_C140 = s
        
    
End Function

Function Registro_C160(COD_PART As String, veic_id As String, qtd_vol As String, peso_brt As String, peso_liq As String, uf_id As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C160|"
    s = s & EsticaTextoSped(COD_PART, 60) & "|"
    s = s & EsticaTextoSped(veic_id, 7) & "|"
    s = s & GravaValorSped(qtd_vol) & "|"
    s = s & GravaValorSped(peso_brt) & "|"
    s = s & GravaValorSped(peso_liq) & "|"
    s = s & EsticaTextoSped(uf_id, 2) & "|"
    s = s & CrLf
    
    
    Registro_C160 = s
        
    
End Function
Function Registro_C170(NUM_ITEM As String, COD_ITEM As String, desc_compl As String, QTD As String, UNID As String, _
    VL_ITEM As String, VL_DESC As String, Ind_mov As String, CST_ICMS As String, CFOP As String, cod_nat As String, _
    VL_BC_ICMS As String, ALIQ_ICMS As String, VL_ICMS As String, VL_BC_ICMS_ST As String, _
    ALIQ_ST As String, VL_ICMS_ST As String, IND_APUR As String, CST_IPI As String, cod_enq As String, _
    VL_BC_IPI As String, aliq_ipi As String, VL_IPI As String, cst_pis As String, vl_bc_pis As String, quant_bc_pis As String, _
    aliq_pis As String, VL_PIS As String, cst_cofins As String, vl_bc_cofins As String, aliq_cofins As String, _
    quant_bc_cofins As String, aliq_confins_r As String, VL_COFINS As String, COD_CTA As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C170|"
    s = s & EsticaTextoSped(NUM_ITEM, 3, 2, "0") & "|"
    s = s & EsticaTextoSped(COD_ITEM, 60) & "|"
    s = s & desc_compl & "|"
    s = s & GravaValorSped(QTD) & "|"
    s = s & EsticaTextoSped(UNID, 6) & "|"
    s = s & GravaValorSped(VL_ITEM) & "|"
    s = s & GravaValorSped(VL_DESC) & "|"
    s = s & EsticaTextoSped(Ind_mov, 1) & "|" ' movimentação fisica item produto :  0: sim 1: não
    s = s & EsticaTextoSped(CST_ICMS, 3) & "|"
    s = s & EsticaTextoSped(CFOP, 4) & "|"
    s = s & EsticaTextoSped(cod_nat, 10) & "|"
    
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & GravaValorSped(ALIQ_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    
    s = s & EsticaTextoSped(IND_APUR, 1) & "|"  '0 mensal 1 decenial
    s = s & EsticaTextoSped(CST_IPI, 2) & "|"
    s = s & EsticaTextoSped(cod_enq, 3) & "|"
    
    s = s & GravaValorSped(VL_BC_IPI) & "|"
    s = s & GravaValorSped(aliq_ipi) & "|"
    s = s & GravaValorSped(VL_IPI) & "|"
    
    s = s & EsticaTextoSped(cst_pis, 2) & "|"
    s = s & GravaValorSped(vl_bc_pis) & "|"
    s = s & GravaValorSped(aliq_pis) & "|"
    s = s & GravaValorSped(quant_bc_pis) & "|"
    s = s & GravaValorSped(aliq_pis) & "|"
    s = s & GravaValorSped(VL_PIS) & "|"
    
    s = s & EsticaTextoSped(cst_cofins, 2) & "|"
    s = s & GravaValorSped(vl_bc_cofins) & "|"
    s = s & GravaValorSped(aliq_cofins) & "|"
    s = s & GravaValorSped(quant_bc_cofins) & "|"
    s = s & GravaValorSped(aliq_cofins) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    
    
    s = s & EsticaTextoSped(COD_CTA, 2) & "|"
    s = s & CrLf
    
    
    Registro_C170 = s
        
    
End Function


Function Registro_C176(COD_MOD_ULT_E As String, NUM_DOC_ULT_E As String, SER_ULT_E As String, _
    DT_ULT_E As String, COD_PART_ULT_E As String, QUANT_ULT_E As String, VL_UNIT_ULT_E As String, _
    VL_UNIT_BC_ST As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C176|"
    s = s & EsticaTextoSped(COD_MOD_ULT_E, 2) & "|"
    s = s & EsticaTextoSped(NUM_DOC_ULT_E, 9) & "|"
    s = s & EsticaTextoSped(SER_ULT_E, 3) & "|"
    s = s & GravaDataSped(DT_ULT_E) & "|"
    s = s & EsticaTextoSped(COD_PART_ULT_E, 60) & "|"
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
    s = s & EsticaTextoSped(CST_ICMS, 3) & "|"
    s = s & EsticaTextoSped(CFOP, 4) & "|"
    s = s & gravaAliquotaSPED(ALIQ_ICMS, 2) & "|"
    s = s & GravaValorSped(VL_OPR) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_RED_BC) & "|"
    s = s & GravaValorSped(VL_IPI) & "|"
    
    s = s & EsticaTextoSped(COD_OBS, 6) & "|"
    
    s = s & CrLf
    
    
    Registro_C190 = s
        
    
End Function

Function Registro_C300(COD_MOD As String, SER As String, SUBSER As String, _
    NUM_DOC_INI As String, NUM_DOC_FIN As String, DT_DOC As String, VL_DOC As String, _
    VL_PIS As String, VL_COFINS As String, COD_CTA As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C300|"
    s = s & EsticaTextoSped(COD_MOD, 2) & "|"
    s = s & EsticaTextoSped(SER, 4) & "|"
    s = s & EsticaTextoSped(SUBSER, 3) & "|"
    s = s & GravaNumeroSped(NUM_DOC_INI, 6) & "|"
    s = s & GravaNumeroSped(NUM_DOC_FIN, 6) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & GravaValorSped(VL_DOC, 2) & "|"
    s = s & GravaValorSped(VL_PIS, 2) & "|"
    s = s & GravaValorSped(VL_COFINS, 2) & "|"
    s = s & EsticaTextoSped(COD_CTA, 0) & "|"
    s = s & CrLf
    
    
    Registro_C300 = s
        
    
End Function

Function Registro_C310(NUM_DOC_CANC As String) As String
    
    Dim s As String
    
    
    s = "|C310|"
    s = s & EsticaTextoSped(NUM_DOC_CANC, 6) & "|"
    
    
    
    Registro_C310 = s
        
    
End Function
Function Registro_C320(CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, _
    VL_OPR As String, VL_BC_ICMS As String, VL_ICMS As String, VL_RED_BC As String, _
    COD_OBS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C320|"
    s = s & EsticaTextoSped(CST_ICMS, 3) & "|"
    s = s & EsticaTextoSped(CFOP, 4) & "|"
    s = s & GravaNumeroSped(ALIQ_ICMS, 6) & "|"
    s = s & GravaValorSped(VL_OPR, 2) & "|"
    s = s & GravaValorSped(VL_BC_ICMS, 2) & "|"
    s = s & GravaValorSped(VL_ICMS, 2) & "|"
    s = s & GravaValorSped(VL_RED_BC, 2) & "|"
    s = s & EsticaTextoSped(COD_OBS, 0) & "|"
    s = s & CrLf
    
    
    Registro_C320 = s
        
    
End Function

Function Registro_C321(COD_ITEM As String, QTD As String, UNID As String, _
    VL_ITEM As String, VL_DESC As String, VL_BC_ICMS As String, VL_ICMS As String, _
    VL_PIS As String, VL_COFINS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C321|"
    s = s & EsticaTextoSped(COD_ITEM, 0) & "|"
    s = s & GravaValorSped(QTD, 3) & "|"
    s = s & EsticaTextoSped(UNID, 6) & "|"
    s = s & GravaValorSped(VL_ITEM, 2) & "|"
    s = s & GravaValorSped(VL_DESC, 2) & "|"
    s = s & GravaValorSped(VL_BC_ICMS, 2) & "|"
    s = s & GravaValorSped(VL_ICMS, 2) & "|"
    s = s & GravaValorSped(VL_PIS, 2) & "|"
    s = s & GravaValorSped(VL_COFINS, 2) & "|"
    s = s & CrLf
    
    
    Registro_C321 = s
        
    
End Function

Function Registro_C350(SER As String, SUB_SER As String, NUM_DOC As String, _
    DT_DOC As String, CNPJ_CPF As String, VL_MERC As String, VL_DOC As String, _
    VL_DESC As String, VL_PIS As String, VL_COFINS As String, COD_CTA As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C350|"
    s = s & EsticaTextoSped(SER, 3) & "|"
    s = s & EsticaTextoSped(SUB_SER, 3) & "|"
    s = s & EsticaTextoSped(NUM_DOC, 6, 2, "0") & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & EsticaTextoSped(CNPJ_CPF, 14) & "|"
    s = s & GravaValorSped(VL_MERC, 2) & "|"
    s = s & GravaValorSped(VL_DOC, 2) & "|"
    s = s & GravaValorSped(VL_DESC, 2) & "|"
    s = s & GravaValorSped(VL_PIS, 2) & "|"
    s = s & GravaValorSped(VL_COFINS, 2) & "|"
    s = s & EsticaTextoSped(COD_CTA, 14) & "|"
    s = s & CrLf
    
    
    Registro_C350 = s
        
    
End Function

Function Registro_C370(NUM_ITEM As String, COD_ITEM As String, QTD As String, _
    UNID As String, VL_ITEM As String, VL_DEC As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C370|"
    s = s & GravaNumeroSped(NUM_ITEM, 3) & "|"
    s = s & EsticaTextoSped(COD_ITEM, 0) & "|"
    s = s & GravaValorSped(QTD, 3) & "|"
    s = s & EsticaTexto(UNID, 6) & "|"
    s = s & GravaValorSped(VL_ITEM, 2) & "|"
    s = s & GravaValorSped(VL_DEC, 2) & "|"
    
    s = s & CrLf
    
    
    Registro_C370 = s
        
    
End Function

Function Registro_C390(CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, VL_OPR As String, _
VL_BC_ICMS As String, VL_ICMS As String, VL_RED_BC As String, COD_OBS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C390|"
    s = s & GravaNumeroSped(CST_ICMS, 3) & "|"
    s = s & GravaNumeroSped(CFOP, 4) & "|"
    s = s & GravaNumeroSped(ALIQ_ICMS, 6) & "|"
    s = s & GravaValorSped(VL_OPR, 2) & "|"
    s = s & GravaValorSped(VL_BC_ICMS, 2) & "|"
    s = s & GravaValorSped(VL_ICMS, 2) & "|"
    s = s & GravaValorSped(VL_RED_BC, 2) & "|"
    s = s & EsticaTexto(COD_OBS) & "|"
    
    s = s & CrLf
    
    
    Registro_C390 = s
        
    
End Function


Function Registro_C400(COD_MOD As String, ECF_MOD As String, ECF_FAB As String, _
    ECF_CX As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C400|"
    s = s & EsticaTextoSped(COD_MOD, 2) & "|"
    s = s & EsticaTextoSped(ECF_MOD, 20) & "|"
    s = s & EsticaTextoSped(ECF_FAB, 21) & "|"
    s = s & EsticaTextoSped(ECF_CX, 3) & "|"
    
    
    s = s & CrLf
    
    
    Registro_C400 = s
        
    
End Function

Function Registro_C405(DT_DOC As String, CRO As String, CRZ As String, _
    NUM_COO_FIN As String, GT_FIN As String, VL_BRT As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C400|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & EsticaTextoSped(CRO, 3, 2, "0") & "|"
    s = s & EsticaTextoSped(CRZ, 3, 2, "0") & "|"
    s = s & EsticaTextoSped(NUM_COO_FIN, 6, 2, "0") & "|"
    s = s & GravaValorSped(GT_FIN) & "|"
    s = s & GravaValorSped(VL_BRT) & "|"
    
    
    s = s & CrLf
    
    
    Registro_C405 = s
        
    
End Function
Function Registro_C410(VL_PIS As String, VL_COFINS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C410|"
    
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    
    s = s & CrLf
    
    
    Registro_C410 = s
        
    
End Function


Function Registro_C420(COD_TOT_PAR As String, VLR_ACUM_TOT As String, NR_TOT As String, _
    CRZ As String, DESCR_NR_TOT As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C420|"
    s = s & EsticaTextoSped(COD_TOT_PAR, 7) & "|"
    s = s & GravaValorSped(VLR_ACUM_TOT) & "|"
    s = s & EsticaTextoSped(NR_TOT, 2, 2, "0") & "|"
    s = s & EsticaTextoSped(CRZ, 3, 2, "0") & "|"
    s = s & DESCR_NR_TOT & "|"
    
    
    
    s = s & CrLf
    
    
    Registro_C420 = s
        
    
End Function
Function Registro_C425(COD_ITEM As String, QTD As String, UNID As String, _
    VL_ITEM As String, VL_PIS As String, VL_COFINS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C425|"
    s = s & EsticaTextoSped(COD_ITEM, 60) & "|"
    's = s & EsticaTextoSped(NR_TOT, 2, 2, "0") & "|"
    s = s & GravaNumeroSped(QTD, 3) & "|"
    s = s & EsticaTextoSped(UNID, 6) & "|"
    s = s & GravaValorSped(VL_ITEM) & "|"
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    
    
    s = s & CrLf
    
    
    Registro_C425 = s
        
    
End Function

Function Registro_C460(COD_MOD As String, COD_SIT As String, NUM_DOC As String, _
    DT_DOC As String, VL_DOC As String, VL_PIS As String, VL_COFINS As String, CPFCNPJ As String, _
    NOME_ADQ As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C425|"
    s = s & EsticaTextoSped(COD_MOD, 2) & "|"
    s = s & EsticaTextoSped(COD_SIT, 2, 2, "0") & "|"
    s = s & EsticaTextoSped(NUM_DOC, 2, 2, "0") & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & GravaValorSped(VL_DOC) & "|"
    
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    s = s & EsticaTextoSped(CPFCNPJ, 14) & "|"
    s = s & EsticaTextoSped(NOME_ADQ, 60) & "|"
    s = s & CrLf
    
    
    Registro_C460 = s
        
    
End Function
Function Registro_C470(COD_ITEM As String, QTD As String, QTD_CANC As String, _
    UNID As String, VL_ITEM As String, CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, _
    VL_PIS As String, VL_COFINS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C470|"
    s = s & EsticaTextoSped(COD_ITEM, 60) & "|"
    s = s & GravaNumeroSped(QTD, 3) & "|"
    s = s & GravaNumeroSped(QTD_CANC, 3) & "|"
    s = s & EsticaTextoSped(UNID, 6) & "|"
    s = s & GravaValorSped(VL_ITEM) & "|"
    s = s & GravaValorSped(CST_ICMS) & "|"
    s = s & EsticaTextoSped(CFOP, 4) & "|"
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
    s = s & EsticaTextoSped(COD_OBS, 6) & "|"
    
    
    s = s & CrLf
    
    
    Registro_C490 = s
        
    
End Function
Function Registro_C495(ALIQ_ICMS As String, COD_ITEM As String, QTD As String, QTD_CANC As String, UNID As String, _
    VL_ITEM As String, VL_DESC As String, VL_CANC As String, VL_ACMO As String, VL_BC_ICMS As String, _
    VL_ICMS As String, VL_ISEN As String, VL_NT As String, VL_ICMS_ST As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C495|"
    s = s & gravaAliquotaSPED(ALIQ_ICMS) & "|"
    s = s & EsticaTextoSped(COD_ITEM) & "|"
    s = s & GravaValorSped(QTD) & "|"
    s = s & EsticaTextoSped(UNID) & "|"
    s = s & GravaValorSped(VL_ITEM) & "|"
    s = s & GravaValorSped(VL_DESC) & "|"
    s = s & GravaValorSped(VL_CANC) & "|"
    s = s & GravaValorSped(VL_ACMO) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_ISEN) & "|"
    s = s & GravaValorSped(VL_NT) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    
    
    
    s = s & CrLf
    
    
    Registro_C495 = s
        
    
End Function

Function Registro_C500(IND_OPER As IND_OPER, ind_emit As ind_emit, COD_PART As String, COD_MOD As String, COD_SIT As String, _
    SER As String, SUBSER As String, COD_CONS As String, NUM_DOC As String, DT_DOC As String, _
    DT_E_S As String, VL_DOC As String, VL_DESC As String, VL_FORN As String, VL_SERV_NT As String, _
    VL_TERC As String, VL_DA As String, VL_BC_ICMS As String, VL_ICMS As String, VL_BC_ICMS_ST As String, _
    VL_ICMS_ST As String, COD_INF As String, VL_PIS As String, VL_COFINS As String, _
    TP_LIGACAO As String, cod_gru_tensao As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C500|"
    s = s & EsticaTexto(CStr(IND_OPER), 1) & "|"
    s = s & EsticaTexto(CStr(ind_emit), 1) & "|"
    s = s & EsticaTexto(COD_PART) & "|"
    s = s & EsticaTexto(COD_MOD) & "|"
    s = s & EsticaTexto(COD_SIT) & "|"
    s = s & GravaNumeroSped(SER, 4) & "|"
    s = s & GravaNumeroSped(SUBSER, 3) & "|"
    s = s & GravaNumeroSped(COD_CONS, 2) & "|"
    s = s & GravaNumeroSped(NUM_DOC, 9) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & GravaDataSped(DT_E_S) & "|"
    s = s & GravaValorSped(VL_DOC) & "|"
    s = s & GravaValorSped(VL_FORN) & "|"
    s = s & GravaValorSped(VL_SERV_NT) & "|"
    s = s & GravaValorSped(VL_TERC) & "|"
    s = s & GravaValorSped(VL_DA) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    s = s & EsticaTexto(COD_INF) & "|"
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    s = s & EsticaTexto(TP_LIGACAO) & "|"
    s = s & EsticaTexto(cod_gru_tensao) & "|"
    
    
    Registro_C500 = s
        
    
End Function
Function Registro_C510(NUM_ITEM As String, COD_ITEM As String, COD_CLASS As String, QTD As String, UNID As String, _
    VL_ITEM As String, VL_DESC As String, CST_ICMS As String, CFOP As String, VL_BC_ICMS As String, _
    ALIQ_ICMS As String, VL_ICMS As String, VL_BC_ICMS_ST As String, ALIQ_ST As String, VL_ICMS_ST As String, _
    IND_REC As String, COD_PART As String, VL_PIS As String, VL_COFINS As String, COD_CTA As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C510|"
    s = s & EsticaTexto(NUM_ITEM, 3) & "|"
    s = s & EsticaTexto(COD_ITEM, 60) & "|"
    s = s & GravaNumeroSped(COD_CLASS, 4) & "|"
    s = s & GravaValorSped(QTD, 3) & "|"
    s = s & EsticaTexto(UNID) & "|"
    s = s & GravaValorSped(VL_ITEM) & "|"
    s = s & GravaValorSped(VL_DESC) & "|"
    s = s & EsticaTexto(CST_ICMS) & "|"
    s = s & EsticaTexto(CFOP) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & EsticaTexto(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & EsticaTexto(ALIQ_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    s = s & EsticaTexto(IND_REC) & "|"
    s = s & EsticaTexto(COD_PART) & "|"
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    s = s & EsticaTexto(COD_CTA) & "|"
    Registro_C510 = s
        
    
End Function
Function Registro_C590(CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, VL_OPR As String, _
 VL_BC_ICMS As String, VL_ICMS As String, VL_BC_ICMS_ST As String, VL_ICMS_ST As String, _
    VL_RED_BC As String, COD_OBS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C590|"
    s = s & GravaNumeroSped(CST_ICMS, 3) & "|"
    s = s & GravaNumeroSped(CFOP, 4) & "|"
    s = s & GravaValorSped(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_OPR) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_RED_BC) & "|"
    
    
    s = s & EsticaTexto(COD_OBS) & "|"
    Registro_C590 = s
        
    
End Function
Function Registro_C600(COD_MOD As String, COD_MUN As String, SER As String, SUBSER As String, _
 COD_CONS As String, QTD_CONS As String, QTD_CANC As String, DT_DOC As String, _
    VL_DOC As String, VL_DESC As String, CONS As String, VL_FORN As String, VL_SERV_NT As String, _
    VL_TERC As String, VL_DA As String, VL_BC_ICMS As String, VL_ICMS As String, _
    VL_BC_ICMS_ST As String, VL_ICMS_ST As String, VL_PIS As String, VL_COFINS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C600|"
    s = s & GravaNumeroSped(COD_MOD, 2) & "|"
    s = s & GravaNumeroSped(COD_MUN, 7) & "|"
    s = s & EsticaTexto(SER, 4) & "|"
    s = s & EsticaTexto(SUBSER, 3) & "|"
    s = s & EsticaTexto(COD_CONS, 2) & "|" '01 comercial
    
    
    s = s & GravaValorSped(QTD_CONS) & "|"
    s = s & GravaValorSped(QTD_CANC) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    
    s = s & GravaValorSped(VL_DOC) & "|"
    s = s & GravaValorSped(VL_DESC) & "|"
    s = s & GravaNumeroSped(CONS, 5) & "|"
    s = s & GravaValorSped(VL_FORN) & "|"
    s = s & GravaValorSped(VL_SERV_NT) & "|"
    s = s & GravaValorSped(VL_TERC) & "|"
    s = s & GravaValorSped(VL_DA) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    
    
    
    Registro_C600 = s
        
    
End Function
Function Registro_C601(NUM_DOC_CANC As String) As String
    
    Dim s As String
    
   
    s = "|C601|"
    s = s & GravaNumeroSped(NUM_DOC_CANC, 9) & "|"
    
    Registro_C601 = s
        
    
End Function

Function Registro_C610(COD_CLASS As String, COD_ITEM As String, QTD As String, _
    UNID As String, VL_ITEM As String, VL_DESC As String, CST_ICMS As String, CFOP As String, _
    ALIQ_ICMS As String, VL_BC_ICMS As String, VL_ICMS As String, _
    VL_BC_ICMS_ST As String, VL_ICMS_ST As String, VL_PIS As String, VL_COFINS As String, COD_CTA As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C610|"
    s = s & GravaNumeroSped(COD_CLASS, 2) & "|"
    s = s & EsticaTexto(COD_ITEM) & "|"
    s = s & GravaValorSped(QTD) & "|"
    s = s & EsticaTexto(UNID) & "|"
    s = s & GravaValorSped(VL_ITEM) & "|"
    s = s & GravaValorSped(VL_DESC) & "|"
    s = s & EsticaTexto(CST_ICMS) & "|"
    s = s & EsticaTexto(CFOP) & "|"
    s = s & GravaValorSped(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_PIS) & "|"
    s = s & GravaValorSped(VL_COFINS) & "|"
    s = s & EsticaTexto(COD_CTA, 2) & "|"
       
    Registro_C610 = s
        
    
End Function

Function Registro_C690_(CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, _
VL_OPR As String, VL_BC_ICMS As String, VL_ICMS As String, VL_BC_ICMS_UF As String, _
VL_ICMS_UF As String, VL_RED_BC As String, COD_OBS As String) As String
    Dim s As String
    
    s = "|C690|"
    s = s & GravaNumeroSped(CST_ICMS, 3) & "|"
    s = s & GravaNumeroSped(CFOP, 4) & "|"
    s = s & GravaValorSped(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_OPR) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_UF) & "|"
    s = s & GravaValorSped(VL_ICMS_UF) & "|"
    s = s & GravaValorSped(VL_RED_BC) & "|"
    s = s & EsticaTextoSped(COD_OBS, 6) & "|"
    s = s & CrLf
    Registro_C690_ = s
        
    
End Function
Function Registro_C700(COD_MOD As String, SER As String, _
NRO_ORD_INI As String, NRO_ORD_FIN As String, DT_DOC_INI As String, DT_DOC_FIN As String, _
NOM_MEST As String, CHV_COD_DIG As String) As String
    Dim s As String
    
    s = "|C700|"
    s = s & EsticaTexto(COD_MOD) & "|"
    s = s & EsticaTexto(SER) & "|"
    s = s & GravaNumeroSped(NRO_ORD_INI, 9) & "|"
    s = s & GravaNumeroSped(NRO_ORD_FIN, 9) & "|"
    s = s & GravaDataSped(DT_DOC_INI) & "|"
    s = s & GravaDataSped(DT_DOC_FIN) & "|"
    s = s & EsticaTexto(NOM_MEST) & "|"
    s = s & EsticaTexto(CHV_COD_DIG) & "|"
    
    s = s & CrLf
    Registro_C700 = s
        
    
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


Function Registro_D100(IND_OPERA As IND_OPER, ind_emit As ind_emit, COD_PART As String, COD_MOD As String, _
    COD_SIT As String, SER As String, SUBSER As String, NUM_DOC As String, CHV_CTE As String, DT_DOC As String, _
    DT_AQUISICAO_PRESTACAO As String, TP_CTe As String, CHV_CTE_REF As String, VL_DOC As String, _
    VL_DESC As String, IND_FRE As IND_FRT, vl_SERV As String, VL_BC_ICMS As String, _
    VL_ICMS As String, VL_NT As String, COD_INF As String, COD_CTA As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|D100|"
    s = s & EsticaTextoSped(CStr(IND_OPERA), 1) & "|"
    s = s & EsticaTextoSped(CStr(ind_emit), 1) & "|"
    s = s & EsticaTextoSped(COD_PART, 60) & "|"
    s = s & EsticaTextoSped(COD_MOD, 2) & "|"
    s = s & GravaNumeroSped(COD_SIT, 2) & "|"
    s = s & EsticaTextoSped(SER, 4) & "|"
    s = s & EsticaTextoSped(SUBSER, 3) & "|"
    s = s & EsticaTextoSped(NUM_DOC, 3) & "|"
    s = s & EsticaTextoSped(CHV_CTE, 44) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & GravaDataSped(DT_AQUISICAO_PRESTACAO) & "|"
    s = s & EsticaTextoSped(TP_CTe, 1) & "|"
    s = s & EsticaTextoSped(CHV_CTE_REF, 44) & "|"
    s = s & GravaValorSped(VL_DOC) & "|"
    s = s & GravaValorSped(VL_DESC) & "|"
    s = s & EsticaTextoSped(CStr(IND_FRE), 1) & "|"
    s = s & GravaValorSped(vl_SERV) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_NT) & "|"
       
    s = s & EsticaTextoSped(COD_INF, 6) & "|"
    s = s & COD_CTA & "|"
    
    s = s & CrLf
    
    
    Registro_D100 = s
        
    
End Function


Function Registro_D110(NUM_ITEN As String, COD_ITEM As String, vl_SERV As String, VL_OUT As String) As String
    'cod_inf Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
    Dim s As String
    s = "|D110|"
    s = s & EsticaTextoSped(NUM_ITEN, 4) & "|"
    s = s & EsticaTextoSped(COD_ITEM, 60) & "|"
    s = s & GravaValorSped(vl_SERV) & "|"
    s = s & GravaValorSped(VL_OUT) & "|"
    
    s = s & CrLf
    Registro_D110 = s
        
    
End Function




Function Registro_D120(COD_MUN_ORIG As String, COD_MUN_DEST As String, placa As String, UF_PLACA As String) As String
    Dim s As String
    
    s = "|D120|"
    s = s & EsticaTextoSped(COD_MUN_ORIG, 7) & "|"
    s = s & EsticaTextoSped(COD_MUN_DEST, 7) & "|"
    s = s & EsticaTextoSped(placa, 7) & "|"
    s = s & EsticaTextoSped(UF_PLACA, 2) & "|"
    s = s & CrLf
    Registro_D120 = s
        
    
End Function

Function Registro_D130(COD_PART As String, COD_REDESPACHADO As String, IND_FRT_RED As String, _
COD_MUN_ORIG As String, COD_MUN_DEST As String, placa As String, VL_LIQ_FRE As String, _
VL_SEC_CAT As String, VL_DESP As String, VL_PEDAGIOS As String, VL_OUTROS As String, _
VL_TOTAL_FRT As String, UF_VEICULO As String) As String
    Dim s As String
    
    s = "|D130|"
    s = s & EsticaTextoSped(COD_PART, 60) & "|"
    s = s & EsticaTextoSped(COD_REDESPACHADO, 60) & "|"
    s = s & EsticaTextoSped(IND_FRT_RED, 1) & "|"  '0 sem redespacho 1- emiente 2-destinatario 9-outros
    s = s & EsticaTextoSped(COD_MUN_ORIG, 7) & "|"
    s = s & EsticaTextoSped(COD_MUN_DEST, 7) & "|"
    s = s & EsticaTextoSped(placa, 7) & "|"
    s = s & GravaValorSped(VL_LIQ_FRE, 2) & "|"
    s = s & GravaValorSped(VL_SEC_CAT, 2) & "|"
    s = s & GravaValorSped(VL_DESP, 2) & "|"
    s = s & GravaValorSped(VL_PEDAGIOS, 2) & "|"
    s = s & GravaValorSped(VL_OUTROS, 2) & "|"
    s = s & GravaValorSped(VL_TOTAL_FRT, 2) & "|"
    s = s & EsticaTextoSped(UF_VEICULO, 2) & "|"
    s = s & CrLf
    Registro_D130 = s
        
    
End Function

Function Registro_D160(DESPACHO As String, CNPJ_REMETENTE As String, IE_REMET As String, _
COD_MUN_ORIGEN As String, CNPJ_DESTINO As String, IE_DESTINO As String, COD_MUN_DEST As String) As String
    Dim s As String
    
    s = "|D160|"
    s = s & DESPACHO & "|"
    s = s & EsticaTextoSped(CNPJ_REMETENTE, 14) & "|"
    s = s & EsticaTextoSped(IE_REMET, 14) & "|"
    s = s & EsticaTextoSped(COD_MUN_ORIGEN, 7) & "|"
    s = s & EsticaTextoSped(CNPJ_DESTINO, 14) & "|"
    s = s & EsticaTextoSped(IE_DESTINO, 14) & "|"
    s = s & EsticaTextoSped(COD_MUN_DEST, 7) & "|"
    s = s & CrLf
    Registro_D160 = s
        
    
End Function

Function Registro_D161(IND_CARG As IND_CARGA, CNPJ_COLETA As String, IE_COLETA As String, _
COD_MUN_COLETA As String, CNPJ_ENTREGA As String, IE_ENTREGA As String, COD_MUN_ENTREGA As String) As String
    Dim s As String
    
    s = "|D161|"
    s = s & IND_CARG & "|"
    s = s & EsticaTextoSped(CNPJ_COLETA, 14) & "|"
    s = s & EsticaTextoSped(IE_COLETA, 14) & "|"
    s = s & EsticaTextoSped(COD_MUN_COLETA, 7) & "|"
    s = s & EsticaTextoSped(CNPJ_ENTREGA, 14) & "|"
    s = s & EsticaTextoSped(IE_ENTREGA, 14) & "|"
    s = s & EsticaTextoSped(COD_MUN_ENTREGA, 7) & "|"
    s = s & CrLf
    Registro_D161 = s
        
    
End Function


Function Registro_D500(IND_OPERACAO As IND_OPER, IND_EMITEnte As ind_emit, COD_PART As String, _
COD_MOD As String, COD_SIT As String, SERIE As String, SUBSER As String, NUM_DOC As String, DT_DOC As String, _
DT_ENTRADA_SAIDA As String, VL_DOC As String, VL_DESC As String, vl_SERV As String, VL_SERV_NT As String, _
VL_TERC As String, VL_DA As String, VL_BC_ICMS As String, VL_ICMS As String, COD_INF As String, VL_PIS As String, _
VL_COFINS As String, COD_CTA As String, TIPO_ASSINANTE As TP_ASSINANTE) As String
    Dim s As String
    
    s = "|D500|"
    s = s & IND_OPERACAO & "|"
    s = s & IND_EMITEnte & "|"
    s = s & EsticaTextoSped(COD_PART, 60) & "|"
    s = s & GravaNumeroSped(COD_MOD, 2) & "|"
    s = s & GravaNumeroSped(COD_SIT, 2) & "|"
    s = s & EsticaTextoSped(SERIE, 4) & "|"
    s = s & EsticaTextoSped(SUBSER, 3) & "|"
    s = s & GravaNumeroSped(NUM_DOC, 9) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & GravaDataSped(DT_ENTRADA_SAIDA) & "|"
    
    s = s & GravaValorSped(VL_DOC, 2) & "|"
    s = s & GravaValorSped(VL_DESC, 2) & "|"
    s = s & GravaValorSped(vl_SERV, 2) & "|"
    s = s & GravaValorSped(VL_SERV_NT, 2) & "|"
    s = s & GravaValorSped(VL_TERC, 2) & "|"
    s = s & GravaValorSped(VL_DA, 2) & "|"
    s = s & GravaValorSped(VL_BC_ICMS, 2) & "|"
    s = s & GravaValorSped(VL_ICMS, 2) & "|"
    s = s & EsticaTextoSped(COD_INF, 6) & "|"
    s = s & GravaValorSped(VL_PIS, 2) & "|"
    s = s & GravaValorSped(VL_COFINS, 2) & "|"
    s = s & COD_CTA & "|"
    
    s = s & EsticaTextoSped(CStr(TIPO_ASSINANTE), 1) & "|"
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
    s = s & EsticaTextoSped(COD_OBS, 6) & "|"
    s = s & CrLf
    Registro_D590 = s
        
    
End Function

Function Registro_D610(COD_CLASS As String, COD_ITEM As String, QTD As String, _
UNID As String, VL_ITEM As String, VL_DESC As String, CST_ICMS As String, _
CFOP As String, ALIQ_ICMS As String, VL_BC_ICMS As String, VL_ICMS As String, _
VL_BC_ICMS_UF As String, VL_ICMS_UF As String, VL_RED_BC As String, _
VL_PIS As String, VL_COFINS As String, COD_CTA As String) As String
    Dim s As String
    
    s = "|D610|"
    s = s & EsticaTextoSped(COD_CLASS, 4) & "|"
    s = s & EsticaTextoSped(COD_ITEM, 60) & "|"
    s = s & GravaNumeroSped(QTD, 3) & "|"
    s = s & EsticaTextoSped(UNID, 6) & "|"
    s = s & GravaNumeroSped(CFOP, 4) & "|"
    s = s & GravaValorSped(VL_ITEM) & "|"
    s = s & GravaValorSped(VL_DESC) & "|"
    s = s & EsticaTextoSped(CST_ICMS, 3) & "|"
    s = s & EsticaTextoSped(CFOP, 4) & "|"
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
    s = s & EsticaTextoSped(CST_ICMS, 3) & "|"
    s = s & EsticaTextoSped(CFOP, 4) & "|"
    s = s & gravaAliquotaSPED(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_OPR) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_ICMS_ST) & "|"
    s = s & GravaValorSped(VL_RED_BC) & "|"
    's = s & GravaValorSped(VL_IPI) & "|"
    s = s & EsticaTextoSped(COD_OBS, 6) & "|"
    s = s & CrLf
    
    
    Registro_C690 = s
        
    
End Function

Function Registro_D696(CST_ICMS As String, CFOP As String, ALIQ_ICMS As String, _
    VL_OPR As String, VL_BC_ICMS As String, VL_ICMS As String, VL_BC_ICMS_UF As String, _
    VL_ICMS_UF As String, VL_RED_BC As String, COD_OBS As String) As String
    
    Dim s As String
    
    'ind_emit 0 = declaracao de importação 1 declraçaõ simplificada de importação
    s = "|C696|"
    s = s & EsticaTextoSped(CST_ICMS, 3) & "|"
    s = s & EsticaTextoSped(CFOP, 4) & "|"
    s = s & gravaAliquotaSPED(ALIQ_ICMS) & "|"
    s = s & GravaValorSped(VL_OPR) & "|"
    s = s & GravaValorSped(VL_BC_ICMS) & "|"
    s = s & GravaValorSped(VL_ICMS) & "|"
    s = s & GravaValorSped(VL_BC_ICMS_UF) & "|"
    s = s & GravaValorSped(VL_ICMS_UF) & "|"
    s = s & GravaValorSped(VL_RED_BC) & "|"
    s = s & EsticaTextoSped(COD_OBS, 6) & "|"
    
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
    s = s & EsticaTextoSped(COD_AJ_APUR, 8) & "|"
    s = s & DESCR_COMPL_AJ & "|"
    s = s & GravaValorSped(vl_AJ_APUR) & "|"
    s = s & CrLf
    Registro_E111 = s
End Function

Function Registro_E112(NUM_DA As String, NUM_PROC As String, IND_PROC As String) As String
    Dim s As String
    s = "|E112|"
    s = s & NUM_DA & "|"
    s = s & EsticaTextoSped(NUM_PROC, 15) & "|"
    s = s & EsticaTextoSped(IND_PROC, 1) & "|" '0 sefaz 1 federal 2 estado 9 outros
    s = s & CrLf
    Registro_E112 = s
End Function

Function Registro_E113(COD_PART As String, COD_MOD As String, SERIE As String, SUBSER As String, _
NUM_DOC As String, DT_DOC As String, COD_ITEM As String, VL_AJUSTE_ITEM As String) As String
    Dim s As String
    s = "|E113|"
    s = s & EsticaTextoSped(COD_PART, 60) & "|"
    s = s & EsticaTextoSped(COD_MOD, 2) & "|"
    s = s & EsticaTextoSped(SERIE, 4) & "|"
    s = s & EsticaTextoSped(SUBSER, 3) & "|"
    s = s & EsticaTextoSped(NUM_DOC, 9) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & EsticaTextoSped(COD_ITEM, 60) & "|"
    
    s = s & GravaValorSped(VL_AJUSTE_ITEM) & "|"
    s = s & CrLf
    Registro_E113 = s
End Function
Function Registro_E115(COD_INF_ADIC As String, VL_INF_ADIC As String, DESC_COMPL_AJ As String) As String
    Dim s As String
    s = "|E115|"
    s = s & EsticaTextoSped(COD_INF_ADIC, 8) & "|"
    
    s = s & DESC_COMPL_AJ & "|"
    
    s = s & CrLf
    Registro_E115 = s
End Function



Function Registro_E116(COD_OR As String, VL_OR As String, DT_VCTO As String, _
COD_REC As String, NUM_PROC As String, IND_PROC As String, PROC As String, txt_COMPL As String, _
MES_REF As String) As String
    Dim s As String
    s = "|E116|"
    s = s & EsticaTextoSped(COD_OR, 4) & "|"
    s = s & GravaValorSped(VL_OR) & "|"
    s = s & GravaDataSped(DT_VCTO) & "|"
    s = s & COD_REC & "|"
    s = s & EsticaTextoSped(NUM_PROC, 15) & "|"
    s = s & EsticaTextoSped(IND_PROC, 1) & "|" '0 sefaz 1 federal 2 estado 9 outros
    s = s & txt_COMPL & "|"
    s = s & MES_REF & "|" 'MMAAAA
    s = s & CrLf
    Registro_E116 = s
End Function



Function Registro_E200(uf As String, DT_INI As String, DT_FIN As String, cod_ant_item As String) As String
Dim s As String
    s = "|E200|"
    s = s & uf & "|"
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
    s = s & EsticaTextoSped(Ind_mov, 1) & "|" '0 sem operação st 1: com operação st
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
    s = s & EsticaTextoSped(COD_AJ_APUR, 8) & "|" 'tabela 5.1.1
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
    s = s & EsticaTextoSped(NUM_PROC, 15) & "|"
    s = s & EsticaTextoSped(IND_PROC, 1) & "|" '0 sefaz 1 federal 2 estadual 9 outros
    s = s & RESUMO_PROCESSO & "|"
    s = s & txt_COMPL & "|"
    s = s & CrLf
    Registro_E230 = s
End Function

Function Registro_E240(COD_PART As String, COD_MOD As String, SERIE As String, SUBSER As String, _
NUM_DOC As String, DT_DOC As String, COD_ITEM As String, VL_AJUSTE_ITEM As String) As String
    Dim s As String
    s = "|E240|"
    s = s & EsticaTextoSped(COD_PART, 60) & "|"
    s = s & EsticaTextoSped(COD_MOD, 2) & "|"
    s = s & EsticaTextoSped(SERIE, 4) & "|"
    s = s & EsticaTextoSped(SUBSER, 3) & "|"
    s = s & EsticaTextoSped(NUM_DOC, 9) & "|"
    s = s & GravaDataSped(DT_DOC) & "|"
    s = s & EsticaTextoSped(COD_ITEM, 60) & "|"
    
    s = s & GravaValorSped(VL_AJUSTE_ITEM) & "|"
    s = s & CrLf
    Registro_E240 = s
End Function


Function Registro_E250(COD_OR As String, VL_OR As String, DT_VCTO As String, _
COD_REC As String, NUM_PROC As String, IND_PROC As String, PROC As String, txt_COMPL As String, _
MES_REF As String) As String
    Dim s As String
    s = "|E116|"
    s = s & EsticaTextoSped(COD_OR, 4) & "|"
    s = s & GravaValorSped(VL_OR) & "|"
    s = s & GravaDataSped(DT_VCTO) & "|"
    s = s & COD_REC & "|"
    s = s & EsticaTextoSped(NUM_PROC, 15) & "|"
    s = s & EsticaTextoSped(IND_PROC, 1) & "|" '0 sefaz 1 federal 2 estado 9 outros
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
    s = s & EsticaTextoSped(CFOP, 4) & "|" 'SINIEF 07/01.
    s = s & EsticaTextoSped(CST_IPI, 2) & "|"  'tabela 4.3.2
    
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
    s = s & EsticaTextoSped(IND_AJ, 1) & "|"  '0 deb 1 cred
    s = s & GravaValorSped(VL_AJ) & "|"
    s = s & EsticaTextoSped(COD_AJ, 3) & "|"  'tabwla 4.5.4
    s = s & EsticaTextoSped(IND_DOC, 1) & "|"  '0 proc judicial 1:proc adm 2:per/dcomp 9:outros
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
    VL_ITEM As String, IND_PROP As String, COD_PART As String, txtCompl As String, COD_CTA As String) As String
    Dim s As String
    s = "|H010|"
    s = s & EsticaTextoSped(COD_ITEM, 60) & "|"
    s = s & EsticaTextoSped(UNID, 6) & "|"
    s = s & GravaValorSped(QTD, 3) & "|"
    s = s & txtCompl & "|"
    s = s & COD_CTA & "|"
    s = s & CrLf
    Registro_H010 = s
End Function


Function Registro_H020(CST_ICMS As String, BC_ICMS As String, VL_ICMS As String) As String
    Dim s As String
    s = "|H020|"
    s = s & EsticaTextoSped(CST_ICMS, 3) & "|"
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
    s = s & EsticaTextoSped(COD_PART, 60) & "|"
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
    s = s & EsticaTextoSped(REG_BLC, 4) & "|"
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

Function EsticaTextoSped(texto As String, Optional Tam As Integer = 0, Optional align As Integer = 0, _
Optional Carac As String = " ", Optional soNum As Boolean = False) As String
    If soNum Then
        texto = F_SoNumero(texto)
    End If
    
    If Tam = 0 Then
        Tam = Len(texto)
    End If
    If Len(texto) >= Tam Then
        EsticaTextoSped = Left(texto, Tam)
    
    Else
        If align = 0 Then
            EsticaTextoSped = texto & String(Tam - Len(texto), Carac)
        ElseIf align = 1 Then
            EsticaTextoSped = String((Tam - Len(texto)) / 2, Carac) & texto & String((Tam - Len(texto)) / 2, Carac)
        Else
            EsticaTextoSped = String(Tam - Len(texto), Carac) & texto
        End If
    End If
    EsticaTextoSped = Trim(EsticaTextoSped)
End Function

