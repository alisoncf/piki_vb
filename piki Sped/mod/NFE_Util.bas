Attribute VB_Name = "NFEUtil"
Global fusoHorario As String
Public sVersaoNfe As String
Public OpcaoEmpresaCRT As CRT

Enum StatusTransacaoNFE
    em_aberto = 0
    concluida = 1
    validada = 2
    assinada = 3
    transmitida = 4
    impressa = 5
    cancelada = 9
    
End Enum

Public Enum modFrete
    por_conta_do_emitente = 0
    por_conta_do_destinatario_terceiro = 1
    por_conta_de_terceiro = 2
    sem_frete = 9
End Enum
Public Enum indPag
    avista = 0
    aprazo = 1
    Outros = 2
End Enum
Enum indIEDest
    contribuinte_icms = 1
    isento = 2 'se for este caso, a tag IE do cliente não deve ser informada
End Enum
Public Enum tipo_N
    N02 = 2
    N03 = 3
    N04 = 4
    N05 = 5
    N06 = 6
    N07 = 7
    N08 = 8
    N09 = 9
    N10 = 10
    N10a = 11
    N10b = 12
    N10c = 13
    N10d = 14
    N10e = 15
    N10f = 16
    N10g = 17
    N10h = 18
End Enum
Public Enum CRT
    SIMPLES_NACIONAL = 1
    SIMP_NAC_EXCESSO_SUBLIMITE_RECEITA_BRUTA = 2
    RECEITA_NORMAL = 3
End Enum

Public Enum tpAmb
    producao = 1
    homologacao = 2
End Enum
Public Enum tpNf
    Entrada = 0
    saida = 1
End Enum

Public Enum tPag
    Dinheiro = 1
    cheque = 2
    Cartao_de_Credito = 3
    Cartao_de_Debito = 4
    Credito_Loja = 5
    Vale_Alimentacao = 10
    Vale_Refeicao = 11
    Vale_Presente = 12
    Vale_Combustivel = 13
    Outros = 99
End Enum


Public Enum tpImp
'Formato de impressão do DANFE (0-sem DANFE;1-DANFe Retrato; 2-DANFe Paisagem;3-DANFe Simplificado; 4-DANFe NFC-e;5-DANFe NFC-e em mensagem eletrônica
    sem_danfe = 0
    danf_retrato = 1
    danf_paisagem = 2
    danf_simplificado = 3
    danf_nfc_e = 4
    danf_nfc_e_msg_eletronica = 5
End Enum




Public Enum tpEmis
    Normal = 1
    Contingência_FS = 2
    Contingência_SCAN = 3
    Contingência_DPEC = 4
    Contingência_FS_DA = 5
End Enum
Public Enum orig
    nacional = 0
    estrangeira_importacao_direta = 1
    estrangeira_mercado_interno = 2
    nacional_com_Conteudo_impor_maior_40_menor_70 = 3
    nacional_com_producao_confirmidade_legislacao_ajuste = 4
    nacional_com_Conteudo_impor_inferior_40 = 5
    entrangeira_importacao_direta_sem_similar_nacional = 6
    entrangeira_compra_interna_sem_similar_nacional_listaCamex_e_GasNatural = 7
    nacional_mercadoria_ou_bem_comConteudoImportadoSuperior_70 = 8
    
End Enum
Public Enum modBCST
    PrecoTabeladoOuMaximo = 0
    Lista_Negativa = 1
    Lista_Positiva = 2
    Lista_Neutra = 3
    Margem_Valor_Agregado_perct = 4
    Pauta = 5
End Enum
Public Enum modBC
    Margem_Valor_Agregado_perc = 0
    Pauta_valor = 1
    Preco_Tabelado_Max_valor = 2
    valor_da_operacao = 3
End Enum

Public Enum TfinNfe
    Nfe_normal = 1
    NFe_Complementar = 2
    NFe_DeAjuste = 3
End Enum

Public Enum tipo_S
    S02 = 2
    S03 = 3
    S04 = 4
    S07 = 7
    S09 = 9
End Enum
Public Enum procEmi
    Aplicativo_contribuinte = 0
    avulsa_pelo_fisco = 1
    avulta_pelo_Contribuinte_site_fisco = 2
    pelo_contruinte_aplicativo_fornecido_fisco = 3
    
End Enum
Public Enum indTot
    vProd_Nao_copoe_total_nota = 0
    vProd_Sim_copoe_total_nota = 1
End Enum
Public Enum indFinal
    nao = 0
    consumidor_final = 1
End Enum
Public Enum indPres
    nao_se_aplica = 0
    Operacao_presencial = 1
    Internet = 2
    Teleatendimento = 3
    entrega_domicilio = 4
    Outros = 9
End Enum
Public Enum idDest
    operacao_interna = 1
    operacao_interestadual = 2
    operacao_exterior = 3
End Enum

Public Type Bnfe
    qtdnf As String
    cUF As String
    cNF As String
    natOp As String
    indPag As indPag
    sMod As String
    serie As String
    nNf As String
    dEmi As String
    dSaiEnt As String
    hSaiEnt As String
    tpNf As String
    cMunFG As String
    tpImp As tpImp
    tpEmis As tpEmis
    cDV As String 'digito verificador chave de acesso -não usa
    tpAmb As tpAmb
    finNfe_ As TfinNfe
    procEmi As String
    verProc As String
    dhCont As String
    xJust As String
    idDest_ As idDest
    indFinal_ As indFinal
    indPres_ As indPres
   
End Type
Public Type C05NFe
    xLgr As String
    nro As String
    Cpl As String
    bairro As String
    cMun As String
    xMun As String
    uf As String
    cep As String
    cPais As String
    xPais As String
    fone As String
End Type
Public Type Dnfe 'identificação do fisco emitente
    CNPJ As String 'não sei pra que serve esse caso, mas deve ser pro próprio fisco emitir
    xOrgao As String
    matr As String
    xAgente As String
    fone As String
    uf As String
    nDAR As String
    dEmi As String
    vDAR As String
    repEmi As String
    dPag As String
End Type
Public Type sNfe
    tipo As tipo_S
    cSt As String
    vBC As String
    pCOFINS As String
    vCOFINS As String

    qBCProd As String
    vAliqProd As String
    
End Type

Public Type Cnfe
    xNome As String
    xFant As String
    IE As String
    iEst As String
    IM As String
    CNAE As String
    CRT As CRT
    CNPJCPF As String
    tipo_pessoa As String
    c05 As C05NFe

End Type
Type E05_Nfe
    xLgr As String
    nro As String
    xCpl As String
    xBairro As String
    cMun As String
    xMun As String
    uf As String
    cep As String
    cPais As String
    xPais As String
    fone As String
End Type
Type ENfe
    xNome As String
    IE As String
    ISUF As String
    email As String
    CNPJ As String
    tipo_pessoa As String
    e05 As E05_Nfe
    indIeDest_ As indIEDest
    IM As String 'inscrição municipal do tomador de serviço
End Type
Public Enum q
    Q02 = 2
    Q03 = 3
    Q04 = 4
    Q05 = 5
    Q07 = 7
    Q10 = 10
    
End Enum
Type qNfe
    tipo As q
    cSt As String
    vBC As String
    pPIS As String
    vPIS As String
    
    qBCProd As String
    vAliqProd As String
    
    
    
    
End Type
Type WNfe
    vBC As String
    vICMS As String
    vBCST As String
    vST As String
    vProd As String
    vFrete As String
    vSeg As String
    vDesc As String
    vII As String
    vIPI As String
    vPIS As String
    vCOFINS As String
    vOutro As String
    vNF As String
End Type
Type HNfe
    nItem As String
    infAdProd As String
End Type
Type nNfe
    tipo As tipo_N
    orig As orig
    cSt As String
    modBC As modBC
    pRedBC As String
    vBC As String
    pICMS As String
    vICMS As String
    modBCST As modBCST
    pMVAST As String
    pRedBCST As String
    vBCST As String
    pICMSST As String
    vICMSST As String
    pBCOp As String
    UFST As String
    pCredSN As String
    vC As String
    redICMSSN As String
    motDesICMS As String
    vBCSTRet As String
    vICMSSTRet As String
    vICMSSTDest As String
    csosn As String
End Type
Type INfe
    cProd As String
    cEAN As String
    xProd As String
    ncm As String
    EXTIPI As String
    CFOP As String
    uCom As String
    qCom As String
    vUnCom As String
    vProd As String
    cEANTrib As String
    uTrib As String
    qTrib As String
    vUnTrib As String
    vFrete As String
    vSeg As String
    vDesc As String
    vOutro As String
    indTot As indTot
    xPed As String
    nItemPed As String
    n As nNfe
    q As qNfe
    s As sNfe
End Type



Public Type NFe
    b As Bnfe
    
    
End Type
Function geraA(versao As String, id As String) As String
    Dim strSQL As String
    
    strSQL = "A|" & versao & "|"
    If id <> "" Then
        strSQL = strSQL & id & "|"
    End If
    geraA = strSQL
End Function
Function geraB(cUF As String, cNF As String, natOp As String, iPag As indPag, Modelo As String, serie As String, nNf As String, dhEmi As String, _
    dhSaiEnt As String, tpNf As String, idDest_ As idDest, cMunFG As String, tpImpressao As tpImp, tpEmiss As tpEmis, cDV As String, tpAmb_ As tpAmb, _
    finNfe_ As TfinNfe, indFinal_ As indFinal, indPres_ As indPres, procEmi As String, verProc As String) As String
    Dim strSQL As String
    
    strSQL = "B|" & cUF & "|" & cNF & "|" & natOp & "|" & iPag & "|"
    
    strSQL = strSQL & Modelo & "|" & serie & "|" & nNf & "|" & dhEmi & "|" _
    & dhSaiEnt & "|" & tpNf & "|" & idDest_ & "|" & cMunFG & "|" & tpImpressao & "|" & tpEmiss & "|" & cDV & "|" & tpAmb_ & "|" _
     & finNfe_ & "|" & indFinal_ & "|" & indPres_ & "|" & procEmi & "|" & verProc
    
    
    geraB = strSQL
End Function

Function geraStringNFE(b As Bnfe, c As Cnfe, e As ENfe, h As HNfe, _
    i() As INfe, w As WNfe, Optional Cont As String = "", _
    Optional xJust As String = "") As String
    Dim iCount As Integer
    Dim strSQL As String
    Dim sVersaoApp As String
    sVersaoNfe = "3.10"
    If sVersaoNfe = "" Then
        sVersaoNfe = "2.00"
    End If
    
    strSQL = "NOTAFISCAL|" & b.qtdnf & CrLf
    strSQL = strSQL & geraA(sVersaoNfe, "") & CrLf
    'strSQL = strSQL & "B|" & b.cUF & "|" & b.cNF & "|" & b.natOp & "|" & b.indPag & "|" & b.sMod & "|" & _
    b.serie & "|" & b.nNF & "|" & b.dEmi & "|" & b.dSaiEnt & "|" & b.hSaiEnt & "|" & b.tpNF & "|" & _
    b.cMunFG & "|" & b.tpImp & "|" & b.tpEmis & "|" & b.cDV & "|" & b.tpAmb & "|" & b.finNfe & "|" & _
    b.procEmi & "|" & b.verProc & "|" & b.dhCont & "| " & b.xJust & "|" & CrLf
    
    strSQL = strSQL & geraB(b.cUF, b.cNF, b.natOp, b.indPag, b.sMod, b.serie, b.nNf, b.dEmi, b.dSaiEnt, b.tpNf, b.idDest_, b.cMunFG, _
        b.tpImp, b.tpEmis, b.cDV, b.tpAmb, b.finNfe_, b.indFinal_, b.indPres_, b.procEmi, b.verProc) & CrLf
    
    
'[0 a N] {
'[seleção entre B13 ou B14 ou B20a ou B20i ou B20j]{
'B13|refNFe|
'[ou]
'B14|cUF|AAMM(ano mês)|CNPJ|Mod|serie|nNF|
'[ou]
'B20a|cUF|AAMM|IE|mod|serie|nNF|
'[seleção entre B20d ou B20e]{
'B20d|CNPJ|
'[ou]
'B20e|CPF|
'}
'[ou]
'B20i|refCTe|
'[ou]
'B20j|mod|nECF|nCOO|
'}
'}
    strSQL = strSQL & "C|" & c.xNome & "|" & c.xFant & "|" & c.IE & _
    "|" & c.iEst & "|" & c.IM & "|" & c.CNAE & "|" & c.CRT & "|" & CrLf
    If c.tipo_pessoa = "F" Then
        strSQL = strSQL & "C02a|" & c.CNPJCPF & "|" & CrLf
    Else
    
        strSQL = strSQL & "C02|" & c.CNPJCPF & "|" & CrLf
    End If

    strSQL = strSQL & "C05|" & c.c05.xLgr & "|" & c.c05.nro & "|" & c.c05.Cpl & "|" & c.c05.bairro & "|" & _
    c.c05.cMun & "|" & c.c05.xMun & "|" & c.c05.uf & "|" & c.c05.cep & "|" & _
    c.c05.cPais & "|" & c.c05.xPais & "|" & c.c05.fone & "|" & CrLf
    
    
'D|CNPJ|xOrgao|matr|xAgente|fone|UF|nDAR|dEmi|vDAR|repEmi|dPag|
'}

    strSQL = strSQL & "E|" & e.xNome & "|" & e.indIeDest_ & "|" & e.IE & "|" & e.ISUF & "|" & e.IM & "|" & e.email & "|" & CrLf
    If e.tipo_pessoa = "J" Then
        strSQL = strSQL & "E02|" & e.CNPJ & "|" & CrLf
    Else
        strSQL = strSQL & "E03|" & e.CNPJ & "|" & CrLf
    End If
    strSQL = strSQL & "E05|" & e.e05.xLgr & "|" & e.e05.nro & "|" & e.e05.xCpl & _
    "|" & e.e05.xBairro & "|" & e.e05.cMun & "|" & e.e05.xMun & "|" & e.e05.uf & "|" & _
    e.e05.cep & "|" & e.e05.cPais & "|" & e.e05.xPais & "|" & e.e05.fone & "|" & CrLf

'[0 ou 1]{
'F|CNPJ|XLgr|Nro|XCpl|XBairro|CMun|XMun|UF| ' f retirada g entrega, não usa
'[seleção entre F02 ou F02a]{
'F02|CNPJ
'[ou]
'F02a|CPF
'}
'}
'[0 ou 1]{
'G|CNPJ|XLgr|Nro|XCpl|XBairro|CMun|XMun|UF|
'[seleção entre G02 ou G02a]{
'G02|CNPJ
'[ou]
'G02a|CPF
'}
'}
'[1 a 990]{
    For iCount = 0 To h.nItem - 1
        strSQL = strSQL & "H|" & iCount + 1 & "|" & h.infAdProd & "|" & CrLf
        strSQL = strSQL & "I|" & i(iCount).cProd & "|" & i(iCount).cEAN & "|" & _
        i(iCount).xProd & "|" & i(iCount).ncm & "|" & i(iCount).EXTIPI & "|" & _
        i(iCount).CFOP & "|" & i(iCount).uCom & "|" & i(iCount).qCom & "|" & _
        i(iCount).vUnCom & "|" & i(iCount).vProd & "|" & i(iCount).cEANTrib & "|" & _
        i(iCount).uTrib & "|" & i(iCount).qTrib & "|" & i(iCount).vUnTrib & "|" & _
        i(iCount).vFrete & "|" & i(iCount).vSeg & "|" & i(iCount).vDesc & "|" & _
        i(iCount).vOutro & "|" & i(iCount).indTot & "|" & i(iCount).xPed & "|" & _
        i(iCount).nItemPed & "|" & CrLf
    


        
        strSQL = strSQL & "M|" & CrLf
        strSQL = strSQL & "N|" & CrLf
        
        If i(iCount).n.tipo = N02 Then
        
            strSQL = strSQL & "N02|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & i(iCount).n.modBC & _
                "|" & (i(iCount).n.vBC) & "|" & (i(iCount).n.pICMS) & "|" & (i(iCount).n.vICMS) & "|"
        ElseIf i(iCount).n.tipo = N03 Then
        
            strSQL = strSQL & "N03|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & i(iCount).n.modBC & _
                "|" & (i(iCount).n.vBC) & "|" & (i(iCount).n.pICMS) & "|" & (i(iCount).n.vICMS) & "|" & _
                i(iCount).n.modBCST & "|" & i(iCount).n.pMVAST & "|" & i(iCount).n.pRedBCST & "|" & _
                (i(iCount).n.vBCST) & "|" & (i(iCount).n.pICMSST) & "|" & (i(iCount).n.vICMSST) & "|"
        ElseIf i(iCount).n.tipo = N04 Then
            strSQL = strSQL & "N04|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & i(iCount).n.modBC & "|" & _
            i(iCount).n.pRedBC & "|" & i(iCount).n.vBC & "|" & i(iCount).n.pICMS & "|" & i(iCount).n.vICMS & "|"
        ElseIf i(iCount).n.tipo = N05 Then
            strSQL = strSQL & "N05|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & i(iCount).n.modBCST & _
            "|" & i(iCount).n.pMVAST & "|" & i(iCount).n.pRedBCST & "|" & i(iCount).n.vBCST & "|" & i(iCount).n.pICMSST & _
            "|" & i(iCount).n.vICMSST & "|"
        ElseIf i(iCount).n.tipo = N06 Then
            strSQL = strSQL & "N06|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & _
            i(iCount).n.vICMS & "|" & i(iCount).n.motDesICMS & "|"
        ElseIf i(iCount).n.tipo = N07 Then
            strSQL = strSQL & "N07|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & _
            i(iCount).n.modBC & "|" & pRedBC & "|" & vBC & "|" & pICMS & "|" & i(iCount).n.vICMS & "|"
        ElseIf i(iCount).n.tipo = N08 Then
            strSQL = strSQL & "N08|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & _
            i(iCount).n.vBCSTRet & "|" & i(iCount).n.vICMSSTRet & "|"
        ElseIf i(iCount).n.tipo = N09 Then
            strSQL = strSQL & "N09|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & _
            i(iCount).n.modBC & "|" & i(iCount).n.pRedBC & "|" & i(iCount).n.vBC & "|" & i(iCount).n.pICMS & "|" & i(iCount).n.vICMS & "|" & _
            i(iCount).n.modBCST & "|" & pMVAST & "|" & pRedBCST & "|" & vBCST & "|" & pICMSST & "|" & i(iCount).n.vICMSST & "|"
        ElseIf i(iCount).n.tipo = N10 Then
            strSQL = strSQL & "N10|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & _
            i(iCount).n.modBC & "|" & pRedBC & "|" & vBC & "|" & pICMS & "|" & i(iCount).n.vICMS & "|" & i(iCount).n.modBCST & "|" & pMVAST & "|" & pRedBCST & "|" & vBCST & "|" & pICMSST & "|" & i(iCount).n.vICMSST & "|"
        ElseIf i(iCount).n.tipo = N10a Then
            strSQL = strSQL & "N10a|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & i(iCount).n.modBC & "|" & pRedBC & "|" & vBC & "|" & pICMS & "|" & i(iCount).n.vICMS & "|" & i(iCount).n.modBCST & "|" & pMVAST & "|" & pRedBCST & "|" & vBCST & "|" & pICMSST & "|" & i(iCount).n.vICMSST & "|" & pBCOp & "|" & UFST & "|"
        ElseIf i(iCount).n.tipo = N10b Then
            strSQL = strSQL & "N10b|" & i(iCount).n.orig & "|" & i(iCount).n.cSt & "|" & i(iCount).n.vBCSTRet & "|" & i(iCount).n.vICMSSTRet & "|" & vBCSTDest & "|" & i(iCount).n.vICMSSTDest & "|"
        ElseIf i(iCount).n.tipo = N10c Then
            strSQL = strSQL & "N10c|" & i(iCount).n.orig & "|" & i(iCount).n.csosn & "|" & i(iCount).n.pCredSN & "|" & vCredICMSSN & "|"
        ElseIf i(iCount).n.tipo = N10d Then
            strSQL = strSQL & "N10d|" & i(iCount).n.orig & "|" & i(iCount).n.csosn & "|"
        ElseIf i(iCount).n.tipo = N10e Then
            strSQL = strSQL & "N10e|" & i(iCount).n.orig & "|" & i(iCount).n.csosn & "|" & i(iCount).n.modBCST & "|" & i(iCount).n.pMVAST & "|" & i(iCount).n.pRedBCST & "|" & vBCST & "|" & pICMSST & "|" & i(iCount).n.vICMSST & "|" & pCredSN & "|" & vCredICMSSN & "|"
        ElseIf i(iCount).n.tipo = N10f Then
            strSQL = strSQL & "N10f|" & i(iCount).n.orig & "|" & i(iCount).n.csosn & "|" & i(iCount).n.modBCST & "|" & i(iCount).n.pMVAST & "|" & i(iCount).n.pRedBCST & "|" & vBCST & "|" & pICMSST & "|" & i(iCount).n.vICMSST & "|"
        ElseIf i(iCount).n.tipo = N10g Then
            strSQL = strSQL & "N10g|" & i(iCount).n.orig & "|" & i(iCount).n.csosn & "|" & i(iCount).n.modBCST & "|" & i(iCount).n.vBCSTRet & "|" & i(iCount).n.vICMSSTRet & "|"
        ElseIf i(iCount).n.tipo = N10h Then
            strSQL = strSQL & "N10h|" & i(iCount).n.orig & "|" & i(iCount).n.csosn & "|" & i(iCount).n.modBC & _
            "|" & vBC & "|" & i(iCount).n.pRedBC & "|" & i(iCount).n.pICMS & "|" & i(iCount).n.vICMS & "|" & i(iCount).n.modBCST & "|" & pMVAST & "|" & pRedBCST & "|" & vBCST & "|" & pICMSST & "|" & i(iCount).n.vICMSST & "|" & pCredSN & "|" & vC & redICMSSN & "|"
        End If

        strSQL = strSQL & CrLf
    
            
        strSQL = strSQL & "Q|" & CrLf
    
        If i(iCount).q.tipo = Q02 Then
            strSQL = strSQL & "Q02|" & i(iCount).q.cSt & "|" & i(iCount).q.vBC & "|" & i(iCount).q.pPIS & "|" & i(iCount).q.vPIS & "|"
        ElseIf i(iCount).q.tipo = Q03 Then
            strSQL = strSQL & "Q03|" & i(iCount).q.cSt & "|" & i(iCount).q.qBCProd & "|" & i(iCount).q.vAliqProd & "|" & i(iCount).q.vPIS & "|"
        ElseIf i(iCount).q.tipo = Q04 Then
            strSQL = strSQL & "Q04|" & i(iCount).q.cSt & "|"
        
        Else
            strSQL = strSQL & "Q05| " & i(iCount).q.cSt & "|" & i(iCount).q.vPIS & "|" & CrLf
            If i(iCount).q.tipo = Q07 Then
                strSQL = strSQL & "Q07|" & i(iCount).q.vBC & "|" & i(iCount).q.pPIS & "|"
            Else
                strSQL = strSQL & "Q10|" & i(iCount).q.qBCProd & "|" & i(iCount).q.vAliqProd & "|"
            End If
            
        End If
    
        strSQL = strSQL & CrLf
        
        strSQL = strSQL & "S|" & CrLf
        If i(iCount).s.tipo = S02 Then
    
            strSQL = strSQL & "S02|" & (i(iCount).s.cSt) & "|" & (i(iCount).s.vBC) & "|" & (i(iCount).s.pCOFINS) & _
                "|" & i(iCount).s.vCOFINS & "|"
        ElseIf i(iCount).s.tipo = S03 Then
            strSQL = strSQL & "S03|" & (i(iCount).s.cSt) & "|" & (i(iCount).s.qBCProd) & "|" & (i(iCount).s.vAliqProd) & "|" & (i(iCount).s.vCOFINS) & "|"
        ElseIf i(iCount).s.tipo = S04 Then
            strSQL = strSQL & "S04|" & i(iCount).s.cSt & "|"
        Else
            strSQL = strSQL & "S05|" & (i(iCount).s.cSt) & "|" & (i(iCount).s.vCOFINS) & "|"
            If i(iCount).s.tipo = S07 Then
                strSQL = strSQL & "S07|" & (i(iCount).s.vBC) & "|" & (i(iCount).s.pCOFINS) & "|"
            Else
                strSQL = strSQL & "S09|" & (i(iCount).s.qBCProd) & "|" & (i(iCount).s.vAliqProd) & "|"
            End If
        End If
    
        strSQL = strSQL & CrLf
    Next

    strSQL = strSQL & "W|"
    strSQL = strSQL & "W02|" & (w.vBC) & "|" & (w.vICMS) & "|" & (w.vBCST) & _
        "|" & (w.vST) & "|" & (w.vProd) & "|" & (w.vFrete) & "|" & _
        (w.vSeg) & "|" & (w.vDesc) & "|" & _
        (w.vII) & "|" & (w.vIPI) & "|" & (w.vPIS) & "|" & _
        (w.vCOFINS) & "|" & _
        (w.vOutro) & "|" & (w.vNF) & "|"
    strSQL = strSQL & CrLf
    strSQL = strSQL & "X|9|"
    strSQL = strSQL & CrLf
    geraStringNFE = strSQL
    '[0 ou 1]{
    'W17|VServ|VBC|VISS|VPIS|VCOFINS|
    '}
    'W23|VRetPIS|VRetCOFINS|VRetCSLL|VBCIRRF|VIRRF|VBCRetPrev|VRetPrev|
'    X|ModFrete|
'    X03|XNome|IE|XEnder|UF|XMun|
'    [Seleção entre X04 ou X05]{
'    X04|CNPJ|
'    [ou]
'    X05|CPF|
'    }
'    [0 ou 1]{
'    X11|VServ|VBCRet|PICMSRet|i(iCount).n.vicmsRet|CFOP|CMunFG|
'    }
'    [0 ou 1]{
'    X18|Placa|UF|RNTC|
'    }
'    Versão 2.0.0 - (24/08/2010) Página 12 de 13
'    [0 a 2]{
'    X22|Placa|UF|RNTC|
'    }
'    [0 a N]{
'    X26|QVol|Esp|Marca|NVol|PesoL|PesoB|
'    [0 a N]{
'    X33|NLacre|
'    }
'    }
'    [0 ou 1]{
'    Y|
'    [0 ou 1]{
'    Y02|NFat|VOrig|VDesc|VLiq|
'    }
'    [0 a N]{
'    Y07|NDup|DVenc|VDup|
'    }
'    [0 ou 1]{
'    Z|InfAdFisco|InfCpl|
'    [0 a 10]{
'    Z04|XCampo|XTexto|
'    }
'    [0 a 10]{
'    Z07|XCampo|XTexto|
'    }
'    [0 a N]{
'    Z10|NProc|IndProc|
'    }
'    }
'    [0 ou 1]{
'    ZA|UFEmbarq|XLocEmbarq|
'    }
'    [0 ou 1]{
'    ZB|XNEmp|XPed|XCont|
'    }
'    [0 ou 1]{
'    ZC01|safra|ref|qTotMes|qTotAnt|qTotGer|vFor|vTotDed|vLiqFor|
'    [1 a 31]{
'    ZC04|dia|qtde|
'    }
'    [0 a 10]{
'    Versão 2.0.0 - (24/08/2010) Página 13 de 13
'    ZC10|xDed|vDed|
'    }
'    }

End Function





Sub ExportaNfe(Seq_sai As String)
    Dim sNfe As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    Dim b As Bnfe
    Dim c As Cnfe
    Dim d As Dnfe
    Dim e As ENfe
    Dim h As HNfe
    Dim i() As INfe
    Dim n As nNfe
    Dim q As qNfe
    Dim s As sNfe
    Dim w As WNfe
    Dim strCfop As String
    Dim Indice As Integer
    Dim sPath As String
    b.qtdnf = "1"
    mdiFrm.cmndlg.Filter = "*.txt"
    mdiFrm.cmndlg.FileName = Seq_sai & ".txt"
    mdiFrm.cmndlg.ShowSave
    sPath = mdiFrm.cmndlg.FileName
    
    strSQL = "select tc.*,ts.*,tci.* "
    strSQL = strSQL & "  from tab_saidasnf as ts, tab_cidades as tci, " & TAB_CLIENTES & " as tc "
    strSQL = strSQL & " where ts.cod_cli=tc.cod_cli and tc.cod_cid=tci.cod_cid "
    strSQL = strSQL & " and seq_sai='" & Seq_sai & "'"
    rPesq.Open strSQL, conn_client
    
    With rPesq
        strCfop = !cfop_sai
        b.natOp = !natop_sai
        b.dhCont = "" 'data hora entrada em con
        b.cDV = "" 'digito verificador
        b.cMunFG = strCodCid_Emp  'código do município
        b.nNf = Format(!nfe_sai, "0000000#") '  numero aleatorio gerado pelo emitente
        'b.cNF = "52" + format(date,"YYMM") + F_SoNumero(strCNPJ_Emp) + "55" + "1" + format(!nfe_sai,"0000000#") +
        b.cUF = "52" 'goiás
        b.dEmi = FormataDataNfe(!emissao_sai & " " & Time)
        b.dSaiEnt = FormataDataNfe(!emissao_sai & " " & Time)
        b.hSaiEnt = ""
        b.tpNf = "1" 'tipo entrada  '0 - saida
        b.indPag = indPag.avista
        b.sMod = "55" 'acho que é isso  - modelo da nota fiscal
        b.serie = "1"
        b.tpImp = retrato
        b.tpEmis = Normal
        b.tpAmb = homologacao
        b.finNfe_ = Nfe_normal
        b.procEmi = 3 'pelo contribuindo com aplicativo fornecido pelo fisco
        b.verProc = "3.10.14" 'versão do aplicativo
        b.xJust = "" 'justificativa entrega contigencia
        b.idDest_ = operacao_interna 'colocar um combo para escolher
        b.indFinal_ = nao
        b.indPres_ = nao_se_aplica
        
        b.cDV = ""
        sCFOP = !cfop_sai
        
        
        'preencher C
        c.xNome = strNome_Emp
        c.xFant = strFantasia_Emp
        
        c.IE = F_SoNumero(strIE_Emp)
        
        
        
        c.iEst = ""
        c.IM = "" 'inscricao municipal
        c.CNAE = "" 'este campo dever ser informado qdo houver insc mun
        c.CNPJCPF = F_SoNumero(strCNPJ_Emp)
        c.CRT = RECEITA_NORMAL  '1 simples nacional, 2 simples nacional excesso de sub rec bruta, 3 regime normal
        c.c05.bairro = strBairro_Emp
        c.c05.cep = F_SoNumero(strCEP_Emp)
        c.c05.cMun = strCodCid_Emp
        c.c05.cPais = "1058" 'código do brasil
        c.c05.xPais = "Brasil"
        c.c05.xLgr = strEndereco_Emp
        c.c05.nro = strNumero_Emp
        c.c05.Cpl = strComplemento_Emp
        
        c.c05.uf = strUF_Emp
        
        e.CNPJ = F_SoNumero(!cpf_cli)
        e.email = "" '!email_cli
        If !ie_Cli <> "" Then
            e.IE = F_SoNumero(!ie_Cli)
        End If
        e.ISUF = ""
        e.tipo_pessoa = !tipo_cli
        e.xNome = !nome_Cli
        If !ie_Cli = "" Then
            e.indIeDest_ = isento
        Else
            e.indIeDest_ = contribuinte_icms
        End If
        If !cep_cli <> "" Then
            e.e05.cep = F_SoNumero(!cep_cli)
        End If
        If !endereco_cli <> "" Then
            e.e05.xLgr = !endereco_cli
        End If
        
        e.e05.cMun = !cod_cid
        e.e05.cPais = "1058"
        If !telefone_cli <> "" Then
            e.e05.fone = F_SoNumero(!telefone_cli)
        End If
        e.e05.nro = Format(!numero_cli)
        e.e05.uf = !uf_cid
        If !bairro_cli <> "" Then
            e.e05.xBairro = !bairro_cli
        End If
        
        e.e05.xCpl = Format(!COMPLEMENTO_cli)
        e.e05.xPais = "Brasil"
        w.vBC = MascaraValorNfe(!bc_sai)
        w.vBCST = MascaraValorNfe("0,00")
        w.vICMS = MascaraValorNfe(!vicms_sai)
        w.vCOFINS = MascaraValorNfe("0,00")
        w.vProd = MascaraValorNfe(!valp_sai)
        w.vFrete = MascaraValorNfe(!frete_sai)
        w.vSeg = MascaraValorNfe(!seguro_sai)
        w.vDesc = MascaraValorNfe(!desconto_sai)
        w.vII = "0.00" ' I.I.
        w.vIPI = MascaraValorNfe(!ipi_sai)
        w.vPIS = MascaraValorNfe(!pis_sai)
        w.vCOFINS = MascaraValorNfe(!cofins_sai)
        w.vOutro = MascaraValorNfe(!despesas_sai)
        w.vNF = MascaraValorNfe(!valor_sai)
        
        
        
        .Close
    End With
    strSQL = "select tp.*,ti.* from " & TAB_PRODUTOS & " as tp, tab_itsainf as ti"
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and seq_sai='" & Seq_sai & "' "
    
    
    rPesq.Open strSQL, conn_client
    With rPesq
        ReDim i(.RecordCount) As INfe
        h.nItem = .RecordCount
        h.infAdProd = ""
        Do While Not .EOF
            Indice = .AbsolutePosition - 1
            i(Indice).cEAN = "" 'código de barras
            i(Indice).cEANTrib = ""
            i(Indice).CFOP = strCfop
            i(Indice).cProd = !cod_pro
            i(Indice).EXTIPI = ""
            i(Indice).indTot = vProd_Sim_copoe_total_nota
            i(Indice).ncm = Format(!ncm_pro, "0000000#")
            i(Indice).nItemPed = ""
            i(Indice).qTrib = MascaraValorNfe(!qtd_its, 3)
            i(Indice).qCom = MascaraValorNfe(!qtd_its, 3)
            i(Indice).vUnCom = MascaraValorNfe(!unit_its)
            i(Indice).vUnTrib = MascaraValorNfe(!unit_its)
            i(Indice).vProd = MascaraValorNfe(!unit_its * !qtd_its)
            i(Indice).vDesc = "" 'MascaraValorNfe("0") 'desconto
            i(Indice).vFrete = "" 'MascaraValorNfe("0") 'frete
            i(Indice).vOutro = "" 'MascaraValorNfe("0")
            i(Indice).xProd = TiraAcento(!nome_pro)
            i(Indice).uCom = !unidade_pro
            i(Indice).uTrib = !unidade_pro
            
            i(Indice).q.cSt = "07" 'ver isso com o contador
            
            
            If !nfe_cst = "40" Or !nfe_cst = "41" Or !nfe_cst = "50" Then
                i(Indice).n.tipo = N06
                i(Indice).n.orig = nacional
                
                i(Indice).n.cSt = !nfe_cst
            
            ElseIf !nfe_cst = "10" Then
                i(Indice).n.tipo = N03
                i(Indice).n.orig = nacional
                i(Indice).n.modBC = Margem_Valor_Agregado_perc
                i(Indice).n.cSt = !nfe_cst
                i(Indice).n.vBC = MascaraValorNfe(!unit_its * !qtd_its) 'valor total do produto
                i(Indice).n.pICMS = MascaraValorNfe(!icms_its)
                i(Indice).n.vICMS = MascaraValorNfe(!icms_its / 100 * !unit_its * !qtd_its) 'criar
                i(Indice).n.modBCST = Margem_Valor_Agregado_perct
                i(Indice).n.vBCST = MascaraValorNfe(!unit_its * !qtd_its) 'valor total do produto
                i(Indice).n.pICMSST = MascaraValorNfe(!icms_its)
                i(Indice).n.vICMSST = MascaraValorNfe(!icms_its / 100 * !unit_its * !qtd_its)  'criar
                i(Indice).n.cSt = !nfe_cst
            ElseIf !nfe_cst = "20" Then
                i(Indice).n.tipo = N04
                i(Indice).n.orig = nacional
                i(Indice).n.modBC = Margem_Valor_Agregado_perc
                i(Indice).n.pRedBC = !predbc_pro
                i(Indice).n.vBC = MascaraValorNfe(!unit_its * !qtd_its)  'valor total do produto
                i(Indice).n.pICMS = MascaraValorNfe(!icms_its)
                i(Indice).n.vICMS = MascaraValorNfe(!icms_its / 100 * !unit_its * !qtd_its) 'criar
                i(Indice).n.cSt = !nfe_cst
            ElseIf !nfe_cst = "60" Then
                i(Indice).n.tipo = N08
                i(Indice).n.orig = nacional
                i(Indice).n.vBCSTRet = MascaraValorNfe(!unit_its * !qtd_its)   'valor total do produto
                i(Indice).n.vICMSSTRet = MascaraValorNfe(17 / 100 * !unit_its * !qtd_its)   'criar
                i(Indice).n.cSt = !nfe_cst
            Else
            
                i(Indice).n.tipo = N02
                
                
                
                i(Indice).n.orig = nacional
                If Not IsNull(!nfe_cst) Then
                    i(Indice).n.cSt = !nfe_cst
                End If
                If i(Indice).n.cSt = "" Then
                    i(Indice).n.cSt = "00"  'tributada integralmente
                
                End If
                
                    i(Indice).n.modBC = modBC.Margem_Valor_Agregado_perc 'perguntar pro contador
                    i(Indice).n.vBC = MascaraValorNfe(!unit_its * !qtd_its) 'valor total do produto
                    i(Indice).n.pICMS = MascaraValorNfe(!icms_its)
                    i(Indice).n.vICMS = MascaraValorNfe(!icms_its / 100 * !unit_its * !qtd_its) 'criar
            End If
            i(Indice).q.tipo = Q04
            i(Indice).q.cSt = "07"
            i(Indice).q.pPIS = ""
            
            i(Indice).s.tipo = S04
            i(Indice).s.cSt = "07"
            
            
            .MoveNext
        Loop
    End With
    
    
    
    sNfe = geraStringNFE(b, c, e, h, i, w)
    
    
    Open sPath For Output As #1
    Print #1, sNfe
    Close #1
    
    On Error Resume Next
    Shell "explorer " & sPath
    Close #1
End Sub

Function FormataCPFNfe(CPF As String) As String
    FormataCPFNfe = Format(F_SoNumero(CPF), "0##########")
End Function
Function FormataCNPJNfe(CNPJ As String) As String
    FormataCNPJNfe = Format(F_SoNumero(CNPJ), "0#############")
End Function


Public Function TiraAcento(vString As String) As String
    
    Dim inString
    Dim inC, i
    
    inString = RTrim(vString)
  
    For i = 1 To Len(inString)
      
       inC = Mid(inString, i, 1)
       If inC = "ã" Or inC = "á" Or inC = "à" Or inC = "â" Then
              inC = "a"
       ElseIf inC = "Ã" Or inC = "Á" Or inC = "À" Or inC = "Â" Then
              inC = "A"
       ElseIf inC = "õ" Or inC = "ó" Or inC = "ô" Then
              inC = "o"
       ElseIf inC = "Õ" Or inC = "Ó" Or inC = "Ô" Then
              inC = "O"
       ElseIf inC = "é" Or inC = "ê" Then
              inC = "e"
       ElseIf inC = "É" Or inC = "Ê" Then
              inC = "E"
       ElseIf inC = "í" Then
              inC = "i"
       ElseIf inC = "Í" Then
              inC = "I"
       ElseIf inC = "ú" Or inC = "ü" Then
              inC = "u"
       ElseIf inC = "Ú" Or inC = "Ü" Then
              inC = "U"
       ElseIf inC = "ç" Then
              inC = "c"
       ElseIf inC = "Ç" Then
              inC = "C"
       ElseIf inC = "º" Then
              inC = "o"
       End If
      
       TiraAcento = TiraAcento & inC
    
    Next i
End Function
Function MascaraValorNfe(valor As String, Optional c_dec As Integer = 2)
    valor = FormatVal(valor, "#######0." & String(c_dec, "0"))
    MascaraValorNfe = mudarCaractere(valor, ",", ".")

End Function

Function FormataDataNfe(valor As String, Optional fuso As String = "03:00") As String
    FormataDataNfe = Format(valor, "yyyy-mm-ddThh:mm:ss-" & fuso)
End Function
Function GravaDataNfe(valor As String, Optional fuso As String = "03:00") As String
    GravaDataNfe = "'" & Format(valor, "yyyy-mm-ddThh:mm:ss-" & fuso) & "' "
End Function

Function Valor2Casas(valor) As String
    Valor2Casas = Format(str(valor), "###,##0.00")
    Valor2Casas = ConverteValorSped(Valor2Casas)
End Function
Function Valor4Casas(valor) As String
    Valor4Casas = Format(str(valor), "###,##0.0000")
    Valor4Casas = ConverteValorSped(Valor4Casas)
End Function

