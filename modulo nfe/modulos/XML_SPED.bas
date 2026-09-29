Attribute VB_Name = "XML_SPED"
Public Type pis
    cst As String
    vBC As String
    pPIS As String
    vPIS As String
End Type
Type ipi
    cEnq As String
    cst As String
    vIPI As String
    pIPI As String
    vBC As String
End Type
Public Type cofins
    cst As String
    vBC As String
    pCOFINS As String
    vCOFINS As String
End Type

Public Type icms
    orig As String
    cst As String
    modBC As String
    vBC As String
    pICMS As String
    vICMS As String
    modBCST As String
    pMVAST As String
    vBCST As String
    pICMSST As String
    vICMSST As String
    
    
End Type

Public Type imposto
    icms_ As icms
    pis_ As pis
    cofins_ As cofins
End Type
Public Type prod
    cProd As String
    cEAN As String
    cNCM As String
    xProd As String
    ncm As String
    CFOP As String
    uCom As String
    qCom As String
    vUnCom As String
    vDesc As String
    vProd As String
    cEANTrib As String
    uTrib As String
    qTrib As String
    vUnTrib As String
    
    indTot As String
    icms_ As icms
    ipi_ As ipi
    pis_ As pis
    codigo As String
    nome_pro As String
    cof_ As cofins
    qtd_convertida As String
    vFrete As String
    vSeguro As String
    vOutro As String
    
    
End Type
Type transporta
    CNPJ As String
    xNome As String
    IE As String
    xEnder As String
    xMun As String
    UF As String
    
    placa As String
    uf_vei As String
    
End Type

Type transp
    modFrete As String
    t As transporta
End Type


Type total
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
    vTotTrib As String
    
End Type


Public Type emit
    CNPJ As String
    xNome As String
    xFant As String
    xLgr As String
    nro As String
    xBairro As String
    cMun As String
    xMun As String
    UF As String
    cep As String
    cPais As String
    xPais As String
    fone As String
    IE As String
    crt As String
    
End Type

Public Type dest
    CNPJ As String
    cpf As String
    xNome As String
    xFant As String
    xLgr As String
    nro As String
    xBairro As String
    cMun As String
    xMun As String
    UF As String
    cep As String
    cPais As String
    xPais As String
    fone As String
    ie_ As String
    crt As String
    IE As String
    email As String
    ende As Ende_NFe.Ende_NFe
    
End Type


Type CabecalhoNF
    id As String
    cUF As String
    cNF As String
    natOp As String
    indPag As String
    mod As String
    serie As String
    nNF As String
    dEmi As String
    tpNf As String
    cMunFG As String
    tpImp As String
    tpEmis As String
    cDV As String
    tpAmb As String
    finNFe As String
    indFnl As indFinal
    procEmi As String
    verProc As String
    e As emit
    d As dest
    idDst As idDest
    dhCont As String
    xJust As String
    indPr As indPres
End Type


Type fatura
    nFat As String
    vOrig As String
    vDesc As String
    vLiq As String
End Type
Type duplicata
    nDup As String
    dVenc As String
    vDup As String
    
End Type
Type Cobranca
    fat As fatura
    d() As duplicata
End Type
Type protNfe
    tpAmb As String
    verApp As String
    dhRecbto As String
    nProt As String
    digVal As String
    cStat As String
    xMotivo As String
End Type
Public Type XMLSped
    cab As CabecalhoNF
    p() As prod
    t As transp
    tot As total
    c As Cobranca
    InfCompl As String
    CH_NFE As String
    infProt As protNfe
End Type



Function ConverteSpedValor(valor As String) As String
    Dim v As String
    v = valor
    v = Replace(v, ".", "#")
    v = Replace(v, ",", "$")
    v = Replace(v, "#", ",")
    v = Replace(v, "$", "")
    ConverteSpedValor = v
End Function

Function ConverteValorSped(valor As String) As String
    Dim v As String
    v = valor
    v = Replace(v, ",", "#")
    v = Replace(v, ".", "$")
    v = Replace(v, "#", ".")
    v = Replace(v, "$", "")
    ConverteValorSped = v
End Function


Function RetornaDifal(UF As String, aliq_icms As String, valor_item As String, desconto As String, seguro As String, frete As String, outrasDA As String) As Double

    If UF = strUF_Emp Then
        RetornaDifal = 0
        Exit Function
    
    
    End If
    
    If CCur(aliq_icms) <> 4 Then
        'RetornaDifal = RetornaVoper(valor_item, desconto, frete, seguro, outrasDA) * 0.0449
         RetornaDifal = RetornaVoper(valor_item, desconto, frete, seguro, outrasDA) * 0
    Else
        'RetornaDifal = RetornaVoper(valor_item, desconto, frete, seguro, outrasDA) * 0.0787
        RetornaDifal = RetornaVoper(valor_item, desconto, frete, seguro, outrasDA) * 0
    End If



End Function

Function RetornaVoper(valor_item As String, desconto_item As String, frete As String, seguro As String, outras_da As String)
    RetornaVoper = CCur(valor_item) - CCur(desconto_item) + CCur(frete) + CCur(seguro) + CCur(outrasDA)
End Function
