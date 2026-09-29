Attribute VB_Name = "XML_SPED"
Public Type pis
    cSt As String
    vBC As String
    pPIS As String
    vPIS As String
End Type
Type ipi
    cEnq As String
    cSt As String
    vIPI As String
    pIPI As String
    vBC As String
End Type
Public Type cofins
    cSt As String
    vBC As String
    pCOFINS As String
    vCOFINS As String
End Type

Public Type icms
    orig As String
    cSt As String
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
    xProd As String
    NCM As String
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
    vOutro As String
    indTot As String
    icms_ As icms
    ipi_ As ipi
    pis_ As pis
    codigo As String
    nome_pro As String
    cof_ As cofins
    qtd_convertida As String
End Type
Type transporta
    CNPJ As String
    xNome As String
    IE As String
    xEnder As String
    xMun As String
    uf As String
    
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
    uf As String
    cep As String
    cPais As String
    xPais As String
    fone As String
    IE As String
    CRT As String
    
End Type

Public Type dest
    CNPJ As String
    xNome As String
    xFant As String
    xLgr As String
    nro As String
    xBairro As String
    cMun As String
    xMun As String
    uf As String
    cep As String
    cPais As String
    xPais As String
    fone As String
    ie_ As String
    CRT As String
    IE As String
    email As String
    
End Type


Type CabecalhoNF
    cUF As String
    cNF As String
    natOp As String
    indPag As String
    mod As String
    serie As String
    nNF As String
    dEmi As String
    tpNF As String
    cMunFG As String
    tpImp As String
    tpEmis As String
    cDV As String
    tpAmb As String
    finNFe As String
    procEmi As String
    verProc As String
    e As emit
End Type


Type fatura
    nFat As String
    vOrig As String
    vLiq As String
End Type
Type duplicata
    nDup As String
    dVenc As String
    vDup As String
    
End Type
Type cobranca
    fat As fatura
    d() As duplicata
End Type

Public Type XMLSped
    cab As CabecalhoNF
    p() As prod
    t As transp
    tot As total
    c As cobranca
    InfCompl As String
    CH_NFE As String
    
    
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


