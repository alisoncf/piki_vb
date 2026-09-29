Attribute VB_Name = "NFe"
Sub AlteraStatusNfe(id As String, status As String)
    strSQL = "update nfe set status='" & status & "'"
    strSQL = strSQL & " where id=" & id
    conn.Execute strSQL
    
End Sub

Function NFE_GravaNFE(num_ven As String, Chave As String, cUF As String, cNF As String, natOp As String, indPag As String, _
    Modelo As String, serie As String, nNf As String, dhEmi As String, dhSaiEnt As String, _
    tpNf As tpNf, idDest As idDest, cMunFG As String, tpImp As tpImp, _
    tpEmis As tpEmis, cDV As String, tpAmb As tpAmb, finNFe As TfinNfe, indFinal As indFinal, _
    indPres As indPres, procEmi As procEmi, verProc As String, dhCont As String, xJust As String) As String
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into nfe(num_ven,chave,cUF , cNF, natOp, indPag, Modelo, serie, nNF, dhEmi, dhSaiEnt, "
    strSQL = strSQL & "tpNF, idDest, cMunFG, tpImp, tpEmis, cDV, tpAmb, finNFe, indFinal, indPres, procEmi, "
    strSQL = strSQL & " verProc, dhCont, xJust)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & "'" & Chave & "',"
    strSQL = strSQL & "'" & cUF & "',"
    strSQL = strSQL & "'" & cNF & "',"
    strSQL = strSQL & "'" & natOp & "', "
    strSQL = strSQL & "'" & indPag & "', "
    strSQL = strSQL & "'" & Modelo & "',"
    strSQL = strSQL & "'" & serie & "',"
    strSQL = strSQL & "'" & nNf & "', "
    strSQL = strSQL & GravaDataNfe(dhEmi) & ", "
    strSQL = strSQL & GravaDataNfe(dhSaiEnt) & ", "
    strSQL = strSQL & "'" & tpNf & "', "
    strSQL = strSQL & "'" & idDest & "', "
    strSQL = strSQL & "'" & cMunFG & "', "
    strSQL = strSQL & "'" & tpImp & "', "
    strSQL = strSQL & "'" & tpEmis & "', "
    strSQL = strSQL & "'" & cDV & "', "
    strSQL = strSQL & "'" & tpAmb & "', "
    strSQL = strSQL & "'" & finNFe & "', "
    strSQL = strSQL & "'" & indFinal & "', "
    strSQL = strSQL & "'" & indPres & "', "
    strSQL = strSQL & "'" & procEmi & "', "
    strSQL = strSQL & "'" & verProc & "', "
    strSQL = strSQL & GravaDataNfe(dhCont) & ", "
    strSQL = strSQL & "'" & xJust & "') "
    
    conn.Execute strSQL
    
    NFE_GravaNFE = F_BuscaUmCampo(" max(id)", "nfe", "")
trata_Erro:
    
    
End Function

Function GravaNFe_Item(id_NFe As String, cProd As String, cEAN As String, xProd As String, ncm As String, NVE As String, EXTIPI As String, CFOP As String, uCom As String, qCom As String, vUnCom As String, vProd As String, cEANTrib As String, uTrib As String, qTrib As String, vUnTrib As String, vFrete As String, vSeg As String, vDesc As String, vOutro As String, indTot As indTot)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item( id_nfe, cProd, cEAN, xProd, NCM, NVE, EXTIPI, CFOP, uCom, qCom, vUnCom, vProd, "
    strSQL = strSQL & " cEANTrib, uTrib, qTrib, vUnTrib, vFrete, vSeg, vDesc, vOutro, indTot) values ("

    strSQL = strSQL & "'" & id_NFe & "',"
    strSQL = strSQL & "'" & cProd & "',"
    strSQL = strSQL & "'" & cEAN & "',"
    strSQL = strSQL & "'" & xProd & "',"
    strSQL = strSQL & "'" & ncm & "',"
    strSQL = strSQL & "'" & NVE & "',"
    strSQL = strSQL & "'" & EXTIPI & "',"
    strSQL = strSQL & "'" & CFOP & "',"
    strSQL = strSQL & "'" & uCom & "',"
    strSQL = strSQL & F_GravaDec(qCom) & ","
    strSQL = strSQL & F_GravaDec(vUnCom) & ","
    strSQL = strSQL & F_GravaDec(vProd) & ","
    strSQL = strSQL & "'" & cEANTrib & "',"
    strSQL = strSQL & "'" & uTrib & "',"
    strSQL = strSQL & F_GravaDec(qTrib) & ","
    strSQL = strSQL & F_GravaDec(vUnTrib) & ","
    strSQL = strSQL & F_GravaDec(vFrete) & ","
    strSQL = strSQL & F_GravaDec(vSeg) & ","
    strSQL = strSQL & F_GravaDec(vDesc) & ","
    strSQL = strSQL & F_GravaDec(vOutro) & ","
    strSQL = strSQL & "'" & (indTot) & "')"

    conn.Execute strSQL
    
    
    GravaNFe_Item = F_BuscaUmCampo("max(id)", "nfe_item", "id_nfe=" & id_NFe)
trata_Erro:
    
End Function
Function GravaNFe_Emit(id_NFe As String, CNPJ As String, CPF As String, xNome As String) As Boolean
Dim strSQL As String
On Error GoTo trata_Erro
    strSQL = "insert into nfe_emit (id_nfe,  CNPJ,  CPF,  xNome)values("
    strSQL = strSQL & id_NFe & ","
    strSQL = strSQL & "'" & CNPJ & "',"
    strSQL = strSQL & "'" & CPF & "',"
    strSQL = strSQL & "'" & xNome & "')"
    conn.Execute strSQL
    GravaNFe_Emit = True
trata_Erro:

End Function
Function GravaNFe_Dest(id_NFe As String, CNPJ As String, CPF As String, idEstrangeiro As String, xNome As String) As Boolean
Dim strSQL As String
On Error GoTo trata_Erro
    strSQL = "insert into nfe_dest (id_nfe,  CNPJ,  CPF,  idEstrangeiro,xnome)values("
    strSQL = strSQL & id_NFe & ","
    strSQL = strSQL & "'" & CNPJ & "',"
    If CNPJ = "" Then
        strSQL = strSQL & "'" & CPF & "',"
    Else
        strSQL = strSQL & "'',"
    End If
    strSQL = strSQL & "'" & idEstrangeiro & "',"
    strSQL = strSQL & "'" & xNome & "')"
    conn.Execute strSQL
    GravaNFe_Dest = True
trata_Erro:

End Function

Function GravaNfe_enderEmit1(id_NFe As String, xLgr As String, nro As String, xCpl As String, _
    xBairro As String, cMun As String, xMun As String, UF As String, cep As String, _
    cPais As String, xPais As String, fone As String, ie As String, IEST As String, _
    im As String, CNAE As String, CRT As CRT) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro

    strSQL = "Insert into nfe_enderdest(id_nfe,  xLgr,  nro,  xCpl,  xBairro,  cMun,  "
    strSQL = strSQL & "xMun,  UF,  CEP,  cPais,  xPais,  fone,  IE,  IEST,  IM,  CNAE,  CRT)values("
    strSQL = strSQL & id_NFe & ","
    strSQL = strSQL & "'" & xLgr & "',"
    strSQL = strSQL & "'" & nro & "',"
    strSQL = strSQL & "'" & xCpl & "',"
    strSQL = strSQL & "'" & xBairro & "',"
    strSQL = strSQL & "'" & cMun & "',"
    strSQL = strSQL & "'" & xMun & "',"
    strSQL = strSQL & "'" & UF & "',"
    strSQL = strSQL & "'" & cep & "',"
    strSQL = strSQL & "'" & cPais & "',"
    strSQL = strSQL & "'" & xPais & "',"
    strSQL = strSQL & "'" & fone & "',"
    strSQL = strSQL & "'" & ie & "',"
    strSQL = strSQL & "'" & IEST & "',"
    strSQL = strSQL & "'" & im & "',"
    strSQL = strSQL & "'" & CNAE & "',"
    strSQL = strSQL & "'" & CRT & "') "
    conn.Execute strSQL
    
    GravaNfe_enderEmit1 = True
trata_Erro:
End Function


Function GravaNfe_enderdest(id_NFe As String, xLgr As String, nro As String, xCpl As String, _
    xBairro As String, cMun As String, xMun As String, UF As String, cep As String, _
    cPais As String, xPais As String, fone As String, ie As String, IndIEDest As String, _
    isuf As String, im As String, email As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro

    strSQL = "Insert into nfe_enderdest(id_nfe,  xLgr,  nro,  xCpl,  xBairro,  cMun,  "
    strSQL = strSQL & "xMun,  UF,  CEP,  cPais,  xPais,  fone,  IE, indIEDest,  ISUF, IM, "
    strSQL = strSQL & "email)values("
    strSQL = strSQL & id_NFe & ","
    strSQL = strSQL & "'" & xLgr & "',"
    strSQL = strSQL & "'" & nro & "',"
    strSQL = strSQL & "'" & xCpl & "',"
    strSQL = strSQL & "'" & xBairro & "',"
    strSQL = strSQL & "'" & cMun & "',"
    strSQL = strSQL & "'" & xMun & "',"
    strSQL = strSQL & "'" & UF & "',"
    strSQL = strSQL & "'" & F_SoNumero(cep) & "',"
    strSQL = strSQL & "'" & cPais & "',"
    strSQL = strSQL & "'" & xPais & "',"
    strSQL = strSQL & "'" & fone & "',"
    strSQL = strSQL & "'" & ie & "',"
    strSQL = strSQL & "'" & IndIEDest & "',"
    
    strSQL = strSQL & "'" & isuf & "',"
    strSQL = strSQL & "'" & im & "',"
    
    strSQL = strSQL & "'" & email & "') "
    conn.Execute strSQL
    
    GravaNfe_enderdest = True
trata_Erro:
End Function



Function GravaNfe_enderEmit(id_NFe As String, xLgr As String, nro As String, xCpl As String, _
    xBairro As String, cMun As String, xMun As String, UF As String, cep As String, _
    cPais As String, xPais As String, fone As String, ie As String, IEST As String, _
    im As String, CNAE As String, CRT As CRT) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro

    strSQL = "Insert into nfe_enderEmit(id_nfe,  xLgr,  nro,  xCpl,  xBairro,  cMun,  "
    strSQL = strSQL & "xMun,  UF,  CEP,  cPais,  xPais,  fone,  IE,  IEST,  IM,  CNAE,  CRT)values("
    strSQL = strSQL & id_NFe & ","
    strSQL = strSQL & "'" & xLgr & "',"
    strSQL = strSQL & "'" & nro & "',"
    strSQL = strSQL & "'" & xCpl & "',"
    strSQL = strSQL & "'" & xBairro & "',"
    strSQL = strSQL & "'" & cMun & "',"
    strSQL = strSQL & "'" & xMun & "',"
    strSQL = strSQL & "'" & UF & "',"
    strSQL = strSQL & "'" & F_SoNumero(cep) & "',"
    strSQL = strSQL & "'" & cPais & "',"
    strSQL = strSQL & "'" & xPais & "',"
    strSQL = strSQL & "'" & fone & "',"
    strSQL = strSQL & "'" & ie & "',"
    strSQL = strSQL & "'" & IEST & "',"
    strSQL = strSQL & "'" & im & "',"
    strSQL = strSQL & "'" & CNAE & "',"
    strSQL = strSQL & "'" & CRT & "') "
    conn.Execute strSQL
    
    GravaNfe_enderEmit = True
trata_Erro:
End Function

Function GravaNfe_item_imposto(id As String, id_NFe As String, vTotTrib As String, Tipo_Icms As String, _
    tipo_pis As String, tipo_cofins As String, tipo_ipi As String) As Boolean
 
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    
    strSQL = "insert into nfe_item_imposto(id , id_nfe, vTotTrib, tipo_icms, tipo_pis, tipo_cofins,tipo_ipi)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    strSQL = strSQL & F_GravaDec(vTotTrib) & "," 'taxa federativa
    strSQL = strSQL & "'" & Tipo_Icms & "',"
    strSQL = strSQL & "'" & tipo_ipis & "',"
    strSQL = strSQL & "'" & tipo_cofins & "',"
    strSQL = strSQL & "'" & tipo_ipi & "')"
    
    conn.Execute strSQL
    GravaNfe_item_imposto = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_icmssn102(id As String, id_NFe As String, orig As orig, csosn As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into nfe_item_imposto_icmssn102(id , id_nfe, orig, csosn)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "')"
    
    
    
    conn.Execute strSQL
    GravaNfe_item_imposto_icmssn102 = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_icmssn101(id As String, id_NFe As String, orig As orig, csosn As String, _
    pCredSN As String, vCredICMSSN As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into nfe_item_imposto_icmssn101(id , id_nfe, orig, csosn,pCredSN,vCredICMSSN)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & F_GravaDec(pCredSN) & ","
    strSQL = strSQL & F_GravaDec(vCredICMSSN) & ")"
    
    
    conn.Execute strSQL
    GravaNfe_item_imposto_icmssn101 = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_icmssn201(id As String, id_NFe As String, orig As orig, csosn As String, _
    modBCST As modBCST, pMVAST As String, pRedBCST As String, vBCST As String, pICMSST As String, _
    vICMSST As String, pCredSN As String, vCredICMSSN As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_icmssn201(id , id_nfe, orig, csosn,modBCST, pCredSN,pMVAST, "
    strSQL = strSQL & " pRedBCST, vBCST, pICMSST, vCredICMSSN,vICMSST)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & "'" & modBCST & "',"
    strSQL = strSQL & F_GravaDec(pCredSN) & ","
    strSQL = strSQL & F_GravaDec(pMVAST) & ","
    strSQL = strSQL & F_GravaDec(pRedBCST) & ","
    strSQL = strSQL & F_GravaDec(vBCST) & ","
    strSQL = strSQL & F_GravaDec(pICMSST) & ","
    strSQL = strSQL & F_GravaDec(vCredICMSSN) & ","
    strSQL = strSQL & F_GravaDec(vICMSST) & ")"
    
    conn.Execute strSQL
    GravaNfe_item_imposto_icmssn201 = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_icmssn202(id As String, id_NFe As String, orig As orig, csosn As String, _
    modBCST As modBCST, pMVAST As String, pRedBCST As String, vBCST As String, pICMSST As String, _
    vICMSST As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_icmssn202(id , id_nfe, orig, csosn,modBCST, pCredSN,pMVAST, "
    strSQL = strSQL & " pRedBCST, vBCST, pICMSST)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & "'" & modBCST & "',"
    strSQL = strSQL & F_GravaDec(pCredSN) & ","
    strSQL = strSQL & F_GravaDec(pMVAST) & ","
    strSQL = strSQL & F_GravaDec(pRedBCST) & ","
    strSQL = strSQL & F_GravaDec(vBCST) & ","
    strSQL = strSQL & F_GravaDec(pICMSST) & ")"
    
    
    conn.Execute strSQL
    GravaNfe_item_imposto_icmssn202 = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_icmssn500(id As String, id_NFe As String, orig As orig, csosn As String, _
    vBCSTRet As String, vICMSSTRet As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into nfe_item_imposto_icmssn500(id , id_nfe, orig, csosn,vBCSTRet,vICMSSTRet)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & F_GravaDec(vBCSTRet) & ","
    strSQL = strSQL & F_GravaDec(vICMSSTRet) & ")"
    
    
    conn.Execute strSQL
    GravaNfe_item_imposto_icmssn500 = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_icmssn900(id As String, id_NFe As String, orig As orig, csosn As String, _
    modBCST As modBCST, pMVAST As String, pRedBCST As String, vBCST As String, pICMSST As String, _
    vICMSST As String, pCredSN As String, vCredICMSSN As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_icmssn900(id , id_nfe, orig, csosn,modBCST, pCredSN,pMVAST, "
    strSQL = strSQL & " pRedBCST, vBCST, pICMSST, vCredICMSSN,vICMSST)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & "'" & modBCST & "',"
    strSQL = strSQL & F_GravaDec(pCredSN) & ","
    strSQL = strSQL & F_GravaDec(pMVAST) & ","
    strSQL = strSQL & F_GravaDec(pRedBCST) & ","
    strSQL = strSQL & F_GravaDec(vBCST) & ","
    strSQL = strSQL & F_GravaDec(pICMSST) & ","
    strSQL = strSQL & F_GravaDec(vCredICMSSN) & ","
    strSQL = strSQL & F_GravaDec(vICMSST) & ")"
    
    conn.Execute strSQL
    GravaNfe_item_imposto_icmssn900 = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_icmsST(id As String, id_NFe As String, orig As orig, csosn As String, _
    cst As String, vBCSTRet As String, vICMSSTRet As String, _
    vBCSTDest As String, vICMSSTDest As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_icmssn201(id , id_nfe, orig, csosn,cst, vBCSTRet,vICMSSTRet, "
    strSQL = strSQL & "vBCSTDest, vICMSSTDest)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBCSTRet) & ","
    strSQL = strSQL & F_GravaDec(vICMSSTRet) & ","
    strSQL = strSQL & F_GravaDec(vBCSTDest) & ","
    strSQL = strSQL & F_GravaDec(vICMSSTDest) & ")"
    
    conn.Execute strSQL
    GravaNfe_item_imposto_icmsST = True
trata_Erro:
    
End Function
Function GravaNfe_IcmsTot(id_NFe As String, vBC As String, vICMS As String, vICMSDeson As String, _
    vBCST As String, vST As String, vProd As String, vFrete As String, vSeg As String, vDesc, _
    vII As String, vIPI As String, vPIS As String, vCOFINS As String, vOutro As String, _
    vNF As String, vTotTrib As String) As Boolean
    On Error GoTo trata_Erro
    strSQL = "insert into nfe_icmstot(id_nfe , vBC, vICMS, vICMSDeson, vBCST, vST, vProd, vFrete, "
    strSQL = strSQL & "vSeg, vDesc, vII, vIPI, vPIS, vCOFINS, vOutro, vNF, vTotTrib)values("
    strSQL = strSQL & id_NFe & ","
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(vICMS) & ","
    strSQL = strSQL & F_GravaDec(vICMSDeson) & ","
    strSQL = strSQL & F_GravaDec(vBCST) & ","
    strSQL = strSQL & F_GravaDec(vST) & ","
    strSQL = strSQL & F_GravaDec(vProd) & ","
    strSQL = strSQL & F_GravaDec(vFrete) & ","
    strSQL = strSQL & F_GravaDec(vSeg) & ","
    strSQL = strSQL & F_GravaDec(vDesc) & ","
    strSQL = strSQL & F_GravaDec(vII) & ","
    strSQL = strSQL & F_GravaDec(vIPI) & ","
    strSQL = strSQL & F_GravaDec(vPIS) & ","
    strSQL = strSQL & F_GravaDec(vCOFINS) & ","
    strSQL = strSQL & F_GravaDec(vOutro) & ","
    strSQL = strSQL & F_GravaDec(vNF) & ","
    strSQL = strSQL & F_GravaDec(vTotTrib) & ")"
    conn.Execute strSQL
    GravaNfe_IcmsTot = True
trata_Erro:
End Function

Function GravaNfe_item_imposto_IPI(id As String, id_NFe As String, clEnq As String, CNPJProd As String, _
    cSelo As String, qSelo As String, cEnq As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_ipi(id , id_nfe, clEnq, CNPJProd,cSelo, qSelo ,cEnq)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & clEnq & "',"
    strSQL = strSQL & "'" & CNPJProd & "',"
    strSQL = strSQL & "'" & cSelo & "',"
    strSQL = strSQL & "'" & qSelo & "',"
    strSQL = strSQL & "'" & cEnq & "')"
    conn.Execute strSQL
    GravaNfe_item_imposto_IPI = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_IPIInt(id As String, id_NFe As String, cst As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_ipiInt(id , id_nfe, cst)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & cst & "')"
    
    conn.Execute strSQL
    GravaNfe_item_imposto_IPIInt = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_IPITrib(id As String, id_NFe As String, cst As String, vBC As String, _
    pIPI As String, qUnid As String, vUnid As String, vIPI As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_ipiTrib(id , id_nfe, cst, vbc,pipi, qunid ,vunid,vipi)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pIPI) & ","
    strSQL = strSQL & F_GravaDec(qUnid) & ","
    strSQL = strSQL & F_GravaDec(vUnid) & ","
    strSQL = strSQL & F_GravaDec(vIPI) & ")"
    conn.Execute strSQL
    GravaNfe_item_imposto_IPITrib = True
trata_Erro:
    
End Function


Function GravaNfe_item_imposto_PISAliq(id As String, id_NFe As String, cst As String, vBC As String, _
    pPIS As String, vPIS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_pisaliq(id , id_nfe, cst, vbc,ppis, vpis)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pPIS) & ","
    
    strSQL = strSQL & F_GravaDec(vPIS) & ")"
    conn.Execute strSQL
    GravaNfe_item_imposto_PISAliq = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_PISNT(id As String, id_NFe As String, cst As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_pisnt(id , id_nfe, cst, vbc,pPIS, qBCPRod  ,vAliqProd,vPIS)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & cst & "')"
    
    conn.Execute strSQL
    GravaNfe_item_imposto_PISNT = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_PISOutr(id As String, id_NFe As String, cst As String, vBC As String, _
    pPIS As String, qBCProd As String, vAliqProd As String, vPIS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_pisoutr(id , id_nfe, cst, vbc,pPIS, qBCPRod  ,vAliqProd,vPIS)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & F_GravaDec(cst) & ","
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pPIS) & ","
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ","
    strSQL = strSQL & F_GravaDec(vPIS) & ")"
    conn.Execute strSQL
    GravaNfe_item_imposto_PISOutr = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_PISST(id As String, id_NFe As String, vBC As String, _
    pPIS As String, qBCProd As String, vAliqProd As String, vPIS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_pisst(id , id_nfe, vbc,pPIS, qBCPRod  ,vAliqProd,vPIS)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pPIS) & ","
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ","
    strSQL = strSQL & "'" & vPIS & "')"
    conn.Execute strSQL
    GravaNfe_item_imposto_PISST = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_PISQtde(id As String, id_NFe As String, cst As String, _
    qBCProd As String, vAliqProd As String, vPIS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_pisQTde(id,id_nfe, cst,qBCPRod  ,vAliqProd,vPIS)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    
    strSQL = strSQL & "'" & cst & "',"
    
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ","
    strSQL = strSQL & F_GravaDec(vPIS) & ")"
    conn.Execute strSQL
    GravaNfe_item_imposto_PISQtde = True
trata_Erro:
    
End Function




Function GravaNfe_item_imposto_COFINSAliq(id As String, id_NFe As String, cst As String, vBC As String, _
    pCOFINS As String, vCOFINS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_cofinsaliq(id , id_nfe, cst, vbc,pCofins,vCofins)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pCOFINS) & ","
    strSQL = strSQL & F_GravaDec(vCOFINS) & ")"
    
    conn.Execute strSQL
    GravaNfe_item_imposto_COFINSAliq = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_COFINSOutr(id As String, id_NFe As String, cst As String, vBC As String, _
    pCOFINS As String, qBCProd As String, vAliqProd As String, _
    vCOFINS As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_cofinsoutr(id , id_nfe, cst, vbc,pCofins,vCofins,qBCProd,vAliqProd)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pCOFINS) & ","
    strSQL = strSQL & F_GravaDec(vCOFINS) & ","
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ")"
    conn.Execute strSQL
    GravaNfe_item_imposto_COFINSOutr = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_COFINSQtde(id As String, id_NFe As String, cst As String, _
    qBCProd As String, vAliqProd As String, _
    vCOFINS As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_cofinsoutr(id , id_nfe, cst, qBCProd,vAliqProd, vcofins)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ","
    strSQL = strSQL & F_GravaDec(vCOFINS) & ")"
    conn.Execute strSQL
    GravaNfe_item_imposto_COFINSQtde = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_COFINSNT(id As String, id_NFe As String, cst As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_cofinsoutr(id , id_nfe, cst)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & cst & "')"
    

    conn.Execute strSQL
    GravaNfe_item_imposto_COFINSNT = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_COFINSST(id As String, id_NFe As String, cst As String, vBC As String, _
    pCOFINS As String, vCOFINS As String, qBCProd As String, vAliqProd As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_cofinsoutr(id , id_nfe, cst, vbc,pCofins,vCofins,qBCProd,vAliqProd)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_NFe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pCOFINS) & ","
    strSQL = strSQL & F_GravaDec(vCOFINS) & ","
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ")"
    conn.Execute strSQL
    GravaNfe_item_imposto_COFINSST = True
trata_Erro:
    
End Function


Function GravaNFe_Pag(id_NFe As String, tPag As String, vPag As String, card As String, _
    CNPJ As String, tBand As String, cAut As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_pag(id_nfe , tPag, vPag, card, CNPJ, tBand, cAut) values ( "
     strSQL = strSQL & id_NFe & ","
     strSQL = strSQL & "'" & tPag & "', "
     strSQL = strSQL & F_GravaDec(vPag) & ", "
     strSQL = strSQL & "'" & card & "',"
     strSQL = strSQL & "'" & CNPJ & "', "
     strSQL = strSQL & "'" & tBand & "',"
     strSQL = strSQL & "'" & cAut & "') "
     conn.Execute strSQL
     
     GravaNFe_Pag = True
trata_Erro:
End Function
Function GravaNFe_ProcRec(id_NFe As String, nProc As String, indProc As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_ProcRec(id_nfe , nproc, indproc) values ( "
     strSQL = strSQL & id_NFe & ","
     
     strSQL = strSQL & "'" & nProc & "',"
     strSQL = strSQL & "'" & indProc & "') "
     conn.Execute strSQL
     
     GravaNFe_ProcRec = True
trata_Erro:
End Function
Function GravaNFe_Transp(id_NFe As String, modFrete As modFrete) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_transp(id_nfe , modFrete) values ( "
     strSQL = strSQL & id_NFe & ","
     
     strSQL = strSQL & "'" & modFrete & "')"
        
     conn.Execute strSQL
     
     GravaNFe_Transp = True
trata_Erro:
End Function

Function GravaNFe_Transp_Reboque(id_NFe As String, placa As String, UF As String, RNTC As String, _
    vagao As String, balsa As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_transp_reboque(id_nfe,  placa,  UF,  RNTC,  vagao,  balsa ) values ( "
     strSQL = strSQL & id_NFe & ","
     
     strSQL = strSQL & "'" & placa & "',"
        strSQL = strSQL & "'" & UF & "',"
        strSQL = strSQL & "'" & RNTC & "',"
        strSQL = strSQL & "'" & vagao & "',"
        strSQL = strSQL & "'" & balsa & "')"
     conn.Execute strSQL
     
     GravaNFe_Transp_Reboque = True
trata_Erro:
End Function

Function GravaNFe_Transp_Transportadora(id_NFe As String, CNPJ As String, CPF As String, xNome As String, _
    ie As String, xEnder As String, xMun As String, UF As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_transp_transportadora(id_nfe,  cnpj,  cpf,  xnome,  ie,  xender, xmun, uf) values ( "
     strSQL = strSQL & id_NFe & ","
     
     strSQL = strSQL & "'" & CNPJ & "',"
        strSQL = strSQL & "'" & CPF & "',"
        strSQL = strSQL & "'" & xNome & "',"
        strSQL = strSQL & "'" & xEnder & "',"
        strSQL = strSQL & "'" & xMun & "',"
        strSQL = strSQL & "'" & UF & "')"
     conn.Execute strSQL
     
     GravaNFe_Transp_Transportadora = True
trata_Erro:
End Function

Function GravaNFe_Transp_VeicTransp(id_NFe As String, placa As String, UF As String, RNTC As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_TranspVeicTransp(id_nfe,  placa,  uf,  rntc) values ( "
     strSQL = strSQL & id_NFe & ","
     
     strSQL = strSQL & "'" & placa & "',"
        strSQL = strSQL & "'" & UF & "',"
        strSQL = strSQL & "'" & RNTC & "')"
        
     conn.Execute strSQL
     
     GravaNFe_Transp_VeicTransp = True
trata_Erro:
End Function
Function GravaNFe_Transp_Vol(id_NFe As String, qVol As String, ESP As String, Marca As String, nVol As String, _
    pesoL As String, pesoB As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_Transp_Vol(id_nfe,  qvol,  esp,  marca, nvol,pesol,pesob) values ( "
     strSQL = strSQL & id_NFe & ","
     strSQL = strSQL & F_GravaDec(qVol) & ", "
     strSQL = strSQL & "'" & ESP & "',"
        strSQL = strSQL & "'" & Marca & "',"
        strSQL = strSQL & "'" & nVol & "',"
        strSQL = strSQL & F_GravaDec(pesoL) & ", "
        strSQL = strSQL & F_GravaDec(pesoB) & ")"
     conn.Execute strSQL
     
     GravaNFe_Transp_Vol = True
trata_Erro:
End Function

Function GravaNFe_Transp_RetTransp(id_NFe As String, vServ As String, vBCRet As String, pICMSRet As String, _
    vICMSRet, CFOP As String, cMunFG As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_Transp_RetTransp(id_nfe,vserv,vBCRet,pICMSRet,vICMSRet,cfop,cMunFG) values ( "
     strSQL = strSQL & id_NFe & ","
     strSQL = strSQL & F_GravaDec(vServ) & ", "
     strSQL = strSQL & F_GravaDec(vBCRet) & ", "
     strSQL = strSQL & F_GravaDec(pICMSRet) & ", "
     strSQL = strSQL & "'" & CFOP & "',"
        strSQL = strSQL & "'" & cMunFG & "')"
        
     conn.Execute strSQL
     
     GravaNFe_Transp_RetTransp = True
trata_Erro:
End Function
Function GravaNFe_InfAdc(id_NFe As String, infAdFisco As String, infCpl As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_InfAdic(id_nfe,infAdFisco ,infCpl) values ( "
     strSQL = strSQL & id_NFe & ","
     strSQL = strSQL & "'" & infAdFisco & "', "
     strSQL = strSQL & "'" & infCpl & "') "
     
     
        
     conn.Execute strSQL
     
     GravaNFe_InfAdc = True
trata_Erro:
End Function

Function GravaNFe_Cobr_Dup(id_NFe As String, nDup As String, dVenc As String, vDup As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_Cobr_Dup(id_nfe,nDup,dVenc,vDup) values ( "
     strSQL = strSQL & id_NFe & ","
     strSQL = strSQL & "'" & nDup & "', "
     strSQL = strSQL & F_GravaData(dVenc) & ", "
     strSQL = strSQL & F_GravaDec(vDup) & ") "
     
        
     conn.Execute strSQL
     
     GravaNFe_Cobr_Dup = True
trata_Erro:
End Function

Function GravaNFe_Cobr_Fat(id_NFe As String, nFat As String, vOrig As String, vDesc As String, vLiq As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_Cobr_Dup(id_nfe,nFat,vOrig,vDesc,vLiq) values ( "
     strSQL = strSQL & id_NFe & ","
     strSQL = strSQL & "'" & nFat & "', "
     strSQL = strSQL & F_GravaDec(vOrig) & ", "
     strSQL = strSQL & F_GravaDec(vDesc) & ", "
     strSQL = strSQL & F_GravaDec(vLiq) & ") "
        
     conn.Execute strSQL
     
     GravaNFe_Cobr_Fat = True
trata_Erro:
End Function

