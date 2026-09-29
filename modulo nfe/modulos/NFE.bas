Attribute VB_Name = "NFe"

Public Sub IniciaConfiguracoesNFe()
    On Error GoTo trata_Erro
    Dim amb As String
    amb = F_BuscaUmCampo("ambiente_nfe", TAB_PARAMETROS, "") ' "2"  '1-produção  2-homologação
    If strTipo_IcmsNFe = "" Then
        strTipo_IcmsNFe = F_BuscaUmCampo("tipoicms_nfe", TAB_PARAMETROS, "")  '
    End If
    
    If strTipo_IcmsNFe = "500" Or strTipo_IcmsNFe = "ICMSSN500" Then
        strTipo_IcmsNFe = "ICMSSN500"
    ElseIf strTipo_IcmsNFe = "102" Or strTipo_IcmsNFe = "ICMSSN102" Then
        strTipo_IcmsNFe = "ICMSSN102"
    ElseIf strTipo_IcmsNFe = "900" Or strTipo_IcmsNFe = "ICMSSN900" Then
        strTipo_IcmsNFe = "ICMSSN900"
    End If
    
    If amb = "" Then
        ambiente = homologacao
    Else
        ambiente = amb
    End If
    sVersaoNfe = F_BuscaUmCampo("versao_nfe", TAB_PARAMETROS, "")
    strEnderecoConsultaChaveDeAcesso = "www.sefaz.go.gov.br"
    strCRT_emp = SIMPLES_NACIONAL
    strCST_COFINS = "49"
    strCST_PIS = "49"
    
        
    ValorMedioImpostoEstadual = 6.21
    ValorMedioImpostoFederal = 4.98
    
    CaminhoArquivoConfigImpressaoNFCe = App.path & "\printnfce.xml"
    CaminhoArquivoConfigImpressaoCopiaVenda = App.path & "\printcopiavenda.xml"
    
    
    
    ImpressoraPadraoCupomNFCe = BuscaImpressoraPadraoNFCe
   On Error Resume Next
    ImpressoraCopiaVenda = BuscaImpressoraPadraoCopiaVenda
    
    
    Exit Sub
trata_Erro:
    
End Sub
Sub AlteraStatusNfe(id As String, status As String)
    strSQL = "update nfe set status='" & status & "'"
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
        
End Sub
Sub AlteraChaveNfe(id As String, Chave As String)
    strSQL = "update nfe set chave='" & Chave & "', "
    strSQL = strSQL & "cDV='" & Right(Chave, 1) & "',"
    strSQL = strSQL & "tpEmis='" & Mid(Chave, 35, 1) & "',"
    strSQL = strSQL & "nNF='" & Mid(Chave, 26, 9) & "'"
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
        
End Sub

Sub AlteraDadosFinaisNFe(id As String, status As String, caminhoXML As String, protocolo As String, recibo As String, DigestValue As String)
    On Error GoTo trata_Erro
    strSQL = "update nfe set status='" & status & "',"
    strSQL = strSQL & "caminhoxml='" & caminhoXML & "',"
    strSQL = strSQL & "protocolo='" & protocolo & "',"
    strSQL = strSQL & "recibo='" & recibo & "', "
    strSQL = strSQL & "digestValue='" & DigestValue & "' "
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
    Exit Sub
    
trata_Erro:
    
End Sub

Sub AlteraInfAdicNFe(id_nfe As String, infCpl As String)
    On Error GoTo trata_Erro
    
    strSQL = "replace into nfe_infadic(id_nfe,infcpl)values(" & id_nfe & ",'" & infCpl & "') "
    
    
    Conn.Execute strSQL
    Exit Sub
    
trata_Erro:
    
End Sub
Function NFE_GravaNFE(num_ven As String, Chave As String, cUF As String, cNF As String, natOp As String, indPag As String, _
    modelo As String, serie As String, nNF As String, dhEmi As String, dhSaiEnt As String, _
    tpNf As tpNf, idDest As idDest, cMunFG As String, tpImp As tpImp, _
    tpEmis As tpEmis, cDV As String, tpAmb As tpAmb, finNFe As TfinNfe, indFinal As indFinal, _
    indPres As indPres, procEmi As procEmi, verProc As String, dhCont As String, xJust As String, Optional status As String = "1", _
    Optional indIntermed As String = "0", Optional cod_emp As String = "01") As String
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into nfe(num_ven,chave,cUF , cNF, natOp, indPag, Modelo, serie, nNF, dhEmi, dhSaiEnt, "
    strSQL = strSQL & "tpNF, idDest, cMunFG, tpImp, tpEmis, cDV, tpAmb, finNFe, indFinal, indPres, procEmi, "
    strSQL = strSQL & " verProc, dhCont, xJust,status,indIntermed,cod_usu,cod_emp)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & "'" & Chave & "',"
    strSQL = strSQL & "'" & cUF & "',"
    strSQL = strSQL & "'" & cNF & "',"
    strSQL = strSQL & "'" & natOp & "', "
    strSQL = strSQL & "'" & indPag & "', "
    strSQL = strSQL & "'" & modelo & "',"
    strSQL = strSQL & "'" & serie & "',"
    strSQL = strSQL & "'" & nNF & "', "
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
    strSQL = strSQL & "'" & xJust & "', "
    
    strSQL = strSQL & "'" & status & "',"  'status
    strSQL = strSQL & "'" & indIntermed & "',"   'status
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & cod_emp & "') "
    Conn.Execute strSQL
    
    NFE_GravaNFE = F_BuscaUmCampo(" max(id)", "nfe", "")
trata_Erro:
    
    
End Function
Function GravaNFe_Ref(id As String, refNfe As String) As Boolean
    Dim strSQL As String
    Dim c As Collection
    Dim iCount As Integer
    Dim sep As String
    On Error GoTo trata_Erro
    strSQL = "delete from nfe_ref where id=" & id
    Conn.Execute strSQL
    
    sep = ",;- " & CrLf
    Set c = F_SeparaTexto(Trim(refNfe), sep, False)
    
    For iCount = 1 To c.Count
        strSQL = "insert into Nfe_Ref(id,refNfe)values("
        strSQL = strSQL & id & ","
        strSQL = strSQL & "'" & Trim(c(iCount)) & "')"
        Conn.Execute strSQL
        
    Next
    GravaNFe_Ref = True
trata_Erro:
    GravaNFe_Ref = False
End Function
Function BuscaNFe_Ref(id As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from nfe_ref where id=" & id
    abreObjetoAdo rPesq, strSQL
    strSQL = ""
    Do While Not rPesq.EOF
        strSQL = strSQL & rPesq!refNfe
        
        If Not rPesq.EOF Then
            strSQL = strSQL & CrLf
        End If
        
        
        rPesq.MoveNext
    Loop
    BuscaNFe_Ref = strSQL
End Function
Function DeletaNFe_Ref(id As String, refNfe As String) As Boolean
    Dim strSQL As String
    strSQL = "delete from Nfe_Ref where id="
    strSQL = strSQL & id & " "
    strSQL = strSQL & " and refNfe='" & refNfe & "')"
    Conn.Execute strSQL
    DeletaNFe_Ref = True
    
End Function

Function NFE_AlteraDadosNfe(id As String, UF As String, natOp As String, indPag As String, _
    modelo As String, dhEmi As String, dhSaiEnt As String, _
    tpNf As tpNf, idDest As idDest, tpImp As tpImp, _
    tpEmis As tpEmis, tpAmb As tpAmb, finNFe As TfinNfe, indFinal As indFinal, _
    indPres As indPres, procEmi As procEmi, verProc As String, dhCont As String, xJust As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update nfe set "
    '(num_ven,chave,cUF , cNF, natOp, indPag, Modelo, serie, nNF, dhEmi, dhSaiEnt, "
    'strSQL = strSQL & "tpNF, idDest, cMunFG, tpImp, tpEmis, cDV, tpAmb, finNFe, indFinal, indPres, procEmi, "
    'strSQL = strSQL & " verProc, dhCont, xJust,status,cod_usu)values("
    'strSQL = strSQL & "num_ven='" & num_ven & "',"
    'strSQL = strSQL & "'" & chave & "',"
    strSQL = strSQL & "cUF='" & UF & "',"
    'strSQL = strSQL & "'" & cNF & "',"
    strSQL = strSQL & "natop='" & natOp & "', "
    strSQL = strSQL & "indpag='" & indPag & "', "
    strSQL = strSQL & "modelo='" & modelo & "',"
    'strSQL = strSQL & "'" & serie & "',"
    'strSQL = strSQL & "'" & nNF & "', "
    strSQL = strSQL & "dhEmi=" & GravaDataNfe(dhEmi) & ", "
    strSQL = strSQL & "dhSaiEnt=" & GravaDataNfe(dhSaiEnt) & ", "
    strSQL = strSQL & "tpnf='" & tpNf & "', "
    strSQL = strSQL & "idDest='" & idDest & "', "
'    strSQL = strSQL & "cmunfg='" & cMunFG & "', "
    strSQL = strSQL & "tpImp='" & tpImp & "', "
    strSQL = strSQL & "tpEmis='" & tpEmis & "', "
    'strSQL = strSQL & "'" & cDV & "', "
    strSQL = strSQL & "tpAmb='" & tpAmb & "', "
    strSQL = strSQL & "finNfe='" & finNFe & "', "
    strSQL = strSQL & "indFinal='" & indFinal & "', "
    strSQL = strSQL & "indPres='" & indPres & "', "
    strSQL = strSQL & "procemi='" & procEmi & "', "
    strSQL = strSQL & "verproc='" & verProc & "', "
    strSQL = strSQL & "dhCOnt=" & GravaDataNfe(dhCont) & ", "
    strSQL = strSQL & "xjust='" & xJust & "',  "
    
    'strSQL = strSQL & "'1', " 'status
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
    
    NFE_AlteraDadosNfe = True
trata_Erro:
    
    
End Function
Function NFE_AlteraDadosNfeCompleta(id As String, Chave As String, serie As String, nNF As String, _
    cNF As String, cDV As String, natOp As String, indPag As String, _
    modelo As String, dhEmi As String, dhSaiEnt As String, _
    tpNf As tpNf, idDest As idDest, tpImp As tpImp, _
    tpEmis As tpEmis, tpAmb As tpAmb, finNFe As TfinNfe, indFinal As indFinal, _
    indPres As indPres, procEmi As procEmi, verProc As String, dhCont As String, xJust As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update nfe set "
    '(num_ven,chave,cUF , cNF, natOp, indPag, Modelo, serie, nNF, dhEmi, dhSaiEnt, "
    'strSQL = strSQL & "tpNF, idDest, cMunFG, tpImp, tpEmis, cDV, tpAmb, finNFe, indFinal, indPres, procEmi, "
    'strSQL = strSQL & " verProc, dhCont, xJust,status,cod_usu)values("
    'strSQL = strSQL & "num_ven='" & num_ven & "',"
    strSQL = strSQL & "'" & Chave & "',"
    strSQL = strSQL & "'" & cUF & "',"
    strSQL = strSQL & "'" & cNF & "',"
    strSQL = strSQL & "natop='" & natOp & "', "
    strSQL = strSQL & "indpag='" & indPag & "', "
    strSQL = strSQL & "modelo='" & modelo & "',"
    strSQL = strSQL & "'" & serie & "',"
    strSQL = strSQL & "'" & nNF & "', "
    strSQL = strSQL & "dhEmi=" & GravaDataNfe(dhEmi) & ", "
    strSQL = strSQL & "dhSaiEnt=" & GravaDataNfe(dhSaiEnt) & ", "
    strSQL = strSQL & "tpnf='" & tpNf & "', "
    strSQL = strSQL & "idDest='" & idDest & "', "
'    strSQL = strSQL & "cmunfg='" & cMunFG & "', "
    strSQL = strSQL & "tpImp='" & tpImp & "', "
    strSQL = strSQL & "tpEmis='" & tpEmis & "', "
    strSQL = strSQL & "'" & cDV & "', "
    strSQL = strSQL & "tpAmb='" & tpAmb & "', "
    strSQL = strSQL & "finNfe='" & finNFe & "', "
    strSQL = strSQL & "indFinal='" & indFinal & "', "
    strSQL = strSQL & "indPres='" & indPres & "', "
    strSQL = strSQL & "procemi='" & procEmi & "', "
    strSQL = strSQL & "verproc='" & verProc & "', "
    strSQL = strSQL & "dhCOnt=" & GravaDataNfe(dhCont) & ", "
    strSQL = strSQL & "xjust='" & xJust & "', "
    
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
    
    NFE_AlteraDadosNfeCompleta = True
trata_Erro:
    
    
End Function

Function GravaNfeItemComb(id_item As String, id_nfe As String, cProdAnp As String, descAnp As String, UFCons As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    If UFCons = "" Then UFCons = "GO"
    
    strSQL = "replace into nfe_Item_comb (id_item,id_nfe,cProdANP,descANP,UFCons)values("
    strSQL = strSQL & "'" & id_item & "', "
    strSQL = strSQL & "'" & id_nfe & "', "
    strSQL = strSQL & "'" & cProdAnp & "', "
    strSQL = strSQL & "'" & descAnp & "', "
    strSQL = strSQL & "'" & UFCons & "') "
    Conn.Execute strSQL
    GravaNfeItemComb = True
    
    
    Exit Function
trata_Erro:
    GravaNfeItemComb = False
End Function

Function GravaNFe_Item(id_nfe As String, cProd As String, cEAN As String, xProd As String, _
    ncm As String, NVE As String, EXTIPI As String, CFOP As String, uCom As String, qCom As String, _
    vUnCom As String, vProd As String, cEANTrib As String, uTrib As String, qTrib As String, _
    vUnTrib As String, vFrete As String, vSeg As String, vDesc As String, vOutro As String, _
    indTot As indTot, cEst As String, infAdProd As String)
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item( id_nfe, cProd, cEAN, xProd, NCM, NVE, EXTIPI, CFOP, uCom, qCom, vUnCom, vProd, "
    strSQL = strSQL & " cEANTrib, uTrib, qTrib, vUnTrib, vFrete, vSeg, vDesc, vOutro, CEST, indTot,infAdProd) values ("

    strSQL = strSQL & "'" & id_nfe & "',"
    strSQL = strSQL & "'" & cProd & "',"
    strSQL = strSQL & "'" & cEAN & "',"
    strSQL = strSQL & "'" & xProd & "',"
    strSQL = strSQL & "'" & Left(ncm, 8) & "',"
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
    strSQL = strSQL & "'" & cEst & "',"
    strSQL = strSQL & "'" & (indTot) & "',"
    strSQL = strSQL & "'" & (infAdProd) & "')"

    Conn.Execute strSQL
    
    
    GravaNFe_Item = F_BuscaUmCampo("max(id)", "nfe_item", "id_nfe=" & id_nfe)
    Exit Function
trata_Erro:
    P_MostraErro
End Function
Function AlteraNFe_Item(id As String, cProd As String, cEAN As String, _
    xProd As String, ncm As String, NVE As String, EXTIPI As String, CFOP As String, uCom As String, _
    qCom As String, vUnCom As String, vProd As String, cEANTrib As String, uTrib As String, qTrib As String, _
    vUnTrib As String, vFrete As String, vSeg As String, vDesc As String, vOutro As String, _
    indTot As indTot, cEst As String, infAdProd As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "update nfe_item set "

    
    strSQL = strSQL & "cprod='" & cProd & "',"
    strSQL = strSQL & "cean='" & cEAN & "',"
    strSQL = strSQL & "xprod='" & xProd & "',"
    strSQL = strSQL & "ncm='" & ncm & "',"
    strSQL = strSQL & "nve='" & NVE & "',"
    strSQL = strSQL & "extipi='" & EXTIPI & "',"
    strSQL = strSQL & "cfop='" & CFOP & "',"
    strSQL = strSQL & "ucom='" & uCom & "',"
    strSQL = strSQL & "qcom=" & F_GravaDec(qCom) & ","
    strSQL = strSQL & "vuncom=" & F_GravaDec(vUnCom) & ","
    strSQL = strSQL & "vprod=" & F_GravaDec(vProd) & ","
    strSQL = strSQL & "ceantrib='" & cEANTrib & "',"
    strSQL = strSQL & "utrib='" & uTrib & "',"
    strSQL = strSQL & "qtrib=" & F_GravaDec(qTrib) & ","
    strSQL = strSQL & "vuntrib=" & F_GravaDec(vUnTrib) & ","
    strSQL = strSQL & "vfrete=" & F_GravaDec(vFrete) & ","
    strSQL = strSQL & "vseg=" & F_GravaDec(vSeg) & ","
    strSQL = strSQL & "vdesc=" & F_GravaDec(vDesc) & ","
    strSQL = strSQL & "voutro=" & F_GravaDec(vOutro) & ","
    strSQL = strSQL & "cest='" & cEst & "',"
    strSQL = strSQL & "infAdProd='" & infAdProd & "',"
    strSQL = strSQL & "indtot='" & (Left(indTot, 1)) & "'"
    strSQL = strSQL & " where id=" & id

    Conn.Execute strSQL
    AlteraNFe_Item = True
    
    
    Exit Function
trata_Erro:
    P_MostraErro
End Function

Function GravaNFe_Emit(id_nfe As String, CNPJ As String, cpf As String, xNome As String, xFant As String) As Boolean
Dim strSQL As String
On Error GoTo trata_Erro
    strSQL = "insert into nfe_emit (id_nfe,  CNPJ,  CPF,  xNome, xFant)values("
    strSQL = strSQL & id_nfe & ","
    strSQL = strSQL & "'" & CNPJ & "',"
    strSQL = strSQL & "'" & cpf & "',"
    strSQL = strSQL & "'" & xNome & "',"
    strSQL = strSQL & "'" & xFant & "')"
    Conn.Execute strSQL
    GravaNFe_Emit = True
trata_Erro:

End Function
Function GravaNFe_Dest(id_nfe As String, CNPJ As String, cpf As String, idEstrangeiro As String, xNome As String) As Boolean
Dim strSQL As String
On Error GoTo trata_Erro
    strSQL = "insert into nfe_dest (id_nfe,  CNPJ,  CPF,  idEstrangeiro,xnome)values("
    strSQL = strSQL & id_nfe & ","
    strSQL = strSQL & "'" & CNPJ & "',"
    
    strSQL = strSQL & "'" & cpf & "',"
    
    strSQL = strSQL & "'" & idEstrangeiro & "',"
    strSQL = strSQL & "'" & Left(xNome, 50) & "')"
    Conn.Execute strSQL
    GravaNFe_Dest = True
trata_Erro:

End Function

Function AlteraNFe_Dest(id_nfe As String, CNPJ As String, cpf As String, idEstrangeiro As String, xNome As String) As Boolean
Dim strSQL As String
On Error GoTo trata_Erro
    Dim s As String
    s = F_SoNumero(CNPJ)
    strSQL = "update nfe_dest set " '(id_nfe,  CNPJ,  CPF,  idEstrangeiro,xnome)values("
    

    If Len(s) > 11 Then
        strSQL = strSQL & "cnpj='" & F_SoNumero(CNPJ) & "',"
    End If
    s = F_SoNumero(cpf)
    If Len(s) = 11 Then
        strSQL = strSQL & "cpf='" & F_SoNumero(cpf) & "',"
    End If
            
    strSQL = strSQL & "idestrangeiro='" & idEstrangeiro & "',"
    strSQL = strSQL & "xnome='" & xNome & "' "
    strSQL = strSQL & " where id_nfe=" & id_nfe
    Conn.Execute strSQL
    AlteraNFe_Dest = True
Exit Function
trata_Erro:
MsgBox Err.Description, vbCritical

End Function



Function GravaNfe_enderEmit1(id_nfe As String, xLgr As String, nro As String, xCpl As String, _
    xBairro As String, cMun As String, xMun As String, UF As String, cep As String, _
    cPais As String, xPais As String, fone As String, IE As String, iEst As String, _
    IM As String, CNAE As String, CRT As CRT) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro

    strSQL = "Insert into nfe_enderdest(id_nfe,  xLgr,  nro,  xCpl,  xBairro,  cMun,  "
    strSQL = strSQL & "xMun,  UF,  CEP,  cPais,  xPais,  fone,  IE,  IEST,  IM,  CNAE,  CRT)values("
    strSQL = strSQL & id_nfe & ","
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
    strSQL = strSQL & "'" & IE & "',"
    strSQL = strSQL & "'" & iEst & "',"
    strSQL = strSQL & "'" & IM & "',"
    strSQL = strSQL & "'" & CNAE & "',"
    strSQL = strSQL & "'" & CRT & "') "
    Conn.Execute strSQL
    
    GravaNfe_enderEmit1 = True
trata_Erro:
End Function


Function GravaNfe_enderdest(id_nfe As String, xLgr As String, nro As String, xCpl As String, _
    xBairro As String, cMun As String, xMun As String, UF As String, cep As String, _
    cPais As String, xPais As String, fone As String, IE As String, IndIEDest As String, _
    isuf As String, IM As String, email As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    If IndIEDest = "9" Then
        IE = "" 'se for não contribuinte, não deve informar IE
    End If


    strSQL = "Insert into nfe_enderdest(id_nfe,  xLgr,  nro,  xCpl,  xBairro,  cMun,  "
    strSQL = strSQL & "xMun,  UF,  CEP,  cPais,  xPais,  fone,  IE, indIEDest,  ISUF, IM, "
    strSQL = strSQL & "email)values("
    strSQL = strSQL & id_nfe & ","
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
    strSQL = strSQL & "'" & IE & "',"
    strSQL = strSQL & "'" & IndIEDest & "',"
    
    strSQL = strSQL & "'" & isuf & "',"
    strSQL = strSQL & "'" & IM & "',"
    
    strSQL = strSQL & "'" & email & "') "
    Conn.Execute strSQL
    
    GravaNfe_enderdest = True
trata_Erro:
End Function
Function AlteraNfe_enderdest(id_nfe As String, xLgr As String, nro As String, xCpl As String, _
    xBairro As String, cMun As String, xMun As String, UF As String, cep As String, _
    cPais As String, xPais As String, fone As String, IE As String, IndIEDest As String, _
    isuf As String, IM As String, email As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro

    strSQL = "update nfe_enderdest set " '(id_nfe,  xLgr,  nro,  xCpl,  xBairro,  cMun,  "
    'strSQL = strSQL & "xMun,  UF,  CEP,  cPais,  xPais,  fone,  IE, indIEDest,  ISUF, IM, "
    'strSQL = strSQL & "email)values("
    
    strSQL = strSQL & "xlgr='" & xLgr & "',"
    strSQL = strSQL & "nro='" & nro & "',"
    strSQL = strSQL & "xcpl='" & xCpl & "',"
    strSQL = strSQL & "xbairro='" & xBairro & "',"
    strSQL = strSQL & "cmun='" & cMun & "',"
    strSQL = strSQL & "xmun='" & xMun & "',"
    strSQL = strSQL & "uf='" & UF & "',"
    strSQL = strSQL & "cep='" & F_SoNumero(cep) & "',"
    strSQL = strSQL & "cpais='" & cPais & "',"
    strSQL = strSQL & "xpais='" & xPais & "',"
    strSQL = strSQL & "fone='" & fone & "',"
    strSQL = strSQL & "ie='" & IE & "',"
    strSQL = strSQL & "indiedest='" & IndIEDest & "',"
    
    strSQL = strSQL & "isuf='" & isuf & "',"
    strSQL = strSQL & "im='" & IM & "',"
    
    strSQL = strSQL & "email='" & email & "' "
    strSQL = strSQL & " where id_nfe=" & id_nfe
    Conn.Execute strSQL
    
    AlteraNfe_enderdest = True
trata_Erro:
End Function



Function GravaNfe_enderEmit(id_nfe As String, xLgr As String, nro As String, xCpl As String, _
    xBairro As String, cMun As String, xMun As String, UF As String, cep As String, _
    cPais As String, xPais As String, fone As String, IE As String, iEst As String, _
    IM As String, CNAE As String, CRT As CRT) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro

    strSQL = "replace into nfe_enderEmit(id_nfe,  xLgr,  nro,  xCpl,  xBairro,  cMun,  "
    strSQL = strSQL & "xMun,  UF,  CEP,  cPais,  xPais,  fone,  IE,  IEST,  IM,  CNAE,  CRT)values("
    strSQL = strSQL & id_nfe & ","
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
    strSQL = strSQL & "'" & IE & "',"
    strSQL = strSQL & "'" & iEst & "',"
    strSQL = strSQL & "'" & IM & "',"
    strSQL = strSQL & "'" & CNAE & "',"
    strSQL = strSQL & "'" & CRT & "') "
    Conn.Execute strSQL
    
    GravaNfe_enderEmit = True
trata_Erro:
End Function

Function GravaNfe_item_imposto(id As String, id_nfe As String, vTotTrib As String, tipo_icms As String, _
    tipo_pis As String, tipo_cofins As String, tipo_ipi As String) As Boolean
 
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    If tipo_icms = "102" Then
        tipo_icms = "ICMSSN102"
    End If
    If tipo_icms = "101" Then
        tipo_icms = "ICMSSN101"
    End If
    If tipo_icms = "103" Then
        tipo_icms = "ICMSSN103"
    End If
    If tipo_icms = "500" Then
        tipo_icms = "ICMSSN500"
    End If
    If tipo_icms = "400" Then
        tipo_icms = "ICMSSN400"
    End If
    If tipo_icms = "900" Then
        tipo_icms = "ICMSSN900"
    End If
    If tipo_icms = "201" Then
        tipo_icms = "ICMSSN201"
    End If
    If tipo_icms = "202" Then
        tipo_icms = "ICMSSN202"
    End If
    If tipo_icms = "203" Then
        tipo_icms = "ICMSSN203"
    End If
    If tipo_icms = "00" Then
        tipo_icms = "ICMS00"
    End If
    strSQL = "insert into nfe_item_imposto(id , id_nfe, vTotTrib, tipo_icms, tipo_pis, tipo_cofins,tipo_ipi)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    strSQL = strSQL & F_GravaDec(vTotTrib) & "," 'taxa federativa
    strSQL = strSQL & "'" & tipo_icms & "',"
    strSQL = strSQL & "'" & tipo_pis & "',"
    strSQL = strSQL & "'" & tipo_cofins & "',"
    strSQL = strSQL & "'" & tipo_ipi & "')"
    
    Conn.Execute strSQL
    GravaNfe_item_imposto = True
trata_Erro:
    
End Function
Function GravaNfe_item_icms(id As String, id_nfe As String, tag As String, csosn As String, orig As String, _
vBC As String, vICMS As String, pICMS As String, modBC As String) As Boolean
 
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    
    strSQL = "insert into nfe_item_icms(id , id_nfe, tag, csosn, orig, vbc,vICms,pIcms,modBC)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    strSQL = strSQL & "'" & tag & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(vICMS) & ","
    strSQL = strSQL & F_GravaDec(pICMS) & ","
    
    strSQL = strSQL & "'" & modBC & "')"
    
    Conn.Execute strSQL
    GravaNfe_item_icms = True
trata_Erro:
    
End Function



Function GravaNfe_item_imposto_icmssn102(id As String, id_nfe As String, orig As String, csosn As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    If CStr(orig) = "" Then
        orig = "0"
    End If
    strSQL = "insert into nfe_item_imposto_icmssn102(id , id_nfe, orig, csosn)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "')"
    
    
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_icmssn102 = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_icmssn101(id As String, id_nfe As String, orig As String, csosn As String, _
    pCredSN As String, vCredICMSSN As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into nfe_item_imposto_icmssn101(id , id_nfe, orig, csosn,pCredSN,vCredICMSSN)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & F_GravaDec(pCredSN) & ","
    strSQL = strSQL & F_GravaDec(vCredICMSSN) & ")"
    
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_icmssn101 = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_icmssn201(id As String, id_nfe As String, orig As String, csosn As String, _
    modBCST As modBCST, pMVAST As String, pRedBCST As String, vBCST As String, pICMSST As String, _
    vICMSST As String, pCredSN As String, vCredICMSSN As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_icmssn201(id , id_nfe, orig, csosn,modBCST, pCredSN,pMVAST, "
    strSQL = strSQL & " pRedBCST, vBCST, pICMSST, vCredICMSSN,vICMSST)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
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
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_icmssn201 = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_icmssn202(id As String, id_nfe As String, orig As String, csosn As String, _
    modBCST As modBCST, pMVAST As String, pRedBCST As String, vBCST As String, pICMSST As String, _
    vICMSST As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    If orig = "" Then
        orig = "0"
    End If
    
    strSQL = "insert into nfe_item_imposto_icmssn202(id , id_nfe, orig, csosn,modBCST, pCredSN,pMVAST, "
    strSQL = strSQL & " pRedBCST, vBCST, pICMSST)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & "'" & modBCST & "',"
    strSQL = strSQL & F_GravaDec(pCredSN) & ","
    strSQL = strSQL & F_GravaDec(pMVAST) & ","
    strSQL = strSQL & F_GravaDec(pRedBCST) & ","
    strSQL = strSQL & F_GravaDec(vBCST) & ","
    strSQL = strSQL & F_GravaDec(pICMSST) & ")"
    
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_icmssn202 = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_icmssn500(id As String, id_nfe As String, orig As String, csosn As String, _
    vBCSTRet As String, vICMSSTRet As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    If CStr(orig) = "" Then
        orig = "0"
    End If
    strSQL = "insert into nfe_item_imposto_icmssn500(id , id_nfe, orig, csosn,vBCSTRet,vICMSSTRet)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & F_GravaDec(vBCSTRet) & ","
    strSQL = strSQL & F_GravaDec(vICMSSTRet) & ")"
    
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_icmssn500 = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_icmssn900(id As String, id_nfe As String, orig As String, csosn As String, _
    vBC As String, pICMS As String, vICMS As String, _
    modBCST As modBCST, pMVAST As String, pRedBCST As String, vBCST As String, pICMSST As String, _
    vICMSST As String, pCredSN As String, vCredICMSSN As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    If modoBC = "" Then
        modoBC = "0"
    End If
    strSQL = "insert into nfe_item_imposto_icmssn900(id , id_nfe, orig, modbc, vbc, picms,vicms,csosn,modBCST, pCredSN,pMVAST, "
    strSQL = strSQL & " pRedBCST, vBCST, pICMSST, vCredICMSSN,vICMSST)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & modoBC & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pICMS) & ","
    strSQL = strSQL & F_GravaDec(vICMS) & ","
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & "'" & modBCST & "',"
    strSQL = strSQL & F_GravaDec(pCredSN) & ","
    strSQL = strSQL & F_GravaDec(pMVAST) & ","
    strSQL = strSQL & F_GravaDec(pRedBCST) & ","
    strSQL = strSQL & F_GravaDec(vBCST) & ","
    strSQL = strSQL & F_GravaDec(pICMSST) & ","
    strSQL = strSQL & F_GravaDec(vCredICMSSN) & ","
    strSQL = strSQL & F_GravaDec(vICMSST) & ")"
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_icmssn900 = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_icmsST(id As String, id_nfe As String, orig As String, csosn As String, _
    cst As String, vBCSTRet As String, vICMSSTRet As String, _
    vBCSTDest As String, vICMSSTDest As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    If orig = "" Then
        orig = "0"
    End If
    
    strSQL = "insert into nfe_item_imposto_icmssn201(id , id_nfe, orig, csosn,cst, vBCSTRet,vICMSSTRet, "
    strSQL = strSQL & "vBCSTDest, vICMSSTDest)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & csosn & "',"
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBCSTRet) & ","
    strSQL = strSQL & F_GravaDec(vICMSSTRet) & ","
    strSQL = strSQL & F_GravaDec(vBCSTDest) & ","
    strSQL = strSQL & F_GravaDec(vICMSSTDest) & ")"
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_icmsST = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_icms00(id As String, id_nfe As String, orig As String, modBC As String, _
    vBC As String, pICMS As String, vICMS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    If orig = "" Then
        orig = "0"
    End If
    
    strSQL = "insert into nfe_item_imposto_icms00(id , id_nfe, CST, orig, modBC, vBC,pIcms,vICMS)values ("
   
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    strSQL = strSQL & "'00',"
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & modBC & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pICMS) & ","
    strSQL = strSQL & F_GravaDec(vICMS) & ")"
    
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_icms00 = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_icms60(id As String, id_nfe As String, orig As String, _
    vBCSTRet As String, vICMSSTRet As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    If orig = "" Then
        orig = "0"
    End If
    
    strSQL = "insert into nfe_item_imposto_icms60(id , id_nfe, CST, orig, vBCSTRet,vICMSSTRet)values ("
   
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    strSQL = strSQL & "'60',"
    strSQL = strSQL & "'" & orig & "',"
  
    strSQL = strSQL & F_GravaDec(vBCSTRet) & ","
    strSQL = strSQL & F_GravaDec(vICMSSTRet) & ")"
    
    
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_icms60 = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_icms20(id As String, id_nfe As String, orig As String, modBC As String, pRedBC As String, _
    vBC As String, pICMS As String, vICMS As String, vICMSDeson As String, motDesICMS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    If orig = "" Then
        orig = "0"
    End If
    
    strSQL = "insert into nfe_item_imposto_icms20(id , id_nfe, CST, orig, modBC, pRedBC, vBC,pIcms,vICMS,vICMSDeson,motDesICMS)values( "
   
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & "'" & orig & "',"
    strSQL = strSQL & "'" & modBC & "',"
    strSQL = strSQL & F_GravaDec(pRedBC) & ","
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pICMS) & ","
    strSQL = strSQL & F_GravaDec(vICMS) & ","
    strSQL = strSQL & F_GravaDec(vICMSDeson) & ","
    strSQL = strSQL & "'" & motDesICMS & "')"
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_icms20 = True
trata_Erro:
    
End Function
Function GravaNfe_IcmsTot(id_nfe As String, vBC As String, vICMS As String, vICMSDeson As String, _
    vBCST As String, vST As String, vProd As String, vFrete As String, vSeg As String, vDesc, _
    vII As String, vIPI As String, vPIS As String, vCOFINS As String, vOutro As String, _
    vNF As String, vTotTrib As String) As Boolean
    On Error GoTo trata_Erro
    
    strSQL = "delete from nfe_icmstot where id_nfe=" & id_nfe
    Conn.Execute strSQL
    
    strSQL = "insert into nfe_icmstot(id_nfe , vBC, vICMS, vICMSDeson, vBCST, vST, vProd, vFrete, "
    strSQL = strSQL & "vSeg, vDesc, vII, vIPI, vPIS, vCOFINS, vOutro, vNF, vTotTrib)values("
    strSQL = strSQL & id_nfe & ","
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
    Conn.Execute strSQL
    GravaNfe_IcmsTot = True
trata_Erro:
End Function


Function GravaNfe_item_imposto_IPI(id As String, id_nfe As String, clEnq As String, CNPJProd As String, _
    cSelo As String, qSelo As String, cEnq As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_ipi(id , id_nfe, clEnq, CNPJProd,cSelo, qSelo ,cEnq)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & clEnq & "',"
    strSQL = strSQL & "'" & CNPJProd & "',"
    strSQL = strSQL & "'" & cSelo & "',"
    strSQL = strSQL & "'" & qSelo & "',"
    strSQL = strSQL & "'" & cEnq & "')"
    Conn.Execute strSQL
    GravaNfe_item_imposto_IPI = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_IPINT(id As String, id_nfe As String, cst As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_ipiNT(id , id_nfe, cst)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & cst & "')"
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_IPINT = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_IPITrib(id As String, id_nfe As String, cst As String, vBC As String, _
    pIPI As String, qUnid As String, vUnid As String, vIPI As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_ipiTrib(id , id_nfe, cst, vbc,pipi, qunid ,vunid,vipi)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pIPI) & ","
    strSQL = strSQL & F_GravaDec(qUnid) & ","
    strSQL = strSQL & F_GravaDec(vUnid) & ","
    strSQL = strSQL & F_GravaDec(vIPI) & ")"
    Conn.Execute strSQL
    GravaNfe_item_imposto_IPITrib = True
trata_Erro:
    
End Function


Function GravaNfe_item_imposto_PISAliq(id As String, id_nfe As String, cst As String, vBC As String, _
    pPIS As String, vPIS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_pisaliq(id , id_nfe, cst, vbc,ppis, vpis)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pPIS) & ","
    
    strSQL = strSQL & F_GravaDec(vPIS) & ")"
    Conn.Execute strSQL
    GravaNfe_item_imposto_PISAliq = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_PISNT(id As String, id_nfe As String, cst As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_pisnt(id , id_nfe, cst)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & cst & "')"
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_PISNT = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_PISOutr(id As String, id_nfe As String, cst As String, vBC As String, _
    pPIS As String, qBCProd As String, vAliqProd As String, vPIS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_pisoutr(id , id_nfe, cst, vbc,pPIS, qBCPRod  ,vAliqProd,vPIS)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & F_GravaDec(cst) & ","
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pPIS) & ","
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ","
    strSQL = strSQL & F_GravaDec(vPIS) & ")"
    Conn.Execute strSQL
    GravaNfe_item_imposto_PISOutr = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_PISST(id As String, id_nfe As String, vBC As String, _
    pPIS As String, qBCProd As String, vAliqProd As String, vPIS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_pisst(id , id_nfe, vbc,pPIS, qBCPRod  ,vAliqProd,vPIS)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pPIS) & ","
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ","
    strSQL = strSQL & "'" & vPIS & "')"
    Conn.Execute strSQL
    GravaNfe_item_imposto_PISST = True
trata_Erro:
    
End Function

Function GravaNfe_item_imposto_PISQtde(id As String, id_nfe As String, cst As String, _
    qBCProd As String, vAliqProd As String, vPIS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_pisQTde(id,id_nfe, cst,qBCPRod  ,vAliqProd,vPIS)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    
    strSQL = strSQL & "'" & cst & "',"
    
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ","
    strSQL = strSQL & F_GravaDec(vPIS) & ")"
    Conn.Execute strSQL
    GravaNfe_item_imposto_PISQtde = True
trata_Erro:
    
End Function




Function GravaNfe_item_imposto_COFINSAliq(id As String, id_nfe As String, cst As String, vBC As String, _
    pCOFINS As String, vCOFINS As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_cofinsaliq(id , id_nfe, cst, vbc,pCofins,vCofins)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pCOFINS) & ","
    strSQL = strSQL & F_GravaDec(vCOFINS) & ")"
    
    Conn.Execute strSQL
    GravaNfe_item_imposto_COFINSAliq = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_COFINSOutr(id As String, id_nfe As String, cst As String, vBC As String, _
    pCOFINS As String, qBCProd As String, vAliqProd As String, _
    vCOFINS As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_cofinsoutr(id , id_nfe, cst, vbc,pCofins,vCofins,qBCProd,vAliqProd)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pCOFINS) & ","
    strSQL = strSQL & F_GravaDec(vCOFINS) & ","
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ")"
    Conn.Execute strSQL
    GravaNfe_item_imposto_COFINSOutr = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_COFINSQtde(id As String, id_nfe As String, cst As String, _
    qBCProd As String, vAliqProd As String, _
    vCOFINS As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_cofinsoutr(id , id_nfe, cst, qBCProd,vAliqProd, vcofins)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ","
    strSQL = strSQL & F_GravaDec(vCOFINS) & ")"
    Conn.Execute strSQL
    GravaNfe_item_imposto_COFINSQtde = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_COFINSNT(id As String, id_nfe As String, cst As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_cofinsnt(id , id_nfe, cst)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & cst & "')"
    

    Conn.Execute strSQL
    GravaNfe_item_imposto_COFINSNT = True
trata_Erro:
    
End Function
Function GravaNfe_item_imposto_COFINSST(id As String, id_nfe As String, cst As String, vBC As String, _
    pCOFINS As String, vCOFINS As String, qBCProd As String, vAliqProd As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into nfe_item_imposto_cofinsoutr(id , id_nfe, cst, vbc,pCofins,vCofins,qBCProd,vAliqProd)values( "
    strSQL = strSQL & id & ","
    strSQL = strSQL & id_nfe & ","
    
    strSQL = strSQL & "'" & cst & "',"
    strSQL = strSQL & F_GravaDec(vBC) & ","
    strSQL = strSQL & F_GravaDec(pCOFINS) & ","
    strSQL = strSQL & F_GravaDec(vCOFINS) & ","
    strSQL = strSQL & F_GravaDec(qBCProd) & ","
    strSQL = strSQL & F_GravaDec(vAliqProd) & ")"
    Conn.Execute strSQL
    GravaNfe_item_imposto_COFINSST = True
trata_Erro:
    
End Function

Function excluir_nfe_pag(id As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "delete from nfe_pag where id=" & id
    Conn.Execute strSQL
    excluir_nfe_pag = True
trata_Erro:

End Function
Function GravaNFe_Pag(id_nfe As String, tPag As String, vPag As String, card As String, _
    CNPJ As String, tBand As String, cAut As String, tpIntegra As String, vTroco As String, Optional id As String = "") As Boolean
     On Error GoTo trata_Erro
     If tPag = "" Then
        tPag = "01"
     End If
     strSQL = "replace into Nfe_pag("
     If id <> "" Then
        strSQL = strSQL & "id,"
     End If
     strSQL = strSQL & " id_nfe , tPag, vPag, vTroco, card, CNPJ, tBand, cAut,tpIntegra) values ( "
     If id <> "" Then
        strSQL = strSQL & id & ", "
     End If
     strSQL = strSQL & id_nfe & ","
     strSQL = strSQL & "'" & Format(tPag, "0#") & "', "
     strSQL = strSQL & F_GravaDec(vPag) & ", "
     strSQL = strSQL & F_GravaDec(vTroco) & ", "
     strSQL = strSQL & "'" & card & "',"
     strSQL = strSQL & "'" & CNPJ & "', "
     strSQL = strSQL & "'" & tBand & "',"
     strSQL = strSQL & "'" & cAut & "', "
     strSQL = strSQL & "'" & tpIntegra & "') "
     Conn.Execute strSQL
     
     GravaNFe_Pag = True
trata_Erro:
End Function
Function GravaNFe_ProcRec(id_nfe As String, nProc As String, indProc As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_ProcRec(id_nfe , nproc, indproc) values ( "
     strSQL = strSQL & id_nfe & ","
     
     strSQL = strSQL & "'" & nProc & "',"
     strSQL = strSQL & "'" & indProc & "') "
     Conn.Execute strSQL
     
     GravaNFe_ProcRec = True
trata_Erro:
End Function
Function GravaNFe_Transp(id_nfe As String, modFrete As modFrete) As Boolean
     On Error GoTo trata_Erro
     strSQL = "replace into Nfe_transp(id_nfe , modFrete) values ( "
     strSQL = strSQL & id_nfe & ","
     
     strSQL = strSQL & "'" & modFrete & "')"
        
     Conn.Execute strSQL
     
     GravaNFe_Transp = True
trata_Erro:
End Function

Function GravaNFe_Transp_Reboque(id_nfe As String, placa As String, UF As String, RNTC As String, _
    vagao As String, balsa As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "replace into Nfe_transp_reboque(id_nfe,  placa,  UF,  RNTC,  vagao,  balsa ) values ( "
     strSQL = strSQL & id_nfe & ","
     
     strSQL = strSQL & "'" & placa & "',"
        strSQL = strSQL & "'" & UF & "',"
        strSQL = strSQL & "'" & RNTC & "',"
        strSQL = strSQL & "'" & vagao & "',"
        strSQL = strSQL & "'" & balsa & "')"
     Conn.Execute strSQL
     
     GravaNFe_Transp_Reboque = True
trata_Erro:
End Function

Function GravaNFe_Transp_Transportadora(id_nfe As String, CNPJ As String, cpf As String, xNome As String, _
    IE As String, xEnder As String, xMun As String, UF As String) As Boolean
     On Error GoTo trata_Erro
     Dim i As Integer
     strSQL = "replace into Nfe_transp_transportadora(id_nfe,  cnpj,  cpf,  xnome,  ie,  xender, xmun, uf) values ( "
     strSQL = strSQL & id_nfe & ","
     
     strSQL = strSQL & "'" & CNPJ & "',"
        strSQL = strSQL & "'" & cpf & "',"
        strSQL = strSQL & "'" & xNome & "',"
        strSQL = strSQL & "'" & IE & "',"
        strSQL = strSQL & "'" & xEnder & "',"
        strSQL = strSQL & "'" & xMun & "',"
        strSQL = strSQL & "'" & UF & "')"
     Conn.Execute strSQL, i
     
     GravaNFe_Transp_Transportadora = True
trata_Erro:
End Function

Function GravaNFe_Transp_VeicTransp(id_nfe As String, placa As String, UF As String, RNTC As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "replace into Nfe_Transp_VeicTransp(id_nfe,  placa,  uf,  rntc) values ( "
     strSQL = strSQL & id_nfe & ","
     
     strSQL = strSQL & "'" & placa & "',"
        strSQL = strSQL & "'" & UF & "',"
        strSQL = strSQL & "'" & RNTC & "')"
        
     Conn.Execute strSQL
     
     GravaNFe_Transp_VeicTransp = True
trata_Erro:
End Function
Function GravaNFe_Transp_Vol(id_nfe As String, qVol As String, ESP As String, Marca As String, nVol As String, _
    pesoL As String, pesoB As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "replace into Nfe_Transp_Vol(id_nfe,  qvol,  esp,  marca, nvol,pesol,pesob) values ( "
     strSQL = strSQL & id_nfe & ","
     strSQL = strSQL & F_GravaDec(qVol) & ", "
     strSQL = strSQL & "'" & ESP & "',"
        strSQL = strSQL & "'" & Marca & "',"
        strSQL = strSQL & "'" & nVol & "',"
        strSQL = strSQL & F_GravaDec(pesoL) & ", "
        strSQL = strSQL & F_GravaDec(pesoB) & ")"
     Conn.Execute strSQL
     
     GravaNFe_Transp_Vol = True
trata_Erro:
End Function

Function GravaNFe_Transp_RetTransp(id_nfe As String, vServ As String, vBCRet As String, pICMSRet As String, _
    vICMSRet, CFOP As String, cMunFG As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "replace into Nfe_Transp_RetTransp(id_nfe,vserv,vBCRet,pICMSRet,vICMSRet,cfop,cMunFG) values ( "
     strSQL = strSQL & id_nfe & ","
     strSQL = strSQL & F_GravaDec(vServ) & ", "
     strSQL = strSQL & F_GravaDec(vBCRet) & ", "
     strSQL = strSQL & F_GravaDec(pICMSRet) & ", "
     strSQL = strSQL & F_GravaDec(vICMSRet) & ", "
     strSQL = strSQL & "'" & CFOP & "',"
        strSQL = strSQL & "'" & cMunFG & "')"
        
     Conn.Execute strSQL
     
     GravaNFe_Transp_RetTransp = True
trata_Erro:
End Function
Function GravaNFe_InfAdc(id_nfe As String, infAdFisco As String, infCpl As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "replace into Nfe_InfAdic(id_nfe,infAdFisco ,infCpl) values ( "
     strSQL = strSQL & id_nfe & ","
     strSQL = strSQL & "'" & infAdFisco & "', "
     strSQL = strSQL & "'" & infCpl & "') "
     
     
        
     Conn.Execute strSQL
     
     GravaNFe_InfAdc = True
trata_Erro:
End Function

Function GravaNFe_Cobr_Dup(id_nfe As String, nDup As String, dVenc As String, vDup As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_Cobr_Dup(id_nfe,nDup,dVenc,vDup) values ( "
     strSQL = strSQL & id_nfe & ","
     strSQL = strSQL & "'" & nDup & "', "
     strSQL = strSQL & F_GravaData(dVenc) & ", "
     strSQL = strSQL & F_GravaDec(vDup) & ") "
     
        
     Conn.Execute strSQL
     
     GravaNFe_Cobr_Dup = True
trata_Erro:
End Function

Function GravaNFe_Cobr_Fat(id_nfe As String, nFat As String, vOrig As String, vDesc As String, vLiq As String) As Boolean
     On Error GoTo trata_Erro
     strSQL = "insert into Nfe_Cobr_Dup(id_nfe,nFat,vOrig,vDesc,vLiq) values ( "
     strSQL = strSQL & id_nfe & ","
     strSQL = strSQL & "'" & nFat & "', "
     strSQL = strSQL & F_GravaDec(vOrig) & ", "
     strSQL = strSQL & F_GravaDec(vDesc) & ", "
     strSQL = strSQL & F_GravaDec(vLiq) & ") "
        
     Conn.Execute strSQL
     
     GravaNFe_Cobr_Fat = True
trata_Erro:
End Function

Sub ImprimirXMLNFeCUPOM(xml As String, obj As Object, Optional bBobina As Boolean = True, _
    Optional via_cartao As String = "", Optional xColunas As String = "", Optional diminuiFonte As Integer = 0, Optional reduzida As Boolean = False)
    
    Dim s As String
    Dim conteudoXML As String
    Dim p As PictureBox
    Dim cQRCode As String
    Dim pi As IPictureDisp
    Dim emit As String
    Dim dest As String
    Dim ide As String
    Dim qrCode As String
    Dim det As String
    Dim iProd As Integer
    Dim sTot As String
    Dim pag As String
    Dim tEmis As String
    Dim infProt As String
    Dim cStat As String
    Dim digVal As String
    Dim nProt As String
    Dim chNFe As String
    Dim dhRecbto As String
    Dim nNF As String
    Dim serie As String
    Dim modelo As String
    Dim dhEmi As String
    Dim natOp As String
    Dim vNF As String
    Dim pOld As Integer
    Dim iCod As Single
    Dim iQtd As Single
    Dim iNome As Single
    Dim iTamNome As Single
    Dim iTotal As Single
    
    
    
    'On Error GoTo trata_Erro
    On Error Resume Next
    P_Status "Emitindo Cupon NFC-e"
    
    If Conn.DefaultDatabase = "lumine" Then
        reduzida = True
    End If
    
    
    pOld = BuscaImpressoraPadrao
    Set Printer = Printers(ImpressoraPadraoCupomNFCe.Indice)
    
    obj.FontName = "Arial"
    cQRCode = App.path & "\qrcode.tmp"
    If UCase(Right(xml, 4)) = ".XML" Then
        conteudoXML = LeArquivoExternoUTF8(xml)
    Else
        conteudoXML = xml
        
    End If
    If conteudoXML = "" Then
        MsgBox "XML não encontrado", vbExclamation
    End If
    
    
    
    dest = XMLUtil.RetornaValorTagXML(conteudoXML, "<dest>")
    emit = XMLUtil.RetornaValorTagXML(conteudoXML, "<emit>")
    sTot = XMLUtil.RetornaValorTagXML(conteudoXML, "<total>")
    ide = XMLUtil.RetornaValorTagXML(conteudoXML, "<ide>")
    emit = XMLUtil.RetornaValorTagXML(conteudoXML, "<emit>")
    pag = XMLUtil.RetornaValorTagXML(conteudoXML, "<pag>")
    s = RetornaValorTagXML(conteudoXML, "<qrCode>")
    qrCode = RetornaValorTagXML(s, "<![CDATA[", "]]")
    
    tEmis = RetornaValorTagXML(ide, "<tpEmis>")
    infProt = RetornaValorTagXML(conteudoXML, "<infProt>")
    cStat = RetornaValorTagXML(infProt, "<cStat>")
    digVal = RetornaValorTagXML(infProt, "<digVal>")
    nProt = RetornaValorTagXML(infProt, "<nProt>")
    chNFe = RetornaValorTagXML(infProt, "<chNFe>")
    dhRecbto = RetornaValorTagXML(infProt, "<dhRecbto>")
    nNF = RetornaValorTagXML(ide, "<nNF>")
    serie = RetornaValorTagXML(ide, "<serie>")
    modelo = RetornaValorTagXML(ide, "<mod>")
    dhEmi = RetornaValorTagXML(ide, "<dhEmi>")
    natOp = RetornaValorTagXML(ide, "<natOp>")
    id = BuscaIdPelaChave(chNFe)
    
    
    
    imprimeLogo obj
    obj.Print
    
    
    obj.fontsize = 8 - diminuiFonte
    s = RetornaValorTagXML(emit, "<xNome>")
    obj.FontBold = True
    P_CentraImpressao obj, s, False
    
    obj.FontBold = False
    s = RetornaValorTagXML(emit, "<CNPJ>")
    P_CentraImpressao obj, "CNPJ:" & F_FormataCNPJ(strCNPJ_Emp), False
    P_CentraImpressao obj, "IE: " & strIE_Emp, False
    
    
    
    s = RetornaValorTagXML(emit, "<xLgr>")
    s = s & " Nº " & RetornaValorTagXML(emit, "<nro>")
    s = s & " " & RetornaValorTagXML(emit, "<xCpl>")
    
    P_CentraImpressao obj, s, False
    
    s = RetornaValorTagXML(emit, "<xBairro>")
    s = s & " CEP:" & RetornaValorTagXML(emit, "<CEP>")
    P_CentraImpressao obj, s, False
    
    s = RetornaValorTagXML(emit, "<xMun>")
    s = s & "/" & RetornaValorTagXML(emit, "<UF>")
    P_CentraImpressao obj, s, False
    P_CentraImpressao obj, "Fone:" & F_AjustaTelefone(strFone_Emp), False
    
    P_line1 obj, vbDot
    obj.Print
    obj.FontBold = True
    obj.fontsize = 7 - diminuiFonte
    P_CentraImpressao obj, "DANFE NFC-e"
    P_CentraImpressao obj, "DOCUMENTO AUXILIAR DA NOTA FISCAL "
    P_CentraImpressao obj, "DE VENDA A CONSUMIDOR ELETRÔNICA"
    P_CentraImpressao obj, "NÃO PERMITE APROVEITAMENTO DE CRÉD. ICMS"
    obj.Print
    obj.FontBold = False
    P_line1 obj, vbDot
    P_line1 obj, vbDot
    obj.FontBold = True
    obj.Print
    P_CentraImpressao obj, "CONSUMIDOR"
    obj.FontBold = False
    Dim cpf As String
    Dim nome As String
    If dest = "" Then
        nome = F_BuscaUmCampo("xnome", "nfe_dest", "id_nfe=" & id)
        cpf = F_BuscaUmCampo("cpf", "nfe_dest", "id_nfe=" & id)
        
        If cpf = "" And nome = "" Then
            P_CentraImpressao obj, "Consumidor Não Identificado"
        Else
            P_ImprX obj, "CPF:" & cpf, 0, False
            P_ImprX obj, nome, 0, False
        End If
        
        
        
    Else
        P_ImprX obj, "CPF:" & RetornaValorTagXML(dest, "<CPF>"), 0, False
        P_ImprX obj, RetornaValorTagXML(dest, "<xNome>"), 0, False
        
        
        
        
        s = RetornaValorTagXML(dest, "<xLgr>")
        If s <> "" Then
            P_ImprX obj, "Endereço: " & s, 0
            
        End If
        s = RetornaValorTagXML(dest, "<nro>")
        If s <> "" Then
            P_ImprX obj, "N" & s, 70
            
        End If
        obj.Print
        s = RetornaValorTagXML(dest, "<xBairro>")
        If s <> "" Then
            P_ImprX obj, "Bairro: " & s, 0, True
            
        End If
        s = RetornaValorTagXML(dest, "<CEP>")
        If s <> "" Then
            P_ImprX obj, "CEP: " & s, 70, False
            
        End If
        s = RetornaValorTagXML(dest, "<xMun>")
        If s <> "" Then
            P_ImprX obj, "CIDADE: " & s & RetornaValorTagXML(dest, "<UF>"), 0, False
            
        End If
    End If
    
    
    
    
    
    P_line1 obj, vbDashDotDot
    obj.Print
    P_CentraImpressao obj, "Produtos Adquiridos"
    'P_line1 obj, vbDot
    obj.Print
    
    iProd = 1
    obj.FontBold = True
    obj.fontsize = 8 - diminuiFonte
    
    
    
    If reduzida = False Then
    
    
        P_ImprX obj, "Código", 0
        P_ImprX obj, "Descrição", 14
        'P_ImprX obj, "Und", 10
        P_ImprX obj, "Qtd", 46
        P_ImprX obj, "Unit", 55
        P_ImprX obj, "Total", 63, False
        P_line1 obj, vbDot
        obj.FontBold = False
        obj.FontName = "Arial Narrow"
        det = "1"
        Do While det <> ""
            det = RetornaValorTagXML(conteudoXML, "<det nItem=" & aspas & iProd & aspas & ">", "</det>")
            If det <> "" Then
                
                P_ImprX obj, RetornaValorTagXML(det, "<cProd>"), 0
                P_ImprX obj, Left(RetornaValorTagXML(det, "<xProd>"), 20), 14
                'P_ImprX obj, "Und", 10
                P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<qCom>"), "###,##0.000"), 46
                P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<vUnCom>")), 55
                
                P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<vProd>")), 63, False
                
                iProd = iProd + 1
            End If
            
            
        Loop
    Else
        
        obj.FontBold = False
        obj.FontName = "Arial Narrow"
        det = "1"
        Do While det <> ""
            det = RetornaValorTagXML(conteudoXML, "<det nItem=" & aspas & iProd & aspas & ">", "</det>")
            If det <> "" Then
                obj.FontBold = True
                P_ImprX obj, "Produto " & iProd, 0, False
                obj.FontBold = False
                P_ImprX obj, "Código: " & RetornaValorTagXML(det, "<cProd>"), 0, False
                P_ImprX obj, "Descrição: " & Left(RetornaValorTagXML(det, "<xProd>"), 40), 0, False
                'P_ImprX obj, "Und", 10
                P_ImprX obj, "Qtd: " & FormatValNFe(RetornaValorTagXML(det, "<qCom>"), "###,##0.000"), 0, False
                P_ImprX obj, "Pç Unitário: " & FormatValNFe(RetornaValorTagXML(det, "<vUnCom>")), 0, False
                
                P_ImprX obj, "Vlr Total: " & FormatValNFe(RetornaValorTagXML(det, "<vProd>")), 0, False
                P_line1 obj, vbDot
                iProd = iProd + 1
            End If
            
            
        Loop
    
    End If
    obj.FontName = "Arial"
    P_line1 obj, vbDot
    obj.Print
    obj.FontBold = True
    P_ImprX obj, "Itens: " & iProd - 1, 0, False
    P_ImprX obj, "Valor dos produtos: " & FormatValNFe(RetornaValorTagXML(sTot, "<vProd>")), 0, False
    P_ImprX obj, "Desconto: " & FormatValNFe(RetornaValorTagXML(sTot, "<vDesc>")), 0, False
    P_ImprX obj, "Valor Total: " & FormatValNFe(RetornaValorTagXML(sTot, "<vNF>")), 0, False
    P_ImprX obj, "Tipo de Pgto: " & RetornaTPag(RetornaValorTagXML(pag, "<tPag>")), 0, False
    
    vNF = RetornaValorTagXML(sTot, "<vNF>")
    vNF = mudarCaractere(vNF, ".", ",")
    If vNF = "" Then
        vNF = "0"
    End If
    P_line1 obj, vbDot
    obj.Print
    P_CentraImpressao obj, "Valor Aproximado dos Tributos (Lei. 12.741)"
    obj.FontBold = False
    P_ImprX obj, "Estadual: " & FormatVal(ValorMedioImpostoEstadual / 100 * CCur(vNF)), 10
    P_ImprX obj, "Federal: " & FormatVal(ValorMedioImpostoFederal / 100 * CCur(vNF)), 40, False
    
    
    obj.Print
    obj.FontBold = False
    P_line1 obj, vbDot
    ImprimeQRCodeArquivo qrCode, cQRCode
    
    
    
    
    
    obj.Print ""
    
    
    If tEmis = "1" Then
        P_CentraImpressao obj, "Emissão Normal"
    Else
        P_CentraImpressao obj, "EMITIDA EM CONTINGÊNCIA"
        P_CentraImpressao obj, "NFC-E DEPENDE DE AUTORIZAÇÃO"
    End If
    If nProt <> "" Then
        If cStat = "100" Then
            s = "Número: " & Format(nNF, "0########") & " Série: " & serie
            
            P_CentraImpressao obj, s
            
            P_CentraImpressao obj, "Emissão: " & dhEmi
            obj.Print
            P_CentraImpressao obj, "Consulte pela Chave de Acesso"
            P_CentraImpressao obj, strEnderecoConsultaChaveDeAcesso
            
            
            P_CentraImpressao obj, "Chave de Acesso "
            
            P_CentraImpressao obj, chNFe
            obj.Print
            P_CentraImpressao obj, "Protocolo de Autorização: " & nProt
            
            obj.Print
        Else
                
        End If
    Else
        s = "Número: " & Format(nNF, "0########") & " Série: " & serie
            
            P_CentraImpressao obj, "NFe não enviada à SEFAZ"
            P_CentraImpressao obj, "RELIZAR CONTINGÊNCIA"
            P_CentraImpressao obj, s
            
            P_CentraImpressao obj, "Emissão: " & dhEmi
            obj.Print
            P_CentraImpressao obj, "Consulte pela Chave de Acesso"
            P_CentraImpressao obj, strEnderecoConsultaChaveDeAcesso
            
            
            P_CentraImpressao obj, "Chave de Acesso "
            
            P_CentraImpressao obj, chNFe
            obj.Print
            P_CentraImpressao obj, "Protocolo de Autorização: " & nProt
            
            obj.Print
        
        
    End If
    
    
    
    obj.Print
    'obj.PaintPicture LoadPicture(cQRCode), 0, obj.CurrentY ', LoadPicture(cQRCode).width * 0.5, LoadPicture(cQRCode).Height * 0.5
    P_CentraImpressao obj, "Consulta Via QR-Code"
    obj.Print
    ImprimeImagemArquivo obj, cQRCode, True, 0.5
    obj.Print
        
        
    
    
    
    obj.Print ""
    
    P_line1 obj, vbDot
    
    P_ImprX obj, "Caixa: " & strCod_Cai, 0
    P_ImprX obj, "Usuário: " & strNome_Usu, 15
    P_ImprX obj, Now, 40
    
    
    If via_cartao <> "" Then
        P_line1 obj, vbDot
        P_CentraImpressao obj, "Comprovante de Transação TEF"
        
        obj.Print via_cartao
        
        If TypeName(obj) = "Printer" Then
            Printer.EndDoc
        End If
        obj.Print via_cartao
        
        
        
    End If
    
    
    
    If TypeName(obj) = "Printer" Then
        Printer.EndDoc
    End If
    
    
     
    Set Printer = Printers(pOld)
    P_Status ""
Exit Sub
trata_Erro:
P_Status
P_MostraErro
End Sub
Function ImprimirXMLNFeDANFE(xml As String, obj As Object, Optional bBobina As Boolean = True, Optional via_cartao As String = "", Optional FontName As String = "Arial") As Integer
    Dim iPagina As Integer
    Dim iQtdPaginas As Integer
    Dim s As String
    Dim conteudoXML As String
    Dim itensPorPagina As Integer
    Dim itensPorPaginaPrimeiraPagina As Integer
    Dim itensDemaisPaginas As Integer
    '0 - não imprimiu
    '1 - visualizou e mandou imprimir
    '-1 erro
    
    ImprimirXMLNFeDANFE = -1
    
    Dim p As Object
    Dim cQRCode As String
    Dim pi As IPictureDisp
    Dim emit As String
    Dim dest As String
    Dim ide As String
    Dim qrCode As String
    Dim det As String
    Dim iProd As Integer
    Dim sTot As String
    Dim pag As String
    Dim tEmis As String
    Dim infProt As String
    Dim cStat As String
    Dim digVal As String
    Dim nProt As String
    Dim chNFe As String
    Dim dhRecbto As String
    Dim nNF As String
    Dim serie As String
    Dim modelo As String
    Dim dhEmi As String
    Dim natOp As String
    Dim vNF As String
    Dim pOld As Integer
    Dim iniY As Single
    Dim iniX As Single
    Dim auxY As Single
    Dim yRet As Single
    Dim yFimQuadroProdutos As Single
    Dim alturaCX As Single
    Dim endereco As String
    Dim alt1 As Single
    Dim alturaBlocoFinal As Single
    Dim alturaBlocoMeio As Single
    Dim alturaBlocoInicial As Single
    
    Dim transp As String
    Dim infAdic As String
    Dim imprimiuFim As Boolean
    Dim tipopag As String
    Dim vPag As String
    Dim strSQL As String
    Dim iQtdItens As Integer
    Dim iRecuo As Single
    Dim iRecuoAmp As Single
    Dim iDemaisPaginas As Integer
    Dim cod_emp As String
    alt1 = 4
    itensPorPagina = 15
    itensDemaisPaginas = 25
    
    
    
    
    iRecuoAmp = 2
    obj.ScaleMode = vbMillimeters
    iniY = 5
    iniX = 5
    alturaCX = 8
    iRecuo = 0.5 'espaço para começar a escrever
    yFimQuadroProdutos = 260
    det = "1" 'importante
    iProd = 1 'importante
    
    
    
    
    
    On Error GoTo trata_Erro
    P_Status "Emitindo DANFE NF-e"
        
   
    If UCase(Right(xml, 4)) = ".XML" Then
        conteudoXML = LeArquivoExternoUTF8(xml)
    Else
        conteudoXML = xml
        
    End If
    If conteudoXML = "" Then
        MsgBox "XML não encontrado", vbExclamation
    End If
    
    
    
    dest = XMLUtil.RetornaValorTagXML(conteudoXML, "<dest>")
    emit = XMLUtil.RetornaValorTagXML(conteudoXML, "<emit>")
    sTot = XMLUtil.RetornaValorTagXML(conteudoXML, "<total>")
    ide = XMLUtil.RetornaValorTagXML(conteudoXML, "<ide>")
    emit = XMLUtil.RetornaValorTagXML(conteudoXML, "<emit>")
    pag = XMLUtil.RetornaValorTagXML(conteudoXML, "<pag>")
    s = RetornaValorTagXML(conteudoXML, "<qrCode>")
    transp = XMLUtil.RetornaValorTagXML(conteudoXML, "<transp>")
    
    tEmis = RetornaValorTagXML(ide, "<tpEmis>")
    infProt = RetornaValorTagXML(conteudoXML, "<infProt")
    cStat = RetornaValorTagXML(infProt, "<cStat>")
    digVal = RetornaValorTagXML(infProt, "<digVal>")
    nProt = RetornaValorTagXML(infProt, "<nProt>")
    chNFe = RetornaValorTagXML(infProt, "<chNFe>")
    dhRecbto = RetornaValorTagXML(infProt, "<dhRecbto>")
    nNF = RetornaValorTagXML(ide, "<nNF>")
    serie = RetornaValorTagXML(ide, "<serie>")
    modelo = RetornaValorTagXML(ide, "<mod>")
    dhEmi = RetornaValorTagXML(ide, "<dhEmi>")
    natOp = RetornaValorTagXML(ide, "<natOp>")
    infAdic = RetornaValorTagXML(conteudoXML, "<infAdic>")
    id = BuscaIdPelaChave(chNFe)
    
    
    iPagina = 1
    strSQL = F_BuscaUmCampo("count(*)", "nfe_item", "id_nfe=" & id)
    
    If Int(strSQL) <= 15 Then
        iQtdPaginas = 1
    Else
       ' iDemaisPaginas = (Int(strSQL) - 15) / itensDemaisPaginas
    
        iQtdPaginas = Int(strSQL) / itensPorPagina
        'iQtdPaginas = iDemaisPaginas + 1
    End If
    
    
    iQtdItens = Int(strSQL)
    If Int(strSQL) Mod itensPorPagina <> 0 Then
        iQtdPaginas = iQtdPaginas + 1
    End If
    If FontName = "" Then
        obj.FontName = "Arial"
    Else
        obj.FontName = FontName
    End If
inicio:
    
    
    
    If TypeName(obj) = "PictureBox" And iPagina > 1 Then
        frmVisualizadorImpressao.addNewPage
    ElseIf TypeName(obj) = "Printer" And iPagina > 1 Then
        Printer.NewPage
    End If
    
    P_Quadrado obj, iniX, iniY, 14, 160
    
    obj.Print
    obj.Print
    obj.fontsize = 7
    
    P_PrintXY obj, "Recebemos de " & RetornaValorTagXML(emit, "<xNome>") & " os produtos constantes na presente Nota Fiscal", 10, 5
    
    
    
    P_Quadrado obj, 160 + iniX, iniY, 14, 40
    obj.Print
    obj.fontsize = 10
    obj.FontBold = True
    P_PrintXY obj, "NF-e", 185, 5, False, 2
    obj.FontBold = False
     'MsgBox obj.ScaleMode
    
        
    P_PrintXY obj, "Nº " & Format(nNF, "00000000#"), 183, 8, False, 2
    P_PrintXY obj, "Série " & Format(serie, "00#"), 185, 12, False, 2
    obj.fontsize = 8
    P_PrintXY obj, "Folha " & iPagina & "/" & iQtdPaginas, 185, 16, False, 2
    
    
    obj.FontBold = False
    obj.fontsize = 9
    P_PrintXY obj, "Data de Recebimento", 5 + iniX, 9
    P_PrintXY obj, "_____/_____/____________", 2 + iniX, 14
    
    obj.FontBold = False
    obj.fontsize = 9
    P_Quadrado obj, iniX, 9, 10, 160
    P_PrintXY obj, "Identificação e Assinatura do Recebedor", 50, 9, False
    P_PrintXY obj, "_________________________________________", 50, 14, False
    
    obj.Print
        
    P_line1 obj, vbDot
    obj.Print
    auxY = obj.CurrentY - 1
    obj.fontsize = 9
    P_Quadrado obj, iniX, auxY, 30, 70, vbSolid, 2, False, "Emitente"
    P_Quadrado obj, iniX + 70, auxY, 30, 50, vbSolid, 2, False, ""
    P_Quadrado obj, iniX + 120, auxY, 30, 80, vbSolid, 2, False, ""
    obj.Print
    obj.Print
    
    obj.FontBold = True
    auxY = obj.CurrentY
    obj.CurrentY = 20
    imprimeLogo obj, iniX + 15, 22, 40
    obj.Print
    obj.fontsize = 9
    obj.CurrentY = 35
    P_ImprQuebraTurbo RetornaValorTagXML(emit, "<xNome>"), obj, 5, 70, False, 2
    obj.FontBold = False
    obj.fontsize = 7
    P_ImprQuebraTurbo RetornaValorTagXML(emit, "<xLgr>") & " N: " & RetornaValorTagXML(emit, "<nro>") & " " & RetornaValorTagXML(emit, "<xBairro>") & "  " & strComplemento_Emp & " CEP:" & strCEP_Emp, obj, 5, 70, False, 2
    P_ImprQuebraTurbo strCidade_Emp & " - " & strUF_Emp & " Fone " & strFone_Emp, obj, 5, 70, False, 2
    
    obj.CurrentY = 25
    
    
    P_ImprX obj, "DANFE", 90, False
    obj.fontsize = 8
    obj.CurrentX = 100
    obj.FontBold = False
    P_ImprQuebraTurbo "DOCUMENTO AUXILIAR DE NOTA FISCAL ELETRÔNICA", obj, 75, 120, False, 2
    
    obj.fontsize = 10
    P_ImprX obj, "0 - Entrada", 85, False
    P_ImprX obj, "1 - Saída", 85, False
    
    s = RetornaValorTagXML(ide, "<tpNF>")
    obj.fontsize = 14
    obj.CurrentY = 40
    P_Quadrado obj, 105, 40, 6, 6, vbDot, 2, False, s
    obj.fontsize = 12
    obj.Print
    
    obj.FontBold = True
    obj.CurrentY = 43
    P_ImprX obj, "Nº " & Format(nNF, "0000000#"), 77, False
    P_ImprX obj, "Série " & Format(serie, "00#"), 77, True
    P_ImprX obj, "Folha " & iPagina & "/" & iQtdPaginas, 100, False
        
    'frmMenu.bcx.caption = chNFe
    
    'obj.PaintPicture frmMenu.bcx.Picture, 132, 28, 65, 8
    
    P_Quadrado obj, iniX + 120, 25, 12, 80
    obj.CurrentY = 37
    obj.FontBold = False
    obj.fontsize = 10
    P_Quadrado obj, iniX + 120, 37, 8, 80, vbSolid, 1, False
    obj.fontsize = 8
    P_PrintXY obj, "chave de acesso", iniX + 148, 37, True
    obj.Print
    P_ImprX obj, FormataChaveDeAcesso(chNFe), 126, False
    obj.Print
    obj.FontBold = False
    obj.fontsize = 8
    obj.CurrentY = 47
    P_ImprQuebraTurbo "Consulta de autenticidade no portal nacional da NF-e www.nfe.fazenda.gov.br/portal ou no site da sefaz autenticadora", _
    obj, 125, 200, False, 2
    
    obj.Print
    
    
    auxY = obj.CurrentY
    
    P_Quadrado obj, iniX, auxY, alturaCX, 120, vbSolid, 2, False, "NATUREZA DA OPERAÇÃO"
    
    obj.CurrentY = auxY + 4
    obj.FontBold = True
    obj.fontsize = 10
    P_ImprX obj, StrConv(natOp, vbUpperCase), iniX + iRecuo, True
    obj.FontBold = False
    obj.fontsize = 8
    P_Quadrado obj, iniX + 120, auxY, alturaCX, 80, vbSolid, 2, False, "PROTOCOLO DE AUTORIZAÇÃO"
    obj.FontBold = True
    obj.fontsize = 10
    P_ImprX obj, nProt & " " & dhRecbto, 130, False
    obj.FontBold = False
    obj.fontsize = 8
    
    obj.Print
    auxY = auxY + alturaCX
    obj.CurrentY = auxY
    P_Quadrado obj, iniX, auxY, alturaCX, 70, vbSolid, 2, False, "INSCRIÇÃO ESTADUAL"
    
    
    P_Quadrado obj, iniX + 70, auxY, alturaCX, 70, vbSolid, 2, False, "INSC.EST.SUBS.TRIB."
    P_Quadrado obj, iniX + 70 + 70, auxY, alturaCX, 60, vbSolid, 2, False, "CNPJ."
    
    auxY = auxY + alt1
        obj.CurrentY = auxY
        obj.FontBold = True
        obj.fontsize = 10
        P_ImprX obj, RetornaValorTagXML(emit, "<IE>"), iniX + iRecuo, True
        P_ImprX obj, RetornaValorTagXML(emit, "<IEST>"), iniX + 120, True
        P_ImprX obj, F_FormataCNPJ(RetornaValorTagXML(emit, "<CNPJ>")), iniX + 150, False
    If iPagina = 1 Then
        
        
        'P_Quadrado obj, iniX + 120, auxY, 10, 80, vbSolid, 2, False, "Protocolo de Autorização"
        auxY = auxY + 4
        obj.CurrentY = auxY
        obj.FontBold = False
        obj.fontsize = 8
        
        P_ImprX obj, "DESTINATÁRIO/ REMETENTE", iniX, False
        auxY = auxY + 4
        obj.CurrentY = auxY
        
        
        P_Quadrado obj, iniX, auxY, alturaCX, 140, vbSolid, 2, False, "NOME/RAZÃO SOCIAL"
        
        'P_Quadrado obj, iniX + 140, auxY, alturaCX, 30, vbSolid, 2, False, "CPF/CNPJ"
        P_Quadrado obj, iniX + 140, auxY, alturaCX, 35, vbSolid, 2, False, "CPF/CNPJ"
        P_Quadrado obj, iniX + 175, auxY, alturaCX, 25, vbSolid, 2, False, "DATA DE EMISSÃO"
        auxY = auxY + 4
        obj.CurrentY = auxY
        obj.FontBold = True
        obj.fontsize = 9
        P_ImprX obj, StrConv(RetornaValorTagXML(dest, "<xNome>"), vbUpperCase), iniX + iRecuo, True
        
        s = RetornaValorTagXML(dest, "<CNPJ>")
        If s = "" Then
            s = F_FormataCPF(RetornaValorTagXML(dest, "<CPF>"))
            
        Else
            s = F_FormataCNPJ(s)
        End If
        P_ImprX obj, s, iniX + 141, True
        P_ImprX obj, Format(Left(dhEmi, 10), "dd/mm/yyyy"), iniX + 175, True
        obj.FontBold = False
        auxY = auxY + 4
        obj.CurrentY = auxY
        obj.fontsize = 8
        
        'auxY = auxY + 1
        obj.CurrentY = auxY
        
        P_Quadrado obj, iniX, auxY, alturaCX, 100, vbSolid, 2, False, "ENDEREÇO"
        P_Quadrado obj, iniX + 100, auxY, alturaCX, 55, vbSolid, 2, False, "BAIRRO"
        P_Quadrado obj, iniX + 155, auxY, alturaCX, 20, vbSolid, 2, False, "CEP"
        
        P_Quadrado obj, iniX + 175, auxY, alturaCX, 25, vbSolid, 2, False, "DATA SAÍDA"
        
        
        auxY = auxY + 4
        obj.CurrentY = auxY
        obj.FontBold = False
        obj.fontsize = 9
        
        
        s = RetornaValorTagXML(dest, "<xLgr>")
        endereco = s
        s = RetornaValorTagXML(dest, "<nro>")
        If s <> "" Then
            endereco = endereco & ", " & s
        End If
        s = RetornaValorTagXML(dest, "<xCpl>")
        
        If s <> "" Then
            endereco = endereco & " " & s
        End If
        
        P_ImprQuebraTurbo StrConv(endereco, vbProperCase), obj, iniX + iRecuo, iniX + iRecuo + 120, False, 1, True
        
        s = RetornaValorTagXML(dest, "<xBairro>")
        P_ImprX obj, StrConv(s, vbUpperCase), iniX + 100 + iRecuo, True
        s = RetornaValorTagXML(dest, "<CEP>")
        P_ImprX obj, F_FormataCEP(s), iniX + 155 + 1, True
        P_ImprX obj, Format(Left(dhRecbto, 10), "dd/mm/yyyy"), iniX + 175, True
        
        auxY = auxY + alt1
        obj.CurrentY = auxY
        
        
        P_Quadrado obj, iniX, auxY, alturaCX, 90, vbSolid, 2, False, "MUNICIPIO"
        P_Quadrado obj, iniX + 90, auxY, alturaCX, 15, vbSolid, 2, False, "UF"
        P_Quadrado obj, iniX + 105, auxY, alturaCX, 35, vbSolid, 2, False, "FONE"
        P_Quadrado obj, iniX + 140, auxY, alturaCX, 35, vbSolid, 2, False, "INSC.ESTADUAL"
        P_Quadrado obj, iniX + 175, auxY, alturaCX, 25, vbSolid, 2, False, "HORA SAÍDA"
        
        auxY = auxY + alt1
        obj.CurrentY = auxY
        P_ImprX obj, StrConv(RetornaValorTagXML(dest, "<xMun>"), vbUpperCase), iniX + iRecuo, True
        P_ImprX obj, StrConv(RetornaValorTagXML(dest, "<UF>"), vbUpperCase), iniX + 90 + iRecuo, True
        P_ImprX obj, StrConv(RetornaValorTagXML(dest, "<fone>"), vbUpperCase), iniX + 105 + iRecuo, True
        P_ImprX obj, StrConv(RetornaValorTagXML(dest, "<IE>"), vbUpperCase), iniX + 140 + iRecuo, True
        P_ImprX obj, Mid(RetornaValorTagXML(ide, "<dhSaiEnt>"), 12, 5), iniX + 175 + iRecuoAmp, True
        
        auxY = auxY + alt1
        obj.CurrentY = auxY
        P_ImprX obj, "FATURA", iniX + iRecuo, False
        auxY = auxY + 4
        obj.CurrentY = auxY
        obj.PaintPicture frmMenu.pcShapeLongo.image, iniX, auxY, 200, 8
        P_ImprX obj, "PAGAMENTO", iniX + iRecuo, False
        obj.DrawWidth = 1
        P_Line obj, iniX, obj.CurrentY, 205, obj.CurrentY
        
        strSQL = RetornaValorTagXML(pag, "<tPag>")
        
        tipopag = RetornaTPag(strSQL)
        vPag = FormatValNFe(RetornaValorTagXML(pag, "<vPag>"))
        
        obj.DrawWidth = 2
        obj.FontBold = True
        P_ImprX obj, tipopag & " " & vPag, iniX + iRecuo, False
        obj.FontBold = False
        
        obj.fontsize = 8
        P_ImprX obj, "CÁLCULO DO IMPOSTO", iniX + iRecuo, False
        
        
        auxY = auxY + 12
        obj.CurrentY = auxY
        obj.PaintPicture frmMenu.pcShapeLongo.image, 160 + iniX, auxY, 40, 9
        
        P_Quadrado obj, iniX, auxY, alturaCX, 40, vbSolid, 2, False, "BASE DE CÁLCULO ICMS"
        P_Quadrado obj, iniX + 40, auxY, alturaCX, 40, vbSolid, 2, False, "VALOR DO ICMS"
        P_Quadrado obj, iniX + 80, auxY, alturaCX, 40, vbSolid, 2, False, "B.C. ICMS S.T."
        P_Quadrado obj, iniX + 120, auxY, alturaCX, 40, vbSolid, 2, False, "VALOR ICMS S.T."
        P_Quadrado obj, iniX + 160, auxY, alturaCX, 40, vbSolid, 2, False, "VALOR TOTAL PRODUTOS"
        
        auxY = auxY + 3
        obj.CurrentY = auxY
        obj.fontsize = 10
        obj.FontBold = True
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vBC>")), iniX + iRecuo, True
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vICMS>")), iniX + 42, True
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vBCST>")), iniX + 82, True
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vST>")), iniX + 122, True
        
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vProd>")), iniX + 197, True, direita
        obj.FontBold = False
        
        auxY = auxY + alt1 + 1
        obj.CurrentY = auxY
        obj.fontsize = 8
        obj.PaintPicture frmMenu.pcShapeLongo.image, 160 + iniX, auxY, 40, 9
        P_Quadrado obj, iniX, auxY, alturaCX, 30, vbSolid, 2, False, "VALOR DO FRETE"
        P_Quadrado obj, iniX + 30, auxY, alturaCX, 30, vbSolid, 2, False, "VLR SEGURO"
        P_Quadrado obj, iniX + 60, auxY, alturaCX, 30, vbSolid, 2, False, "VLR DESCONTO"
        P_Quadrado obj, iniX + 90, auxY, alturaCX, 30, vbSolid, 2, False, "OUTRAS DESP."
        P_Quadrado obj, iniX + 120, auxY, alturaCX, 40, vbSolid, 2, False, "IPI"
        obj.FontBold = True
        P_Quadrado obj, iniX + 160, auxY, alturaCX, 40, vbSolid, 2, False, "VALOR TOTAL DA NOTA"
        
        obj.fontsize = 11
        obj.FontBold = False
        auxY = auxY + 3
        obj.CurrentY = auxY
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vFrete>")), iniX + 2, True
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vSeg>")), iniX + 32, True
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vDesc>")), iniX + 62, True
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vOutro>")), iniX + 92, True
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vIPI>")), iniX + 122, True
        obj.FontBold = True
        P_ImprX obj, FormatValNFe(RetornaValorTagXML(sTot, "<vNF>")), iniX + 197, True, direita
        obj.FontBold = False
        
        
        auxY = auxY + 5
        obj.CurrentY = auxY
        
        
        obj.fontsize = 8
        P_ImprX obj, "TRANSPORTADOR/VOLUMES TRANSPORTADOS", iniX, False
        
        auxY = auxY + 4
        obj.CurrentY = auxY
        P_Quadrado obj, iniX, auxY, alturaCX, 80, vbSolid, 2, False, "NOME/RAZÃO SOCIAL"
        P_Quadrado obj, iniX + 80, auxY, alturaCX, 30, vbSolid, 2, False, "FRETE POR CONTA"
        P_Quadrado obj, iniX + 110, auxY, alturaCX, 15, vbSolid, 2, False, "C.ANTT"
        P_Quadrado obj, iniX + 125, auxY, alturaCX, 25, vbSolid, 2, False, "PLACA VEÍCULO"
        P_Quadrado obj, iniX + 150, auxY, alturaCX, 10, vbSolid, 2, False, "UF"
        P_Quadrado obj, iniX + 160, auxY, alturaCX, 40, vbSolid, 2, False, "CNPJ/CPF"
        
        obj.fontsize = 11
        obj.FontBold = False
        auxY = auxY + 3
        obj.CurrentY = auxY
        s = RetornaValorTagXML(transp, "<modFrete>")
        If s = "0" Then
            s = "0 - Emitente"
        ElseIf s = "1" Then
            s = "1 - Destinatário/Rem."
        ElseIf s = "2" Then
            s = "2 - Terceiros"
        Else
            s = "9 - Sem Frete"
        End If
        obj.fontsize = 9
        P_ImprX obj, RetornaValorTagXML(transp, "<xNome>"), iniX + 2, True
        P_ImprX obj, s, iniX + 82, True
        P_ImprX obj, (RetornaValorTagXML(transp, "<RNTC>")), iniX + 110, True
        P_ImprX obj, (RetornaValorTagXML(sTot, "<placa>")), iniX + 125, True
        P_ImprX obj, (RetornaValorTagXML(sTot, "<UF>")), iniX + 150, True
        obj.FontBold = True
        s = RetornaValorTagXML(transp, "<CNPJ>")
        If s = "" Then
            s = RetornaValorTagXML(transp, "<CPF>")
        End If
        P_ImprX obj, s, iniX + 197, True, direita
        
        obj.FontBold = False
        obj.fontsize = 8
        
        auxY = auxY + 5
        obj.CurrentY = auxY
        P_Quadrado obj, iniX, auxY, alturaCX, 80, vbSolid, 2, False, "ENDEREÇO"
        P_Quadrado obj, iniX + 80, auxY, alturaCX, 70, vbSolid, 2, False, "MUNICIPIO"
        P_Quadrado obj, iniX + 150, auxY, alturaCX, 10, vbSolid, 2, False, "UF"
        P_Quadrado obj, iniX + 160, auxY, alturaCX, 40, vbSolid, 2, False, "IE."
        
        obj.fontsize = 9
        obj.FontBold = False
        auxY = auxY + 3
        obj.CurrentY = auxY
        
        yRet = P_ImprQuebraTurbo(StrConv(RetornaValorTagXML(transp, "<xEnder>"), vbProperCase), obj, iniX + iRecuo, iniX + 2 + 50, False, 1, True)
        P_ImprX obj, (RetornaValorTagXML(transp, "<xMun>")), iniX + 82, True
        P_ImprX obj, (RetornaValorTagXML(transp, "<UF>")), iniX + 152, True
        P_ImprX obj, (RetornaValorTagXML(transp, "<IE>")), iniX + 162, True
    Else
        auxY = auxY + 5
    End If
    auxY = auxY + 5
    obj.CurrentY = auxY
    
    
    obj.fontsize = 7
    obj.FontBold = True
    P_CentraImpressao obj, "DADOS DOS PRODUTOS/SERVIÇOS", False, iniX
    auxY = auxY + 4
    obj.CurrentY = auxY
    
    
   ' obj.PaintPicture frmMenu.pcShapeLongo.image, iniX, auxY, 200, 6
    P_Quadrado obj, iniX - 3, auxY, alturaCX - 3, 3, vbSolid, 1, False, "N"
    P_Quadrado obj, iniX, auxY, alturaCX - 3, 15, vbSolid, 1, False, "Código"
    P_Quadrado obj, iniX + 15, auxY, alturaCX - 3, 65, vbSolid, 1, False, "Descrição do Produto"
    P_Quadrado obj, iniX + 80, auxY, alturaCX - 3, 15, vbSolid, 1, False, "NCM/SH"
    P_Quadrado obj, iniX + 95, auxY, alturaCX - 3, 8, vbSolid, 1, False, "CST"
    P_Quadrado obj, iniX + 103, auxY, alturaCX - 3, 8, vbSolid, 1, False, "CFOP"
    P_Quadrado obj, iniX + 111, auxY, alturaCX - 3, 9, vbSolid, 1, False, "UND."
    P_Quadrado obj, iniX + 120, auxY, alturaCX - 3, 10, vbSolid, 1, False, "Qtd."
    P_Quadrado obj, iniX + 130, auxY, alturaCX - 3, 10, vbSolid, 1, False, "Unit."
    P_Quadrado obj, iniX + 140, auxY, alturaCX - 3, 10, vbSolid, 1, False, "Total."
    P_Quadrado obj, iniX + 150, auxY, alturaCX - 3, 10, vbSolid, 1, False, "Desc."
    obj.fontsize = 6
    P_Quadrado obj, iniX + 160, auxY, alturaCX - 3, 10, vbSolid, 1, False, "BCICMS"
    P_Quadrado obj, iniX + 170, auxY, alturaCX - 3, 10, vbSolid, 1, False, "VICMS"
    P_Quadrado obj, iniX + 180, auxY, alturaCX - 3, 10, vbSolid, 1, False, "VIPI"
    obj.fontsize = 5
    P_Quadrado obj, iniX + 190, auxY, alturaCX - 3, 5, vbSolid, 1, False, "ICMS"
    P_Quadrado obj, iniX + 195, auxY, alturaCX - 3, 5, vbSolid, 1, False, "IPI"
    obj.DrawWidth = 1
    obj.CurrentY = obj.CurrentY + 1
    P_Line obj, iniX, obj.CurrentY, iniX, yFimQuadroProdutos
    P_Line obj, iniX + 15, obj.CurrentY, iniX + 15, yFimQuadroProdutos
    P_Line obj, iniX + 80, obj.CurrentY, iniX + 80, yFimQuadroProdutos
    P_Line obj, iniX + 95, obj.CurrentY, iniX + 95, yFimQuadroProdutos
    P_Line obj, iniX + 95, obj.CurrentY, iniX + 95, yFimQuadroProdutos
    P_Line obj, iniX + 103, obj.CurrentY, iniX + 103, yFimQuadroProdutos
    P_Line obj, iniX + 111, obj.CurrentY, iniX + 111, yFimQuadroProdutos
    P_Line obj, iniX + 120, obj.CurrentY, iniX + 120, yFimQuadroProdutos
    P_Line obj, iniX + 130, obj.CurrentY, iniX + 130, yFimQuadroProdutos
    P_Line obj, iniX + 140, obj.CurrentY, iniX + 140, yFimQuadroProdutos
    P_Line obj, iniX + 150, obj.CurrentY, iniX + 150, yFimQuadroProdutos
    P_Line obj, iniX + 160, obj.CurrentY, iniX + 160, yFimQuadroProdutos
    P_Line obj, iniX + 170, obj.CurrentY, iniX + 170, yFimQuadroProdutos
    P_Line obj, iniX + 180, obj.CurrentY, iniX + 180, yFimQuadroProdutos
    P_Line obj, iniX + 190, obj.CurrentY, iniX + 190, yFimQuadroProdutos
    P_Line obj, iniX + 195, obj.CurrentY, iniX + 195, yFimQuadroProdutos
    P_Line obj, iniX + 200, obj.CurrentY, iniX + 200, yFimQuadroProdutos
    P_Line obj, iniX, yFimQuadroProdutos, 205 + xIni, yFimQuadroProdutos
    
    auxY = auxY + 6
    obj.CurrentY = auxY
    
    Dim y As Single
    Dim yTamanho As Single
    Dim codigo As String
    Dim cod_fab As String
    obj.fontsize = 7
    Do While det <> ""
        det = RetornaValorTagXML(conteudoXML, "<det nItem=" & aspas & iProd & aspas & ">", "</det>")
        If det <> "" Then
            obj.FontBold = False
            obj.fontsize = 5
            
            
            
            P_ImprX obj, CStr(iProd), 2
            obj.fontsize = 7
            codigo = RetornaValorTagXML(det, "<cProd>")
            cod_fab = F_BuscaUmCampo("codfab_pro", TAB_PRODUTOS, "cod_pro='" & codigo & "'")
            obj.fontsize = 6
            P_ImprX obj, codigo, iniX + 2
            
            Set p = obj
            p.CurrentY = obj.CurrentY
           ' Y = P_ImprQuebraTurbo(RetornaValorTagXML(det, "<xProd>") & " (" & cod_fab & ")", p, iniX + 16, iniX + 80, False, 1, True)
           
            
            
            'If Y <= 250 Then
                y = P_ImprQuebraTurbo(RetornaValorTagXML(det, "<xProd>") & " (" & cod_fab & ")", obj, iniX + 16, iniX + 80, False, 1, True)
           'Else
            'MsgBox ""
            'GoTo fim
            'End If
            
            
            
           
           
           
           'erro
           
            obj.fontsize = 7
            P_ImprX obj, Left(RetornaValorTagXML(det, "<NCM>"), 20), iniX + 81
            P_ImprX obj, RetornaValorTagXML(det, "<orig>") & RetornaValorTagXML(det, "<CSOSN>"), iniX + 96
            P_ImprX obj, RetornaValorTagXML(det, "<CFOP>"), iniX + 104
            P_ImprX obj, RetornaValorTagXML(det, "<uCom>"), iniX + 113
            P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<qCom>"), "###,##0.000"), 126
            P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<vUnCom>")), 136
            
            P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<vProd>")), 146
            P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<vDesc>")), 156
            P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<vBC>")), 166
            P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<vICMS>")), 176
            P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<vIPI>")), 186
            obj.fontsize = 6
            P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<pICMS>"), ""), 196
            P_ImprX obj, FormatValNFe(RetornaValorTagXML(det, "<pIPI>"), ""), 201, False
            iProd = iProd + 1
            obj.CurrentY = y
            obj.Print
            
        End If
        
        If iProd = 31 Then
         '   MsgBox ""
        End If
        
        obj.fontsize = 7
        
        
            If iPagina = 1 Then
                
            Else
                
            End If
            
        
            If ((iProd - 1) Mod (itensPorPagina) = 0 And iPagina = 1) Or ((iProd - 1) Mod (itensPorPagina) = 0 And iPagina > 1) Or (iProd > iQtdItens) And (det <> "") Then
            'If obj.CurrentY >= obj.ScaleHeight - 20 Then
                iPagina = iPagina + 1
        
            
        
fim:
            auxY = 260
            obj.CurrentY = auxY
            obj.fontsize = 10
            P_CentraImpressao obj, "DADOS ADICIONAIS", False, iniX
            auxY = auxY + 5
            obj.CurrentY = auxY
            P_Quadrado obj, iniX, auxY, 30, 120, vbSolid, 1, False, "INFORMAÇÕES COMPLEMENTARES"
            P_Quadrado obj, iniX + 120, auxY, 30, 80, vbSolid, 1, False, "RESERVADO AO FISCO"
            auxY = auxY + 5
            obj.CurrentY = auxY
            
            obj.fontsize = 8
            obj.FontBold = False
            P_ImprQuebraTurbo RetornaValorTagXML(infAdic, "<infCpl>"), obj, iniX + 1, iniX + 120
            
           ' P_ImprX obj, RetornaValorTagXML(infAdic, "<infCpl>"), iniX + 120
            
            imprimiuFim = True
            
            If iPagina <= iQtdPaginas Then
                GoTo inicio
            Else
                If TypeName(obj) = "PictureBox" And iPagina > 1 Then
                    frmVisualizadorImpressao.addNewPage
                ElseIf TypeName(obj) = "Printer" And iPagina > 1 Then
                    Printer.NewPage
                End If
            End If
        Else
        
        End If
    Loop
    
    If Not imprimiuFim Then
        GoTo fim
    End If
    
termina:
    
    If TypeName(obj) = "Printer" Then
        Printer.EndDoc
        ImprimirXMLNFeDANFE = 0
    Else
        ImprimirXMLNFeDANFE = 1
    End If
    
    
     
    
    P_Status ""
Exit Function
trata_Erro:
P_Status
P_MostraErro
End Function


Function FormatValNFe(sValor, Optional sFormat As String = "###,###,##0.00") As String
    On Error GoTo trata_Erro
    sValor = mudarCaractere(CStr(sValor), ".", ",")
    
    
    FormatValNFe = Format(sValor, sFormat)
    If FormatValNFe = "" Then FormatValNFe = Format("0,00", sFormat)
    Exit Function
    
trata_Erro:
    FormatValNFe = Format("0,00", sFormat)
End Function

Function GeraLoteNFe(id_nfe As String) As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into Nfe_lote(horario,id_nfe)values("
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & id_nfe & "')"
    Conn.Execute strSQL
    
    GeraLoteNFe = MaxKeySQL("nfe_lote", "id", "", "")
trata_Erro:

End Function
Function BuscaChaveNFe(id As String) As String
    BuscaChaveNFe = F_BuscaUmCampo("chave", "nfe", "id=" & id)
End Function
Function BuscaIdPelaChave(Chave As String) As String
    BuscaIdPelaChave = F_BuscaUmCampo("id", "nfe", "chave='" & Chave & "'")
End Function
Function BuscaModeloPelaChave(Chave As String) As String
    BuscaModeloPelaChave = F_BuscaUmCampo("modelo", "nfe", "chave='" & Chave & "'")
End Function

Function BuscaModeloPeloID(id As String) As String
    BuscaModeloPeloID = F_BuscaUmCampo("modelo", "nfe", "id=" & id)
End Function

Function BuscaIdPelaVenda(venda As String) As String
    BuscaIdPelaVenda = F_BuscaUmCampo("id", "nfe", "num_Ven='" & venda & "'")
End Function

Function BuscaAnoMesNFe(id As String) As String
    Dim s As String
    s = F_BuscaUmCampo("dhEmi", "nfe", "id=" & id)
    If s <> "" Then
        BuscaAnoMesNFe = Format(Left(s, 10), "YYYYMM")
    End If
    
    
    
End Function
Function buscaStatusNFe(id As String) As String
    buscaStatusNFe = F_BuscaUmCampo("status", "nfe", "id=" & id)
End Function

Sub AtualizarRetorno(id As String, Optional bMsg As Boolean = True)
    Dim s As String
    Dim chNFe As String
    Dim DigestValue As String
    Dim autorizacao As String
    Dim sRecibo As String
    Dim Inf As infProt
    Dim Chave As String
    On Error GoTo trata_Erro
    Chave = BuscaChaveNFe(id)
    s = XmlRetornoNFe(Chave)
    
    
        If s = "" Then
                Exit Sub
            
        
        End If
    
    DigestValue = RetornaValorTagXML(s, "<DigestValue>")
    Inf = RetornaInfProt(s)
    If Inf.cStat = "100" Then
        
        AlteraDadosFinaisNFe id, "2", mudarCaractere(RetornoCaminhoNFePeloID(id), "\", "/"), Inf.nProt, Inf.digVal, DigestValue
        
        If bMsg Then MsgBox Inf.xMotivo, vbInformation
    
    Else
        If Inf.cStat = "101" Then
            AlteraDadosFinaisNFe ado!id, "5", mudarCaractere(RetornoCaminhoNFePeloID(ado!id), "\", "/"), Inf.nProt, Inf.digVal, DigestValue
            If bMsg Then MsgBox Inf.xMotivo, vbInformation
        End If
        
    End If
    
    
    'MsgBox s
trata_Erro:

End Sub
Function VerificaNCMVenda(num_ven As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sMsg As String
    On Error GoTo trata_Erro:
    
    strSQL = "select ti.cod_Pro,tp.ncm_pro, desc_ite from " & TAB_ITVENDAS & " as ti inner join "
    strSQL = strSQL & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and num_ven='" & num_ven & "' "
    strSQL = strSQL & " and (ncm_pro='' or ncm_pro is null) "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            sMsg = sMsg & "Produto: " & !cod_pro & "-" & !desc_ite & CrLf
            .MoveNext
        Loop
    End With
    If sMsg <> "" Then
        If sMsg <> "" Then
            VerificaNCMVenda = sMsg
            
        End If
    Else
        VerificaNCMVenda = ""
    
    End If
    
    
    
    Exit Function
trata_Erro:
End Function
Function BuscaVendaPeloID(id As String) As String
    BuscaVendaPeloID = F_BuscaUmCampo("num_ven", "nfe", "id=" & id)
End Function
Sub P_BaixaSaidaNFe(id As String, Optional saida As Boolean = True)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sinal As String
    On Error GoTo trata_Erro
    
    If saida Then
        sinal = "-"
    Else
        sinal = "+"
    End If
    
    
    strSQL = "select * from nfe_item where id_nfe=" & id
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            Do While Not .EOF
                strSQL = "update " & TAB_PRODUTOS & " set ef_pro=ef_pro" & sinal & str(!qCom)
                strSQL = strSQL & " where cod_pro='" & !cProd & "' "
                Conn.Execute strSQL
                
                
                strSQL = "update estoque set ef=ef" & sinal & str(!qCom)
                strSQL = strSQL & " where cod_pro='" & !cProd & "' and cod_emp='" & !cod_emp & "' "
                Conn.Execute strSQL
                .MoveNext
            Loop
            .Close
            Exit Sub
        End If
    End With
    
    
P_BaixaSaidaNFeAntiga id, saida
    
    Exit Sub
trata_Erro:
        
        
    
End Sub
Sub P_BaixaSaidaNFeAntiga(id As String, Optional saida As Boolean = True)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sinal As String
    On Error GoTo trata_Erro
    
    
        Dim c As New ADODB.Connection
    
    Dim rItens As New ADODB.Recordset
    Dim cod_For As String
    Dim iItem As Integer
    Dim ESP As String
    If F_AbreBancoODBCAdo(c, "muitomais") Then
    
    End If

    
    If saida Then
        sinal = "-"
    Else
        sinal = "+"
    End If
    
    
    strSQL = "select * from nfe_item where id_nfe=" & id
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            Do While Not .EOF
                strSQL = "update tab_produtos set ef_pro=ef_pro" & sinal & str(!qCom)
                strSQL = strSQL & " where cod_pro='" & !cProd & "' "
                Conn.Execute strSQL
                
                
                'strSQL = "update estoque set ef=ef" & sinal & str(!qCom)
                'strSQL = strSQL & " where cod_pro='" & !cProd & "' and cod_emp='" & !cod_emp & "' "
                'conn.Execute strSQL
                .MoveNext
            Loop
            .Close
            Exit Sub
        End If
    End With
    
    
            
    
    Exit Sub
trata_Erro:
        
        
    
End Sub

Function BuscaCaminhoPeloId(id As String) As String

    Dim caminho As String
    caminho = F_BuscaUmCampo("caminhoxml", "nfe", "id='" & id & "'")

    BuscaCaminhoPeloId = caminho
End Function
Function BuscaEmailClientePeloId(id As String, Optional Id0Chave1 As Integer = 0)
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select email_cli from " & TAB_CLIENTES & " as tc inner join " & TAB_VENDAS & " as tv "
    strSQL = strSQL & " on tv.cod_cli=tc.cod_Cli inner join nfe on nfe.num_ven=tv.num_ven "
    If Id0Chave1 = 0 Then
        strSQL = strSQL & " where nfe.id=" & id
    Else
        strSQL = strSQL & " where nfe.chave='" & id & "'"
    End If
    strSQL = strSQL & " and tipo_operacao=1"
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        BuscaEmailClientePeloId = TiraNull(rPesq!email_cli)
    End If
    rPesq.Close
    
End Function


Sub PreencheTipoPagNFe(cmbo As ComboBox)
    P_Preenche cmbo, "nfe_tpag", "concat(tpag,' - ',nome)"
End Sub

Sub MudaCorEmpresa(objeto As Object)
    On Error Resume Next
    If strCor_Emp <> "" Then
        objeto.BackColor = strCor_Emp
    End If

End Sub
Function buscaPRedBCICMS20() As Single
    Dim r As String
    r = BuscaParametroEstrutura("PRedICMS20").valor
    If r = "" Then
        buscaPRedBCICMS20 = 0.3529
    Else
        buscaPRedBCICMS20 = CSng(r)
    End If
    
    
    
    
End Function
Function grava_item_ipi() As String
    


End Function
