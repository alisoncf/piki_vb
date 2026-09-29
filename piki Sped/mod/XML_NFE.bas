Attribute VB_Name = "XML_NFE"
Public ambiente As tpAmb
Public strXML_NFE_CabXML As String

Public Const aspas As String = """"
Public Const espaco As String = " "
Public VersaoNFe As String
Dim cItens As New Collection











Public Enum TMod
    NFe = 55
    nfc_e = 65

End Enum

Enum eIdDest
    interna = 1
    interestadual = 2
    exterior = 3
End Enum




Public emitente As emi_nfe


Function AleatorioNFe() As String
    Dim s As String
    s = Format((Rnd() * 99999999), "0#######")
    s = Format((Rnd() * 99999999), "0#######")
    AleatorioNFe = s
End Function
Function UltimaNFe() As String
    UltimaNFe = MaxKeySQL("nfe", "nnf", "", "")
End Function

Sub IniciaComponentesXMLNFE(chave As String)

    emitente.CNPJ = F_SoNumero(strCNPJ_Emp)
    emitente.xNome = strNome_Emp
    emitente.xFant = strFantasia_Emp
    emitente.ende.cep = F_SoNumero(strCEP_Emp)
    emitente.ende.cMun = strCodigoMunicipioIBGE
    emitente.ende.cPais = strCodigoPaisIBGE
    emitente.ende.fone = F_SoNumero(strFone_Emp)
    emitente.ende.nro = strNumero_Emp
    emitente.ende.UF = strUF_Emp
    emitente.ende.xBairro = strBairro_Emp
    emitente.ende.xCpl = strComplemento_Emp
    emitente.ende.xLgr = strLogradouro_Emp
    emitente.ende.xMun = strCidade_Emp
    emitente.ende.xPais = strPais
    emitente.ie = F_SoNumero(strIE_Emp)
    emitente.IEST = strIEST_Emp
    emitente.im = strIM_Emp
    emitente.CNAE = strCNAE_Emp
    emitente.cCRT = strCRT_emp
    
    'If emitente.cCRT = "" Then
    '    emitente.cCRT = SIMPLES_NACIONAL
    'End If
    
    
    VersaoNFe = "4.0"

    strXML_NFE_CabXML = "<?xml version=" & aspas & "1.0" & aspas & espaco & "encoding=" & aspas & "UTF-8" & _
    aspas & "?>"
    strXML_NFE_CabXML = strXML_NFE_CabXML & "<NFe xmlns=" & aspas & "http://www.portalfiscal.inf.br/nfe" & aspas & "><infNFe Id=" & _
    aspas & chave & aspas & " versao=" & aspas & VersaoNFe & aspas & ">"
    
    
    
End Sub

Function RetornaCabNFeXML(chave As String, versao As String) As String
    Dim s As String
    
End Function




Function EMIT_NFE(emi As emi_nfe) As String
    Dim s As String
    
    
    'cNF As String, natOP As String, ipag As indPag, modelo As TMod, serie As String, nNf As String, _
    'dhEmi As String, tpNF As TtpNF, idDest As String, tipoImp As tpImp, tpEmis As String, cDv As String, _
    'tipoAmb As tpAmb, fin As TfinNfe, indiFin As indFinal, indP As indPres, procEmi As String, verProc As String
    
    
    
    s = "<emit>"
    s = s & "<CNPJ>" & emi.CNPJ
    s = s & "</CNPJ>"
    s = s & "<xNome>" & emi.xNome
    s = s & "</xNome>"
    s = s & "<xFant>" & emi.xFant
    s = s & "</xFant>"
    
    s = s & "<enderEmit>"
    s = s & "<xLgr>" & emi.ende.xLgr & "</xLgr>"
    s = s & "<nro>" & emi.ende.nro & "</nro>"
    s = s & "<xCpl>" & emi.ende.xCpl & "</xCpl>"
    s = s & "<xBairro>" & emi.ende.xBairro & "</xBairro>"
    s = s & "<cMun>" & emi.ende.cMun & "</cMun>"
    s = s & "<xMun>" & emi.ende.xMun & "</xMun>"
    s = s & "<UF>" & emi.ende.UF & "</UF>"
    s = s & "<CEP>" & emi.ende.cep & "</CEP>"
    s = s & "<cPais>" & emi.ende.cPais & "</cPais>"
    s = s & "<xPais>" & emi.ende.xPais & "</xPais>"
    s = s & "<fone>" & emi.ende.fone & "</fone>"
    s = s & "</enderEmit>"
    s = s & "<IE>" & emi.ie & "</IE>"
    s = s & "<CRT>" & emi.cCRT & "</CRT>"
    s = s & "</emit>"
    
    EMIT_NFE = s
End Function
Function DEST_NFE(d As dest) As String
    Dim s As String
    
    
    'cNF As String, natOP As String, ipag As indPag, modelo As TMod, serie As String, nNf As String, _
    'dhd As String, tpNF As TtpNF, idDest As String, tipoImp As tpImp, tpds As String, cDv As String, _
    'tipoAmb As tpAmb, fin As TfinNfe, indiFin As indFinal, indP As indPres, procd As String, verProc As String
    
    
    
    s = "<dest>"
    If d.CNPJ <> "" Then
        s = s & "<CNPJ>" & d.CNPJ
        s = s & "</CNPJ>"
    Else
        s = s & "<CPF>" & d.CPF
        s = s & "</CPF>"
    End If
    s = s & "<xNome>" & d.xNome
    s = s & "</xNome>"
    
    
    s = s & "<enderdest>"
    s = s & "<xLgr>" & d.ende.xLgr & "</xLgr>"
    s = s & "<nro>" & d.ende.nro & "</nro>"
    s = s & "<xCpl>" & d.ende.xCpl & "</xCpl>"
    s = s & "<xBairro>" & d.ende.xBairro & "</xBairro>"
    s = s & "<cMun>" & d.ende.cMun & "</cMun>"
    s = s & "<xMun>" & d.ende.xMun & "</xMun>"
    s = s & "<UF>" & d.ende.UF & "</UF>"
    s = s & "<CEP>" & d.ende.cep & "</CEP>"
    s = s & "<cPais>" & d.ende.cPais & "</cPais>"
    s = s & "<xPais>" & d.ende.xPais & "</xPais>"
    s = s & "<fone>" & d.ende.fone & "</fone>"
    s = s & "<indIEDest>" & d.ende.IndIEDest & "</indIEDest>"
    s = s & "<IE>" & d.ende.ie & "</IE>"
    s = s & "<ISUF>" & d.ende.isuf & "</ISUF>"
    s = s & "<IM>" & d.ende.im & "</IM>"
    s = s & "<email>" & d.ende.email & "</email>"
    
    s = s & "</enderdest>"
    s = s & "</dest>"
    DEST_NFE = s
End Function
Function TranspNFe(id As String) As String
    Dim s As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from nfe_transp where id_nfe=" & id
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            s = "<transp>"
            s = s & "<modFrete>9</modFrete>"
            s = s & "</transp>"
        
            Exit Function
        End If
        s = "<transp>"
        s = s & "<modFrete>" & !modFrete & "</modFrete>"
        s = s & "</transp>"
        
        If !modFrete <> "9" Then
        End If
        .Close
    End With
    TranspNFe = s
    
    
End Function
Function infAdicNFe(id As String) As String
    Dim s As String
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from nfe_infAdicNfe where id_nfe=" & id
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            s = "<infAdic>"
        
            s = s & "<infCpl/>"
            s = s & "</infAdic></infNFe>"
            infAdicNFe = s
            Exit Function
        End If
        
        s = "<infAdic>"
        
        s = s & "<infCpl>" & !infCpl & "</infCpl>"
        s = s & "</infAdic></infNFe>"
    End With
infAdicNFe = s

End Function
Function NFePag(id As String) As String
    Dim s As String
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from nfe_pag where id_nfe=" & id
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
        
        s = "<pag>"
        s = s & "<tPag>" & !tPag & "</tPag>"
        s = s & "<vPag>" & Valor2Casas(!vPag) & "</vPag>"
        s = s & "<card>" & !card & "</card>"
        s = s & "<CNPJ>" & !CNPJ & "</CNPJ>"
        s = s & "<tBand>" & !tBand & "</tBand>"
        s = s & "<cAut>" & !cAut & "</cAut>"
        s = s & "</pag>"
        End If
    End With

NFePag = s
End Function
Function itemNFe(id As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rAux As New ADODB.Recordset
    Dim s As String
    Dim t_vBC As Currency
    Dim t_vICMS As Currency
    Dim t_vICMSDeson As Currency
    Dim t_vBCST As Currency
    Dim t_vST As Currency
    Dim t_vProd As Currency
    Dim t_vFrete As Currency
    Dim t_vSeg As Currency
    Dim t_vDesc As Currency
    Dim t_vII As Currency
    Dim t_vIPI As Currency
    Dim t_vPIS As Currency
    Dim t_vCOFINS As Currency
    Dim t_vOutro As Currency
    Dim t_vNF As Currency
    Dim t_vTotTrib As Currency
    s = ""
    
    strSQL = "select a.*,b.* from nfe_item as a inner join nfe_item_imposto as b on a.id_nfe=b.id_nfe"
    strSQL = strSQL & " where a.id_nfe=" & id
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            s = s & "<det nItem=" & aspas & .AbsolutePosition & aspas & ">"
            s = s & "<prod>"
            s = s & "<cProd>" & !cProd & "</cProd>"
            s = s & "<cEAN>" & !cEAN & "</cEAN>"
            s = s & "<xProd>" & !xProd & "</xProd>"
            s = s & "<NCM>" & "</NCM>"
            s = s & "<CFOP>" & !CFOP & "</CFOP>"
            s = s & "<uCom>" & !uCom & "</uCom>"
            s = s & "<qCom>" & Valor4Casas(!qCom) & "</qCom>"
            s = s & "<vUnCom>" & Valor2Casas(!vUnCom) & "</vUnCom>"
            s = s & "<vProd>" & Valor2Casas(!vProd) & "</vProd>"
            t_vProd = t_vProd + !vProd
            s = s & "<cEANTrib>" & !cEANTrib & "</cEANTrib>"
            s = s & "<uTrib>" & !uTrib & "</uTrib>"
            s = s & "<qTrib>" & Valor4Casas(!qTrib) & "</qTrib>"
            s = s & "<vUnTrib>" & Valor2Casas(!vUnTrib) & "</vUnTrib>"
            s = s & "<vDesc>" & Valor2Casas(!vDesc) & "</vDesc>"
            t_vDesc = t_vDesc & !vDesc
            s = s & "<indTot>" & !indTot & "</indTot>"
            s = s & "</prod>"
            s = s & "<imposto>"
            s = s & "<vTotTrib>" & !vTotTrib & "</vTotTrib>"
            t_vTotTrib = t_vTotTrib + !vTotTrib
            
            
            
            If !Tipo_Icms = "102" Or !Tipo_Icms = "ICMSSN102" Then
                strSQL = "select * from nfe_item_imposto_icmssn102 where id=" & !id
                abreObjetoAdo rAux, strSQL
                With rAux
                    s = s & "<ICMS><ICMSSN102>"
                    s = s & "<orig>" & !orig & "</orig>"
                    s = s & "<CSOSN>" & !csosn & "</CSOSN></ICMSSN102></ICMS>"
                    .Close
                End With
            End If
            If !Tipo_Icms = "500" Or !Tipo_Icms = "ICMSSN500" Then
                strSQL = "select * from nfe_item_imposto_icmssn500 where id=" & !id
                abreObjetoAdo rAux, strSQL
                With rAux
                    s = s & "<ICMS><ICMSSN500>"
                    s = s & "<orig>" & !orig & "</orig>"
                    s = s & "<CSOSN>" & !csosn & "</CSOSN></ICMSSN500></ICMS>"
                .Close
                End With
            End If
            
            If !tipo_pis = "PISOutr" Then
                strSQL = "select * from nfe_item_imposto_pisoutr where id=" & !id
                abreObjetoAdo rAux, strSQL
                With rAux
                    s = s & "<PIS>"
                    s = s & "<PISOutr>"
                    s = s & "<CST>" & !cst & "</CST>"
                    s = s & "<vBC>" & Valor2Casas(!vBC) & "</vBC>"
                    s = s & "<pPIS>" & Valor4Casas(!pPIS) & "</pPIS>"
                    s = s & "<vPIS>" & Valor2Casas(!vPIS) & " </vPIS>"
                    s = s & "</PISOutr>"
                    s = s & "</PIS>"
                    t_vPIS = t_vPIS + !vPIS
                    .Close
                End With
            End If
            If !tipo_cofins = "COFINSOutr" Then
                strSQL = "select * from nfe_item_imposto_cofinsoutr where id=" & !id
                abreObjetoAdo rAux, strSQL
                With rAux
                    s = s & "<COFINS>"
                    s = s & "<COFINSOutr>"
                    s = s & "<CST>" & !cst & "</CST>"
                    s = s & "<vBC>" & Valor2Casas(!vBC) & "</vBC>"
                    s = s & "<pCOFINS>" & Valor4Casas(!pCOFINS) & "</pCOFINS>"
                    s = s & "<vCOFINS>" & Valor2Casas(!vCOFINS) & " </vCOFINS>"
                    s = s & "</COFINSOutr>"
                    s = s & "</COFINS>"
                    t_vCOFINS = t_vCOFINS + !vCOFINS
                    .Close
                End With
            End If
            s = s & "</imposto>"
            
            
            .MoveNext
            
        Loop
        t_vNF = t_vProd - t_vDesc
        s = s & "<total>"
        s = s & "<ICMSTot>"
        s = s & "<vBC>" & Valor2Casas(t_vBC) & "</vBC>"
        s = s & "<vICMS>" & Valor2Casas(t_vICMS) & "</vICMS>"
        s = s & "<vICMSDeson>" & Valor2Casas(t_vICMSDeson) & "</vICMSDeson>"
        s = s & "<vBCST>" & Valor2Casas(t_vBC) & "</vBCST>"
        s = s & "<vST>" & Valor2Casas(t_vST) & "</vST>"
        s = s & "<vProd>" & Valor2Casas(t_vProd) & "</vProd>"
        s = s & "<vFrete>" & Valor2Casas(t_vFrete) & "</vFrete>"
        s = s & "<vSeg>" & Valor2Casas(t_vSeg) & "</vSeg>"
        s = s & "<vDesc>" & Valor2Casas(t_vDesc) & "</vDesc>"
        s = s & "<vII>" & Valor2Casas(t_vII) & "</vII>"
        s = s & "<vIPI>" & Valor2Casas(t_vIPI) & "</vIPI>"
        s = s & "<vPIS>" & Valor2Casas(t_vIPI) & "</vPIS>"
        s = s & "<vCOFINS>" & Valor2Casas(t_vCOFINS) & "</vCOFINS>"
        s = s & "<vOutro>" & Valor2Casas(t_vOutro) & "</vOutro>"
        s = s & "<vNF>" & Valor2Casas(t_vNF) & "</vNF>"
        s = s & "<vTotTrib>" & Valor2Casas(t_vTotTrib) & "</vTotTrib>"
        s = s & "</ICMSTot>"
        s = s & "</total>"
                    
        
        .Close
    End With
    s = s & "</det>"
    itemNFe = s

End Function

Function GeraChaveNfE(Modelo As String, serie As String, nNf As String, _
    cNF As String, tipoEmis As tpEmis) As String
    Dim s As String
    Dim cDV As String
    'Dim cNf As String
    'UF - Código da UF do emitente do Documento Fiscal;
    'AAMM - Ano e Mês de emissão da NF-e;
    'CNPJ - CNPJ do emitente;
    'mod - Modelo do Documento Fiscal;
    'serie - Série do Documento Fiscal;
    'nNF - Número do Documento Fiscal;
    'tpEmis – forma de emissão da NF-e;
    'cNF - Código Numérico que compõe a Chave de Acesso;
    'cDV - Dígito Verificador da Chave de Acesso.
    
    
    s = strCodigoEstadoIBGE & Format(Date, "YYMM") & FormataCNPJNfe(strCNPJ_Emp) & _
        Modelo & Format(serie, "00#") & Format(nNf, "0########") & tipoEmis & cNF
    
    
    
    cDV = CalculaDigitoVerificadorModulo11(s)
    GeraChaveNfE = s & cDV
End Function

Function CalculaDigitoVerificadorModulo11(numero As String) As String

    Dim i As Integer
    Dim soma As Integer
    Dim dv As Integer
    Dim resto As Integer
    Dim pond As Integer
    pond = 0
    Dim item As Integer
    For i = Len(numero) To 1 Step -1
        If pond = 9 Or pond = 0 Then
            pond = 2
        Else
            pond = pond + 1
        End If
        item = (Int(Mid(numero, i, 1)))
        soma = soma + (item * pond)
'        MsgBox "Índice" & i & "=" & item & "*" & pond & "=" & (item * pond) & "| soma=" & soma
    Next
    
    
    
    
    resto = (soma Mod 11)
    If resto = 0 Or resto = 1 Then
        dv = 0
    Else
        dv = 11 - resto
    End If
    CalculaDigitoVerificadorModulo11 = dv

End Function
Function RetornaDVModulo11(numero As String) As String

    Dim i As Integer
    Dim soma As Integer
    Dim dv As Integer
    Dim resto As Integer
    For i = 1 To Len(numero)
    
        soma = soma + (Int(Mid(numero, i, 1))) * ((Len(numero) + 2) - i)
    Next
    
    
    
    resto = (soma Mod 11)
    If resto = 0 Or resto = 1 Then
        dv = 0
    Else
        dv = 11 - resto
    End If
    RetornaDVModulo11 = dv

End Function


Function GeraNovaXmlNFc_E(fin As TfinNfe, Destino As eIdDest, tipoConsumidor As indFinal, _
    tipo_presenca As indPres, tipo_pag As indPag, natOp As String, tipoMovimeto As tpNf) As String
    Dim Aleatorio As String
    Dim strNumNF As String
    Dim s As String
    Dim ide As tIde_Nfe
    Dim emi As emi_nfe
    ambiente = "2" 'homologação
    
    
    Aleatorio = AleatorioNFe
    
    strNumNF = UltimaNFe 'apenas as que já foram transmitidas
    
    ide.cNF = Aleatorio
    'ide.cNf = strNumNF
    ide.dhEmi = FormataDataNfe(Now)
    ide.fin = fin 'ver complementar e outras
    ide.idDest = Destino
    ide.indiFin = tipoConsumidor
    ide.indP = tipo_presenca
    ide.iPag = tipo_pag
    ide.Modelo = nfc_e
    ide.natOp = natOp
    ide.nNf = strNumNF
    ide.procEmi = "0" 'emitido pelo aplicativo proprio
    ide.serie = "1"
    
    ide.tipoAmb = ambiente
    ide.tipoImp = danf_nfc_e
    ide.tipoEmis = Normal
    ide.tpNf = tipoMovimeto
    ide.verProc = "2017"
    
    
    emi.cCRT = strCRT_emp
    emi.CNAE = strCNAE_Emp
    emi.CNPJ = Format(F_SoNumero(strCNPJ_Emp), "0#############")
    emi.ende.cep = F_SoNumero(strCEP_Emp)
    emi.ende.cMun = strCodigoMunicipioIBGE
    emi.ende.cPais = strCodigoPaisIBGE
    emi.ende.fone = F_SoNumero(strFone_Emp)
    emi.ende.nro = strNumero_Emp
    emi.ende.UF = strUF_Emp
    emi.ende.xBairro = strBairro_Emp
    emi.ende.xCpl = strComplemento_Emp
    emi.ende.xLgr = strLogradouro_Emp
    emi.ende.xMun = strCidade_Emp
    emi.ende.xPais = strPais
    emi.ie = F_SoNumero(strIE_Emp)
    emi.IEST = strIEST_Emp
    emi.im = strIM_Emp
    emi.xFant = strFantasia_Emp
    emi.xNome = strNome_Emp
    
    
    IniciaComponentesXMLNFE GeraChaveNfE("65", "1", strNumNF, ide.cNF, Normal)
    
    
    
    
    
    
    s = strXML_NFE_CabXML & CrLf
    
    s = s & IDE_NFE(ide) & CrLf

    s = s & EMIT_NFE(emi) & CrLf

    'ListaProduto.Add

    Set cProdutos = ListaProduto

    's = s & prod_nfe.getXmlProduto()
    




    GeraNovaXmlNFc_E = s



    




End Function

Function GeraXML_NFcE_Dados(id As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rAux As New ADODB.Recordset
    Dim chave As String
    Dim ide As IdeNFe.tIde_Nfe
    Dim emi As emi_nfe
    Dim ende As Ende_NFe.Ende_NFe
    Dim endedest As Ende_NFe.Ende_NFe
    Dim dest As dest
    Dim caminhoXML As String
    Dim s As String
    Dim cod_cli As String
    caminhoXML = App.Path & "\xml"
    strSQL = "select * from nfe where id=" & id
    
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    With rPesq
        chave = !chave
        IniciaComponentesXMLNFE chave
        ide = Get_IDE_Nfe(!cNF, !natOp, !indPag, !Modelo, !serie, !nNf, !dhEmi, !tpNf, !idDest, _
        !tpImp, !tpEmis, !cDV, !tpAmb, !finNFe, !indFinal, !indPag, !procEmi, !verProc)
    
        
        ende = Get_EnderEmi(strLogradouro_Emp, strNumero_Emp, strComplemento_Emp, strBairro_Emp, strCodigoMunicipioIBGE, _
        strCidade_Emp, strUF_Emp, F_SoNumero(strCEP_Emp), strCodigoPaisIBGE, strPais, F_SoNumero(strFone_Emp), strIEST_Emp, _
        "", F_SoNumero(strIE_Emp), "", strIM_Emp, strEmail_emp, strCNAE_Emp, CStr(strCRT_emp))
        
        emi = Get_Emi_NFe(F_SoNumero(strCNPJ_Emp), strNome_Emp, strFantasia_Emp, ende, F_SoNumero(strIE_Emp), _
        strIEST_Emp, strIM_Emp, strCNAE_Emp, strCRT_emp)
        emi.ende = ende
    
        strSQL = "select d.*,e.* from nfe_dest as d inner join nfe_enderdest as e on e.id_nfe=d.id_nfe"
        strSQL = strSQL & " where d.id_nfe=" & id
        
        .Close
    End With
    
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
        'dest.cMun = !cMun
            dest.CNPJ = !CNPJ
            dest.CPF = !CPF
            dest.cPais = !cPais
            dest.email = !email
            dest.fone = !fone
            dest.ie = !ie
            
            dest.UF = !UF
            dest.xNome = !xNome
            dest.ende.cep = !cep
            dest.ende.cMun = !cMun
            dest.ende.cPais = !cPais
            dest.ende.email = !email
            dest.ende.ie = !ie
            dest.ende.im = !im
            dest.ende.isuf = !isuf
            dest.ende.IndIEDest = !IndIEDest
            
        End If
    
        
    
    .Close
    End With
    
    
    
    
    
    
    s = strXML_NFE_CabXML & CrLf
    
    s = s & IDE_NFE(ide) & CrLf

    s = s & EMIT_NFE(emi) & CrLf
    s = s & DEST_NFE(dest) & CrLf
    
    s = s & itemNFe(id) & CrLf
    
    s = s & TranspNFe(id) & CrLf
    s = s & NFePag(id) & CrLf
    
    s = s & "</infNFe>"
    s = s & "</NFe>"


    GeraXML_NFcE_Dados = s

    

    If Dir(caminhoXML, vbDirectory) = "" Then
        FileSystem.MkDir caminhoXML
    End If
    caminhoXML = caminhoXML & "\NFe" & chave & ".xml"
    
   ' GravaArquivoExterno caminhoXML, s
    
    If Dir(caminhoXML, vbArchive) <> "" Then
        Kill caminhoXML
    End If
    
    GravaArquivoExternoUTF8 caminhoXML, s
    Shell "explorer " & caminhoXML


End Function





