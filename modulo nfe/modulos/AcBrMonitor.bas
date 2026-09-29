Attribute VB_Name = "AcBrMonitor"
Public CaminhoArquivoEntradaMonitor As String
Public CaminhoMonitor As String
Public CaminhoArquivoSaidaMonitor As String
Public CaminhoArquivosXMLRetornoNFe As String
Public CaminhoArquivosXMLRetornoNFCe As String
Public CaminhoArquivosXMLRetornoRaiz As String
Public CaminhoArquivosPDFRetornoNFe As String
Public CaminhoArquivosPDFRetornoNFCe As String
Public CaminhoArquivosPDFRetornoRaiz As String
Public CaminhoArquivosXMLInut As String
Public CaminhoXMLCopia As String




Sub KillArquivoSaidaMonitor()
    On Error Resume Next
    Kill CaminhoArquivoSaidaMonitor
End Sub
Sub KillArquivoEntradaMonitor()
    On Error Resume Next
    Kill CaminhoArquivoEntradaMonitor
End Sub


Function GeraNfeIni(id As String) As String
    Dim ado As New ADODB.Recordset
    Dim adoImposto As New ADODB.Recordset
    Dim strSQL As String
    Dim indIEst As String
    Dim iItem As Integer
    Dim sTipoIcms As String
    Dim finNFe As String
    Dim infCpl As String
    Dim rPis As New ADODB.Recordset
    Dim i As Integer
    On Error GoTo trata_Erro
    Dim iGtin As Integer
    
    P_Msg "Localizando a nota"
    
    strSQL = "select * from nfe where id=" & id
    abreObjetoAdo ado, strSQL
    P_Msg "localizei"
    
    If ado.RecordCount = 0 Then
        MsgBox "Não encontrado", vbExclamation
        ado.Close
        Exit Function
    End If
    If sVersaoNfe = "" Then
        sVersaoNfe = "4.00"
    End If


    Dim s As String
    
    
    s = "[infNFe]" & CrLf
    s = s & "versao = " & sVersaoNfe & CrLf
    
    s = s & "[Identificacao]" & CrLf
    
    s = s & "cNF=" & ado!cNF & CrLf
    
    s = s & "natOp=" & ado!natOp & CrLf
    
    s = s & "indPag = " & ado!indPag & CrLf
    
    s = s & "mod=" & ado!modelo & CrLf
    
    s = s & "serie=" & ado!serie & CrLf
    
    s = s & "nNF=" & ado!nNF & CrLf
    
    s = s & "dhEmi=" & Format(Now, "dd/mm/yyyy HH:MM:SS") & CrLf
    
    s = s & "dhSaiEnt=" & Format(Now, "dd/mm/yyyy HH:MM:SS") & CrLf
    
    s = s & "tpNF=" & ado!tpNf & CrLf
    
    s = s & "idDest=" & ado!idDest & CrLf
    
    s = s & "tpImp=" & ado!tpImp & CrLf
    
    s = s & "tpEmis=" & ado!tpEmis & CrLf
    
    s = s & "finNFe=" & ado!finNFe & CrLf
    
    s = s & "indFinal=" & ado!indFinal & CrLf
    
    s = s & "indPres=" & ado!indPres & CrLf
    
    s = s & "procEmi=" & ado!procEmi & CrLf
    
    s = s & "verProc=" & ado!verProc & CrLf
    
    s = s & "dhCont=" & FormataDataNfeParaNormal(ado!dhCont) & CrLf
    
    s = s & "xJust=" & ado!xJust & CrLf
    
    finNFe = ado!finNFe
    
    P_Msg "localizando nfe_ref "
    strSQL = "select * from nfe_ref where id=" & ado!id
    abreObjetoAdo adoImposto, strSQL
    P_Msg "localizei"
    With adoImposto
        Do While Not .EOF
            s = s & "[NFRef" & Format(.AbsolutePosition, "00#") & "]" & CrLf
            s = s & "Tipo=NFe" & CrLf
            s = s & "refNFe=" & adoImposto!refNfe & CrLf
            .MoveNext
        Loop
    End With
    
        
    
    
'    [NFRefXXX]
'    Tipo=
'    refNFe=
'    cUF=
'    AAMM=
'    CNPJ=
'    mod=
'    Serie=
'    nNF=
'    CNPJCPF=
'    IE=
'    refCTe=
'    ModECF=
'    nECF=
'    nCOO=
    
    P_Msg "localizando nfe_emit"
    
    strSQL = "select * from nfe_emit where id_nfe=" & id
    abreObjetoAdo ado, strSQL
    
    P_Msg "localizando localizei"
    
    s = s & "[emitente]" & CrLf
    
    s = s & "CNPJCPF=" & F_SoNumero(strCNPJ_Emp) & CrLf  'ado!CNPJ & crlf
    
    s = s & "xNome=" & ado!xNome & CrLf
    s = s & "xFant=" & ado!xFant & CrLf
    
    P_Msg "localizando endereço"
    strSQL = "select * from nfe_enderemit where id_nfe=" & id
    abreObjetoAdo ado, strSQL
    
    
    s = s & "IE=" & F_SoNumero(strIE_Emp) & CrLf 'ado!IE & crlf
    s = s & "IEST=" & ado!iEst & CrLf
    s = s & "IM=" & ado!IM & CrLf
    s = s & "CNAE=" & ado!CNAE & CrLf
    s = s & "CRT=" & ado!CRT & CrLf
    s = s & "xLgr=" & ado!xLgr & CrLf
    s = s & "nro=" & ado!nro & CrLf
    s = s & "xCpl=" & ado!xCpl & CrLf
    s = s & "xBairro=" & ado!xBairro & CrLf
    s = s & "cMun=" & ado!cMun & CrLf
    s = s & "xMun=" & ado!xMun & CrLf
    s = s & "UF=" & UCase(ado!UF) & CrLf
    s = s & "CEP=" & ado!cep & CrLf
    s = s & "cPais=" & ado!cPais & CrLf
    s = s & "xPais=" & ado!xPais & CrLf
    s = s & "Fone=" & ado!fone & CrLf
    's = s & "cUF=" & "52" & CrLf
    's = s & "cMunFG=" & ado!cMunFG & CrLf
    
    
    
    
'    [Avulsa]
'    CNPJ=
'    xOrgao=
'    matr=
'    xAgente=
'    fone=
'    UF=
'    nDAR=
'    dEmi=
'    vDAR=
'    repEmi=
'    dPag=
    Dim indFinal As String
    Dim cpf As String
    P_Msg "localizando destino"
    strSQL = "select * from nfe_dest where id_nfe=" & id
    abreObjetoAdo ado, strSQL
    
    
    s = s & "[Destinatario]" & CrLf
    
    s = s & "idEstrangeiro=" & ado!idEstrangeiro & CrLf
    s = s & "xNome=" & ado!xNome & CrLf
    cpf = Format(ado!cpf)
    If cpf <> "" Then
        indFinal = "1"
        If F_CPF(ado!cpf) Then
            
            s = s & "CNPJCPF=" & ado!cpf & CrLf
            
        End If
    ElseIf Not IsNull(ado!CNPJ) Then
        If F_CGC(ado!CNPJ) Then
            
                s = s & "CNPJCPF=" & ado!CNPJ & CrLf
            
        End If
    End If
    
    
    
    
    
    
    
    
    
    strSQL = "select * from nfe_enderdest where id_nfe=" & id
    abreObjetoAdo ado, strSQL
    
    If indFinal = "1" Then
        s = s & "indIEDest=9" & CrLf 'ado!IndIEDest & crlf
    Else
        s = s & "indIEDest=" & ado!IndIEDest & CrLf 'ado!IndIEDest & crlf
    End If
    s = s & "IE=" & ado!IE & CrLf
    s = s & "ISUF=" & ado!isuf & CrLf
    s = s & "Email=" & ado!email & CrLf
    s = s & "xLgr=" & ado!xLgr & CrLf
    s = s & "nro=" & ado!nro & CrLf
    s = s & "xCpl=" & ado!xCpl & CrLf
    s = s & "xBairro=" & ado!xBairro & CrLf
    s = s & "cMun=" & ado!cMun & CrLf
    s = s & "xMun=" & ado!xMun & CrLf
    s = s & "UF=" & UCase(ado!UF) & CrLf
    s = s & "CEP=" & ado!cep & CrLf
    s = s & "cPais=" & ado!cPais & CrLf
    s = s & "xPais=" & ado!xPais & CrLf
    s = s & "Fone=" & ado!fone & CrLf
    
'    [Retirada]
'    CNPJCPF=
'    xLgr=
'    nro=
'    xCpl=
'    xBairro=
'    cMun=
'    xMun=
'    UF=
'
    
    
    
'    [Entrega]
'    CNPJCPF=
'    xLgr=
'    nro=
'    xCpl=
'    xBairro=
'    cMun=
'    xMun=
'    UF=
    
    
    
    
'    [autXMLXXX]
'    CNPJCPF=
    
    P_Msg "localizando produtos... "
    
    strSQL = "select a.*,b.*, p.balanca_pro,comb.cProdAnp,comb.descAnp,comb.UFCons from nfe_item as a inner join nfe_item_imposto as b "
    strSQL = strSQL & " on a.id=b.id "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as p on  a.cprod=p.cod_pro "
    
    strSQL = strSQL & " left join nfe_item_comb as comb on a.id=comb.id_item "
    strSQL = strSQL & " where a.id_nfe=" & id
    strSQL = strSQL & " order by a.id"
    abreObjetoAdo ado, strSQL
    
    
    With ado
        Do While Not .EOF
        
        
        
            s = s & "[Produto" & Format(.AbsolutePosition, "00#") & "]" & CrLf
            s = s & "cProd=" & !cProd & CrLf
            iGtin = Len(!cEAN)
            
            If iGtin <> 8 And iGtin <> 12 And iGtin And iGtin <> 13 And iGtin <> 14 Or !balanca_pro = "S" Then
                s = s & "cEAN=SEM GTIN" & CrLf
                
            Else
                s = s & "cEAN=" & !cEAN & CrLf
            End If
            s = s & "xProd=" & !xProd & CrLf
            s = s & "NCM=" & !ncm & CrLf
            s = s & "EXTIPI=" & !EXTIPI & CrLf
            s = s & "CEST=" & !cEst & CrLf
            
            If !tipo_icms = "20" Or !tipo_icms = "ICMS20" Then
                s = s & "CBENEF=" & "GO821019" & CrLf
            End If
            
            s = s & "CFOP=" & !CFOP & CrLf
            s = s & "uCom=" & !uCom & CrLf
            s = s & "qCom=" & !qCom & CrLf
            s = s & "vUnCom=" & !vUnCom & CrLf
            s = s & "vProd=" & !vProd & CrLf
            iGtin = Len(!cEANTrib)
            If iGtin <> 8 And iGtin <> 12 And iGtin And iGtin <> 13 And iGtin <> 14 Or !balanca_pro = "S" Then
                s = s & "cEANTrib=SEM GTIN" & CrLf
            Else
                If !cEANTrib = "" Then
                    s = s & "cEANTrib=SEM GTIN" & CrLf
                Else
                    s = s & "cEANTrib=" & !cEANTrib & CrLf
                End If
                
                
                
                
            End If
            s = s & "uTrib=" & !uTrib & CrLf
            s = s & "qTrib=" & !qTrib & CrLf
            s = s & "vUnTrib=" & !vUnTrib & CrLf
            s = s & "vFrete=" & !vFrete & CrLf
            s = s & "vSeg=" & !vSeg & CrLf
            s = s & "vDesc=" & !vDesc & CrLf
            s = s & "vOutro=" & !vOutro & CrLf
            s = s & "indTot=" & !indTot & CrLf
            's = s & "xPed=" & !xPed & CrLf
            's = s & "nItemPed=" & CrLf
            's = s & "nFCI=" & CrLf
            's = s & "nRECOPI=" & CrLf
            's = s & "pDevol=" & CrLf
            's = s & "vIPIDevol=" & CrLf
            's = s & "vTotTrib=" & CrLf
            s = s & "infAdProd=" & TiraNull(!infAdProd) & CrLf
            
            
            
            If Not IsNull(!cProdAnp) Then
                s = s & "[Combustivel" & Format(.AbsolutePosition, "00#") & "]" & CrLf
                s = s & "cProdAnp=" & !cProdAnp & CrLf
                s = s & "descAnp=" & !descAnp & CrLf
                s = s & "UFCons=" & !UFCons & CrLf
                
                
            End If
            
            
            
            
            
            '[NVEXXXYYY]
            'NVE=
                      
            
            
'            [DIXXXYYY]
'            nDi=
'            dDi=
'            xLocDesemb=
'            UFDesemb=
'            dDesemb=
'            tpViaTransp=
'            vAFRMM=
'            tpIntermedio=
'            CNPJ=
'            UFTerceiro=
'            cExportador=
             Dim tabela As String
             
             tabela = BuscaTabelaICMS(F_SoNumero(!tipo_icms))
      
            
            
            If tabela <> "" Then
            
                strSQL = "select * from " & tabela & " where id=" & ado!id
            Else
            
            
                If !tipo_icms = "102" Or !tipo_icms = "ICMSSN102" Then
                    strSQL = "select * from nfe_item_imposto_icmssn102 where id=" & ado!id
                
                
                ElseIf !tipo_icms = "500" Then
                    strSQL = "select * from nfe_item_imposto_icmssn500 where id=" & ado!id
                Else
                    strSQL = "select * from nfe_item_imposto_" & !tipo_icms & "  where id=" & ado!id
                End If
            
            End If
            abreObjetoAdo adoImposto, strSQL
            
            
            
            
            If ado.RecordCount > 0 Then
                s = s & "[ICMS" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
                
                If adoImposto.RecordCount > 0 Then
                    s = s & "orig=" & adoImposto!orig & CrLf
                
                    If !tipo_icms = "900" Or !tipo_icms = "ICMSSN900" Then
                        If finNFe <> "4" Then
                            s = s & "CST=00" & CrLf
                        Else
                            s = s & "CSOSN=" & adoImposto!csosn & CrLf
                        End If
                    ElseIf !tipo_icms = "ICMS00" Or !tipo_icms = "ICMS20" Or !tipo_icms = "ICMS60" Or !tipo_icms = "00" Or !tipo_icms = "20" Or !tipo_icms = "60" Then
                        s = s & "CST=" & F_SoNumero(!tipo_icms) & CrLf
                    Else
                        s = s & "CSOSN=" & adoImposto!csosn & CrLf
                    End If
                    
                End If
                s = s & "modBC=" & BuscaValorSeExistirCampo(adoImposto, "modBC") & CrLf
                s = s & "pRedBC=" & BuscaValorSeExistirCampo(adoImposto, "pRedBC") & CrLf
                s = s & "vBC=" & BuscaValorSeExistirCampo(adoImposto, "vbc") & CrLf
                s = s & "pICMS=" & BuscaValorSeExistirCampo(adoImposto, "pICMS") & CrLf
                s = s & "vICMS=" & BuscaValorSeExistirCampo(adoImposto, "vICMS") & CrLf
                s = s & "modBCST=0" & CrLf
                s = s & "pMVAST=" & CrLf
                s = s & "pRedBCST=" & CrLf
                s = s & "vBCST=0" & CrLf
                s = s & "pICMSST=0" & CrLf
                s = s & "vICMSST=0" & CrLf
                s = s & "UFST=" & CrLf
                s = s & "pBCOp=" & CrLf
                s = s & "vBCSTRet=" & CrLf
                s = s & "vICMSSTRet=" & CrLf
                s = s & "motDesICMS=" & CrLf
                s = s & "pCredSN=" & BuscaValorSeExistirCampo(adoImposto, "PCredSN") & CrLf
                s = s & "vCredICMSSN=" & BuscaValorSeExistirCampo(adoImposto, "vCredICMSSN") & CrLf
                s = s & "vBCSTDest=" & CrLf
                s = s & "vICMSSTDest=" & CrLf
                s = s & "vICMSDeson=" & CrLf
                s = s & "vICMSOp=" & CrLf
                s = s & "pDif=" & CrLf
                s = s & "vICMSDif=" & CrLf
              
            
            End If
            
            If !tipo_ipi = "IPITrib" Then
                strSQL = "select a.*,b.* from nfe_item_imposto_ipi as a left join nfe_item_imposto_ipitrib as b "
                strSQL = strSQL & " on a.id=b.id "
                strSQL = strSQL & " Where a.id = " & ado!id
                abreObjetoAdo rPis, strSQL
                If rPis.RecordCount > 0 Then
                    s = s & "[IPI" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
                    s = s & "CST=" & TiraNull(rPis!cst) & CrLf
                    s = s & "clEnq=999" & TiraNull(rPis!clEnq) & CrLf
                    s = s & "CNPJProd=" & TiraNull(rPis!CNPJProd) & CrLf
                    s = s & "cSelo=" & TiraNull(rPis!cSelo) & CrLf
                    s = s & "qSelo=" & TiraNull(rPis!qSelo) & CrLf
                    s = s & "cEnq=" & TiraNull(rPis!cEnq) & CrLf
                    's = s & "vBC=" & TiraNull(!vBC) & CrLf
                    's = s & "qUnid=" & TiraNull(!qUnid) & CrLf
                    's = s & "vUnid=" & TiraNull(!vUnid) & CrLf
                    's = s & "pIPI=" & TiraNull(!pIPI) & CrLf
                    's = s & "vIPI=" & TiraNull(!vIPI) & CrLf
                    
                End If
            Else
                strSQL = "select a.*,b.* from nfe_item_imposto_ipi as a left join nfe_item_imposto_ipint as b "
                strSQL = strSQL & " on a.id=b.id "
                strSQL = strSQL & " Where a.id = " & ado!id
                abreObjetoAdo rPis, strSQL
                If rPis.RecordCount > 0 Then
                    s = s & "[IPI" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
                    s = s & "CST=" & TiraNull(rPis!cst) & CrLf
                    s = s & "clEnq=" & TiraNull(rPis!clEnq) & CrLf
                    s = s & "CNPJProd=" & TiraNull(rPis!CNPJProd) & CrLf
                    s = s & "cSelo=" & TiraNull(rPis!cSelo) & CrLf
                    s = s & "qSelo=" & TiraNull(rPis!qSelo) & CrLf
                    s = s & "cEnq=" & TiraNull(rPis!cEnq) & CrLf
               End If
                    
            
            End If
'            [IIXXX]
'
'            vBC=
'
'            vDespAdu=
'
'            vII=
'
'            vIOF=
'
'
            strSQL = "select * from nfe_item_imposto_" & !tipo_pis & " where id=" & ado!id
            abreObjetoAdo rPis, strSQL
            
            If rPis.RecordCount > 0 Then
                s = s & "[PIS" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
                
                For i = 2 To rPis.Fields.Count - 1
                    
                    s = s & rPis.Fields(i).Name & "=" & rPis.Fields(i).Value & CrLf
                    
                Next
            Else
            
                s = s & "[PIS" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
                s = s & "CST=" & strCST_PIS & CrLf
                s = s & "vBC=0" & CrLf
                s = s & "pPIS=0" & CrLf
                s = s & "qBCProd=" & CrLf
                s = s & "vAliqProd=" & CrLf
                s = s & "vPIS=0" & CrLf
            End If
            
            
            
'            s = s & "[PISST" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
'            s = s & "vBC=" & CrLf
'            s = s & "pPis=" & CrLf
'            s = s & "qBCProd=" & CrLf
'            s = s & "vAliqProd=" & CrLf
'            s = s & "vPIS=" & CrLf
            
            
            
            strSQL = "select * from nfe_item_imposto_" & !tipo_cofins & " where id=" & ado!id
            abreObjetoAdo rPis, strSQL
            
            If rPis.RecordCount > 0 Then
                s = s & "[COFINS" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
                
                For i = 2 To rPis.Fields.Count - 1
                    
                    s = s & rPis.Fields(i).Name & "=" & rPis.Fields(i).Value & CrLf
                    
                Next
            Else
                s = s & "[COFINS" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
                s = s & "CST=" & strCST_COFINS & CrLf
                s = s & "vBC=0" & CrLf
                s = s & "pCOFINS=0" & CrLf
                s = s & "qBCProd=" & CrLf
                s = s & "vAliqProd=" & CrLf
                s = s & "vCOFINS=0" & CrLf
            End If
'            s = s & "[COFINSSTXXX]" & CrLf
'            s = s & "vBC=" & CrLf
'            s = s & "pCOFINS=" & CrLf
'            s = s & "qBCProd=" & CrLf
'            s = s & "vAliqProd=" & CrLf
'            s = s & "vCOFINS=" & CrLf
            
            
            .MoveNext
        Loop
        .Close
    End With
    
    strSQL = "select * from nfe_icmstot as i left join nfe_transp as p  on i.id_nfe=p.id_nfe where p.id_nfe=" & id
    abreObjetoAdo ado, strSQL
    
    s = s & "[Total]" & CrLf
    s = s & "vBC=" & ado!vBC & CrLf
    s = s & "vICMS=" & ado!vICMS & CrLf
    s = s & "vICMSDeson=" & ado!vICMSDeson & CrLf
    s = s & "vBCST=" & ado!vBCST & CrLf
    s = s & "vST=" & ado!vST & CrLf
    s = s & "vProd=" & ado!vProd & CrLf
    s = s & "vFrete=" & ado!vFrete & CrLf
    s = s & "vSeg=" & ado!vSeg & CrLf
    s = s & "vDesc=" & ado!vDesc & CrLf
    s = s & "vII=" & ado!vII & CrLf
    s = s & "vIPI=" & ado!vIPI & CrLf
    s = s & "vPIS=" & ado!vPIS & CrLf
    s = s & "vCOFINS=" & ado!vCOFINS & CrLf
    s = s & "vOutro=" & ado!vOutro & CrLf
    If ado!vProd <> ado!vNF Then
        If ado!vDesc > 0 Then
            
            s = s & "vNF=" & (ado!vProd - ado!vDesc) & CrLf
        Else
            s = s & "vNF=" & (ado!vProd) & CrLf
        End If
    Else
        s = s & "vNF=" & ado!vNF & CrLf
    End If
    's = s & "vTotTrib=" & ado!vTotTrib & CrLf
    
    
    '    [retTrib]
    '    vRetPIS=
    '    vRetCOFINS=
    '    vRetCSLL=
    '    vBCIRRF=
    '    vIRRF=
    '    vBCRetPrev=
    '    vRetPrev=
        
        
    
    
    s = s & "[Transportador]" & CrLf
    s = s & "modFrete = " & ado!modFrete & CrLf
    Dim adoFrete As New ADODB.Recordset
    If ado!modFrete <> "9" Then
    
        strSQL = "select * from nfe_transp as n left join nfe_transp_reboque as r on n.id_nfe=r.id_nfe "
        strSQL = strSQL & " left join nfe_transp_rettransp as rt on n.id_nfe=tr.id_nfe "
        strSQL = strSQL & " left join nfe_transp_transportadora as t on n.id_nfe=t.id_nfe "
        strSQL = strSQL & " left join nfe_transp_veictransp as vt on n.id_nfe=vt.id_nfe "
        strSQL = strSQL & " left join nfe_transp_vol as v on n.id_nfe=v.id_nfe "
        strSQL = strSQL & " where n.id_nfe=" & id
        abreObjetoAdo adoFrete, strSQL
        
        With adoFrete
            CNPJCPF = IIf(TiraNull(!CNPJ) = "", !cpf, !CNPJ)
            xNome = !xNome
            IE = !IE
            xEnder = !xEnder
            xMun = !xMun
            UF = !UF
            vServ = !vServ
            vBCRet = !vBCRet
            pICMSRet = !pICMSRet
            vICMSRet = !vICMSRet
            CFOP = !CFOP
            cMunFG = !cMunFG
            placa = !placa
            ufplaca = !ufplaca
            RNTC = !RNTC
            vagao = !vaga
            balsa = !balsa
            
    End With
        
        
        
     '   [Reboque001]
     '   placa = !placa
     '   UF=
      '  RNTC=
    
        
        
        
       ' [VolumeXXX]
       ' qVol=
       ' esp=
       ' Marca=
       ' nVol=
       ' pesoL=
       ' pesoB=
        
        
        
    '    [LacreXXXYYY]
    '    nLacre=
        
        
    End If
    
    
            
    '    [fatura]
    '    nFat=
    '    vOrig=
    '    vDesc=
    '    vLiq=
        
    '    [DuplicataXXX]
    '    nDup=
    '    dVenc=
    '    vDup=

    
    
    strSQL = "select * from nfe_pag where id_nfe=" & id
    abreObjetoAdo ado, strSQL
    With ado
        Do While Not .EOF
            s = s & "[PAG" & Format(.AbsolutePosition, "00#") & "]" & CrLf
            s = s & "tpag=" & !tPag & CrLf
            s = s & "vPag=" & !vPag & CrLf
            s = s & "CNPJ=" & !CNPJ & CrLf
            s = s & "tBand=" & !tBand & CrLf
            s = s & "cAut=" & !cAut & CrLf
            strSQL = TiraNull(!tpIntegra)
            If strSQL <> "" And strSQL <> "0" Then
                s = s & "tpIntegra=" & !tpIntegra & CrLf
            End If
            
            
            .MoveNext
        Loop
    End With
    
    
    strSQL = "select * from nfe_infadic where id_nfe=" & id
    abreObjetoAdo ado, strSQL
    If ado.RecordCount > 0 Then
        If Not IsNull(ado!infCpl) Then
            infCpl = ado!infCpl
        End If
        
    End If
    
    
    s = s & "[DadosAdicionais]" & CrLf
    
    
    
    
    
    s = s & "infAdFisco=" & CrLf
    
    s = s & "infCpl=" & infCpl & CrLf
    
    
    
    
'    s = s & "[InfAdicXXX]" & CrLf
'    s = s & "xCampo="
'    s = s & "xTexto="
'
'    s = s & "[ObsFiscoXXX]"
'    xCampo=
'    xTexto=
'
    
    
    
'    [procRefXXX]
'    nProc=
'    indProc=
'
    
    
    
'    [Exporta]
'    UFSaidaPais=
'    xLocExporta=
'    xLocDespacho=
    
    
    
    
'    [Compra]
'    xNEmp=
'    xPed=
'    xCont=
    
    
    
    
    
'
'
'    [deducXXX]
'
'    xDed=
'
'    vDed=

GeraNfeIni = s
Exit Function
trata_Erro:
MsgBox Err.Description, vbExclamation
End Function


Sub CriarEnviarNFe(cIniNfE As String, nLote As String)
    Dim Comando As String
    If CaminhoArquivoEntradaMonitor = "" Then
        CaminhoArquivoEntradaMonitor = "c:\acbrmonitorplus\ent.txt"
    End If
    
    Comando = "NFe.CriarEnviarNFe(" & aspas & cIniNfE & aspas & ", " & nLote & ",0,1)"
    If GravaArquivoExternoUTF8(CaminhoArquivoEntradaMonitor, Comando) Then
    End If

End Sub
Function PedidoInut(cCNPJ As String, cJustificativa As String, nAno As String, nModelo As String, nSerie As String, nNumInicial, nNumFinal As String) As String
    Dim Comando As String
    Dim retorno As String
    If CaminhoArquivoEntradaMonitor = "" Then
        CaminhoArquivoEntradaMonitor = "c:\acbrmonitorplus\ent.txt"
    End If
    
    Comando = "NFE.InutilizarNFe(" & cCNPJ & "," & cJustificativa & "," & nAno & ", "
    Comando = Comando & nModelo & "," & nSerie & "," & nNumInicial & "," & nNumFinal & ")"
    If GravaArquivoExternoUTF8(CaminhoArquivoEntradaMonitor, Comando) Then
    End If
    
retorno:
    If AguardarArquivoRetorno(500) = True Then
        retorno = LeArquivoExternoUTF8(CaminhoArquivoSaidaMonitor)
        KillArquivoSaidaMonitor
    End If
    
    PedidoInut = retorno
End Function

Sub setarCertificado(cod_emp As String)
    Dim Comando As String
    Dim retorno As String
    If CaminhoArquivoEntradaMonitor = "" Then
        CaminhoArquivoEntradaMonitor = "c:\acbrmonitorplus\ent.txt"
    End If
    Dim emp As sped_empresa.sped_empresa
    emp = sped_empresa.BuscaSpedEmpresa(cod_emp)
    If emp.codigo = "" Then
        Exit Sub
    End If
       
    
    Comando = "nfe.setCertificado(" & aspas & emp.caminho_certificado & aspas & "," & aspas & emp.senha_certificado & aspas & ")" & CrLf
    'If GravaArquivoExternoUTF8(CaminhoArquivoEntradaMonitor, Comando) Then
    '    Sleep (100)
    'End If
    
    retorno = GravaComandoEntradaAguardaOk(Comando)
    
    
    
    'Comando = "nfe.setToken(" & aspas & emp.token & aspas & ")" & CrLf
    Comando = "nfe.setCSC(" & aspas & emp.token & aspas & "," & aspas & emp.idCSC & aspas & ")" & CrLf
    
    retorno = GravaComandoEntradaAguardaOk(Comando)
    
    Comando = "nfe.lerini() "
    retorno = GravaComandoEntradaAguardaOk(Comando)
    'Comando = "nfe.setIdToken(" & aspas & emp.idCSC & aspas & ")" & CrLf
    'Comando = "nfe.setIdCSC(" & aspas & emp.idCSC & aspas & ")" & CrLf
    

    'If GravaArquivoExternoUTF8(CaminhoArquivoEntradaMonitor, Comando) Then
    '    Sleep (100)
    'End If
   ' Comando = "nfe.lerini() "
   ' If GravaArquivoExternoUTF8(CaminhoArquivoEntradaMonitor, Comando) Then
    
    
   ' End If


End Sub



Function TransmitirNFcE(id As String, lote As String, Optional bImprimirDanfe As Boolean = True) As String
    Dim strSQL As String
    Dim cIni As String
    Dim Chave As String
    Dim AnoMes As String
    Dim retorno As String
    Dim conteudo As String
    Dim caminhoXML As String
    Dim Inf As infProt
    Dim i As Integer
    Dim iImpressora As Integer
    Dim modelo As String
    Dim venda As String
    Dim comprovante_cartao As String
    Dim verifica_ncm As String
    Dim caminho As String
    Dim cod_emp As String
    Dim rPesq As New ADODB.Recordset
    'Dim conteudo As String
     P_Status "preparando"
    Chave = NFe.BuscaChaveNFe(id)
    AnoMes = NFe.BuscaAnoMesNFe(id)
    iImpressora = BuscaImpressoraPadrao
    On Error GoTo trata_Erro
    P_Status "Transmitindo nfe"
    If Chave = "" Then
        MsgBox "Chave não encontrada para a nf " & id, vbExclamation
        
        Exit Function
    End If
    
       
    P_Status "buscando venda"
    venda = F_BuscaUmCampo("num_ven", "nfe", "id=" & id)
    cod_emp = F_BuscaUmCampo("cod_emp", TAB_VENDAS, "num_ven='" & venda & "'")
    
    If cod_emp = "" Then
        cod_emp = "01"
    End If
    P_Msg "preparando"
    P_Status "venda encontrada"
    If ParametroEstrutura.BuscaParametroEstrutura("setarCertificado").valor = "S" Then
         P_Status "setando certificado"
        setarCertificado cod_emp
    End If
    verifica_ncm = VerificaNCMVenda(venda)
    If verifica_ncm <> "" Then
        If MsgBox("Problemas com ncm: " & CrLf & verifica_ncm & " Deseja continuar?", vbYesNo + vbCritical) = vbNo Then
            Set Printer = Printers(iImpressora)
            P_Status ""
            Exit Function
        End If
    End If
    
    
    If venda <> "" Then
        comprovante_cartao = F_BuscaUmCampo("mensagem", "tab_venfin", "num_Ven='" & venda & "'")
    End If
    
    
    
    caminho = RetornoCaminhoNFeXML(Chave, AnoMes)
    retorno = RetornoConteudoNFeXML(Chave, AnoMes)
    If retorno <> "" Then
        conteudo = RetornoConteudoXML(retorno)
        
        Inf = RetornaInfProt(conteudo)
        
         P_Status "copiando xml " & caminho & " -> " & CaminhoXMLCopia
        CopiaXML caminho, CaminhoXMLCopia
        
        If Inf.cStat = "100" Then
            MsgBox "Essa nota já foi gerada com sucesso. Não é possível retransmitir!", vbExclamation
            Set Printer = Printers(iImpressora)
            P_Status ""
            Exit Function
        End If
        
        
                
        
        
    End If
    
    
    
    
    
    
    If lote = "" Then
        lote = GeraLoteNFe(id)
    End If
    
    
    
    cIni = CaminhoMonitor & "NFe" & id & ".ini"
    
    P_Msg "gerando"
    strSQL = AcBrMonitor.GeraNfeIni(id)
    P_Msg ""
    '
    P_Msg "buscando impressora"
    iImpressora = BuscaImpressoraPadrao
    P_Msg "impressora encontrada"
inicio:
    On Error Resume Next
    P_Msg "Gerando arquivo para Conexão NFe"
    
    Kill CaminhoArquivoSaidaMonitor
    If Not GravaArquivoExternoUTF8(cIni, strSQL) Then
        MsgBox "erro na gravação do arquivo", vbExclamation
    End If
    
    iImpressora = BuscaImpressoraPadrao
    Set Printer = Printers(ImpressoraPadraoCupomNFCe.Indice)
    
    P_Msg "criar e enviar nfe - monitor"
    AcBrMonitor.CriarEnviarNFe cIni, lote
    
    P_Msg "Arquivo gerado... aguardando retorno"
    
    If AguardarArquivoRetorno(500) = True Then
        P_Msg "retornou..."
        modelo = BuscaModeloPeloID(id)
        
        TransmitirNFcE = LeArquivoExternoUTF8(CaminhoArquivoSaidaMonitor)
        
        conteudo = XmlRetornoNFe(Chave)
        caminho = RetornoCaminhoNFeXML(Chave)
        Inf = RetornaInfProt(conteudo)
        If Inf.cStat = "100" Then
            AlteraStatusNfe id, "2"
            CopiaXML caminho, CaminhoXMLCopia
            P_Msg "alterando venda e saídas"
            
            AlteraStatusVen venda, "2"
            P_BaixaSaidaNFe id
            P_Msg "enviar email cliente"
            EnviaEmailNFe BuscaEmailClientePeloId(id), caminho
            P_Msg "enviou..."
            
            
            
            If modelo = "65" Then
                If isDesignTime Or bUsaVendaSupermercado = False Then
                
                
                    ImprimirXMLNFeCUPOM conteudo, frmVisualiza.Pic, True, comprovante_cartao
             
                    frmVisualiza.Show 1
                    If frmVisualiza.tag <> "" Then
                        ImprimirXMLNFeCUPOM conteudo, Printer, True, comprovante_cartao
                     
                    End If
                    Unload frmVisualiza
                Else
                    ImprimirXMLNFeCUPOM conteudo, Printer, True, comprovante_cartao
                End If
            
                
                
            
            End If
            
            
        ElseIf Inf.cStat = "" Then
        
            
            retorno = RetornoConteudoNFeXML(Chave, AnoMes)
            If retorno <> "" Then
                conteudo = RetornoConteudoXML(retorno)
    
                Inf = RetornaInfProt(conteudo)
    
                If Inf.cStat = "100" Then
                    MsgBox "Essa nota já foi gerada com sucesso. Não é possível retransmitir!", vbExclamation
                    Set Printer = Printers(iImpressora)
                    P_Status ""
                    Exit Function
                End If
    
    
        End If
            
            
            
            modelo = BuscaModeloPeloID(id)
        
            If modelo = "65" Then
                If MsgBox(TransmitirNFcE & Chr(13) & "Deseja transmitir novamente?", vbYesNo + vbQuestion) = vbYes Then
                    GoTo inicio
                End If
                MsgBox "Verique o status da NFe", vbInformation
            Else
                P_Status ""
                MsgBox TransmitirNFcE
            End If
        
        Else
            P_Status ""
            MsgBox TransmitirNFcE
        End If
        
        
        
        
        Kill CaminhoArquivoSaidaMonitor
        
    End If
    AtualizarRetorno id, False
     P_Status ""
     P_Msg
     Set Printer = Printers(iImpressora)

Exit Function

trata_Erro:
MsgBox Err.Description, vbCritical
Set Printer = Printers(iImpressora)

End Function

Function AguardarXMLRetorno(caminho As String, tempo As Integer) As String
    Dim encontrou As Boolean
    Dim agora As String
    agora = Now
    Dim i As infProt
    
    
    
    Do While encontrou = False
        Sleep tempo
        If Dir(caminho, vbArchive) <> "" Then
            AguardarXMLRetorno = LeArquivoExternoUTF8(caminho)
             
            
            Exit Do
            
        End If
        If DateDiff("s", agora, Now) > 60 Then
            MsgBox "Arquivo XML  não encontrado. Tente novamente", vbExclamation
            Exit Do
            Exit Function
        End If
    Loop
End Function

Function AguardarArquivoRetorno(tempo As Integer) As Boolean
    Dim encontrou As Boolean
    Dim agora As String
    agora = Now
    P_Status "Aguardando arquivo de retorno..."
    
    If CaminhoArquivoSaidaMonitor = "" Then
        CaminhoArquivoSaidaMonitor = "c:\ACBrMonitorPLUS\sai.txt"
    End If
    Do While encontrou = False
        Sleep tempo
        If Dir(CaminhoArquivoSaidaMonitor, vbArchive) <> "" Then
            'txtStatusServ.Text = LeArquivoExternoUTF8(CaminhoArquivoSaidaMonitor)
            encontrou = True
            AguardarArquivoRetorno = True
            P_Status
            Exit Do
        
        End If
        If DateDiff("s", agora, Now) > 60 Then
            P_Status
            MsgBox "Arquivo de saída não encontrado. Tente novamente", vbExclamation
            
            Exit Do
            Exit Function
        End If
    Loop
    P_Status
End Function

Function GravaComandoEntradaAguardaOk(Comando) As String
    Dim retorno As String
    On Error GoTo trata_Erro
    If GravaArquivoExternoUTF8(CaminhoArquivoEntradaMonitor, Comando) Then
        
    End If
    retorno = AguardarArquivoRetornoString(100)
    Kill CaminhoArquivoEntradaMonitor
    GravaComandoEntradaAguardaOk = retorno
trata_Erro:
End Function

Function AguardarArquivoRetornoString(tempo As Integer) As String
    Dim encontrou As Boolean
    Dim agora As String
    Dim retorno
    agora = Now
    P_Status "Aguardando arquivo de retorno..."
    
    If CaminhoArquivoSaidaMonitor = "" Then
        CaminhoArquivoSaidaMonitor = "c:\ACBrMonitorPLUS\sai.txt"
    End If
    Do While encontrou = False
        Sleep tempo
        If Dir(CaminhoArquivoSaidaMonitor, vbArchive) <> "" Then
            'txtStatusServ.Text = LeArquivoExternoUTF8(CaminhoArquivoSaidaMonitor)
            retorno = LeArquivoExternoUTF8(CaminhoArquivoSaidaMonitor)
            encontrou = True
            AguardarArquivoRetornoString = retorno
            P_Status
            Exit Do
        
        End If
        If DateDiff("s", agora, Now) > 60 Then
            P_Status
            MsgBox "Arquivo de saída não encontrado. Tente novamente", vbExclamation
            
            Exit Do
            Exit Function
        End If
    Loop
    P_Status
End Function

Sub ImprimirDanfe(caminho As String, Optional protocolo As String = "")
    Dim Comando As String
    If protocolo = "" Then
    
        Comando = "NFe.ImprimirDANFE(" & caminho & ")"    '"NFe.ImprimirDANFEPDF(cArqXML,[cProtocolo],[cMarcaDaqgua],[bViaConsumidor],[bSimplificado])"
    Else
        Comando = "NFe.ImprimirDANFE(" & caminho & "," & protocolo & ")"
    End If
    GravaArquivoExternoUTF8 CaminhoArquivoEntradaMonitor, Comando

End Sub
Sub ImprimirDanfePDF(caminho As String, protocolo As String)
    Dim Comando As String
    If protocolo = "" Then
    
        Comando = "NFe.ImprimirDANFEPDF(" & caminho & ")"    '"NFe.ImprimirDANFEPDF(cArqXML,[cProtocolo],[cMarcaDaqgua],[bViaConsumidor],[bSimplificado])"
    Else
        Comando = "NFe.ImprimirDANFEPDF(" & caminho & "," & protocolo & ")"
    End If
    GravaArquivoExternoUTF8 CaminhoArquivoEntradaMonitor, Comando

End Sub

Function CancelarNFe(Chave As String, justificativa As String, cCNPJ As String, nEvento As String) As String
    Dim Comando As String
    setarCertificado strCod_Emp
    Comando = "NFe.CancelarNFe(" & aspas & Chave & aspas & "," & aspas & justificativa & aspas & IIf(cCNPJ <> "", "," & cCNPJ, "") & IIf(nEvento <> "", "," & nEvento, "") & ")"
    GravaArquivoEntradaMonitor Comando
End Function
Sub GravaArquivoEntradaMonitor(Comando As String)
    On Error Resume Next
    Kill CaminhoArquivoEntradaMonitor
    On Error GoTo trata_Erro
    GravaArquivoExternoUTF8 CaminhoArquivoEntradaMonitor, Comando
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub
Function LeArquivoSaidaMonitor(Optional Excluir As Boolean = True) As String

    LeArquivoSaidaMonitor = LeArquivoExternoUTF8(CaminhoArquivoSaidaMonitor)

    If LeArquivoSaidaMonitor <> "" Then
        If Excluir = True Then
            Kill CaminhoArquivoSaidaMonitor
        End If
    End If

End Function
Public Sub IniciaAcbr()
    Dim pasta As String
    CaminhoArquivoEntradaMonitor = "C:\ACBrMonitorPLUS\ent.txt"
    CaminhoArquivoSaidaMonitor = "C:\ACBrMonitorPLUS\sai.txt"
    CaminhoMonitor = "C:\ACBrMonitorPLUS\"
    CaminhoArquivosXMLRetornoNFCe = "C:\ACBrMonitorPLUS\Arqs\" & strCNPJ_Emp & "\NFCe\" & Format(Date, "yyyymm") & "\NFCe\"
    CaminhoArquivosXMLRetornoNFe = "C:\ACBrMonitorPLUS\Arqs\" & strCNPJ_Emp & "\NFe\" & Format(Date, "yyyymm") & "\NFe\"
    CaminhoArquivosXMLRetornoRaiz = "C:\ACBrMonitorPLUS\Arqs\" & strCNPJ_Emp & "\"
    
    CaminhoArquivosPDFRetornoNFCe = "C:\ACBrMonitorPLUS\PDF\" & strCNPJ_Emp & "\NFCe\" & Format(Date, "yyyymm") & "\NFCe\"
    CaminhoArquivosPDFRetornoNFe = "C:\ACBrMonitorPLUS\PDF\"
    CaminhoArquivosPDFRetornoRaiz = "C:\ACBrMonitorPLUS\PDF\" & strCNPJ_Emp & "\"
    
    
    
    
    
    CaminhoXMLCopia = BuscaCaminhoXMLCopia
    
    
    
End Sub
Function BuscaCaminhoXMLCopia(Optional mes As String = "") As String
    Dim strSQL As String
    Dim pasta As String
    On Error GoTo trata_Erro
    If mes = "" Then
        pasta = Format(Date, "yyyymm")
    Else
        pasta = mes
    End If
    caminho = F_BuscaUmCampo("CaminhoCopiaXml", TAB_PARAMETROS, "")
    caminho = caminho & "\" & pasta
    
    If Dir(caminho, vbDirectory) = "" Then
        MkDir caminho
    End If
    BuscaCaminhoXMLCopia = caminho
    Exit Function
trata_Erro:
    caminho = App.path
End Function
Function XmlRetornoNFe(Chave As String, Optional tipoNFe As modelo_Nfe = 65, Optional mes As String = "") As String
    'mes em branco, procura no mÊs atual
    Dim caminho As String
    Dim pastaNF As String
    Dim modelo As String
    
    modelo = BuscaModeloPelaChave(Chave)
    
    
    
    If modelo = "65" Then
        pasta = "NFCe\"
    Else
        pasta = "NFe\"
    End If
    
    
    
    If mes = "" Then
        If modelo = "55" Then
            caminho = CaminhoArquivosXMLRetornoNFe & Chave & "-nfe.xml"
        Else
            caminho = CaminhoArquivosXMLRetornoNFCe & Chave & "-nfe.xml"
        End If
    Else
        caminho = CaminhoArquivosXMLRetornoRaiz & pasta & mes & "\" & pasta & Chave & "-nfe.xml"
    End If
    
    If Dir(caminho) <> "" Then
        XmlRetornoNFe = LeArquivoExternoUTF8(caminho)
    End If

End Function

Function PDFRetornoNFe(Chave As String, Optional tipoNFe As modelo_Nfe = 65, Optional mes As String = "", _
    Optional AbrirPDF As Boolean = True, Optional BImprimirPDF As Boolean = False) As String
    'mes em branco, procura no mÊs atual
    Dim caminho As String
    Dim pastaNF
    Dim modelo As String
    modelo = BuscaModeloPelaChave(Chave)
    
    
    If modelo = "65" Then
        pasta = "NFCe\"
    Else
        pasta = "NFe\"
    End If
    
    
    If modelo = "65" Then
        If mes = "" Then
            If tipoNFe = NFe_55 Then
                caminho = CaminhoArquivosPDFRetornoNFe & Chave & "-nfe.PDF"
            Else
                caminho = CaminhoArquivosPDFRetornoNFCe & Chave & "-nfe.PDF"
            End If
        Else
            caminho = CaminhoArquivosPDFRetornoRaiz & pasta & mes & "\" & pasta & Chave & "-nfe.PDF"
        End If
    Else
        caminho = CaminhoMonitor & "\pdf\" & Chave & "-nfe.PDF"
    End If
    
    If Dir(caminho) <> "" Then
         'XmlRetornoNFe = LeArquivoExternoUTF8(caminho)
         If AbrirPDF = True Then
            shell "explorer " & caminho
         Else
            If BImprimirPDF = True Then
                ImprimirPDF caminho
            End If
            
         End If
         PDFRetornoNFe = caminho
         
    End If
    

    
    
    
End Function


Function RetornoConteudoNFeXMLPelaChave(Chave As String)
    
End Function

Function RetornoCaminhoNFePeloID(id As String, Optional AnoMes As String = "", Optional PastaTipoNFe As String = "NFCe") As String
    Dim Chave As String
    Chave = BuscaChaveNFe(id)
    RetornoCaminhoNFePeloID = RetornoConteudoNFeXML(Chave, AnoMes)
    
End Function
Function CopyXmlNFe(Chave As String) As String
    Dim s As String
    Dim r As String
    On Error GoTo trata_Erro
    s = RetornoCaminhoNFeXML(Chave)
    If s <> "" Then
        If Dir(App.path & "\xml", vbDirectory) = "" Then
            MkDir App.path & "\xml"
        End If
        r = App.path & "\xml\" & RetornaNomeArquivoResumido(s)
        FileCopy s, r
        CopyXmlNFe = r
    End If
    
    Exit Function
trata_Erro:
End Function
Function RetornoConteudoNFeXML(Chave As String, Optional AnoMes As String = "") As String
    Dim caminho As String
    
    
    
    caminho = RetornoCaminhoNFeXML(Chave, AnoMes)
    
    If Dir(caminho) <> "" Then
        RetornoConteudoNFeXML = caminho
    End If

End Function

Function RetornoCaminhoNFeXML(Chave As String, Optional AnoMes As String = "") As String
    Dim caminho As String
    Dim id As String
    Dim modelo As String
    Dim PastaTipoNFe As String
    
    id = NFe.BuscaIdPelaChave(Chave)
    modelo = NFe.BuscaModeloPeloID(id)
    If modelo = "65" Then
        PastaTipoNFe = "NFCe"
    Else
        PastaTipoNFe = "NFe"
    End If
    If AnoMes = "" Then
        AnoMes = NFe.BuscaAnoMesNFe(id)   'Format(Now, "YYYYMM")
    End If
    
    caminho = CaminhoArquivosXMLRetornoRaiz & PastaTipoNFe & "\" & AnoMes & "\" & PastaTipoNFe & "\" & Chave & "-nfe.xml"
    
    If Dir(caminho) = "" Then
        caminho = BuscaCaminhoXMLCopia(AnoMes) & "\" & Chave & "-nfe.xml"
        
    End If
    RetornoCaminhoNFeXML = caminho

End Function


Function ConsultarNFe(Chave As String, Optional AnoMes As String, Optional PastaTipoNFe As String = "NFCe") As String
    Dim c As String
    'Dim anoMes As String
    Dim tipo As String
    
    
    
    
    
    c = RetornoConteudoNFeXML(Chave, AnoMes)
    If c <> "" Then
        GravaArquivoEntradaMonitor "NFe.ConsultarNFe(" & c & ") "
    
        
        If AguardarArquivoRetorno(500) = True Then
            ConsultarNFe = LeArquivoSaidaMonitor(True)
            
        End If
    End If
    
End Function
Function ConsultarDestNFe(cCNPJ As String, nIndicadorNFe As String, nIndicadorEmissor As String, nUltNSU As String) As String
    Dim c As String
    'Dim anoMes As String
    Dim tipo As String
    
    
    
    
    c = cCNPJ
    If nIndicadorNFe <> "" Then
        c = c & "," & nIndicadorNFe
    End If
    If nIndicadorEmissor <> "" Then
        c = c & "," & nIndicadorEmissor
    End If
    
    If nUltNSU <> "" Then
        c = c & "," & nUltNSU
    End If
        GravaArquivoEntradaMonitor "NFe.ConsultarNFeDest(" & c & ") "
    
        
        If AguardarArquivoRetorno(500) = True Then
            ConsultarDestNFe = LeArquivoSaidaMonitor(True)
            
        End If
    
    
End Function






Function XmlRetornoNFeCaminho(caminho As String) As String
    
    
    
    If Dir(caminho) <> "" Then
        XmlRetornoNFeCaminho = LeArquivoExternoUTF8(caminho)
    End If

End Function


Sub BuscaPDFDanfe(id As String)
    Dim s As String
    Dim b As Boolean
    Dim modelo As String
    Dim caminho As String
    Dim Chave As String
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo trata_Erro
    modelo = BuscaModeloPeloID(id)
    Chave = BuscaChaveNFe(id)
    
    
  
    If modelo = "55" Then
        caminho = CaminhoArquivosPDFRetornoNFe & Chave & "-nfe.pdf"
    Else
        caminho = CaminhoArquivosPDFRetornoNFCe & Chave & "-nfe.pdf"
    End If
    
    caminho = mudarCaractere(caminho, "/", "\")
    
    
    If Dir(caminho) <> "" Then
    
        shell "explorer " & caminho
        Exit Sub
    
    
    End If
    
    strSQL = "select * from nfe where id=" & id
    abreObjetoAdo rPesq, strSQL
    
    
    s = RetornoConteudoNFeXML(rPesq!Chave)
    
    ImprimirDanfePDF s, ""
    
    s = AguardarArquivoRetorno(500)
    
    
    
    
    s = PDFRetornoNFe(ado!Chave, ado!modelo, Format(Left(ado!dhEmi, 10), "yyyymm"), True, False)
    If s = "" Then
        MsgBox "O PDF da NFe não foi encontrado!", vbExclamation
        Exit Sub
    End If
        
    
    Exit Sub
trata_Erro:
    
    
        
    
    
End Sub
Sub EnviarDanfeXmlPorEmail(Chave As String)
    Dim s As String
    Dim anexos(2) As String
    Dim m As String
    Dim p As String
    Dim i As Integer
    Dim email As Mail
    Dim emailCliente As String
    
    
    
    
    
    emailCliente = BuscaEmailClientePeloId(Chave, 1)
    
    
    emailCliente = InputBox("Confirme o e-mail a ser enviado", "Email", emailCliente)
    If emailCliente = "" Then
        Exit Sub
    End If
    email = BuscaMailNome("NFe")
    If email.id = "" Then
        MsgBox "Não foi possível encontrar a configuração de Email.", vbExclamation
        Exit Sub
    End If
    s = CopyXmlNFe(Chave)
    
    If s <> "" Then
        anexos(0) = s
    End If
    
    m = NFe.BuscaModeloPelaChave(Chave)
    If m = "55" Then
        
        
        
        s = RetornoConteudoNFeXML(Chave)
        p = bioPDF.ApontarPrinterToPDF(Chave & ".pdf")
        i = ImprimirXMLNFeDANFE(s, Printer)
        Printer.EndDoc
        anexos(1) = p
    End If
    
    
    
    SendMailGmail email.sendusername, email.sendpassword, emailCliente, email.subject, email.body, anexos
    
    
    
    
End Sub
Sub EnviaEmailNFe(cEmailDestino As String, cArqXml As String, Optional cEnviaPdf As Integer = 1)
    
    Dim Comando As String
    Comando = "NFe.EnviarEmail(" & cEmailDestino & "," & cArqXml & "1,Nota Fiscal Eletrônica " & strFantasia_Emp

    If GravaArquivoExternoUTF8(CaminhoArquivoEntradaMonitor, Comando) Then
    End If
    

End Sub
