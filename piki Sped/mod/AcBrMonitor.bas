Attribute VB_Name = "AcBrMonitor"
Public CaminhoArquivoEntradaMonitor As String
Public CaminhoMonitor As String
Public CaminhoArquivoSaidaMonitor As String
Public CaminhoArquivosXMLRetorno As String
Function GeraNfeIni(id As String) As String
    Dim ado As New ADODB.Recordset
    Dim adoImposto As New ADODB.Recordset
    Dim strSQL As String
    
    
    strSQL = "select * from nfe where id=" & id
    abreObjetoAdo ado, strSQL
    
    If ado.RecordCount = 0 Then
        MsgBox "Não encontrado", vbExclamation
        ado.Close
        Exit Function
    End If
    


    Dim s As String
    s = "[infNFe]" & CrLf
    s = s & "versao = 3.10" & CrLf
    
    s = s & "[Identificacao]" & CrLf
    
    s = s & "cNF=" & ado!cNF & CrLf
    
    s = s & "natOp=" & ado!natOp & CrLf
    
    s = s & "indPag = " & ado!indPag & CrLf
    
    s = s & "mod=" & ado!Modelo & CrLf
    
    s = s & "serie=" & ado!serie & CrLf
    
    s = s & "nNF=" & ado!nNf & CrLf
    
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
    
    s = s & "dhCont=" & ado!dhCont & CrLf
    
    s = s & "xJust=" & ado!xJust & CrLf
    
        
    
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
    
    
    strSQL = "select * from nfe_emit where id_nfe=" & id
    abreObjetoAdo ado, strSQL
    
    
    s = s & "[emitente]" & CrLf
    
    s = s & "CNPJCPF=" & ado!CNPJ & CrLf
    
    s = s & "xNome=" & ado!xNome & CrLf
    s = s & "xFant=" & ado!xFant & CrLf
    strSQL = "select * from nfe_enderemit where id_nfe=" & id
    abreObjetoAdo ado, strSQL
    
    
    s = s & "IE=" & ado!ie & CrLf
    s = s & "IEST=" & ado!IEST & CrLf
    s = s & "IM=" & ado!im & CrLf
    s = s & "CNAE=" & ado!CNAE & CrLf
    s = s & "CRT=" & ado!CRT & CrLf
    s = s & "xLgr=" & ado!xLgr & CrLf
    s = s & "nro=" & ado!nro & CrLf
    s = s & "xCpl=" & ado!xCpl & CrLf
    s = s & "xBairro=" & ado!xBairro & CrLf
    s = s & "cMun=" & ado!cMun & CrLf
    s = s & "xMun=" & ado!xMun & CrLf
    s = s & "UF=" & ado!UF & CrLf
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
    
    
    strSQL = "select * from nfe_dest where id_nfe=" & id
    abreObjetoAdo ado, strSQL
    
    s = s & "[Destinatario]" & CrLf
    
    s = s & "idEstrangeiro=" & ado!idEstrangeiro & CrLf
    
    
    
    If Not IsNull(ado!CPF) Then
        If F_CPF(ado!CPF) Then
            s = s & "CPF=" & ado!CPF & CrLf
        End If
    End If
    s = s & "xNome=" & ado!xNome & CrLf
    
    strSQL = "select * from nfe_enderdest where id_nfe=" & id
    abreObjetoAdo ado, strSQL
    
    
    s = s & "indIEDest=" & ado!IndIEDest & CrLf
    s = s & "IE=" & ado!ie & CrLf
    s = s & "ISUF=" & ado!isuf & CrLf
    s = s & "Email=" & ado!email & CrLf
    s = s & "xLgr=" & ado!xLgr & CrLf
    s = s & "nro=" & ado!nro & CrLf
    s = s & "xCpl=" & ado!xCpl & CrLf
    s = s & "xBairro=" & ado!xBairro & CrLf
    s = s & "cMun=" & ado!cMun & CrLf
    s = s & "xMun=" & ado!xMun & CrLf
    s = s & "UF=" & ado!UF & CrLf
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
    
    strSQL = "select * from nfe_item where id_nfe=" & id & " order by id"
    abreObjetoAdo ado, strSQL
    
    With ado
        Do While Not .EOF
            s = s & "[Produto" & Format(.AbsolutePosition, "00#") & "]" & CrLf
            s = s & "cProd=" & !cProd & CrLf
            s = s & "cEAN=" & !cEAN & CrLf
            s = s & "xProd=" & !xProd & CrLf
            s = s & "NCM=" & !ncm & CrLf
            s = s & "EXTIPI=" & !EXTIPI & CrLf
            s = s & "CFOP=" & !CFOP & CrLf
            s = s & "uCom=" & !uCom & CrLf
            s = s & "qCom=" & !qCom & CrLf
            s = s & "vUnCom=" & !vUnCom & CrLf
            s = s & "vProd=" & !vProd & CrLf
            s = s & "cEANTrib=" & !cEANTrib & CrLf
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
            's = s & "infAdProd=" & CrLf
            
            
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
                    
      
            strSQL = "select * from nfe_item_imposto_icmssn102 where id=" & ado!id
            abreObjetoAdo adoImposto, strSQL
            If ado.RecordCount > 0 Then
                s = s & "[ICMS" & Format(adoImposto.AbsolutePosition, "00#") & "]" & CrLf
                s = s & "orig=" & adoImposto!orig & CrLf
                s = s & "CST=" & CrLf
                s = s & "CSOSN=" & adoImposto!csosn & CrLf
                s = s & "modBC=" & CrLf
                s = s & "pRedBC=" & CrLf
                s = s & "vBC=" & CrLf
                s = s & "pICMS=" & CrLf
                s = s & "vICMS=" & CrLf
                s = s & "modBCST=" & CrLf
                s = s & "pMVAST=" & CrLf
                s = s & "pRedBCST=" & CrLf
                s = s & "vBCST=" & CrLf
                s = s & "pICMSST=" & CrLf
                s = s & "vICMSST=" & CrLf
                s = s & "UFST=" & CrLf
                s = s & "pBCOp=" & CrLf
                s = s & "vBCSTRet=" & CrLf
                s = s & "vICMSSTRet=" & CrLf
                s = s & "motDesICMS=" & CrLf
                s = s & "pCredSN=" & CrLf
                s = s & "vCredICMSSN=" & CrLf
                s = s & "vBCSTDest=" & CrLf
                s = s & "vICMSSTDest=" & CrLf
                s = s & "vICMSDeson=" & CrLf
                s = s & "vICMSOp=" & CrLf
                s = s & "pDif=" & CrLf
                s = s & "vICMSDif=" & CrLf
            
            End If
            
            
'            s = s & "[IPI" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
'            s = s & "CST=" & CrLf
'            s = s & "clEnq=" & CrLf
'            s = s & "CNPJProd=" & CrLf
'            s = s & "cSelo=" & CrLf
'            s = s & "qSelo=" & CrLf
'            s = s & "cEnq=" & CrLf
'            s = s & "vBC=" & CrLf
'            s = s & "qUnid=" & CrLf
'            s = s & "vUnid=" & CrLf
'            s = s & "pIPI=" & CrLf
'            s = s & "vIPI=" & CrLf
            
            
            
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
            
            
            s = s & "[PIS" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
            s = s & "CST=49" & CrLf
            s = s & "vBC=0" & CrLf
            s = s & "pPIS=0" & CrLf
            s = s & "qBCProd=" & CrLf
            s = s & "vAliqProd=" & CrLf
            s = s & "vPIS=0" & CrLf
            
            
'            s = s & "[PISST" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
'            s = s & "vBC=" & CrLf
'            s = s & "pPis=" & CrLf
'            s = s & "qBCProd=" & CrLf
'            s = s & "vAliqProd=" & CrLf
'            s = s & "vPIS=" & CrLf
            
            s = s & "[COFINS" & Format(ado.AbsolutePosition, "00#") & "]" & CrLf
            s = s & "CST=49" & CrLf
            s = s & "vBC=0" & CrLf
            s = s & "pCOFINS=0" & CrLf
            s = s & "qBCProd=" & CrLf
            s = s & "vAliqProd=" & CrLf
            s = s & "vCOFINS=0" & CrLf
            
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
    
    strSQL = "select * from nfe_icmstot where id_nfe=" & id
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
    s = s & "vNF=" & ado!vNF & CrLf
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
    s = s & "modFrete = 9" & CrLf
'    CNPJCPF=
'    xNome=
'    IE=
'    xEnder=
'    xMun=
'    UF=
'    vServ=
'    vBCRet=
'    pICMSRet=
'    vICMSRet=
'    CFOP=
'    cMunFG=
'    Placa=
'    UFPlaca=
'    RNTC=
'    vagao=
'    balsa=
    
    
    
    
'    [ReboqueXXX]
'    placa=
'    UF=
'    RNTC=
'
    
    
    
'    [VolumeXXX]
'    qVol=
'    esp=
'    Marca=
'    nVol=
'    pesoL=
'    pesoB=
    
    
    
'    [LacreXXXYYY]
'    nLacre=
    
    
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
            .MoveNext
        Loop
    End With
    
    
    
    s = s & "[DadosAdicionais]" & CrLf
    
    s = s & "infAdFisco=" & CrLf
    
    s = s & "CrLfinfCpl=" & CrLf
    
    
    
    
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
End Function


Sub CriarEnviarNFe(cIniNfE As String, nLote As String)
    Dim comando As String
    
    comando = "NFe.CriarEnviarNFe(" & aspas & cIniNfE & aspas & ", " & nLote & ", 1, 1,1,1)"
    If GravaArquivoExterno(CaminhoArquivoEntradaMonitor, comando) Then
    End If

End Sub
Function TransmitirNFcE(id As String) As String
        Dim strSQL As String
    Dim cIni As String
    cIni = CaminhoMonitor & "NFe" & i & ".ini"
    strSQL = AcBrMonitor.GeraNfeIni(id)
    
    'Kill CaminhoArquivoEntradaMonitor
    If Not GravaArquivoExternoUTF8(cIni, strSQL) Then
    
    End If
    
    
    AcBrMonitor.CriarEnviarNFe cIni, "1"
    If AguardarArquivoRetorno(500) = True Then
        TransmitirNFcE = LeArquivoExternoUTF8(CaminhoArquivoSaidaMonitor)
        MsgBox TransmitirNFcE
        Kill CaminhoArquivoSaidaMonitor
    End If
     

End Function


Function AguardarArquivoRetorno(tempo As Integer) As Boolean
    Dim encontrou As Boolean
    Dim agora As String
    agora = Now
    
    
    If CaminhoArquivoSaidaMonitor = "" Then
        CaminhoArquivoSaidaMonitor = "C:\ACBrMonitorPLUS\sai.txt"
    End If
    Do While encontrou = False
        Sleep tempo
        If Dir(CaminhoArquivoSaidaMonitor, vbArchive) <> "" Then
            'txtStatusServ.Text = LeArquivoExternoUTF8(CaminhoArquivoSaidaMonitor)
            encontrou = True
            AguardarArquivoRetorno = True
            Exit Do
            
        End If
        If DateDiff("s", agora, Now) > 60 Then
            MsgBox "Arquivo de saída não encontrado. Tente novamente", vbExclamation
            Exit Do
            Exit Function
        End If
    Loop
End Function
Sub ImprimirDanfe(caminho As String, protocolo As String)
    Dim comando As String
    If protocolo = "" Then
    
        comando = "NFe.ImprimirDANFEPDF(" & caminho & ")"    '"NFe.ImprimirDANFEPDF(cArqXML,[cProtocolo],[cMarcaDaqgua],[bViaConsumidor],[bSimplificado])"
    Else
        comando = "NFe.ImprimirDANFEPDF(" & caminho & "," & protocolo & ")"
    End If
    GravaArquivoExternoUTF8 CaminhoArquivoEntradaMonitor, comando

End Sub

Function CancelarNFe(chave As String, justificativa As String, cCNPJ As String, nEvento As String) As String
    Dim comando As String
    comando = "NFe.CancelarNFe(" & aspas & chave & aspas & "," & aspas & justificativa & aspas & IIf(cCNPJ <> "", "," & cCNPJ, "") & IIf(nEvento <> "", "," & nEvento, "") & ")"
    GravaArquivoEntradaMonitor comando
End Function
Sub GravaArquivoEntradaMonitor(comando As String)
    On Error Resume Next
    Kill CaminhoArquivoEntradaMonitor
    On Error GoTo trata_Erro
    GravaArquivoExternoUTF8 CaminhoArquivoEntradaMonitor, comando
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub
Function LeArquivoSaidaMonitor(Optional excluir As Boolean = True) As String

    LeArquivoSaidaMonitor = LeArquivoExternoUTF8(CaminhoArquivoSaidaMonitor)

    If LeArquivoSaidaMonitor <> "" Then
        If excluir = True Then
            Kill CaminhoArquivoSaidaMonitor
        End If
    End If

End Function
Sub IniciaAcbr()
    CaminhoArquivoEntradaMonitor = "C:\ACBrMonitorPLUS\ent.txt"
    CaminhoArquivoSaidaMonitor = "C:\ACBrMonitorPLUS\sai.txt"
    CaminhoMonitor = "C:\ACBrMonitorPLUS\"
    CaminhoArquivosXMLRetorno = "C:\ACBrMonitorPLUS\Arqs\" & strCNPJ_Emp & "\NFCe\" & Format(Date, "yyyymm") & "\NFCe\"
End Sub
Function XmlRetornoNFe(chave As String) As String
    Dim caminho As String
    caminho = CaminhoArquivosXMLRetorno & chave & "-nfe.xml"
    
    If Dir(caminho) <> "" Then
        XmlRetornoNFe = LeArquivoExternoUTF8(caminho)
    End If

End Function
Function RetornoCaminhoNFeXML(chave As String) As String
    Dim caminho As String
    caminho = CaminhoArquivosXMLRetorno & chave & "-nfe.xml"
    
    If Dir(caminho) <> "" Then
        RetornoCaminhoNFeXML = caminho
    End If

End Function

Function ConsultarNFe(chave As String) As String
    Dim c As String
    c = RetornoCaminhoNFeXML(chave)
    If c <> "" Then
        GravaArquivoEntradaMonitor "NFe.ConsultarNFe(" & c & ") "
    
        
        If AguardarArquivoRetorno(500) = True Then
            ConsultarNFe = LeArquivoSaidaMonitor
            
        End If
    End If
    
End Function






Function XmlRetornoNFeCaminho(caminho As String) As String
    
    
    
    If Dir(caminho) <> "" Then
        XmlRetornoNFeCaminho = LeArquivoExternoUTF8(caminho)
    End If

End Function


