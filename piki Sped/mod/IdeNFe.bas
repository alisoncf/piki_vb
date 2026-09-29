Attribute VB_Name = "IdeNFe"
Type tIde_Nfe
    cNF As String
    natOp As String
    iPag As indPag
    Modelo As TMod
    serie As String
    nNf As String
    dhEmi As String
    tpNf As tpNf
    idDest As String
    tipoImp As tpImp
    tipoEmis As tpEmis
    cDV As String
    tipoAmb As tpAmb
    fin As TfinNfe
    indiFin As indFinal
    indP As indPres
    procEmi As String
    verProc As String
End Type
Function Get_IDE_Nfe(cNF As String, natOp As String, iPag As indPag, Modelo As TMod, serie As String, _
    nNf As String, _
    dhEmi As String, _
    tpNf As tpNf, _
    idDest As String, _
    tipoImp As tpImp, _
    tipoEmis As tpEmis, _
    cDV As String, _
    tipoAmb As tpAmb, _
    fin As TfinNfe, _
    indiFin As indFinal, _
    indP As indPres, _
    procEmi As String, _
    verProc As String) As IdeNFe.tIde_Nfe
    
    Get_IDE_Nfe.cDV = cDV
    Get_IDE_Nfe.cNF = cNF
    Get_IDE_Nfe.dhEmi = dhEmi
    Get_IDE_Nfe.fin = fin
    Get_IDE_Nfe.idDest = idDest
    Get_IDE_Nfe.indiFin = indiFin
    Get_IDE_Nfe.indP = indP
    Get_IDE_Nfe.iPag = iPag
    Get_IDE_Nfe.Modelo = Modelo
    Get_IDE_Nfe.natOp = natOp
    Get_IDE_Nfe.nNf = nNf
    Get_IDE_Nfe.procEmi = procEmi
    Get_IDE_Nfe.serie = serie
    Get_IDE_Nfe.tipoAmb = tipoAmb
    Get_IDE_Nfe.tipoEmis = tipoEmis
    Get_IDE_Nfe.tipoImp = tipoImp
    Get_IDE_Nfe.tpNf = tpNf
    Get_IDE_Nfe.verProc = verProc
    

End Function
Function IDE_NFE(ide As tIde_Nfe) As String
    Dim s As String
    'cNF As String, natOP As String, ipag As indPag, modelo As TMod, serie As String, nNf As String, _
    'dhEmi As String, tpNF As TtpNF, idDest As String, tipoImp As tpImp, tpEmis As String, cDv As String, _
    'tipoAmb As tpAmb, fin As TfinNfe, indiFin As indFinal, indP As indPres, procEmi As String, verProc As String
    
    s = "<ide>"
    s = s & "<cUF>" & strCodigoEstadoIBGE & "</cUF>"
    s = s & "<cNF>" & ide.cNF & "</cNF>"
    s = s & " <natOp>" & ide.natOp & "</natOp>"
    s = s & "<indPag>" & ide.iPag & "</indPag>"
    s = s & "<mod>" & ide.Modelo & "</mod>"
    s = s & "<serie>" & ide.serie & "</serie>"
    s = s & "<nNF>" & ide.nNf & "</nNF>"
    s = s & "<dhEmi>" & ide.dhEmi & "</dhEmi>"
    s = s & "<tpNF>" & ide.tpNf & "</tpNF>"
    s = s & "<idDest>" & ide.idDest & "</idDest>"
    s = s & "<cMunFG>" & strCodigoMunicipioIBGE & "</cMunFG>"
    s = s & "<tpImp>" & ide.tipoImp & "</tpImp>"
    s = s & "<tpEmis>" & ide.tipoEmis & "</tpEmis>"
    s = s & "<cDV>" & ide.cDV & "</cDV>"
    s = s & "<tpAmb>" & ide.tipoAmb & "</tpAmb>"
    s = s & "<finNFe>" & ide.fin & "</finNFe>"
    s = s & "<indFinal>" & ide.indiFin & "</indFinal>"
    s = s & "<indPres>" & ide.indP & "</indPres>"
    s = s & "<procEmi>" & ide.procEmi & "</procEmi>"
    s = s & "<verProc>" & ide.verProc & " </verProc>"

    s = s & "</ide>"
    
    IDE_NFE = s


End Function

