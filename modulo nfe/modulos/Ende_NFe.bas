Attribute VB_Name = "Ende_NFe"
Public Type Ende_NFe
    xLgr As String 'logradouro
    nro As String 'numero
    xCpl As String 'complemento
    xBairro As String
    cMun As String 'código do município IBGE
    xMun As String 'nome da cidade
    UF As String
    cep As String
    cPais As String 'código do pais '1058 pro brasil
    xPais As String 'nome do país
    fone As String
    iEst As String
    IndIEDest As String
    IE As String
    isuf As String
    IM As String
    email As String
    CNAE As String
    CRT As String
    
End Type





Function Get_EnderEmi(xLgr As String, _
    nro As String, _
    xCpl As String, _
    xBairro As String, _
    cMun As String, _
    xMun As String, _
    UF As String, _
    cep As String, _
    cPais As String, _
    xPais As String, _
    fone As String, _
    iEst As String, _
    IndIEDest As String, _
    IE As String, _
    isuf As String, _
    IM As String, _
    email As String, _
    CNAE As String, _
    CRT As String) As Ende_NFe.Ende_NFe

    Get_EnderEmi.xLgr = xLgr
    Get_EnderEmi.nro = nro
    Get_EnderEmi.xCpl = xCpl
    Get_EnderEmi.xBairro = xBairro
    Get_EnderEmi.cMun = cMun
    Get_EnderEmi.xMun = xMun
    Get_EnderEmi.UF = UF
    Get_EnderEmi.cep = cep
    Get_EnderEmi.cPais = cPais
    Get_EnderEmi.xPais = xPais
    Get_EnderEmi.fone = fone
    Get_EnderEmi.iEst = iEst
    Get_EnderEmi.IndIEDest = IndIEDest
    Get_EnderEmi.iEst = iEst
    Get_EnderEmi.IM = IM
    Get_EnderEmi.isuf = isuf
    Get_EnderEmi.CRT = CRT
    Get_EnderEmi.CNAE = CNAE
    Get_EnderEmi.email = email
    
End Function
