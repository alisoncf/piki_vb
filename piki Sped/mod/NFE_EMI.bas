Attribute VB_Name = "NFE_EMI"
Public Type emi_nfe
    CNPJ As String
    xNome As String
    xFant As String
    ende As Ende_NFe.Ende_NFe
    IE As String
    iEst As String
    IM As String
    CNAE As String
    cCRT As CRT 'código do regime tributário
    
    
End Type



Function Get_Emi_NFe(CNPJ As String, _
    xNome As String, _
    xFant As String, _
    ende As Ende_NFe.Ende_NFe, _
    IE As String, _
    iEst As String, _
    IM As String, _
    CNAE As String, _
    cCRT As CRT) As emi_nfe
    
    Get_Emi_NFe.CNPJ = CNPJ
    Get_Emi_NFe.xNome = xNome
    Get_Emi_NFe.xFant = xFant
    'Get_Emi_NFe.ende = ender
    Get_Emi_NFe.IE = IE
    Get_Emi_NFe.iEst = iEst
    Get_Emi_NFe.IM = IM
    Get_Emi_NFe.CNAE = CNAE
    Get_Emi_NFe.cCRT = cCRT  'código do regime tributário
    
    
    
End Function


