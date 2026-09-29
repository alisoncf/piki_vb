Attribute VB_Name = "Sintegra"
Public Enum CodigoPosse
    DoInformanteEmSeuPoder = 1
    DoInformanteEmPoderDeTerceiros = 2
    DeTerceirosEmPoderDoInformante = 3
    

End Enum

Function GeraTipo10(CNPJ As String, IE As String, nome As String, municipio As String, _
    uf As String, fax As String, dataI As String, dataF As String, codEstrutura As String, _
    codNatureza As String, CodFinalidade As String) As String
    Dim str As String
    
    str = "10"
    str = str & EsticaTexto(CNPJ, 14, 0, " ", True)
    str = str & EsticaTexto(IE, 14, 0, " ", True)
    str = str & EsticaTexto(nome, 35)
    str = str & EsticaTexto(municipio, 30)
    str = str & EsticaTexto(uf, 2)
    str = str & EsticaTexto(fax, 10, 2, "0", True)
    str = str & GravaDataSintegra(dataI)
    str = str & GravaDataSintegra(dataF)
    str = str & EsticaTexto(codEstrutura, 1)
    str = str & EsticaTexto(codNatureza, 1)
    str = str & EsticaTexto(CodFinalidade, 1)
    
    GeraTipo10 = str
    
End Function

Function geraTipo10Padrao(dataI As String, dataF As String) As String
    geraTipo10Padrao = GeraTipo10(strCNPJ_Emp, strIE_Emp, strNome_Emp, strCidade_Emp, strUF_Emp, strFax_Emp, dataI, dataF, "3", "3", "1")
End Function
Function geraTipo11Padrao() As String
    geraTipo11Padrao = GeraTipo11(strLogradouro_Emp, strNumero_Emp, strComplemento_Emp, strBairro_Emp, _
    strCEP_Emp, strNomeResp_Emp, strFone_Emp)
End Function
Function GeraTipo11(Logradouro As String, numero As String, complemento As String, bairro As String, _
    cep As String, nomeContato As String, telefone As String) As String
    Dim str As String
    
    str = "11"
    str = str & EsticaTexto(Logradouro, 34)
    str = str & EsticaTexto(numero, 5, 2, "0", True)
    str = str & EsticaTexto(complemento, 22)
    str = str & EsticaTexto(bairro, 15)
    str = str & EsticaTexto(cep, 8, 0, "0", True)
    str = str & EsticaTexto(nomeContato, 28)
    str = str & EsticaTexto(telefone, 12, 2, "0", True)
    GeraTipo11 = str
End Function
Function GravaDataSintegra(data As String) As String
    
    GravaDataSintegra = Format(data, "yyyymmdd")
End Function
Function GravaValorSintegra(valor As String, Optional casas As String = "000") As String
    
    valor = Format(valor, "###,###,##0." & casas)
    GravaValorSintegra = F_SoNumero(valor)
    GravaValorSintegra = EsticaTexto(valor, 13, 2, "0", True)
End Function
Function GeraItemTipo74(data As String, CodigoProduto As String, qtd As String, valor As String, _
    CodPosse As CodigoPosse, cnpjPosse As String, iePosse As String, ufPosse As String) As String

    Dim str As String
    str = "74"
    str = str & GravaDataSintegra(data)
    str = str & EsticaTexto(CodigoProduto, 14)
    str = str & GravaValorSintegra(qtd)
    str = str & GravaValorSintegra(valor, "00")
    str = str & EsticaTexto(CStr(CodPosse), 1)
    str = str & EsticaTexto(cnpjPosse, 14)
    str = str & EsticaTexto(iePosse, 14)
    str = str & EsticaTexto(ufPosse, 2)
    str = str & EsticaTexto(" ", 45)
    GeraItemTipo74 = str
    
End Function
Function gravaAliquotaSintegra(Aliquota As String, tamanho As Integer) As String
    If Aliquota = "FF" Or Aliquota = "II" Then
        Aliquota = "17"
    End If
    Aliquota = Format(Aliquota, "##0.00")
    gravaAliquotaSintegra = EsticaTexto(Aliquota, tamanho, 2, "0", True)
End Function
Function GeraItemTipo75(dataI As String, dataF As String, CodigoProduto As String, ncm As String, Descricao As String, _
    unidade As String, ipi As String, icms As String, reducaoBCIcms As String, BCIcms As String) As String

    

    Dim str As String
    str = "75"
    str = str & GravaDataSintegra(dataI)
    str = str & GravaDataSintegra(dataF)
    str = str & EsticaTexto(CodigoProduto, 14)
    str = str & EsticaTexto(ncm, 8)
    str = str & EsticaTexto(Descricao, 53)
    str = str & EsticaTexto(unidade, 6)
    str = str & gravaAliquotaSintegra(ipi, 5)
    str = str & gravaAliquotaSintegra(icms, 4)
    str = str & gravaAliquotaSintegra(reducaoBCIcms, 5)
    str = str & gravaAliquotaSintegra(BCIcms, 13)
    
    GeraItemTipo75 = str
End Function
Function GeraItemTipo90(CNPJ As String, IE As String, tipoTotal As String, total As String, NumRegs90 As String) As String

    '2 14 14 2 8 1

    Dim str As String
    str = "90"
    str = str & EsticaTexto(CNPJ, 14)
    str = str & EsticaTexto(IE, 14)
    str = str & EsticaTexto(tipoTotal, 2)
    str = str & EsticaTexto(total, 8, 2, "0", True)
    str = str & EsticaTexto(NumRegs90, 1)
    
    GeraItemTipo90 = str
    
End Function


Function GeraItemTipo90_74(CNPJ As String, IE As String, total74 As String, total75 As String) As String

    '2 14 14 2 8 1

    Dim str As String
    str = "90"
    str = str & EsticaTexto(CNPJ, 14)
    str = str & EsticaTexto(IE, 14)
    str = str & EsticaTexto("74", 2)
    str = str & EsticaTexto(total74, 8, 2, "0", True)
    str = str & EsticaTexto("75", 2)
    str = str & EsticaTexto(total74, 8, 2, "0", True)
    'str = str & EsticaTexto(NumRegs90, 1)
    
    GeraItemTipo90_74 = CStr(str)
    
End Function



Function GeraItemTipo99(totalRegs As String) As String
    Dim str As String
    str = "99"
    str = str & EsticaTexto(totalRegs, 8, 2, "0", True)
    
    GeraItemTipo99 = str
End Function


