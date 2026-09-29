Attribute VB_Name = "ICMS_NFe"
Type ICMS_Todos
    tipo As String
    orig As String
    csosn As String
    cst As String
    vBCSTRet As String
    vICMSSTRet As String
    vBCSTDest As String
    vICMSSTDest As String
    

End Type



Public Type ICMSST
    corig As orig
    cst As String
    vBCSTRet As String
    vICMSSTRet As String
    vBCSTDest As String
    vICMSSTDest As String
    
    
End Type


Public Type ICMS00
    corig As orig
    cst As String
    cModBC As modBC
    vBC As String
    pICMS As String
    vICMS As String
    
End Type

Public Type ICMS10
    corig As orig
    cst As String
    cModBC As modBC
    vBC As String
    pICMS As String
    vICMS As String
    cmodBCST As modBCST
    pMVAST As String
    pRedBCST As String
    vBCST As String
    pICMSST As String
    vICMSST As String
    
    
End Type
Public Type ICMS20
    corig As orig
    cst As String
    cModBC As modBC
    pRedBC As String
    vBC As String
    pICMS As String
    vICMS As String
    vICMSDeson As String
    motDesICMS As String
End Type

Public Type ICMS30
    corig As orig
    cst As String
    cmodBCST As modBCST
    pMVAST As String
    pRedBCST As String
    vBCST As String
    pICMSST As String
    vICMSST As String
    vICMSDeson As String
    motDesICMS As String 'motivo da desoneração
    
End Type

Public Type ICMS40

    corig As orig
    cst As String
    vICMSDeson As String
    motDesICMS As String
    
End Type
Public Type ICMS51

    corig As orig
    cst As String
    cModBC As modBC
    pRedBC As String
    vBC As String
    pICMS As String
    vICMSOp As String
    pDif As String
    vICMSDif As String
    vICMS As String
    
End Type
Public Type ICMS60

    corig As orig
    cst As String
    vBCSTRet As String
    vICMSSTRet As String
    
End Type

Public Type ICMS70

    corig As orig
    cst As String
    cModBC As modBC
    pRedBC As String
    vBC As String
    pICMS As String
    vICMS As String
    cmodBCST As modBCST
    pMVAST As String
    pRedBCST As String
    vBCST As String
    pICMSST As String
    vICMSST As String
    vICMSDeson As String
    motDesICMS As String
    
    
End Type

Public Type ICMS90

    corig As orig
    cst As String
    cModBC As modBC
    pRedBC As String
    vBC As String
    pICMS As String
    vICMS As String
    cmodBCST As modBCST
    pMVAST As String
    pRedBCST As String
    vBCST As String
    pICMSST As String
    vICMSST As String
    vICMSDeson As String
    motDesICMS As String
    
    
End Type
Public Type ICMSPART
    corig As orig
    cst As String
    cModBC As modBC
    vBC As String
    pRedBC As String
    pICMS As String
    vICMS As String
    cmodBCST As modBCST
    pMVAST As String
    pRedBCST As String
    vBCST As String
    pICMSST As String
    vICMSST As String
    pBCOp As String
    UFST As String
End Type


Public Type ICMSSN101
    orig As String
    csosn As String
    vCredICMSSN As String
End Type
Public Type ICMSSN102
    orig As String
    csosn As String
End Type
Public Type ICMSSN201
    orig As String
    csosn As String
    cmodBCST As modBCST
    pMVAST As String
    pRedBCST As String
    vBCST As String
    pICMSST As String
    vICMSST As String
    pCredSN As String
    vCredICMSSN As String
End Type

Public Type ICMSSN202
    orig As String
    csosn As String
    cmodBCST As modBCST
    pMVAST As String
    pRedBCST As String
    vBCST As String
    pICMSST As String
    vICMSST As String
    
End Type


Public Type ICMS500
    orig As String
    csosn As String
    vBCSTRet As String
    vICMSSTRet As String
    
End Type

Public Type ICMSSN900
    orig As String
    csosn As String
    cModBC As modBC
    vBC As String
    pRedBC As String
    pICMS As String
    vICMS As String
    modBCST As String
    pMVAST As String
    pRedBCST As String
    vBCST As String
    pICMSST As String
    vICMSST As String
    pCredSN As String
    vCredICMSSN As String
End Type

