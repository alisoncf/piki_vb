Attribute VB_Name = "prod_nfe"
Option Explicit
Public ListaProduto As Collection
Public PTest As TProd
Public Type TProd
    cProd As String
    cEAN As String
    xProd As String
    ncm As String
    cCFOP As String
    uCom As String
    qCom As String
    vUnCom As String
    vProd As String
    cEANTrib As String
    uTrib As String
    qTrib As String
    vUnTrib As String
    vDesc As String
    iIndTot As indTot
End Type


Public Function setProduto(cProd As String, cEAN As String, _
    xProd As String, ncm As String, cCFOP As String, _
    uCom As String, qCom As String, vUnCom As String, qCon As String, _
    vProd As String, cEANTrib As String, uTrib As String, qTrib As String, _
    vUnTrib As String, vDesc As String, iIndTot As indTot) As TProd
    
    
    setProduto.cCFOP = cCFOP
    setProduto.cEAN = cEAN
    setProduto.cEANTrib = cEANTrib
    setProduto.cProd = cProd
    setProduto.iIndTot = iIndTot
    setProduto.ncm = ncm
    setProduto.qCom = qCom
    setProduto.qTrib = qTrib
    setProduto.vUnCom = vUnCom
    setProduto.uTrib = uTrib
    setProduto.vDesc = vDesc
    setProduto.vProd = vProd
    setProduto.vUnTrib = vUnTrib
    setProduto.xProd = xProd
    
    
End Function

Function getXmlProduto(prod As TProd) As String

    Dim s As String
    
    'ListaProduto.Add prod
    
       
    
    
    s = "<prod>"
    s = s & "<cProd>" & prod.cProd & "</cProd>"
    s = s & "<cEAN>" & prod.cEAN & "<cEAN>"
    s = s & "<xProd>" & prod.xProd & "</xProd>"
    s = s & "<NCM>" & prod.ncm & "</NCM>"
    s = s & "<CFOP>" & prod.cCFOP & "</CFOP>"
    s = s & "<uCom>" & prod.uCom & "</uCom>"
    s = s & "<qCom>" & prod.qCom & "</qCom>"
    s = s & "<vUnCom>" & prod.vUnCom & "</vUnCom>"
    s = s & "<vProd>" & prod.vProd & "</vProd>"
    s = s & "<cEANTrib>" & prod.cEANTrib & "<cEANTrib>"
    s = s & "<uTrib>" & prod.uTrib & "</uTrib>"
    s = s & "<qTrib>" & prod.qTrib & "</qTrib>"
    s = s & "<vUnTrib>" & prod.vUnTrib & "</vUnTrib>"
    s = s & "<vDesc>" & prod.vDesc & "</vDesc>"
    s = s & "<indTot>" & prod.iIndTot & "</indTot>"
    
    s = s & "</prod>"
    
    getXmlProduto = s


End Function

Public Sub ZeraListaProduto()
    Set ListaProduto = Null
End Sub

Public Sub AddListaProduto(produt As Object, Lista As Collection)
    ListaProduto.Add produt
End Sub


Sub iniciaListaProduto(Lista As Collection)
    Set Lista = New Collection
End Sub
