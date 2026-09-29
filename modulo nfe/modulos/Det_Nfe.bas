Attribute VB_Name = "Det_Nfe"
Public Type itemNFe
    produt As TProd
    impost As TImposto
End Type

Public Type detNfe
    num_itens As Integer
    item(64) As itemNFe
End Type


Sub IniciaDet(det As detNfe, numItens)
    
    det.num_itens = numItens
    
    
End Sub
Sub AddItem(item As itemNFe, det As detNfe, i As Integer)
    det.item(i) = item
End Sub

Sub setProdutoItemNFe(item As itemNFe, prod As TProd)
    item.produt = prod
End Sub
Sub setImpostoItemNFe(item As itemNFe, impost As TImposto)
    item.impost = impost
End Sub


Function getXmlDet(det As detNfe) As String
    Dim i As Integer
    
    For i = 1 To det.num_itens
    Next
    
End Function
