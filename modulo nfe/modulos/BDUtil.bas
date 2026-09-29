Attribute VB_Name = "BDUtil"
Function BuscaValorSeExistirCampo(ado As ADODB.Recordset, campo As String) As String
    On Error GoTo trata_erro
    Dim i As Integer
    For i = 0 To ado.Fields.Count - 1
        If LCase(ado.Fields(i).Name) = LCase(campo) Then
            BuscaValorSeExistirCampo = ado.Fields(i).Value
            Exit For
        End If
    Next
    
    Exit Function
trata_erro:
    
End Function

