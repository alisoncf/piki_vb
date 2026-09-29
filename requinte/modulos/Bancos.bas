Attribute VB_Name = "Bancos"
Option Explicit
Type TBanco
    cod_bco As String
    nome_bco As String
End Type

Function F_SelectBanco(sExpr As String) As TBanco
        F_SelectBanco = F_BuscaBanco(sExpr)
End Function
Function F_BuscaBanco(sExpr As String) As TBanco
        frmEdiBan.tag = "x"
        frmEdiBan.txtbxBusca.Text = sExpr
        
    
    frmEdiBan.Show 1
    If frmEdiBan.tag <> "" Then
        With frmEdiBan.ado
            
            F_BuscaBanco.cod_bco = !cod_bco
            
            F_BuscaBanco.nome_bco = !nome_bco
        End With
    End If
    Unload frmEdiBan
End Function

