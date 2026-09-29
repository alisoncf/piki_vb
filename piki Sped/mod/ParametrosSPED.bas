Attribute VB_Name = "ParametrosSPED"
Public SpedModeloCUPOM As String
Public SpedSituacaoRegular As String
Public spedCFOPPreferencialCUPOMRetida As String
Public spedCFOPPreferencialCUPOMTributada  As String

Sub iniciaParametrosSPED()
    SpedModeloCUPOM = "2D"
    SpedSituacaoRegular = "00" 'vide tabela sped_situacao_doc
    spedCFOPPreferencialCUPOMRetida = "5405"
    spedCFOPPreferencialCUPOMTributada = "5102"
End Sub
