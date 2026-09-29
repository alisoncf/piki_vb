Attribute VB_Name = "Sped_IC"
    Public Type sped_apuracao_icms
        id As String
        dt_ini As String
        dt_fin As String
        VL_TOT_DEBITOS As String
        VL_AJ_DEBITOS As String
        VL_TOT_AJ_DEBITOS As String
        VL_ESTORNOS_CRED As String
        VL_TOT_CREDITOS As String
        VL_AJ_CREDITOS As String
        VL_TOT_AJ_CREDITOS As String
        VL_ESTORNOS_DEB As String
        VL_SLD_CREDOR_ANT As String
        VL_SLD_APURADO As String
        VL_TOT_DED As String
        VL_ICMS_RECOLHER As String
        VL_SLD_CREDOR_TRANSPORTAR As String
        DEB_ESP As String
    End Type

Public Sub gravasped_apuracao_icms(dt_ini As String, dt_fin As String, VL_TOT_DEBITOS As String, VL_AJ_DEBITOS As String, VL_TOT_AJ_DEBITOS As String, VL_ESTORNOS_CRED As String, VL_TOT_CREDITOS As String, VL_AJ_CREDITOS As String, VL_TOT_AJ_CREDITOS As String, VL_ESTORNOS_DEB As String, VL_SLD_CREDOR_ANT As String, VL_SLD_APURADO As String, VL_TOT_DED As String, VL_ICMS_RECOLHER As String, VL_SLD_CREDOR_TRANSPORTAR As String, DEB_ESP As String)
    Dim strSQL As String
    strSQL = " insert into sped_apuracao_icms(dt_ini,dt_fin,VL_TOT_DEBITOS,VL_AJ_DEBITOS,VL_TOT_AJ_DEBITOS,VL_ESTORNOS_CRED,VL_TOT_CREDITOS,VL_AJ_CREDITOS,VL_TOT_AJ_CREDITOS,VL_ESTORNOS_DEB,VL_SLD_CREDOR_ANT,VL_SLD_APURADO,VL_TOT_DED,VL_ICMS_RECOLHER,VL_SLD_CREDOR_TRANSPORTAR,DEB_ESP)values("
    strSQL = strSQL & F_GravaData(dt_ini) & ","
    strSQL = strSQL & F_GravaData(dt_fin) & ","
    strSQL = strSQL & str(VL_TOT_DEBITOS) & ","
    strSQL = strSQL & str(VL_AJ_DEBITOS) & ","
    strSQL = strSQL & str(VL_TOT_AJ_DEBITOS) & ","
    strSQL = strSQL & str(VL_ESTORNOS_CRED) & ","
    strSQL = strSQL & str(VL_TOT_CREDITOS) & ","
    strSQL = strSQL & str(VL_AJ_CREDITOS) & ","
    strSQL = strSQL & str(VL_TOT_AJ_CREDITOS) & ","
    strSQL = strSQL & str(VL_ESTORNOS_DEB) & ","
    strSQL = strSQL & str(VL_SLD_CREDOR_ANT) & ","
    strSQL = strSQL & str(VL_SLD_APURADO) & ","
    strSQL = strSQL & str(VL_TOT_DED) & ","
    strSQL = strSQL & str(VL_ICMS_RECOLHER) & ","
    strSQL = strSQL & str(VL_SLD_CREDOR_TRANSPORTAR) & ","
    strSQL = strSQL & str(DEB_ESP) & ","
    strSQL = strSQL & ")"
    conn.Execute strSQL

End Sub

Sub Alterasped_apuracao_icms(id As String, dt_ini As String, dt_fin As String, VL_TOT_DEBITOS As String, VL_AJ_DEBITOS As String, VL_TOT_AJ_DEBITOS As String, VL_ESTORNOS_CRED As String, VL_TOT_CREDITOS As String, VL_AJ_CREDITOS As String, VL_TOT_AJ_CREDITOS As String, VL_ESTORNOS_DEB As String, VL_SLD_CREDOR_ANT As String, VL_SLD_APURADO As String, VL_TOT_DED As String, VL_ICMS_RECOLHER As String, VL_SLD_CREDOR_TRANSPORTAR As String, DEB_ESP As String)
Dim strSQL As String
    strSQL = " update sped_apuracao_icms set "
    strSQL = strSQL & "dt_ini=" & F_GravaData(dt_ini) & ","
    strSQL = strSQL & "dt_fin=" & F_GravaData(dt_fin) & ","
    strSQL = strSQL & "VL_TOT_DEBITOS=" & str(VL_TOT_DEBITOS) & ","
    strSQL = strSQL & "VL_AJ_DEBITOS=" & str(VL_AJ_DEBITOS) & ","
    strSQL = strSQL & "VL_TOT_AJ_DEBITOS=" & str(VL_TOT_AJ_DEBITOS) & ","
    strSQL = strSQL & "VL_ESTORNOS_CRED=" & str(VL_ESTORNOS_CRED) & ","
    strSQL = strSQL & "VL_TOT_CREDITOS=" & str(VL_TOT_CREDITOS) & ","
    strSQL = strSQL & "VL_AJ_CREDITOS=" & str(VL_AJ_CREDITOS) & ","
    strSQL = strSQL & "VL_TOT_AJ_CREDITOS=" & str(VL_TOT_AJ_CREDITOS) & ","
    strSQL = strSQL & "VL_ESTORNOS_DEB=" & str(VL_ESTORNOS_DEB) & ","
    strSQL = strSQL & "VL_SLD_CREDOR_ANT=" & str(VL_SLD_CREDOR_ANT) & ","
    strSQL = strSQL & "VL_SLD_APURADO=" & str(VL_SLD_APURADO) & ","
    strSQL = strSQL & "VL_TOT_DED=" & str(VL_TOT_DED) & ","
    strSQL = strSQL & "VL_ICMS_RECOLHER=" & str(VL_ICMS_RECOLHER) & ","
    strSQL = strSQL & "VL_SLD_CREDOR_TRANSPORTAR=" & str(VL_SLD_CREDOR_TRANSPORTAR) & ","
    strSQL = strSQL & "DEB_ESP=" & str(DEB_ESP)
    strSQL = strSQL & " where id=" & id
    conn.Execute strSQL
End Sub

