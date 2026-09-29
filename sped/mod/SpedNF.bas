Attribute VB_Name = "SpedNF"

Function GravaSpedNf(IND_OPER As String, ind_emit As String, cod_part As String, _
    cod_mod As String, cod_sit As String, SER As String, num_doc As String, _
    chv_nfe As String, dt_doc As String, DT_E_S As String, vl_doc As String, IND_PGTO As String, _
    VL_DESC As String, VL_ABAT_NT As String, VL_MERC As String, IND_FRT As String, _
    VL_FRT As String, VL_SEG As String, VL_OUT_DA As String, VL_BC_ICMS As String, _
    vl_icms As String, VL_BC_ICMS_ST As String, VL_ICMS_ST As String, VL_IPI As String, _
    vl_pis As String, vl_cofins As String, VL_PIS_ST As String, VL_COFINS_ST As String, status As String, data As String, cod_emp As String, _
    Optional tipo_entrada) As String

Dim strSQL As String
Dim rPesq As New ADODB.Recordset

strSQL = "insert into sped_nf(cod_emp,IND_OPER,  IND_EMIT,  COD_PART,  COD_MOD,  COD_SIT,  SER,  NUM_DOC,  "
strSQL = strSQL & " CHV_NFE,  DT_DOC,  DT_E_S,  VL_DOC,  IND_PGTO,  VL_DESC,  VL_ABAT_NT,  "
strSQL = strSQL & " VL_MERC,  IND_FRT,  VL_FRT,  VL_SEG,  VL_OUT_DA,  VL_BC_ICMS,  VL_ICMS,  "
strSQL = strSQL & " VL_BC_ICMS_ST,  VL_ICMS_ST,  VL_IPI,  VL_PIS,  VL_COFINS,  VL_PIS_ST,  VL_COFINS_ST,status,data,cod_usu)values("
strSQL = strSQL & "'" & cod_emp & "',"
strSQL = strSQL & "'" & IND_OPER & "',"
strSQL = strSQL & "'" & ind_emit & "',"
strSQL = strSQL & "'" & cod_part & "',"
strSQL = strSQL & "'" & cod_mod & "',"
strSQL = strSQL & "'" & cod_sit & "',"
strSQL = strSQL & "'" & SER & "',"
strSQL = strSQL & "'" & num_doc & "',"
strSQL = strSQL & "'" & chv_nfe & "',"
strSQL = strSQL & "'" & (dt_doc) & "',"
strSQL = strSQL & "'" & (DT_E_S) & "',"
strSQL = strSQL & F_GravaDec(vl_doc) & ","
strSQL = strSQL & "'" & IND_PGTO & "',"
strSQL = strSQL & F_GravaDec(VL_DESC) & ","
strSQL = strSQL & F_GravaDec(VL_ABAT_NT) & ","
strSQL = strSQL & F_GravaDec(VL_MERC) & ","
strSQL = strSQL & "'" & IND_FRT & "',"
strSQL = strSQL & F_GravaDec(VL_FRT) & ","
strSQL = strSQL & F_GravaDec(VL_SEG) & ","
strSQL = strSQL & F_GravaDec(VL_OUT_DA) & ","
strSQL = strSQL & F_GravaDec(VL_BC_ICMS) & ","
strSQL = strSQL & F_GravaDec(vl_icms) & ","
strSQL = strSQL & F_GravaDec(VL_BC_ICMS_ST) & ","
strSQL = strSQL & F_GravaDec(VL_ICMS_ST) & ","
strSQL = strSQL & F_GravaDec(VL_IPI) & ","
strSQL = strSQL & F_GravaDec(vl_pis) & ","
strSQL = strSQL & F_GravaDec(vl_cofins) & ","
strSQL = strSQL & F_GravaDec(VL_PIS_ST) & ","
strSQL = strSQL & F_GravaDec(VL_COFINS_ST) & ","
strSQL = strSQL & "'" & status & "', "
strSQL = strSQL & F_GravaData(data) & ","
strSQL = strSQL & "'" & strCod_Usu & "')"
Conn.Execute strSQL


strSQL = "select id,num_doc from sped_nf order by id desc limit 1"

abreObjetoAdo rPesq, strSQL

With rPesq
    If .RecordCount > 0 Then
        If !num_doc = num_doc Then
            GravaSpedNf = !id
        End If
    End If
End With
End Function

Sub AlteraSpedNf(id As String, IND_OPER As String, ind_emit As String, cod_part As String, _
    cod_mod As String, cod_sit As String, SER As String, num_doc As String, _
    chv_nfe As String, dt_doc As String, DT_E_S As String, vl_doc As String, IND_PGTO As String, _
    VL_DESC As String, VL_ABAT_NT As String, VL_MERC As String, IND_FRT As String, _
    VL_FRT As String, VL_SEG As String, VL_OUT_DA As String, VL_BC_ICMS As String, _
    vl_icms As String, VL_BC_ICMS_ST As String, VL_ICMS_ST As String, VL_IPI As String, _
    vl_pis As String, vl_cofins As String, VL_PIS_ST As String, VL_COFINS_ST As String, status As String, tipo_entrada As String)

Dim strSQL As String

Dim rPesq As New ADODB.Recordset

If tipo_entrada = "" Then
    tipo_entrada = "01"
End If

strSQL = "update sped_nf set "
strSQL = strSQL & "IND_OPER='" & IND_OPER & "',"
strSQL = strSQL & "IND_EMIT='" & ind_emit & "',"
strSQL = strSQL & "COD_PART='" & cod_part & "',"
strSQL = strSQL & "COD_mod='" & cod_mod & "',"
strSQL = strSQL & "cod_sit='" & cod_sit & "',"
strSQL = strSQL & "ser='" & SER & "',"
strSQL = strSQL & "num_doc='" & num_doc & "',"
strSQL = strSQL & "chv_nfe='" & chv_nfe & "',"
strSQL = strSQL & "dt_doc=" & F_GravaData(dt_doc) & ","
strSQL = strSQL & "dt_e_s=" & F_GravaData(DT_E_S) & ","
strSQL = strSQL & "vl_doc=" & F_GravaDec(vl_doc) & ","
strSQL = strSQL & "IND_PGTO='" & IND_PGTO & "',"
strSQL = strSQL & "VL_DESC=" & F_GravaDec(VL_DESC) & ","
strSQL = strSQL & "VL_ABAT_NT=" & F_GravaDec(VL_ABAT_NT) & ","
strSQL = strSQL & "VL_MERC=" & F_GravaDec(VL_MERC) & ","
strSQL = strSQL & "ind_frt='" & IND_FRT & "',"
strSQL = strSQL & "vl_frt=" & F_GravaDec(VL_FRT) & ","
strSQL = strSQL & "vl_seg=" & F_GravaDec(VL_SEG) & ","
strSQL = strSQL & "VL_OUT_DA=" & F_GravaDec(VL_OUT_DA) & ","
strSQL = strSQL & "VL_BC_ICMS=" & F_GravaDec(VL_BC_ICMS) & ","
strSQL = strSQL & "VL_ICMS=" & F_GravaDec(vl_icms) & ","
strSQL = strSQL & "VL_BC_ICMS_ST=" & F_GravaDec(VL_BC_ICMS_ST) & ","
strSQL = strSQL & "VL_ICMS_ST=" & F_GravaDec(VL_ICMS_ST) & ","
strSQL = strSQL & "VL_IPI=" & F_GravaDec(VL_IPI) & ","
strSQL = strSQL & "VL_PIS=" & F_GravaDec(vl_pis) & ","
strSQL = strSQL & "VL_COFINS=" & F_GravaDec(vl_cofins) & ","
strSQL = strSQL & "VL_PIS_ST=" & F_GravaDec(VL_PIS_ST) & ","
strSQL = strSQL & "VL_COFINS_ST=" & F_GravaDec(VL_COFINS_ST) & ","
strSQL = strSQL & "tipo_entrada='" & tipo_entrada & "', "
strSQL = strSQL & "cod_usu='" & strCod_Usu & "', "
strSQL = strSQL & "status='" & status & "' "
strSQL = strSQL & " where id=" & id
Conn.Execute strSQL




End Sub


Function InsereItemSpedNF(Id_nf As String, NUM_ITEM As String, cod_pro As String, cod_item As String, DESCR_COMPL As String, qtd As String, _
    unid As String, qtd_sai As String, unid_sai As String, vl_item As String, vUnCom As String, VL_DESC As String, Ind_mov As String, cst_icms As String, _
    CFOP As String, cod_nat As String, VL_BC_ICMS As String, aliq_icms As String, vl_icms As String, _
    VL_BC_ICMS_ST As String, ALIQ_ST As String, VL_ICMS_ST As String, IND_APUR As String, CST_IPI As String _
    , cod_enq As String, VL_BC_IPI As String, aliq_ipi As String, VL_IPI As String, cst_pis As String, vl_bc_pis As String, _
    aliq_pis As String, quant_bc_pis As String, ALIQ_PIS_r As String, vl_pis As String, cst_cofins As String, _
    vl_bc_cofins As String, aliq_cofins As String, quant_bc_cofins As String, ALIQ_COFINS_r As String, _
    vl_cofins As String, COD_CTA As String, cEAN As String, cNCM As String, qtd_convert As String, custo As String, preco As String, alterapreco As String) As Boolean
    
    Dim cCusto As String
    If VL_IPI = "" Then VL_IPI = "0"
    If VL_ICMS_ST = "" Then VL_ICMS_ST = "0"
    If vl_item = "" Then vl_item = "0"
    If VL_DESC = "" Then VL_DESC = "0"
    If custo = "0" Then
        cCusto = (CCur(VL_IPI) + CCur(VL_ICMS_ST) + (CCur(vl_item)) - (CCur(VL_DESC) / CCur(qtd_convert))) / CCur(qtd_convert)
    Else
        cCusto = CCur(custo)
    End If
    
    strSQL = "Insert into sped_nf_itens (id_nf,  NUM_ITEM,  cod_pro, COD_ITEM,  DESCR_COMPL,  QTD,  UNID, QTD_saida,  UNID_saida,  VL_ITEM, vUnCom,  "
    strSQL = strSQL & " VL_DESC,  IND_MOV,  CST_ICMS,  CFOP,  COD_NAT,  VL_BC_ICMS,  ALIQ_ICMS,  "
    strSQL = strSQL & " VL_ICMS,  VL_BC_ICMS_ST,  ALIQ_ST,  VL_ICMS_ST,  IND_APUR,  CST_IPI,  COD_ENQ,  "
    strSQL = strSQL & " VL_BC_IPI,  ALIQ_IPI,  VL_IPI,  CST_PIS,  VL_BC_PIS,  ALIQ_PIS,  QUANT_BC_PIS,  "
    strSQL = strSQL & " ALIQ_PIS_r,  VL_PIS,  CST_COFINS,  VL_BC_COFINS,  ALIQ_COFINS,  QUANT_BC_COFINS,  "
    strSQL = strSQL & " ALIQ_COFINS_r,  VL_COFINS,  COD_CTA,  cEAN,cNCM, qtd_convert,custo,preco, alterapreco)values("
    strSQL = strSQL & Id_nf & ","
    strSQL = strSQL & "'" & NUM_ITEM & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & "'" & cod_item & "',"
    strSQL = strSQL & "'" & DESCR_COMPL & "',"
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & "'" & unid & "',"
    strSQL = strSQL & F_GravaDec(qtd_sai) & ","
    strSQL = strSQL & "'" & unid_sai & "',"
    strSQL = strSQL & F_GravaDec(vl_item) & ","
    strSQL = strSQL & F_GravaDec(vUnCom) & ","
    strSQL = strSQL & F_GravaDec(VL_DESC) & ","
    strSQL = strSQL & "'" & Ind_mov & "',"
    strSQL = strSQL & "'" & cst_icms & "',"
    strSQL = strSQL & "'" & CFOP & "',"
    strSQL = strSQL & "'" & cod_nat & "',"
    strSQL = strSQL & F_GravaDec(VL_BC_ICMS) & ","
    strSQL = strSQL & F_GravaDec(aliq_icms) & ","
    strSQL = strSQL & F_GravaDec(vl_icms) & ","
    strSQL = strSQL & F_GravaDec(VL_BC_ICMS_ST) & ","
    strSQL = strSQL & F_GravaDec(ALIQ_ST) & ","
    strSQL = strSQL & F_GravaDec(VL_ICMS_ST) & ","
    strSQL = strSQL & "'" & IND_APUR & "',"
    strSQL = strSQL & "'" & CST_IPI & "',"
    strSQL = strSQL & "'" & cod_enq & "',"
    strSQL = strSQL & F_GravaDec(VL_BC_IPI) & ","
    strSQL = strSQL & F_GravaDec(aliq_ipi) & ","
    strSQL = strSQL & F_GravaDec(VL_IPI) & ","
    strSQL = strSQL & "'" & cst_pis & "',"
    strSQL = strSQL & F_GravaDec(vl_bc_pis) & ","
    strSQL = strSQL & F_GravaDec(aliq_pis) & ","
    strSQL = strSQL & F_GravaDec(quant_bc_pis) & ","
    strSQL = strSQL & F_GravaDec(ALIQ_PIS_r) & ","
    strSQL = strSQL & F_GravaDec(vl_pis) & ","
    strSQL = strSQL & "'" & cst_cofins & "',"
    strSQL = strSQL & F_GravaDec(vl_bc_cofins) & ","
    strSQL = strSQL & F_GravaDec(aliq_cofins) & ","
    strSQL = strSQL & F_GravaDec(quant_bc_cofins) & ","
    strSQL = strSQL & F_GravaDec(ALIQ_COFINS_r) & ","
    strSQL = strSQL & F_GravaDec(vl_cofins) & ","
    strSQL = strSQL & "'" & COD_CTA & "', "
    strSQL = strSQL & "'" & cEAN & "', "
    strSQL = strSQL & "'" & cNCM & "', "
    strSQL = strSQL & F_GravaDec(qtd_convert) & ","
    strSQL = strSQL & F_GravaDec(cCusto) & ","
    strSQL = strSQL & F_GravaDec(preco) & ","
    strSQL = strSQL & "'" & alterapreco & "')"
    Conn.Execute strSQL
    
        InsereItemSpedNF = True
        Exit Function
trata_Erro:
        InsereItemSpedNF = False
    
    End Function
Sub AlteraParticipante(id As String, cod_part As String)
    strSQL = "update sped_nf set cod_part='" & cod_part & "'"
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
    
End Sub
Sub AlteraCodProItemNf(id As String, cod_pro As String)
    strSQL = "update sped_nf_itens set cod_pro='" & cod_pro & "'"
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
    
End Sub
Sub AlteraCodEmp(id As String, cod_emp As String)
    On Error GoTo trata_Erro
    If cod_emp = "" Then
        cod_emp = "01"
    End If
    strSQL = "update sped_nf set cod_emp='" & cod_emp & "'"
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
    Exit Sub
trata_Erro:
  '  MostraErro

End Sub
Sub AlteraTipoNota(id As String, tipo_entrada As String)
    On Error GoTo trata_Erro
    If tipo_entrada = "" Then
        tipo_entrada = "01"
    End If
    strSQL = "update sped_nf set tipo_entrada='" & tipo_entrada & "'"
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
    Exit Sub
trata_Erro:
  '  MostraErro

End Sub


Sub AlteraStatusNf(id As String, status As String)
    strSQL = "update sped_nf set status='" & status & "'"
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
    
End Sub


Sub ZeraCodigoItemSpedNF(id As String)
    Dim strSQL As String
    strSQL = "update sped_nf_itens set cod_pro=''"
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
End Sub


Sub AlteraItemSpedNF(id As String, NUM_ITEM As String, cod_pro As String, cod_item As String, DESCR_COMPL As String, qtd As String, _
    unid As String, qtd_sai As String, unid_sai As String, vl_item As String, VL_DESC As String, Ind_mov As String, cst_icms As String, _
    CFOP As String, cod_nat As String, VL_BC_ICMS As String, aliq_icms As String, vl_icms As String, _
    VL_BC_ICMS_ST As String, ALIQ_ST As String, VL_ICMS_ST As String, IND_APUR As String, CST_IPI As String _
    , cod_enq As String, VL_BC_IPI As String, aliq_ipi As String, VL_IPI As String, cst_pis As String, vl_bc_pis As String, _
    aliq_pis As String, quant_bc_pis As String, ALIQ_PIS_r As String, vl_pis As String, cst_cofins As String, _
    vl_bc_cofins As String, aliq_cofins As String, quant_bc_cofins As String, ALIQ_COFINS_r As String, _
    vl_cofins As String, COD_CTA As String, qtd_convert As String, custo As String, precofuturo As String, alterapreco As String)


    strSQL = "update sped_nf_itens set "
    
    strSQL = strSQL & "NUM_ITEM =" & NUM_ITEM & ","
    strSQL = strSQL & "COD_pro='" & cod_pro & "',"
    If cod_item <> "" Then
        strSQL = strSQL & "COD_ITEM='" & cod_item & "',"
    End If
    strSQL = strSQL & "DESCR_COMPL='" & DESCR_COMPL & "',"
    strSQL = strSQL & "qtd=" & F_GravaDec(qtd) & ","
    strSQL = strSQL & "unid='" & unid & "',"
    strSQL = strSQL & "qtd_saida=" & F_GravaDec(qtd_sai) & ","
    strSQL = strSQL & "unid_saida='" & unid_sai & "',"
    strSQL = strSQL & "VL_ITEM=" & F_GravaDec(vl_item) & ","
    strSQL = strSQL & "VL_DESC=" & F_GravaDec(VL_DESC) & ","
    strSQL = strSQL & "IND_MOV ='" & Ind_mov & "',"
    strSQL = strSQL & "CST_ICMS ='" & cst_icms & "',"
    strSQL = strSQL & "CFOP='" & CFOP & "',"
    strSQL = strSQL & "COD_NAT='" & cod_nat & "',"
    strSQL = strSQL & "VL_BC_ICMS=" & F_GravaDec(VL_BC_ICMS) & ","
    strSQL = strSQL & "ALIQ_ICMS=" & F_GravaDec(aliq_icms) & ","
    strSQL = strSQL & "VL_ICMS=" & F_GravaDec(vl_icms) & ","
    strSQL = strSQL & "VL_BC_ICMS_ST=" & F_GravaDec(VL_BC_ICMS_ST) & ","
    strSQL = strSQL & "ALIQ_ST=" & F_GravaDec(ALIQ_ST) & ","
    strSQL = strSQL & "VL_ICMS_ST=" & F_GravaDec(VL_ICMS_ST) & ","
    strSQL = strSQL & "IND_APUR='" & IND_APUR & "',"
    strSQL = strSQL & "CST_IPI ='" & CST_IPI & "',"
    strSQL = strSQL & "COD_ENQ ='" & cod_enq & "',"
    strSQL = strSQL & "VL_BC_IPI=" & F_GravaDec(VL_BC_IPI) & ","
    strSQL = strSQL & "ALIQ_IPI=" & F_GravaDec(aliq_ipi) & ","
    strSQL = strSQL & "VL_IPI=" & F_GravaDec(VL_IPI) & ","
    strSQL = strSQL & "CST_PIS ='" & cst_pis & "',"
    strSQL = strSQL & "VL_BC_PIS=" & F_GravaDec(vl_bc_pis) & ","
    strSQL = strSQL & "ALIQ_PIS=" & F_GravaDec(aliq_pis) & ","
    strSQL = strSQL & "QUANT_BC_PIS=" & F_GravaDec(quant_bc_pis) & ","
    strSQL = strSQL & "ALIQ_PIS_r=" & F_GravaDec(ALIQ_PIS_r) & ","
    strSQL = strSQL & "VL_PIS=" & F_GravaDec(vl_pis) & ","
    strSQL = strSQL & "CST_COFINS ='" & cst_cofins & "',"
    strSQL = strSQL & "VL_BC_COFINS=" & F_GravaDec(vl_bc_cofins) & ","
    strSQL = strSQL & "ALIQ_COFINS=" & F_GravaDec(aliq_cofins) & ","
    strSQL = strSQL & "QUANT_BC_COFINS=" & F_GravaDec(quant_bc_cofins) & ","
    strSQL = strSQL & "ALIQ_COFINS_r=" & F_GravaDec(ALIQ_COFINS_r) & ","
    strSQL = strSQL & "VL_COFINS=" & F_GravaDec(vl_cofins) & ","
    strSQL = strSQL & "qtd_convert=" & F_GravaDec(qtd_convert) & ","
    strSQL = strSQL & "custo=" & F_GravaDec(custo) & ","
    strSQL = strSQL & "preco=" & F_GravaDec(precofuturo) & ","
    strSQL = strSQL & "COD_CTA='" & COD_CTA & "', "
    strSQL = strSQL & " alterapreco='" & alterapreco & "' "
    strSQL = strSQL & " WHERE id=" & id
    Conn.Execute strSQL
    
    
    
    End Sub
    
Function QuebraProdutoSpedXml(xml As String) As XML_SPED.prod()
    Dim c As New Collection
    Dim i As Single
    Dim a As Single, b As Single
    Dim Linha As String
    Dim p As XML_SPED.prod
    Dim vp(999) As XML_SPED.prod
    Dim vp2() As XML_SPED.prod
    Dim k As Single
    i = 1
    Do While i > 0
        Linha = QuebraTagsXML(xml, "<det ", "</det>", i)
        
        
        p.cEAN = QuebraTagsXML(Linha, "<cEAN>", "</cEAN>", 1)
        p.cEANTrib = QuebraTagsXML(Linha, "<cEANTrib>", "</cEANTrib>", 1)
        p.CFOP = QuebraTagsXML(Linha, "<CFOP>", "</CFOP>", 1)
        p.cProd = QuebraTagsXML(Linha, "<cProd>", "</cProd>", 1)
        
        p.indTot = QuebraTagsXML(Linha, "<indTot>", "</indTot>", 1)
        p.ncm = QuebraTagsXML(Linha, "<NCM>", "</NCM>", 1)
        p.cNCM = QuebraTagsXML(Linha, "<NCM>", "</NCM>", 1)
        p.qCom = QuebraTagsXML(Linha, "<qCom>", "</qCom>", 1)
        p.qTrib = QuebraTagsXML(Linha, "<qTrib>", "</qTrib>", 1)
        p.uCom = QuebraTagsXML(Linha, "<uCom>", "</uCom>", 1)
        p.uTrib = QuebraTagsXML(Linha, "<uTrib>", "</uTrib>", 1)
        p.uCom = QuebraTagsXML(Linha, "<uCom>", "</uCom>", 1)
        p.vOutro = QuebraTagsXML(Linha, "<vOutro>", "</vOutro>", 1)
        p.vProd = QuebraTagsXML(Linha, "<vProd>", "</vProd>", 1)
        p.vUnCom = QuebraTagsXML(Linha, "<vUnCom>", "</vUnCom>", 1)
        p.vDesc = QuebraTagsXML(Linha, "<vDesc>", "</vDesc>", 1)
        p.vUnTrib = QuebraTagsXML(Linha, "<vUnTrib>", "</vUnTrib>", 1)
        p.xProd = QuebraTagsXML(Linha, "<xProd>", "</xProd>", 1)
        p.vFrete = QuebraTagsXML(Linha, "<vFrete>", "</vFrete>", 1)
        p.vSeguro = QuebraTagsXML(Linha, "<vSeguro>", "</vSeguro>", 1)
        
        p.icms_.cst = QuebraTagsXML(Linha, "<CST>", "</CST>", 1)
                
        p.icms_.modBC = QuebraTagsXML(Linha, "<modBC>", "</modBC>", 1)
        p.icms_.modBCST = QuebraTagsXML(Linha, "<modBCST>", "</modBCST>", 1)
        p.icms_.orig = QuebraTagsXML(Linha, "<orig>", "</orig>", 1)
        p.icms_.pICMS = QuebraTagsXML(Linha, "<pICMS>", "</pICMS>", 1)
        p.icms_.vBC = QuebraTagsXML(Linha, "<vBC>", "</vBC>", 1)
        p.icms_.vICMS = QuebraTagsXML(Linha, "<vICMS>", "</vICMS>", 1)
        p.icms_.pICMSST = QuebraTagsXML(Linha, "<pICMSST>", "</pICMSST>", 1)
        p.icms_.vICMSST = QuebraTagsXML(Linha, "<vICMSST>", "</vICMSST>", 1)
         
         
        p.ipi_.cEnq = QuebraTagsXML(Linha, "<cEnq>", "</cEnq>", 1)
        p.ipi_.cst = QuebraTagsXML(Linha, "<CST>", "</CST>", 1)
        p.ipi_.vBC = QuebraTagsXML(Linha, "<vBC>", "</vBC>", 1)
        p.ipi_.pIPI = QuebraTagsXML(Linha, "<pIPI>", "</pIPI>", 1)
        p.ipi_.vIPI = QuebraTagsXML(Linha, "<vIPI>", "</vIPI>", 1)
                        
        p.pis_.cst = QuebraTagsXML(Linha, "<CST>", "</CST>", 1)
        p.pis_.vBC = QuebraTagsXML(Linha, "<vBC>", "</vBC>", 1)
        p.pis_.pPIS = QuebraTagsXML(Linha, "<pPIS>", "</pPIS>", 1)
        p.pis_.vPIS = QuebraTagsXML(Linha, "<vPIS>", "</vPIS>", 1)
         
        p.cof_.cst = QuebraTagsXML(Linha, "<CST>", "</CST>", 1)
        p.cof_.vBC = QuebraTagsXML(Linha, "<vBC>", "</vBC>", 1)
        p.cof_.pCOFINS = QuebraTagsXML(Linha, "<pCOFINS>", "</pCOFINS>", 1)
        p.cof_.vCOFINS = QuebraTagsXML(Linha, "<vCOFINS>", "</vCOFINS>", 1)
                 
                 
                 
         
        k = i
        
        vp(b) = p
        
        i = RetornaPosicaoTag(xml, "</det>", i + Len(Linha))
        
        If i = k Then
            i = 0
            
        Else
            b = b + 1
        End If
        
    Loop
    ReDim QuebraProdutoSpedXml(b) As XML_SPED.prod
    ReDim vp2(b) As XML_SPED.prod
    For i = 1 To b
        vp2(i - 1) = vp(i - 1)
        
    Next
    
    qtp_prod_sped = b
    QuebraProdutoSpedXml = vp2
    
    
End Function
Function QuebraCobrancaSpedXml(xml As String) As XML_SPED.Cobranca
    Dim c As New Collection
    Dim i As Single
    Dim a As Single, b As Single
    Dim Linha As String
    Dim linha1 As String
    Dim p As XML_SPED.duplicata
    Dim f As XML_SPED.fatura
    Dim vp(99) As XML_SPED.duplicata
    Dim vp2() As XML_SPED.duplicata
    Dim k As Single
    i = 1
    Linha = QuebraTagsXML(xml, "<cobr>", "</cobr>", i)
    f.nFat = QuebraTagsXML(Linha, "<nFat>", "</nFat>", 1)
    f.vLiq = QuebraTagsXML(Linha, "<vLiq>", "</vLiq>", 1)
    f.vOrig = QuebraTagsXML(Linha, "<vOrig>", "</vOrig>", 1)
    
    Do While i > 0
        
        linha1 = QuebraTagsXML(Linha, "<dup>", "</dup>", i)
        
        m = RetornaPosicaoTag(Linha, "</dup>", i + 1)
        p.dVenc = QuebraTagsXML(linha1, "<dVenc>", "</dVenc>", 1)
        p.nDup = QuebraTagsXML(linha1, "<nDup>", "</nDup>", 1)
        p.vDup = QuebraTagsXML(linha1, "<vDup>", "</vDup>", 1)
        
        
        
        k = i
        
        vp(b) = p
        
        i = RetornaPosicaoTag(Linha, "</dup>", i + 1)
        
        If i = k Then
            i = 0
            
        Else
            b = b + 1
        End If
        
    Loop
    ReDim QuebraCobrancaSpedXml.d(b) As XML_SPED.duplicata
    ReDim vp2(b) As XML_SPED.duplicata
    For i = 1 To b
        vp2(i - 1) = vp(i - 1)
        
    Next
    
    
    QuebraCobrancaSpedXml.d = vp2
    QuebraCobrancaSpedXml.fat = f
    
    
End Function


    
Function Le_XML(caminhoXML As String) As XMLSped
    Dim xml As String
    Dim textLine As String
    Dim i As Single
    Dim b1 As String
    Dim b2 As String
    Dim v As String
    Dim iSeek As Single
    Dim vProd() As XML_SPED.prod
    Dim Cobr As XML_SPED.Cobranca
    Dim a As Single
    Dim b As Single
    
    AbreArquivoExternoLeitura caminhoXML, 1
    
    Do While Not EOF(1)
        Line Input #1, textLine
        xml = xml & textLine
    Loop

    textLine = QuebraTagsXML(xml, "<infNFe", ">")
    Le_XML.cab.id = QuebraTagsXML(textLine, "Id=", " ", 1)
    If Le_XML.cab.id = "" Then
        Le_XML.cab.id = QuebraTagsXML(textLine, "Id=", "", 1)
    End If
        
        
    textLine = QuebraTagsXML(xml, "<ide>", "</ide>")
    
    Le_XML.cab.cDV = QuebraTagsXML(textLine, "<cDV>", "</cDV>", 1)
    Le_XML.cab.cMunFG = QuebraTagsXML(textLine, "<cMunFG>", "</cMunFG>", 1)
    Le_XML.cab.cNF = QuebraTagsXML(textLine, "<cNF>", "</cNF>", 1)
    Le_XML.cab.cUF = QuebraTagsXML(textLine, "<cUF>", "</cUF>", 1)
    Le_XML.cab.dEmi = QuebraTagsXML(textLine, "<dhEmi>", "</dhEmi>", 1)
    Le_XML.cab.finNFe = QuebraTagsXML(textLine, "<finNFe>", "</finNFe>", 1)
    Le_XML.cab.indPag = QuebraTagsXML(textLine, "<indPag>", "</indPag>", 1)
    Le_XML.cab.mod = QuebraTagsXML(textLine, "<mod>", "</mod>", 1)
    Le_XML.cab.natOp = QuebraTagsXML(textLine, "<natOp>", "</natOp>", 1)
    Le_XML.cab.nNF = QuebraTagsXML(textLine, "<nNF>", "</nNF>", 1)
    Le_XML.cab.procEmi = QuebraTagsXML(textLine, "<procEmi>", "</procEmi>", 1)
    Le_XML.cab.serie = QuebraTagsXML(textLine, "<serie>", "</serie>", 1)
    Le_XML.cab.tpAmb = QuebraTagsXML(textLine, "<tpAmb>", "</tpAmb>", 1)
    Le_XML.cab.tpEmis = QuebraTagsXML(textLine, "<tpEmis>", "</tpEmis>", 1)
    Le_XML.cab.tpImp = QuebraTagsXML(textLine, "<tpImp>", "</tpImp>", 1)
    Le_XML.cab.tpNf = QuebraTagsXML(textLine, "<tpNF>", "</tpNF>", 1)
    Le_XML.cab.verProc = QuebraTagsXML(textLine, "<verProc>", "</verProc>", 1)
    If textLine <> "" Then
        Le_XML.cab.idDst = QuebraTagsXML(textLine, "<idDest>", "</idDest>", 1)
    
        Le_XML.cab.indFnl = QuebraTagsXML(textLine, "<indFinal>", "</indFinal>", 1)
        Le_XML.cab.indPr = QuebraTagsXML(textLine, "<indPres>", "</indPres>", 1)
        Le_XML.cab.dhCont = QuebraTagsXML(textLine, "<dhCont>", "</dhCont>", 1)
        Le_XML.cab.xJust = QuebraTagsXML(textLine, "<xJust>", "</xJust>", 1)
    End If
    
    
    textLine = QuebraTagsXML(xml, "<emit>", "</emit>")
    Le_XML.cab.e.cep = QuebraTagsXML(textLine, "<CEP>", "</CEP>", 1)
    Le_XML.cab.e.cMun = QuebraTagsXML(textLine, "<cMun>", "</cMun>", 1)
    Le_XML.cab.e.CNPJ = QuebraTagsXML(textLine, "<CNPJ>", "</CNPJ>", 1)
    Le_XML.cab.e.cPais = QuebraTagsXML(textLine, "<cPais>", "</cPais>", 1)
    Le_XML.cab.e.CRT = QuebraTagsXML(textLine, "<CRT>", "</CRT>", 1)
    Le_XML.cab.e.fone = QuebraTagsXML(textLine, "<fone>", "</fone>", 1)
    Le_XML.cab.e.IE = QuebraTagsXML(textLine, "<IE>", "</IE>", 1)
    Le_XML.cab.e.nro = QuebraTagsXML(textLine, "<nro>", "</nro>", 1)
    Le_XML.cab.e.UF = QuebraTagsXML(textLine, "<UF>", "</UF>", 1)
    Le_XML.cab.e.xBairro = QuebraTagsXML(textLine, "<xBairro>", "</xBairro>", 1)
    Le_XML.cab.e.xFant = QuebraTagsXML(textLine, "<xFant>", "</xFant>", 1)
    Le_XML.cab.e.xLgr = QuebraTagsXML(textLine, "<xLgr>", "</xLgr>", 1)
    Le_XML.cab.e.xMun = QuebraTagsXML(textLine, "<xMun>", "</xMun>", 1)
    Le_XML.cab.e.xNome = QuebraTagsXML(textLine, "<xNome>", "</xNome>", 1)
    Le_XML.cab.e.xPais = QuebraTagsXML(textLine, "<xPais>", "</xPais>", 1)
    
    
    textLine = QuebraTagsXML(xml, "<dest>", "</dest>")
    Le_XML.cab.d.cep = QuebraTagsXML(textLine, "<CEP>", "</CEP>", 1)
    Le_XML.cab.d.cMun = QuebraTagsXML(textLine, "<cMun>", "</cMun>", 1)
    Le_XML.cab.d.CNPJ = QuebraTagsXML(textLine, "<CNPJ>", "</CNPJ>", 1)
    Le_XML.cab.d.cPais = QuebraTagsXML(textLine, "<cPais>", "</cPais>", 1)
    Le_XML.cab.d.CRT = QuebraTagsXML(textLine, "<CRT>", "</CRT>", 1)
    Le_XML.cab.d.fone = QuebraTagsXML(textLine, "<fone>", "</fone>", 1)
    Le_XML.cab.d.IE = QuebraTagsXML(textLine, "<IE>", "</IE>", 1)
    Le_XML.cab.d.nro = QuebraTagsXML(textLine, "<nro>", "</nro>", 1)
    Le_XML.cab.d.UF = QuebraTagsXML(textLine, "<UF>", "</UF>", 1)
    Le_XML.cab.d.xBairro = QuebraTagsXML(textLine, "<xBairro>", "</xBairro>", 1)
    Le_XML.cab.d.xFant = QuebraTagsXML(textLine, "<xFant>", "</xFant>", 1)
    Le_XML.cab.d.xLgr = QuebraTagsXML(textLine, "<xLgr>", "</xLgr>", 1)
    Le_XML.cab.d.xMun = QuebraTagsXML(textLine, "<xMun>", "</xMun>", 1)
    Le_XML.cab.d.xNome = QuebraTagsXML(textLine, "<xNome>", "</xNome>", 1)
    Le_XML.cab.d.xPais = QuebraTagsXML(textLine, "<xPais>", "</xPais>", 1)
    
    
    textLine = QuebraTagsXML(xml, "<total>", "</total>")
    Le_XML.tot.vBC = QuebraTagsXML(textLine, "<vBC>", "</vBC>")
    Le_XML.tot.vICMS = QuebraTagsXML(textLine, "<vICMS>", "</vICMS")
    Le_XML.tot.vBCST = QuebraTagsXML(textLine, "<vBCST>", "</vBCST>")
    Le_XML.tot.vProd = QuebraTagsXML(textLine, "<vProd>", "</vProd>")
    Le_XML.tot.vFrete = QuebraTagsXML(textLine, "<vFrete>", "</vFrete>")
    Le_XML.tot.vSeg = QuebraTagsXML(textLine, "<vSeg>", "</vSeg>")
    Le_XML.tot.vDesc = QuebraTagsXML(textLine, "<vDesc>", "</vDesc>")
    Le_XML.tot.vII = QuebraTagsXML(textLine, "<vII>", "</vII>")
    Le_XML.tot.vIPI = QuebraTagsXML(textLine, "<vIPI>", "</vIPI>")
    Le_XML.tot.vPIS = QuebraTagsXML(textLine, "<vPIS>", "</vPIS>")
    Le_XML.tot.vCOFINS = QuebraTagsXML(textLine, "<vCOFINS>", "</vCOFINS>")
    Le_XML.tot.vOutro = QuebraTagsXML(textLine, "<vOutro>", "</vOutro>")
    Le_XML.tot.vNF = QuebraTagsXML(textLine, "<vNF>", "</vNF>")
    
    
    textLine = QuebraTagsXML(xml, "<transp>", "</transp>")
    Le_XML.t.modFrete = QuebraTagsXML(textLine, "<modFrete>", "</modFrete>")
    Le_XML.t.t.CNPJ = QuebraTagsXML(textLine, "<CNPJ>", "</CNPJ>")
    Le_XML.t.t.xNome = QuebraTagsXML(textLine, "<xNome>", "</xNome>")
    Le_XML.t.t.xEnder = QuebraTagsXML(textLine, "<xEnder>", "</xEnder>")
    Le_XML.t.t.xMun = QuebraTagsXML(textLine, "<xMun>", "</xMun>")
    
    
    textLine = QuebraTagsXML(xml, "<cobr>", "</cobr>")
    Le_XML.t.t.CNPJ = QuebraTagsXML(textLine, "<CNPJ>", "</CNPJ>")
    
    Cobr = QuebraCobrancaSpedXml(xml)
    
    ReDim Le_XML.c.d(UBound(Cobr.d))
    
 '   ReDim Le_XML.c.fat(UBound(Cobr.fat))
    Le_XML.c.d = Cobr.d
    Le_XML.c.fat = Cobr.fat
    
    vProd = QuebraProdutoSpedXml(xml)
    
    ReDim Le_XML.p(UBound(vProd))
    Le_XML.p = vProd
        
    
    textLine = QuebraTagsXML(xml, "<infProt", "</infProt>")
    
    Le_XML.infProt.cStat = QuebraTagsXML(textLine, "<cStat>", "</cStat>")
    Le_XML.infProt.dhRecbto = QuebraTagsXML(textLine, "<dhRecbto>", "</dhRecbto>")
    Le_XML.infProt.digVal = QuebraTagsXML(textLine, "<digVal>", "</digVal>")
    Le_XML.infProt.nProt = QuebraTagsXML(textLine, "<nProt>", "</nProt>")
    Le_XML.infProt.tpAmb = QuebraTagsXML(textLine, "<tpAmb>", "</tpAmb>")
    Le_XML.infProt.verApp = QuebraTagsXML(textLine, "<verApp>", "</verApp>")
    Le_XML.infProt.xMotivo = QuebraTagsXML(textLine, "<xMotivo>", "</xMotivo>")
    
    Le_XML.CH_NFE = QuebraTagsXML(xml, "<chNFe>", "</chNFe>")
    
    
    
    
FechaArquivoExterno 1

End Function
    
Function RetornaPosicaoTag(xml As String, tag As String, Optional inicio As Single = 1) As Single
    Dim a As Single
    Dim b As Single
    If inicio = 0 Then inicio = 1
    a = InStr(inicio, xml, tag)
    RetornaPosicaoTag = a
End Function
Function QuebraTagsXML(XML_Conteudo As String, tagI As String, tagF As String, Optional inicio As Single = 1) As String
    Dim a As Single
    Dim b As Single
    On Error GoTo trata_Erro:
    If inicio = 0 Then inicio = 1
    a = InStr(inicio, XML_Conteudo, tagI)
    If a > 0 Then
        If tagF <> "" Then
            b = InStr(a, XML_Conteudo, tagF)
            
            QuebraTagsXML = Mid(XML_Conteudo, a + Len(tagI), (b - a) - Len(tagI))
        Else
            QuebraTagsXML = Mid(XML_Conteudo, a + Len(tagI))
        End If
    End If
Exit Function
trata_Erro:
'MsgBox tagI & tagF & " Não encontrado no xml", vbExclamation
End Function
Function GravaCupomReferenciado(Id_nf As String, cod_mod As String, ECF_FAB As String, _
    ECF_CX As String, num_doc As String, dt_doc As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into sped_cupom_referenciado("
    strSQL = strSQL & "id_nf,"
    strSQL = strSQL & "cod_mod,"
    strSQL = strSQL & "ecf_fab ,"
    strSQL = strSQL & "ecf_cx,"
    strSQL = strSQL & "num_doc,"
    strSQL = strSQL & "dt_doc)values("
    strSQL = strSQL & Id_nf & ","
    strSQL = strSQL & "'" & cod_mod & "',"
    strSQL = strSQL & "'" & ECF_FAB & "',"
    strSQL = strSQL & "'" & ECF_CX & "',"
    strSQL = strSQL & "'" & num_doc & "',"
    strSQL = strSQL & F_GravaData(dt_doc) & ")"
    
    Conn.Execute strSQL
    GravaCupomReferenciado = True
trata_Erro:
    
    
End Function
Function AlteraCupomReferenciado(id As String, Id_nf As String, cod_mod As String, ECF_FAB As String, _
    ECF_CX As String, num_doc As String, dt_doc As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update sped_cupom_referenciado set "
    strSQL = strSQL & "id_nf=" & Id_nf & ", "
    strSQL = strSQL & "cod_mod='" & cod_mod & "', "
    strSQL = strSQL & "ecf_fab='" & ECF_FAB & "',"
    strSQL = strSQL & "ecf_cx='" & ECF_CX & "',"
    strSQL = strSQL & "num_doc='" & num_doc & "',"
    strSQL = strSQL & "dt_doc=" & F_GravaData(dt_doc) & " "
    strSQL = strSQL & " where id=" & id
    
    Conn.Execute strSQL
    AlteraCupomReferenciado = True
trata_Erro:
    
    
End Function

Function GravaInformacaoComplementarNF(Id_nf As String, codigo As String)
    Dim strSQL As String
    strSQL = "insert into sped_nf_ic(id_nf,codigo)values(" & Id_nf & "'" & codigo & "')"
    Conn.Execute strSQL
End Function
Function ExcluiInformacaoComplementarNF(Id_nf As String, codigo As String)
    Dim strSQL As String
    strSQL = "delete from sped_nf_ic"
    strSQL = strSQL & " where id_nf=" & Id_nf & " and codigo='" & codigo & "'"
    Conn.Execute strSQL
End Function
Function EfetivaSpedNf(id As String, Optional alterapreco As Boolean = True) As Boolean
    
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    
    strSQL = "select tp.pv_pro,n.*,tp.cod_pro,i.preco,i.custo,i.qtd_convert as quant,i.cod_pro as codigo from sped_nf as n left join sped_nf_itens as i on n.id=i.id_nf "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on i.cod_pro=tp.cod_pro "
    strSQL = strSQL & " where n.id=" & id
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "NF com id: " & id & " não encontrada!", vbExclamation
            EfetivaSpedNf = False
            Exit Function
        End If
        Do While Not .EOF
            
            If !pv_pro < !preco Then
                        strSQL = "update " & TAB_PRODUTOS
                        strSQL = strSQL & " set pv_pro=" & F_GravaDec(!preco) & "  "
                        strSQL = strSQL & " where cod_pro='" & !codigo & "'"
                        Conn.Execute strSQL
                    

                        
                        P_AltPreco !codigo, !preco, !pv_pro
                    End If
            strSQL = "update " & TAB_PRODUTOS & " set "
            If !tipo_entrada = "01" Then
                    
                strSQL = strSQL & " ef_pro=ef_pro+" & F_GravaDec(!quant) & ","
            End If
            strSQL = strSQL & "cr_pro=" & F_GravaDec(!custo) & ", "
            strSQL = strSQL & "er_pro=er_pro+" & F_GravaDec(!quant) & " "
            
            
            
            strSQL = strSQL & " where cod_pro='" & !codigo & "'"
            Conn.Execute strSQL
                    
                    
            If bUsaEstoqueSeparado = True Then
                        
                If F_BuscaUmCampo("cod_pro", "estoque", "cod_emp='" & !cod_emp & "' and cod_pro='" & !codigo & "'") = "" Then
                    strSQL = "Insert into estoque(cod_emp,cod_pro,er,ef)values("
                    strSQL = strSQL & "'" & !cod_emp & "',"
                    strSQL = strSQL & "'" & !codigo & "',"
                    strSQL = strSQL & F_GravaDec(!quant) & ", "
                    strSQL = strSQL & F_GravaDec(!quant) & ") "
                    Conn.Execute strSQL
                Else
                    strSQL = "update estoque set "
                    strSQL = strSQL & "cod_emp='" & !cod_emp & "',"
                    strSQL = strSQL & "cod_pro='" & !codigo & "',"
                    
                    strSQL = strSQL & "er=er+" & F_GravaDec(!quant) & " "
                    
                    
                    If !tipo_entrada = "01" Then
                        strSQL = strSQL & ",ef=ef+" & F_GravaDec(!quant) & "  "
                    End If
                    strSQL = strSQL & " where cod_emp='" & !cod_emp & "' and cod_pro='" & !codigo & "'"
                    Conn.Execute strSQL
                End If
                        
                        
                    
            End If
            
            
            
            .MoveNext
        Loop
    End With

    strSQL = "update sped_nf set status=4 where id=" & id
    Conn.Execute strSQL
    MsgBox "Nota Fiscal efetivada!", vbExclamation
End Function

Function GravaSpedNfEnergia(IND_OPER As String, ind_emit As String, cod_mod As String, cod_part As String, cod_sit As String, _
    SER As String, SUBSER As String, COD_CONS As String, num_doc As String, vl_doc As String, _
    dt_doc As String, DT_E_S As String, VL_DESC As String, VL_FORN As String, _
    VL_SERV_NT As String, VL_TERC As String, VL_DA As String, VL_BC_ICMS As String, _
    vl_icms As String, VL_ICMS_ST As String, COD_INF As String, vl_pis As String, _
    vl_cofins As String, TP_LIGACAO As String, COD_GRUPO_TENSAO As String) As Boolean

    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    strSQL = "Insert into sped_nf_energia(IND_OPER,  IND_EMIT, COD_PART, COD_MOD,  COD_SIT,  SER,  SUB,  COD_CONS,  "
    strSQL = strSQL & " NUM_DOC,  DT_DOC,  DT_E_S,  VL_DOC, VL_DESC,  VL_FORN,  VL_SERV_NT,  VL_TERC,  VL_DA,  "
    strSQL = strSQL & " VL_BC_ICMS,  VL_ICMS,  VL_ICMS_ST,  COD_INF,  VL_PIS,  VL_COFINS,  TP_LIGACAO,  "
    strSQL = strSQL & " COD_GRUPO_TENSAO)values("
    
    strSQL = strSQL & "'" & IND_OPER & "',"
    strSQL = strSQL & "'" & ind_emit & "',"
    strSQL = strSQL & "'" & cod_mod & "',"
    strSQL = strSQL & "'" & cod_part & "',"
    strSQL = strSQL & "'" & cod_sit & "',"
    strSQL = strSQL & "'" & SER & "',"
    strSQL = strSQL & "'" & SUBSER & "',"
    strSQL = strSQL & "'" & COD_CONS & "',"
    strSQL = strSQL & "'" & num_doc & "',"
    strSQL = strSQL & F_GravaData(dt_doc) & "',"
    strSQL = strSQL & F_GravaData(DT_E_S) & "',"
    strSQL = strSQL & F_GravaDec(vl_doc) & ","
    strSQL = strSQL & F_GravaDec(VL_DESC) & ","
    strSQL = strSQL & F_GravaDec(VL_FORN) & ","
    strSQL = strSQL & F_GravaDec(VL_SERV_NT) & ","
    strSQL = strSQL & F_GravaDec(VL_TERC) & ","
    strSQL = strSQL & F_GravaDec(VL_DA) & ","
    strSQL = strSQL & F_GravaDec(VL_BC_ICMS) & ","
    strSQL = strSQL & F_GravaDec(vl_icms) & ","
    strSQL = strSQL & F_GravaDec(VL_ICMS_ST) & ","
    strSQL = strSQL & "'" & COD_INF & "',"
    strSQL = strSQL & F_GravaDec(vl_pis) & ","
    strSQL = strSQL & F_GravaDec(vl_cofins) & ","
    strSQL = strSQL & "'" & TP_LIGACAO & "',"
    strSQL = strSQL & "'" & COD_GRUPO_TENSAO & "')"
    
    Conn.Execute strSQL
    GravaSpedNfEnergia = True
    Exit Function
trata_Erro:
    GravaSpedNfEnergia = False
End Function
Function AlteraSpedNfEnergia(id As String, IND_OPER As String, ind_emit As String, cod_part As String, cod_mod As String, cod_sit As String, _
    SER As String, SUBSER As String, COD_CONS As String, num_doc As String, _
    dt_doc As String, DT_E_S As String, vl_doc As String, VL_DESC As String, VL_FORN As String, _
    VL_SERV_NT As String, VL_TERC As String, VL_DA As String, VL_BC_ICMS As String, _
    vl_icms As String, VL_ICMS_ST As String, COD_INF As String, vl_pis As String, _
    vl_cofins As String, TP_LIGACAO As String, COD_GRUPO_TENSAO As String) As Boolean

    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    strSQL = "update sped_nf_energia set "
    
    strSQL = strSQL & "IND_OPER='" & IND_OPER & "',"
    strSQL = strSQL & "IND_EMIT='" & ind_emit & "',"
    strSQL = strSQL & "COD_MOD='" & cod_mod & "',"
    strSQL = strSQL & "COD_part='" & cod_part & "',"
    strSQL = strSQL & "COD_SIT='" & cod_sit & "',"
    strSQL = strSQL & "SER='" & SER & "',"
    strSQL = strSQL & "SUBSER='" & SUBSER & "',"
    strSQL = strSQL & "COD_CONS='" & COD_CONS & "',"
    strSQL = strSQL & "NUM_DOC='" & num_doc & "',"
    strSQL = strSQL & "dt_doc=" & F_GravaData(dt_doc) & "',"
    strSQL = strSQL & "dt_e_s=" & F_GravaData(DT_E_S) & "',"
    strSQL = strSQL & "VL_doc=" & F_GravaDec(vl_doc) & ","
    strSQL = strSQL & "VL_desc=" & F_GravaDec(VL_DESC) & ","
    strSQL = strSQL & "VL_forn=" & F_GravaDec(VL_FORN) & ","
    strSQL = strSQL & "VL_serv_nt=" & F_GravaDec(VL_SERV_NT) & ","
    strSQL = strSQL & "VL_TERC=" & F_GravaDec(VL_TERC) & ","
    strSQL = strSQL & "VL_DA=" & F_GravaDec(VL_DA) & ","
    strSQL = strSQL & "VL_BC_ICMS=" & F_GravaDec(VL_BC_ICMS) & ","
    strSQL = strSQL & "VL_ICMS=" & F_GravaDec(vl_icms) & ","
    strSQL = strSQL & "VL_ICMS_ST=" & F_GravaDec(VL_ICMS_ST) & ","
    strSQL = strSQL & "cod_inf='" & COD_INF & "',"
    strSQL = strSQL & "VL_PIS=" & F_GravaDec(vl_pis) & ","
    strSQL = strSQL & "VL_COFINS=" & F_GravaDec(vl_cofins) & ","
    strSQL = strSQL & "tp_ligacao='" & TP_LIGACAO & "',"
    strSQL = strSQL & "cod_grupo_tensao='" & COD_GRUPO_TENSAO & "'"
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL
    AlteraSpedNfEnergia = True
    Exit Function
trata_Erro:
    AlteraSpedNfEnergia = False
End Function

Function CFOPEntradaOuSaida(CFOP As String) As Single
    Dim esquerda As Single
    esquerda = F_GravaDec(Left(CFOP, 1))
    
    
    If esquerda < 1 Or esquerda > 7 Then
        CFOPEntradaOuSaida = -1
    ElseIf esquerda > 4 Then
        CFOP = 1
    Else
        CFOP = 0
    End If
    
End Function
Function ConverteCfop(CFOP As String) As String
    
    Dim direita As String
    Dim esquerda As Single
    direita = Right(CFOP, 3)
    
    esquerda = F_GravaDec(Left(CFOP, 1))
    
    If esquerda < 1 Or esquerda > 7 Then
        ConverteCfop = ""
        Exit Function
    End If
    
    If esquerda > 4 Then
        esquerda = esquerda - 4
    Else
        esquerda = esquerda + 4
    End If
    ConverteCfop = CStr(esquerda) & direita
End Function
Sub ControlarPrecoItem(Id_nf As String, id_item As String)
    frmCalculoPrecoItem.Id_nf = Id_nf
    frmCalculoPrecoItem.id_item = id_item
    frmCalculoPrecoItem.Show 1
End Sub
Function GravaNfeItensPreco(id_item As String, pg As String, pf As String, dia As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into nfe_itens_preco(id_item,pg,pf,dia)values("
    strSQL = strSQL & id_item & ","
    strSQL = strSQL & F_GravaDec(pg) & ","
    strSQL = strSQL & F_GravaDec(pf) & ","
    strSQL = strSQL & F_GravaDataTempo(dia) & ")"
    Exit Function
trata_Erro:
    MostraErro
    Conn.Execute strsqin
End Function
Function SelectNFs(dataI As String, dataF As String, Optional cod_part As String = "", Optional status As String = "", Optional tipo_op As String = "E") As String
    Dim strSQL As String
    strSQL = "select tf.nome, "
    strSQL = strSQL & " te.* "
    strSQL = strSQL & " from sped_nf as te left join sped_participante as tf  "
    strSQL = strSQL & " on tf.codigo=te.COD_PART "
    strSQL = strSQL & " where 1>0 "
    If status <> "" Then
        strSQL = strSQL & " and status='" & Left(cmbxStatus.Text, 1) & "'"
    End If
    If tipo_op = "E" Then
        strSQL = strSQL & " and ind_oper='0' "
    Else
        strSQL = strSQL & " and ind_oper='1' "
    End If

    If dataI <> "" Then
        strSQL = strSQL & " and dt_e_s  >=" & F_GravaData(txtbxDataIni.Text)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and dt_e_s  <=" & F_GravaData(txtbxdataFin.Text)
    End If
    
    SelectNFs = strSQL
    
    
End Function
Sub VisualizarSpedNF(id As String)
    
        
        frmEntradaSped.tag = id
        frmEntradaSped.bVisualizar = True
        frmEntradaSped.Show 1
        
    
End Sub
Sub P_PreencheSpedCFOP(cmbx As Object)
    P_Preenche cmbx, "sped_cfop", "concat(codigo,' - ',descricao)"
End Sub



Function GravaConferenciaEntrada(id_item As String, qtd As String, Optional id As String = "") As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    If id = "" Then
        strSQL = "insert into conferencia_entrada(id_item,qtd,cod_usu)values("
        strSQL = strSQL & "'" & id_item & "',"
        strSQL = strSQL & F_GravaDec(qtd) & ","
        strSQL = strSQL & "'" & strCod_Usu & "')"
        Conn.Execute strSQL
    Else
        strSQL = "update conferencia_entrada set qtd="
        strSQL = strSQL & F_GravaDec(qtd) & ","
        strSQL = strSQL & " cod_usu='" & strCod_Usu & "' ,"
        strSQL = strSQL & " data=now() "
        strSQL = strSQL & " where id=" & id
        Conn.Execute strSQL
    End If
    GravaConferenciaEntrada = True
trata_Erro:
    
End Function

