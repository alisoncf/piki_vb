Attribute VB_Name = "NFe_DAO"
Function GravaNFCeVenda(num_ven As String, Tipo_Icms As TipoICMS) As String
    
    
    Dim strSQL As String
    Dim sID_Nfe As String
    Dim sId_Item As String
    Dim sNum As String
    Dim cNF As String
    Dim nNf As String
    Dim rPesq As New ADODB.Recordset
    Dim rItem As New ADODB.Recordset
    Dim fPag As indPag
    Dim iDest As idDest
    Dim chave As String
    Dim dv As String
    Dim sQtd As String
    Dim TipoICMS As String
    Dim cTotalProd As Currency
    Dim cTotalDesc As Currency
    Dim tPag As String
    Dim msg As String
    Dim cod_cli As String
    strSQL = "select tv.*, tc.*,tc.cod_cid as cidcli,tf.tipo_cai,tf.cod_tip,tf.tpag from " & TAB_VENDAS
    strSQL = strSQL & " as tv inner join " & TAB_CLIENTES & " as tc "
    strSQL = strSQL & " on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " inner join tab_formaspag as tf on tv.cod_for = tf.cod_For "
    
    
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            
        End If
        cNF = AleatorioNFe
        nNf = UltimaNFe
        
        chave = GeraChaveNfE("65", "1", nNf, cNF, Normal)
        dv = Right(chave, 1)
        If !cod_tip = "01" Then
            fPag = indPag.avista
        ElseIf !cod_tip = "02" Then
            fPag = indPag.aprazo
        Else
            fPag = indPag.Outros
        End If
        If !uf_cli = "GO" Or !uf_cli = "" Then
            iDest = operacao_interna
        Else
            iDest = operacao_interestadual
        End If
        
        sID_Nfe = NFe.NFE_GravaNFE(!num_ven, chave, IIf(!uf_cli <> "", !uf_cli, "GO"), cNF, "Venda de Produtos", CStr(fPag), "65", "1", nNf, CStr(Now), _
        CStr(Now), tpNf.saida, iDest, !cod_cid, danf_nfc_e, Normal, CStr(dv), ambiente, Nfe_normal, consumidor_final, _
          Operacao_presencial, Aplicativo_contribuinte, "2017", "", "")
        
        If sID_Nfe = "" Then
            MsgBox "erro na geração da NFe", vbExclamation
            Exit Function
        End If
        
        
        If Not GravaEmitenteComEndereco(sID_Nfe) Then
            MsgBox "Erro "
        End If
        cod_cli = !cod_cli
        
        If cod_cli <> "-1" Then
            If Not NFe.GravaNFe_Dest(sID_Nfe, FormataCNPJNfe(!cnpj_Cli), FormataCPFNfe(!cpf_cli), "", !nome_Cli) Then
                MsgBox "Erro"
            End If
            If Not NFe.GravaNfe_enderdest(sID_Nfe, Format(!Logradouro_cli), !NUMERO_cli, !COMPLEMENTO_cli, !bairro_cli, _
                !cod_cid, !cidade_cli, !uf_cli, !cep_cli, strCodigoPaisIBGE, strPais, !telefone_cli, !ie_Cli, _
                 "2", "", "", !email_Cli) Then
            
                MsgBox "Erro"
            End If
        Else
        
        End If
        
        
        strSQL = "select tp.*,ti.* from " & TAB_ITVENDAS & " as ti inner join " & TAB_PRODUTOS & " as tp "
        strSQL = strSQL & " on tp.cod_pro=ti.cod_pro and ti.num_ven='" & num_ven & "' "
        abreObjetoAdo rItem, strSQL
        With rItem
            Do While Not .EOF
                cTotalProd = cTotalProd + (!unit_ite * !qtd_ite)
                cTotalDesc = cTotalDesc + (!desconto_ite)
                sQtd = CStr(!qtd_ite)
                sId_Item = NFe.GravaNFe_Item(sID_Nfe, !cod_pro, !ean, !nome_pro, !ncm_pro, "", "", !cfop_pro, _
                    !unidade_pro, sQtd, CStr(!unit_ite), !unit_ite * !qtd_ite, !ean, !unidade_pro, sQtd, _
                    CStr(!unit_ite), "0", "0", !desconto_ite, "0", vProd_Sim_copoe_total_nota)
                
                If NFe.GravaNfe_item_imposto(sId_Item, sID_Nfe, CStr(!final_ite * crcTxFederativa_Par / 100), _
                    "ICMSSN102", "PISOutr", "COFINSOutr", "") Then
                    
                
                    If Not NFe.GravaNfe_item_imposto_icmssn102(sId_Item, sID_Nfe, !cod_ori, ICMSSN102) Then
                        MsgBox "erro na gravação do item_Imposto_icmssn102"
                        
                    End If
                    If Not NFe.GravaNfe_item_imposto_PISOutr(sId_Item, sID_Nfe, !cst_pis, "0", "0", "0", "0", "0") Then
                        MsgBox "erro na gravação do item_Imposto_pisOutr"
                        
                    End If
                    If Not NFe.GravaNfe_item_imposto_COFINSOutr(sId_Item, sID_Nfe, !cst_cofins, "0", "0", "0", "0", "0") Then
                        MsgBox "erro na gravação do item_Imposto_pisOutr"
                        
                    End If
                                    
                Else
                    MsgBox "erro na gravação do item_Imposto"
                End If
                
                
                .MoveNext
            Loop
        End With
        If rPesq!uf_cli = "GO" Then
            
            If Not GravaNfe_IcmsTot(sID_Nfe, "0", "0", "0", "0", "0", CStr(cTotalProd), "0", _
                        "0", CStr(cTotalDesc), "0", "0", "0", "0", "0", _
                        CStr((cTotalProd - cTotalDesc)), CStr((cTotalProd - cTotalDesc) * crcTxFederativa_Par / 100)) Then
            End If
                        
        Else
                If Not GravaNfe_IcmsTot(sID_Nfe, "0", "0", "0", "0", "0", CStr(cTotalProd), "0", _
                        "0", CStr(cTotalDesc), "0", "0", "0", "0", "0", _
                        CStr(!valor_ven), CStr(!valor_ven * crcTxFederativa_Par / 100)) Then
                End If
        End If
            
        If Not GravaNFe_Transp(sID_Nfe, sem_frete) Then
                MsgBox "erro no transporte"
        End If
            
        If Not GravaNFe_Pag(sID_Nfe, !tPag, CStr(cTotalProd - cTotalDesc), "", "", "", "") Then
            MsgBox "erro no pag"
        End If
            
            
        msg = "Tributos incidentes Lei Federal 12.741/12 - Total R$ " & _
                FormatVal(!final_ven * crcTxFederativa_Par / 100) & " Federal (0,00) Estadual (0,00) Municipal"
            
            
            
            If Not GravaNFe_InfAdc(sID_Nfe, "", msg) Then
                MsgBox "erro infAdc"
            End If
            
        .Close
       End With
    
    
        GravaNFCeVenda = sID_Nfe
        
'       MsgBox "Nota gravada!", vbInformation
    
    
    
    

End Function

Function GravaEmitenteComEndereco(id_NFe As String) As Boolean
    If Not NFe.GravaNFe_Emit(id_NFe, F_SoNumero(strCNPJ_Emp), "", strNome_Emp) Then
            GravaEmitenteComEndereco = False
    ElseIf Not NFe.GravaNfe_enderEmit(id_NFe, strLogradouro_Emp, strNumero_Emp, _
        strComplemento_Emp, strBairro_Emp, strCodigoMunicipioIBGE, strCidade_Emp, strUF_Emp, _
        strCEP_Emp, strCodigoPaisIBGE, strPais, F_SoNumero(strFone_Emp), F_SoNumero(strIE_Emp), _
        F_SoNumero(strIEST_Emp), F_SoNumero(strIM_Emp), F_SoNumero(strCNAE_Emp), strCRT_emp) Then
            GravaEmitenteComEndereco = False
    Else
        GravaEmitenteComEndereco = True
    End If
End Function


