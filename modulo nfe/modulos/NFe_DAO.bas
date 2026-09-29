Attribute VB_Name = "NFe_DAO"

Function GravaNFCeVenda(num_ven As String, tipo_icms As String, _
    Optional natOp As String = "Venda de produtos", Optional modelo As modelo_Nfe, _
    Optional IdentificaCliente As Boolean = True, Optional Tipo_Operacao As tipo_operacao_venda = venda) As String
    
    
    Dim strSQL As String
    Dim sId_NFe As String
    Dim sId_Item As String
    Dim sNum As String
    Dim cNF As String
    Dim nNF As String
    Dim rPesq As New ADODB.Recordset
    Dim rItem As New ADODB.Recordset
    Dim fPag As indPag
    Dim iDest As idDest
    Dim Chave As String
    Dim dv As String
    Dim sQtd As String
    Dim sTipoIcms As String
    Dim cTotalProd As Currency
    Dim cTotalProdCredIcms As Currency
    Dim cTotalDesc As Currency
    Dim tPag As String
    Dim msg As String
    Dim cod_cli As String
    Dim sCfop As String
    Dim sCST As String
    Dim cpf As String
    Dim CNPJ As String
    Dim sId As String
    Dim nomecli As String
    Dim CFOP As String
    Dim finalItem As Currency
    Dim isCombustivel As Boolean
    Dim bRevenda As Boolean
    Dim ncm As String
    Dim pICMS As Single

    Dim vICMS As Single
    Dim vICMSSTRet As Single
    Dim vBCICMSSTRet As Single
    Dim vOutrosItem As Single
    Dim vOutros As Single
    Dim vBC As Single
    Dim vBCItem As Single
    Dim vBCItemPIS As Single
    Dim vBCPis As Single
    Dim vBCItemCOFINS As Single
    Dim vBCCOFINS As Single
    
    Dim vicmsItem As Single
    Dim pRedBC As Single
    Dim vPIS As Single
    Dim vCOFINS As Single
    Dim pPIS As Single
    Dim pCOFINS As Single
    If strCod_Cai = "" Then
        strCod_Cai = "01" 'a série é o número do caixa
    End If
    
        P_Status "Gerando Nota Fiscal..."
        
        If Tipo_Operacao = tipo_operacao_venda.Devolucao Then
            natOp = "Devolução de Mercadoria"
        ElseIf Tipo_Operacao = Remessa Then
            natOp = "Remessa de Venda de Mercadoria"
        
        End If
        
    
    
    
'    On Error GoTo trata_Erro
    If modelo = 0 Then
        modelo = NFCe_65
    End If
    
    If TAB_CLIENTES = "" Then
        TAB_CLIENTES = "tab_clientes"
    End If
    If TAB_VENDAS = "" Then
        TAB_VENDAS = "tab_vendas"
    End If
    If TAB_ITVENDAS = "" Then
        TAB_ITVENDAS = "tab_itvendas"
    End If
    
    
    
    sId = BuscaIdPelaVenda(num_ven)
    If sId <> "" Then
        
        
        
        strSQL = buscaStatusNFe(sId)
        If strSQL = "5" Then
            
        
        ElseIf strSQL = "2" Then
            If MsgBox("Já existe uma NFe para esta venda que já foi transmitida. Deseja realmente criar outra NFe?", vbYesNo + vbQuestion) = vbNo Then
                P_Status
                Exit Function
            End If
        Else
            If MsgBox("Já existe uma NFe para esta venda que encontra-se em processamento. Deseja realmente sobrepor essa NFe?", vbYesNo + vbQuestion) = vbYes Then
                'chave = BuscaChaveNFe(sId)
'                MsgBox "Operação impossível!", vbExclamation
 '               Exit Function
            Else
                P_Status
                Exit Function
            End If
        End If
        
    End If
    
    
    strSQL = "select tv.*, tc.*,tc.cod_cid as cidcli,tc.cod_cid as nomecid, uf_cid,tf.tipo_cai,tf.cod_tip,tf.tpag from " & TAB_VENDAS
    strSQL = strSQL & " as tv inner join " & TAB_CLIENTES & " as tc "
    strSQL = strSQL & " on tv.cod_cli=tc.cod_cli "
    
    strSQL = strSQL & " left join tab_formaspag as tf on tv.cod_for = tf.cod_For "
    strSQL = strSQL & " left join " & TAB_CIDADES & " as tcid on tc.cod_cid=tcid.cod_cid "
    
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    'strSQL = strSQL & " and tipo_operacao=1"
    abreObjetoAdo rPesq, strSQL
    P_Status "buscando venda..."
    
    
    
    
    With rPesq
        If .RecordCount = 0 Then
            
        End If
        
        If modelo = NFe_55 And !tipo_cli = "J" And !indiedest_cli = "1" Then
            bRevenda = True
        End If
        
        InicializaEmpresa strCod_Emp
        
        cNF = AleatorioNFe
        
        
        
        'If F_BuscaUmCampo("id", "nfe", "cNF='" & cNF & "'") <> "" Then
            
            'MsgBox "Já existe um CNF com este número!" & cNF, vbExclamation
            'Exit Function
        '    cNF = AleatorioNFe
        'End If
        
        nNF = UltimaNFe(modelo, !cod_emp)
        If Chave = "" Then
            Chave = GeraChaveNfE(CStr(modelo), strSerie_Nfe, nNF, cNF, Normal)
        End If
        dv = Right(Chave, 1)
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
        
       P_Status "gerando Nfe..."
        sId_NFe = NFe.NFE_GravaNFE(!num_ven, Chave, IIf(!uf_cli <> "", !uf_cli, "GO"), cNF, natOp, CStr(fPag), CStr(modelo), Int(strSerie_Nfe), nNF, CStr(Now), _
        CStr(Now), IIf(Tipo_Operacao = tipo_operacao_venda.Devolucao, tpNf.entrada, tpNf.saida), iDest, !cod_cid, IIf(modelo = NFCe_65, tpImp.danf_nfc_e, tpImp.danf_retrato), Normal, CStr(dv), ambiente, Nfe_normal, IIf(modelo = NFCe_65, consumidor_final, nao), _
           IIf(modelo = "65", Operacao_presencial, Operacao_presencial), Aplicativo_contribuinte, "2017", "", "", "1", "0", !cod_emp)
        
        If sId_NFe = "" Then
            MsgBox "erro na geração da NFe", vbExclamation
            Exit Function
        End If
        
        P_Status "buscando Nfe... emitente"
        If Not GravaEmitenteComEndereco(sId_NFe) Then
            MsgBox "Erro "
        End If
        cod_cli = !cod_cli
        sTipoIcms = TiraNull(!tipoicms_cli)
       ' sCFOP = TiraNull(!cfop_cli)
        
        
        If sTipoIcms = "101" Then
            tipo_icms = "ICMSSN101"
        End If
        
        
        
        If cod_cli = "-1" Or cod_cli = "0" Or cod_cli = "00000" Or cod_cli = "000000" Then
            nomecli = !nomecli_ven
            
            strSQL = F_SoNumero(TiraNull(!cpfcli_ven))
            If Len(strSQL) = 11 Then
                cpf = strSQL
            ElseIf Len(strSQL) = 14 Then
                CNPJ = strSQL
            Else
                
            
            End If
        Else
            If !tipo_cli = "J" Then
                CNPJ = F_SoNumero(!cnpj_Cli)
            Else
                cpf = F_SoNumero(!cpf_cli)
            End If
            nomecli = !nome_cli
        End If
            
            
         If IdentificaCliente = True Then
            If cpf <> "" Or CNPJ <> "" Then
                If Not NFe.GravaNFe_Dest(sId_NFe, CNPJ, cpf, "", nomecli) Then
                    MsgBox "Erro"
                    Exit Function
                End If
                If Not NFe.GravaNfe_enderdest(sId_NFe, "", "", "", "", _
                    "", "", "", "", "", "", "", "", _
                    "1", "", "", "") Then
        
                    MsgBox "Erro gravando endereço do Destinatário"
                End If
            
            Else
            If Not NFe.GravaNFe_Dest(sId_NFe, CNPJ, cpf, "", nomecli) Then
                    MsgBox "Erro"
                    Exit Function
                End If
        P_Status "buscando Nfe... destinatário"
                If Not NFe.GravaNfe_enderdest(sId_NFe, Format(!Logradouro_cli), Format(!numero_cli), Format(!complemento_cli), Format(!bairro_cli), _
                    !cod_cid, Format(!nomecid), Format(!uf_cli), Format(!cep_cli), strCodigoPaisIBGE, strPais, Format(!telefone_cli), Format(!ie_Cli), _
                      !indiedest_cli, "", "", Format(!email_cli)) Then
            
                    MsgBox "Erro gravando endereço do Destinatário"
                End If
            End If
        
        Else
            
            If Not NFe.GravaNFe_Dest(sId_NFe, "", "", "", "Venda a Consumidor") Then
                MsgBox "Erro"
                Exit Function
            End If
            If Not NFe.GravaNfe_enderdest(sId_NFe, "", "", "", "", _
                "", "", "", "", "", "", "", "", _
                 "1", "", "", "") Then
        
                MsgBox "Erro gravando endereço do Destinatário"
            End If
        End If
        
        If TAB_PRODUTOS = "" Then
            TAB_PRODUTOS = "tab_produtos"
        End If
        
        If tipo_icms = "ICMSSN900" Then
            pICMS = 0.17
            
            If IdentificaCliente And rPesq!uf_cli <> "GO" And rPesq!indiedest_cli = "1" Then
                pICMS = 0.12
            End If
                
                
        Else
            pcims = 0
            vBC = 0
                
        End If
        
        Dim descProporcional As Double
        Dim proporcao As Double
        Dim qItFinal As Double
        P_Status "buscando Nfe... itens..."
        If Not bUsaEstoqueSeparado Then
            strSQL = "select tp.*,ti.*,cofins.tipo as tipo_cofins,pis.tipo as tipo_pis,ipi.tipo as tipo_ipi,"
            strSQL = strSQL & " icms.*,icms.tipo as tipo_icms,comb.cProdAnp,comb.descAnp "
            strSQL = strSQL & " from " & TAB_ITVENDAS & " as ti inner join " & TAB_PRODUTOS & " as tp "
            strSQL = strSQL & " on tp.cod_pro=ti.cod_pro and ti.num_ven='" & num_ven & "' "
            strSQL = strSQL & " left join sped_cst_cofins as cofins on cofins.codigo=tp.cst_cofins "
            strSQL = strSQL & " left join sped_cst_pis as pis on pis.codigo=tp.cst_pis "
            strSQL = strSQL & " left join sped_cst_ipi as ipi on ipi.codigo=tp.cst_ipi "
            strSQL = strSQL & " left join nfe_tipo_icms as icms on icms.codigo=tp.cst_icms "
            strSQL = strSQL & " left join produtos_comb as comb on tp.cod_pro=comb.cod_pro"
        Else
            strSQL = "select tp.cod_ori,tp.cod_gen,tp.ncm_pro,tp.unidade_pro, tp.cod_pro,tp.nome_pro,tp.codfab_pro,tp.ean_pro,"
            strSQL = strSQL & "e.cfop_revenda as cfop_revenda,"
            strSQL = strSQL & "e.cfop_consumidor as cfop_consumidor,"
            strSQL = strSQL & "e.cst_icms_revenda,e.cst_icms_consumidor,e.cst_pis,"
            strSQL = strSQL & " e.cst_cofins as cst_cofins,e.picms,e.ppis,e.pcofins,"
            strSQL = strSQL & "ti.*,cofins.tipo as tipo_cofins,pis.tipo as tipo_pis,ipi.tipo as tipo_ipi,"
            strSQL = strSQL & " icms.*,icms.tipo as tipo_icms,comb.cProdAnp,comb.descAnp "
            strSQL = strSQL & " from " & TAB_ITVENDAS & " as ti inner join " & TAB_PRODUTOS & " as tp "
            strSQL = strSQL & " on tp.cod_pro=ti.cod_pro and ti.num_ven='" & num_ven & "' "
            strSQL = strSQL & " inner join estoque as e on tp.cod_pro=e.cod_pro and e.cod_emp=" & !cod_emp
            strSQL = strSQL & " left join sped_cst_cofins as cofins on cofins.codigo=e.cst_cofins "
            strSQL = strSQL & " left join sped_cst_pis as pis on pis.codigo=e.cst_pis "
            strSQL = strSQL & " left join sped_cst_ipi as ipi on ipi.codigo=e.cst_ipi "
            strSQL = strSQL & " left join nfe_tipo_icms as icms on icms.codigo=e.cst_icms_consumidor "
            strSQL = strSQL & " left join produtos_comb as comb on tp.cod_pro=comb.cod_pro"
        
        End If
        
        abreObjetoAdo rItem, strSQL
        Dim devolv As String
        With rItem
            Do While Not .EOF
                vicmsItem = 0
                devolv = TiraNull(!devolvido_ite)
                If !qtd_ite - CCur((FormatVal(!devolvido_ite))) > 0 Then
                    If IsNull(!devolvido_ite) Then
                        qItFinal = !qtd_ite
                    Else
                        qItFinal = !qtd_ite - CCur(!devolvido_ite)
                    End If
                    'cTotalProd = cTotalProd + (!unit_ite * (!qtd_ite - CCur(F_AjustaValor(TiraNull(!devolvido_ite)))))
                    cTotalProd = cTotalProd + !final_ite
                    'proporcao = ((!qtd_ite - CCur(F_AjustaValor(TiraNull(!devolvido_ite)))) / !qtd_ite)
                    proporcao = qItFinal / !qtd_ite
                    
                    finalItem = !final_ite * proporcao
                    
                    'descProporcional = (!Desconto_ite * proporcao)
                    descProporcional = 0
                    cTotalDesc = cTotalDesc + descProporcional
                    sQtd = CStr(!qtd_ite - CCur(F_AjustaValor(TiraNull(!devolvido_ite))))
                    
                    'sId_Item = NFe.GravaNFe_Item(sId_NFe, !cod_pro, Format(!ean_pro), !desc_ite, TiraNull(!ncm_pro), "", "", sCFOP, _
                     '   !unidade_pro, sQtd, CStr(!unit_ite), !unit_ite * (!qtd_ite - CCur(F_AjustaValor(TiraNull(!devolvido_ite)))), TiraNull(!ean_pro), !unidade_pro, sQtd, _
                      '  CStr(!unit_ite), "0", "0", CStr(descProporcional), "0", vProd_Sim_copoe_total_nota, "", "")
                    isCombustivel = Not IsNull(!cProdAnp)
                    If Tipo_Operacao = venda Then
                        If bRevenda Then
                            sCfop = !cfop_revenda
                            If Not bUsaEstoqueSeparado Then
                                sCST = !cst_revenda
                            Else
                                sCST = !cst_icms_revenda
                            End If
                            tipo_icms = sCST
                        Else
                            If Not bUsaEstoqueSeparado Then
                                sCfop = !cfop_pro
                            Else
                                sCfop = !cfop_consumidor
                            End If
                            
                            If strCRT_emp = SIMP_NAC_EXCESSO_SUBLIMITE_RECEITA_BRUTA Then
                                sCST = "60"
                                tipo_icms = "60"
                            Else
                                If Not bUsaEstoqueSeparado Then
                                
                                    sCST = !cst_icms
                                Else
                                    sCST = !cst_icms_consumidor
                                End If
                                tipo_icms = sCST
                            End If
                            
                        End If
                    Else
                        If Tipo_Operacao = tipo_operacao_venda.retorno_remessa Then
                            sCfop = "1949"
                            sCST = "102"
                        ElseIf Tipo_Operacao = tipo_operacao_venda.Devolucao Then
                            sCfop = "6904"
                            sCST = "102"
                        Else
                            sCfop = "5904"
                            sCST = "102"
                        End If
                    End If
                    vProd = !final_ite
                    If TiraNull(!ncm_pro) = "" Then
                        ncm = ""
                    End If
                    
                    sId_Item = NFe.GravaNFe_Item(sId_NFe, !cod_pro, TiraNull(!ean_pro), !desc_ite, TiraNull(!ncm_pro), "", "", sCfop, _
                        !unidade_pro, sQtd, CStr(!final_ite / qItFinal), !final_ite, TiraNull(!ean_pro), !unidade_pro, sQtd, _
                        CStr(!final_ite / qItFinal), "0", "0", CStr(descProporcional), CStr(vOutros), vProd_Sim_copoe_total_nota, "", "")
                    
                    
                    If sId_Item = "" Then
                        MsgBox "Erro na gravação do item " & !nome_pro, vbExclamation
                        Exit Function
                    End If
                    
                    If isCombustivel Then
                        If NFe.GravaNfeItemComb(sId_Item, sId_NFe, !cProdAnp, !descAnp, rPesq!uf_cli) Then
                        End If
                    End If
                    
                    If strCRT_emp = SIMPLES_NACIONAL Then
                        tipo_icms = sCST
                    ElseIf strCRT_emp = RECEITA_NORMAL Then
                        tipo_icms = sCST
                        pICMS = !pICMS / 100
                        pPIS = !pPIS / 100
                        pCOFINS = !pCOFINS / 100
                        sTipoIcms = sCST
                    ElseIf strCRT_emp = SIMP_NAC_EXCESSO_SUBLIMITE_RECEITA_BRUTA Then
                        If Not IsNull(!cProdAnp) Then
                            tipo_icms = sCST
                        Else
                            If rPesq!indiedest_cli = "1" And modelo = NFe_55 Then
                                tipo_icms = "20"
                                sTipoIcms = "20"
                            Else
                                tipo_icms = "00"
                                sTipoIcms = "00"
                                
                            End If
                            If isCombustivel Then
                                tipo_icms = "60"
                            End If
                            pICMS = 0.17
                        End If
                    Else
                        tipo_icms = sCST
                        sTipoIcms = sCST
                    End If
                    
                    If NFe.GravaNfe_item_imposto(sId_Item, sId_NFe, finalItem * pICMS, _
                          CStr(tipo_icms), TiraNull(!tipo_pis), TiraNull(!tipo_cofins), TiraNull(!tipo_ipi)) Then
                        
                        If tipo_icms = "ICMSSN102" Or tipo_icms = "102" Then
                            If Not NFe.GravaNfe_item_imposto_icmssn102(sId_Item, sId_NFe, TiraNull(!cod_ori), TipoICMS.ICMSSN102) Then
                                MsgBox "erro na gravação do item_Imposto_icmssn102"
                            End If
                        ElseIf tipo_icms = "ICMSSN101" Or tipo_icms = "101" Then
                            Dim aux As Single
                            aux = buscaNfeCredInfoAddPadrao("", TipoICMS.ICMSSN101)
                            If Not NFe.GravaNfe_item_imposto_icmssn101(sId_Item, sId_NFe, TiraNull(!cod_ori), CStr(TipoICMS.ICMSSN101), CStr(aux), CStr((!final_ite * (aux / 100)))) Then
                                MsgBox "erro na gravação do item_Imposto_icmssn101"
                            End If

                        ElseIf tipo_icms = "ICMSSN500" Or tipo_icms = "500" Then
                            If Not NFe.GravaNfe_item_imposto_icmssn500(sId_Item, sId_NFe, TiraNull(!cod_ori), TipoICMS.ICMSSN500, "0", "0") Then
                                MsgBox "erro na gravação do item_Imposto_icmssn500"
                            End If
                        ElseIf tipo_icms = "ICMSSN900" Or tipo_icms = "900" Then
                            If Not NFe.GravaNfe_item_imposto_icmssn900(sId_Item, sId_NFe, TiraNull(!cod_ori), TipoICMS.ICMSSN900, CStr(finalItem), CStr(pICMS * 100), _
                                pICMS * finalItem, "0", "0", "0", "0", "0", "0", "0", "0") Then
                                MsgBox "erro na gravação do item_Imposto_icmssn900"
                            End If
                        ElseIf tipo_icms = "ICMS00" Or tipo_icms = "00" Then
                            vBCItem = (finalItem)
                            vicmsItem = vBCItem * pICMS
                            vBC = vBC + vBCItem
                                
                            vICMS = vICMS + vicmsItem
                            
                            If NFe.GravaNfe_item_imposto_icms00(sId_Item, sId_NFe, TiraNull(!cod_ori), "3", CStr(vBCItem), CStr(pICMS * 100), _
                                 CStr(vBCItem * pICMS)) Then
                            Else
                                MsgBox "erro na gravação do item_Imposto_icms00"
                            End If
                        ElseIf tipo_icms = "ICMS20" Or tipo_icms = "20" Then
                            
                            'pRedBC = 0.3529
                            pRedBC = buscaPRedBCICMS20
                            vBCItem = (finalItem) - (finalItem * pRedBC)
                            vicmsItem = vBCItem * pICMS

                            If NFe.GravaNfe_item_imposto_icms20(sId_Item, sId_NFe, TiraNull(!cod_ori), "3", CStr(pRedBC * 100), _
                                 CStr(vBCItem), CStr(pICMS * 100), _
                                CStr(vicmsItem), "0", "") Then
                                vBC = vBC + vBCItem
                                vICMS = vICMS + vicmsItem
                            Else
                                MsgBox "erro na gravação do item_Imposto_icms20"
                            End If
                        ElseIf tipo_icms = "ICMS60" Or tipo_icms = "60" Then

                            If NFe.GravaNfe_item_imposto_icms60(sId_Item, sId_NFe, TiraNull(!cod_ori), "0", "0") Then
                               ' vBCICMSSTRet = vBCICMSSTRet + (finalItem)
                               ' vICMSSTRet = vICMSSTRet + (finalItem * pICMS)
                            Else
                                MsgBox "erro na gravação do item_Imposto_icms00"
                            End If
                        End If


                        If !tipo_pis = "PISOutr" Then

                            If Not NFe.GravaNfe_item_imposto_PISOutr(sId_Item, sId_NFe, !cst_pis, "0", "0", "0", "0", "0") Then
                                MsgBox "erro na gravação do item_Imposto_pisOutr"

                            End If
                        ElseIf !tipo_pis = "PISNT" Then
                            If Not NFe.GravaNfe_item_imposto_PISNT(sId_Item, sId_NFe, !cst_pis) Then
                                MsgBox "erro na gravação do item_Imposto_pisOutr"

                            End If
                        ElseIf !tipo_pis = "PISAliq" Then
                            
                            vBCItem = !final_ite
                            vBCItemPIS = !final_ite - vicmsItem
                            If Not NFe.GravaNfe_item_imposto_PISAliq(sId_Item, sId_NFe, !cst_pis, CStr(vBCItemPIS), pPIS * 100, pPIS * vBCItemPIS) Then
                                MsgBox "erro na gravação do item_Imposto_pisOutr"

                            End If
                            vPIS = vPIS + (pPIS * vBCItemPIS)
                        End If

                        If !tipo_cofins = "COFINSOutr" Then
                            If Not NFe.GravaNfe_item_imposto_COFINSOutr(sId_Item, sId_NFe, !cst_cofins, "0", "0", "0", "0", "0") Then
                                MsgBox "erro na gravação do item_Imposto_pisOutr"

                            End If
                        
                        ElseIf !tipo_cofins = "COFINSNT" Then
                            If Not NFe.GravaNfe_item_imposto_COFINSNT(sId_Item, sId_NFe, !cst_cofins) Then
                                MsgBox "erro na gravação do item_Imposto_pisOutr"

                            End If
                        ElseIf !tipo_cofins = "COFINSAliq" Then
                            vBCItem = !final_ite
                            vBCItemCOFINS = !final_ite - vicmsItem
                            If Not NFe.GravaNfe_item_imposto_COFINSAliq(sId_Item, sId_NFe, !cst_cofins, CStr(vBCItemCOFINS), pCOFINS * 100, pCOFINS * vBCItemCOFINS) Then
                                MsgBox "erro na gravação do item_Imposto_pisOutr"

                            End If
                            vCOFINS = vCOFINS + (pCOFINS * vBCItemCOFINS)
                        
                        End If

                        If TiraNull(!tipo_ipi) <> "" Then
                            If Not NFe.GravaNfe_item_imposto_IPI(sId_Item, sId_NFe, "", "", "", "", "999") Then
                                MsgBox "Erro na gravação do IPI"
                            End If

                            If TiraNull(!tipo_ipi) = "IPITrib" Then
                                If Not NFe.GravaNfe_item_imposto_IPITrib(sId_Item, sId_NFe, !CST_IPI, "0", "0", "0", "0", "0") Then
                                    MsgBox "Erro na gravação do IPITrib"
                                End If
                            ElseIf TiraNull(!tipo_ipi) = "IPINT" Then
                                If Not NFe.GravaNfe_item_imposto_IPINT(sId_Item, sId_NFe, "99") Then
                                    MsgBox "Erro na gravação do IPINT"
                                End If
                            End If


                        
                        End If
                    Else 'item imposto
                   
                        MsgBox "erro na gravação do item_Imposto"
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        
       
        If tipo_icms = "ICMSSN900" Then
            pICMS = 0.17
            vBC = cTotalProd - cTotalDesc
            If IdentificaCliente And rPesq!uf_cli <> "GO" And rPesq!indiedest_cli = "1" Then
                pICMS = 0.17
            End If
                
                
        Else
          '  pcims = 0
          '  vBC = 0
                
        End If
            
            
            
                If Not GravaNfe_IcmsTot(sId_NFe, CStr(vBC), CStr(vICMS), "0", "0", "0", CStr(cTotalProd), "0", _
                        "0", CStr(cTotalDesc), "0", "0", CStr(vPIS), CStr(vCOFINS), CStr(vOutros), _
                         CStr((cTotalProd - cTotalDesc)), CStr(vBC)) Then
                End If
                
            
            
'        Else
'                If Not GravaNfe_IcmsTot(sId_NFe, vBC, "0", "0", "0", "0", CStr(cTotalProd), "0", _
'                        "0", CStr(cTotalDesc), "0", "0", "0", "0", "0", _
'                        CStr(!valor_ven), CStr(!valor_ven * crcTxFederativa_Par / 100)) Then
'                End If
'        End If
            
        If Not GravaNFe_Transp(sId_NFe, sem_frete) Then
                MsgBox "erro no transporte"
        End If
        
        Dim i As Integer
        
        
        If Tipo_Operacao = tipo_operacao_venda.venda Then
        
            If !tPag = "03" Or !tPag = "04" Then
                strSQL = "select tv.*,tf.* from " & TAB_VENFIN & " as tv left join tab_financeiras as tf "
                strSQL = strSQL & " on tv.cod_fin=tf.cod_Fin "
                strSQL = strSQL & " where num_ven='" & !num_ven & "' "
                
                abreObjetoAdo rItem, strSQL
                With rItem
                    Do While Not .EOF
'                        For i = 1 To !parcelas_fin
'
'                            If Not GravaNFe_Pag(sId_NFe, rPesq!tPag, CStr((cTotalProd - cTotalDesc) / CCur(!parcelas_fin)), "", _
'                            !CNPJ, !tBand, TiraNull(!autorizacao_fin), TiraNull(!tpIntegra), "0") Then
'                                MsgBox "erro no pag"
'                            End If
'                        Next
                        If Not GravaNFe_Pag(sId_NFe, rPesq!tPag, CStr(!valor_fin), "", _
                            !CNPJ, !tBand, TiraNull(!autorizacao_fin), TiraNull(!tpIntegra), "0") Then
                                MsgBox "erro no pag"
                            End If
                        .MoveNext
                        
                        
                    Loop
                    .Close
                End With
            Else
                If Not GravaNFe_Pag(sId_NFe, Format(!tPag), CStr(cTotalProd - cTotalDesc), "", "", "", "", "", "0") Then
                        MsgBox "erro no pag"
                End If
            
            End If
        Else
            If Not GravaNFe_Pag(sId_NFe, "90", "0", "", "", "", "", "", "0") Then
                        MsgBox "erro no pag"
                End If
            
        End If
        
            
            
            
        
        
            
            
                msg = buscaInfoAddPadrao(sTipoIcms)
                Dim sufixo As String
                
                msg = Replace(msg, "@credicms", BuscaSomaCredIcms(sId_NFe, tipo_icms))
            
        If modelo = NFCe_65 And msg = "" Then
            msg = "Tributos incidentes Lei Federal 12.741/12 - Total R$ " & _
                FormatVal(!valor_ven * crcTxFederativa_Par / 100)
                msg = msg & CrLf & "Documento emitidido por ME ou EPP optante pelos Simples Nacional; Não gera direito a crédito fiscal de IPI. "
        
            
            
        End If
    
        
        
        
            
            
        If Not GravaNFe_InfAdc(sId_NFe, "", msg) Then
            MsgBox "erro infAdc"
        End If
            
        .Close
       End With
    
    
        GravaNFCeVenda = sId_NFe
        
'       MsgBox "Nota gravada!", vbInformation
    P_Status ""
    
    Exit Function
trata_Erro:
    MsgBox Err.Description, vbCritical

End Function

Function GravaEmitenteComEndereco(id_nfe As String) As Boolean
    If Not NFe.GravaNFe_Emit(id_nfe, F_SoNumero(strCNPJ_Emp), "", strNome_Emp, strFantasia_Emp) Then
            GravaEmitenteComEndereco = False
    ElseIf Not NFe.GravaNfe_enderEmit(id_nfe, strLogradouro_Emp, strNumero_Emp, _
        strComplemento_Emp, strBairro_Emp, strCodigoMunicipioIBGE, strCidade_Emp, strUF_Emp, _
        strCEP_Emp, strCodigoPaisIBGE, strPais, F_SoNumero(strFone_Emp), F_SoNumero(strIE_Emp), _
        F_SoNumero(strIEST_Emp), F_SoNumero(strIM_Emp), F_SoNumero(strCNAE_Emp), strCRT_emp) Then
            GravaEmitenteComEndereco = False
    Else
        GravaEmitenteComEndereco = True
    End If
End Function


Function GravaNfePedidoInut(Inicial As String, final As String, justificativa As String) As Boolean
    Dim strSQL As String
    strSQL = "insert into nfe_pedidoinut(inicial, final, datahora, status, "
    strSQL = strSQL & "cod_usu, justificativa,modelo, serie)values("
    strSQL = strSQL & "'" & Inicial & "',"
    strSQL = strSQL & "'" & final & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'1',"
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & justificativa & "',"
    strSQL = strSQL & "'" & modelo & "',"
    strSQL = strSQL & "'" & serei & "')"
    Conn.Execute strSQL
    
    GravaNfePedidoInut = True
    
    Exit Function
End Function

Function BuscaTabelaICMS(codigo As String) As String
    BuscaTabelaICMS = F_BuscaUmCampo("tabela", "nfe_tipo_icms", "codigo='" & codigo & "'")

End Function

Function GravaXML(id As String, xml As String) As Boolean
    Dim t As String
    t = "xml" & Year(Date)

    'if F_BuscaUmCampo("id","xml" & year(date)


End Function
Function buscaNfeCredInfoAddPadrao(id As String, tipo_icms As String) As String

    If id <> "" Then
        buscaNfeCredInfoAddPadrao = F_BuscaUmCampo("credicms", "nfe_infaddpadrao", "id='" & id & "'")
    End If
    If tipo_icms <> "" Then
        buscaNfeCredInfoAddPadrao = F_BuscaUmCampo("credicms", "nfe_infaddpadrao", "tipo_icms='" & tipo_icms & _
        "' and cod_emp='" & strCod_Emp & "'")
    End If
End Function
Function BuscaSomaCredIcms(id_nfe As String, Optional sufixo As String = "icmssn101") As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    BuscaSomaCredIcms = "0"
    If sufixo = "101" Then
        sufixo = "ICMSSN101"
    ElseIf sufixo = "201" Then
        sufixo = "ICMSSN201"
    End If
    
    
    strSQL = "select sum(vcredicmssn) as total from nfe_item_imposto_" & sufixo
    strSQL = strSQL & " where id_nfe=" & id_nfe
    abreObjetoAdo rPesq, strSQL
    BuscaSomaCredIcms = CStr(rPesq!total)
    FechaObjeto rPesq
    
    
trata_Erro:

End Function
Function buscaInfoAddPadrao(tipo_icms As String) As String
    
    
    If tipo_icms <> "" Then
        buscaInfoAddPadrao = F_BuscaUmCampo("texto", "nfe_infaddpadrao", "tipo_icms='" & tipo_icms & _
        "' and cod_emp='" & strCod_Emp & "'")
    End If
End Function


