Attribute VB_Name = "Cheques"
Option Explicit
Type cheque
    seq_che As String 'caso seja um ch já gravado
    num_che As String
    cod_bco As String
    cco_che As String
    age_che As String
    nome_che As String
    cnpj_che As String
    Venc_che As String
    Valor_che As Currency
    cpf_che As String
    cod_loc As String
    nome_loc As String
End Type
Type chePag
    num_che As String
    cod_con As String
    Venc_che As String
    Valor_che As Currency
End Type

Function F_GravaChNovo(seq_doc As String, num_che As String, cod_bco As String, _
    age_che As String, cco_che As String, nome_che As String, _
    Valor_che As String, Venc_che As String, cod_ter As String, _
    tipo_che As String, num_ven As String, seq_itd As String, Parc_Che As String, seq_cai As String, cpf_che As String) As String
            
    Dim strSQL As String
    strSQL = "insert into tab_cheques(Seq_Doc,num_che,cod_bco,age_che,cco_che,"
    strSQL = strSQL & "valor_che,venc_che,cod_ter,tipo_Che,num_ven,seq_che,Parc_Che,nome_che,cpf_che,status_Che,hora_che)Values("
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & "'" & num_che & "',"
    strSQL = strSQL & "'" & cod_bco & "',"
    strSQL = strSQL & "'" & age_che & "',"
    strSQL = strSQL & "'" & cco_che & "',"
        
    strSQL = strSQL & F_GravaDec(Valor_che) & ","
    strSQL = strSQL & F_GravaData(Venc_che) & ","
    strSQL = strSQL & "'" & cod_ter & "',"
    strSQL = strSQL & "'" & tipo_che & "',"
    strSQL = strSQL & "'" & num_ven & "',"
    
    
    F_GravaChNovo = MaxKeySQL("Tab_cheques", "seq_che", "", "000000#")
    strSQL = strSQL & "'" & F_GravaChNovo & "', "
    strSQL = strSQL & Parc_Che & ", "
    strSQL = strSQL & "'" & nome_che & "',"
    strSQL = strSQL & "'" & cpf_che & "',1,"
    strSQL = strSQL & "'" & Time & "')"
    
    banco.Execute strSQL
    
End Function
Function F_GravaCh(seq_doc As String, num_che As String, cod_bco As String, _
    age_che As String, cco_che As String, nome_che As String, _
    Valor_che As String, Venc_che As String, cod_ter As String, _
    tipo_che As String, num_ven As String, seq_itd As String, Parc_Che As String, seq_cai As String, cpf_che As String, hora_che) As String
            
    Dim strSQL As String
    strSQL = "insert into tab_cheques1(Seq_Doc,num_che,cod_bco,age_che,cco_che,"
    strSQL = strSQL & "valor_che,venc_che,cod_ter,tipo_Che,num_ven,seq_che,Parc_Che,nome_che,cpf_che,hora_che,status_Che)Values("
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & "'" & num_che & "',"
    strSQL = strSQL & "'" & cod_bco & "',"
    strSQL = strSQL & "'" & age_che & "',"
    strSQL = strSQL & "'" & cco_che & "',"
        
    strSQL = strSQL & F_GravaDec(Valor_che) & ","
    strSQL = strSQL & F_GravaData(Venc_che) & ","
    strSQL = strSQL & "'" & cod_ter & "',"
    strSQL = strSQL & "'" & tipo_che & "',"
    strSQL = strSQL & "'" & num_ven & "',"
    
    
    F_GravaCh = MaxKeySQL("tab_cheques1", "seq_che", "", "000000#")
    strSQL = strSQL & "'" & F_GravaCh & "', "
    strSQL = strSQL & Parc_Che & ", "
    strSQL = strSQL & "'" & nome_che & "',"
    strSQL = strSQL & "'" & cpf_che & "',"
    strSQL = strSQL & "'" & hora_che & "',"
    strSQL = strSQL & " 1)"
    
    banco.Execute strSQL
    
End Function

Function F_GravaChePagNovo(seq_doc As String, num_che As String, cod_con As String, _
    Valor_che As String, Venc_che As String, cod_ter As String, _
    Parc_Che As String) As String

    On Error GoTo trata_Erro
            
    Dim strSQL As String
    strSQL = "insert into tab_cheques(cod_con,Seq_Doc,num_che,cod_bco,age_che,cco_che,nome_Che,"
    strSQL = strSQL & "valor_che,venc_che,cod_ter,tipo_Che,num_ven,seq_che,Parc_Che,status_che,hora_che)Values("
    strSQL = strSQL & "'" & cod_con & "',"
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & "'" & num_che & "',"
    strSQL = strSQL & "'" & "',"
    strSQL = strSQL & "'" & "',"
    strSQL = strSQL & "'" & "',"
    strSQL = strSQL & "'" & "',"
    
    strSQL = strSQL & F_GravaDec(Valor_che) & ","
    strSQL = strSQL & F_GravaData(Venc_che) & ","
    strSQL = strSQL & "'" & cod_ter & "',"
    strSQL = strSQL & "'P',"
    strSQL = strSQL & "'',"
    
    
    F_GravaChePagNovo = MaxKeySQL("Tab_cheques", "seq_che", "", "000000#")
    strSQL = strSQL & "'" & F_GravaChePagNovo & "', "
    strSQL = strSQL & Parc_Che & ",1,"
    strSQL = strSQL & "'" & Time & "') "
    banco.Execute strSQL
    Exit Function
trata_Erro:
    F_GravaChePagNovo = ""
End Function

Function F_GravaChePag(seq_doc As String, num_che As String, cod_con As String, _
    Valor_che As String, Venc_che As String, cod_ter As String, _
    Parc_Che As String, hora_che) As String

    On Error GoTo trata_Erro
            
    Dim strSQL As String
    strSQL = "insert into tab_cheques1(cod_con,Seq_Doc,num_che,cod_bco,age_che,cco_che,nome_Che,"
    strSQL = strSQL & "valor_che,venc_che,cod_ter,tipo_Che,num_ven,seq_che,Parc_Che,hora_che,status_che)Values("
    strSQL = strSQL & "'" & cod_con & "',"
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & "'" & num_che & "',"
    strSQL = strSQL & "'" & "',"
    strSQL = strSQL & "'" & "',"
    strSQL = strSQL & "'" & "',"
    strSQL = strSQL & "'" & "',"
    
    strSQL = strSQL & F_GravaDec(Valor_che) & ","
    strSQL = strSQL & F_GravaData(Venc_che) & ","
    strSQL = strSQL & "'" & cod_ter & "',"
    strSQL = strSQL & "'P',"
    strSQL = strSQL & "'',"
    
    
    F_GravaChePag = MaxKeySQL("tab_cheques1", "seq_che", "", "000000#")
    strSQL = strSQL & "'" & F_GravaChePag & "', "
    
    strSQL = strSQL & Parc_Che & ","
    strSQL = strSQL & "'" & hora_che & "',"
    strSQL = strSQL & "1)"
    banco.Execute strSQL
    Exit Function
trata_Erro:
    F_GravaChePag = ""
End Function

Function F_selectChPag(tipo_ter As String, Optional cod_ter As String = "", Optional dataI As String = "", _
    Optional dataF As String = "", Optional num_che As String = "", Optional histdoc As String = "", _
    Optional status_che As String = "1", Optional cod_con As String = "", _
    Optional seq_che As String = "", _
    Optional seq_doc As String = "", Optional num_doc As String = "", _
    Optional sOrderBy As String = "venc_che") As String
    
Dim strSQL As String
    
    Dim sTab As String
    
    Dim sCod As String
    Dim sNome As String
    Dim sCPF As String
    
    If tipo_ter = "" Then
        tipo_ter = "F"
    End If
    If tipo_ter = "F" Then
        sTab = "tab_fornecedores"
        sCod = "cod_for"
        sNome = "Nome_for"
        sCPF = "Cpf_for"
    ElseIf tipo_ter = "R" Then
        sTab = "tab_representantes"
        sCod = "cod_rep"
        sNome = "Nome_rep"
        sCPF = "Cpf_Rep"
    ElseIf tipo_ter = "U" Then 'funcionário
        sTab = "tab_funcionarios"
        sCod = "cod_fun"
        sNome = "Nome_fun"
        sCPF = "cpf_fun"
    End If


    strSQL = "select tc." & sNome & " as nome,tcc.nome_con, tx.*, td.valor_Doc,td.parc_doc,td.num_doc,td.data_doc,  "
        strSQL = strSQL & " if(status_che='1','Aberto',if(status_che='2','Baixado','DEVOLVIDO')) as status "
        strSQL = strSQL & " from " & sTab & "  as tc,tab_contas as tcc, tab_doc as td,tab_cheques as tx  "
        strSQL = strSQL & " where tc." & sCod & "=td.cod_ter and td.seq_doc=tx.seq_doc "
        strSQL = strSQL & " and tcc.cod_con=tx.cod_con "
        strSQL = strSQL & " and td.tipo_Doc='P' "
        If status_che <> "" Then
        strSQL = strSQL & " and status_che='" & Left(status_che, 1) & "' "
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and tx.venc_che < " & F_GravaData(dataF)
    End If
    If dataI <> "" Then
            strSQL = strSQL & " and tx.venc_che > " & F_GravaData(dataI)
        
    End If
        
    If cod_ter <> "" Then
            strSQL = strSQL & " and td.cod_ter='" & cod_ter & "'"
    End If
    If cod_con <> "" Then
        strSQL = strSQL & " and tx.cod_con='" & cod_con & "' "
    End If
    
    'If nome_che <> "" Then
    '    strSQL = strSQL & " and nome_che like '" & Nome & "%'"
    '
    'End If
    If num_che <> "" Then
        strSQL = strSQL & " and num_che='" & num_che & "' "
    End If
    If num_doc <> "" Then
        strSQL = strSQL & " and num_doc='" & num_doc & "' "
    End If
    If seq_che <> "" Then
        strSQL = strSQL & " and tx.seq_che='" & seq_che & "' "
    End If
    If seq_doc <> "" Then
        strSQL = strSQL & " and tx.seq_doc='" & seq_doc & "' "
    End If
    'If cod_loc <> "" Then
    '    strSQL = strSQL & " and tx.cod_loc='" & cod_loc & "'  "
    'End If
    If sOrderBy <> "" Then
        strSQL = strSQL & " order by " & sOrderBy
    End If
        
    F_selectChPag = strSQL
End Function

Function F_SelectChRec(tipo_ter As String, Optional cod_cli As String = "", Optional dataI As String = "", _
    Optional dataF As String = "", Optional num_che As String = "", Optional histdoc As String = "", _
    Optional status_che As String = "1", Optional cpf_che As String = "", _
    Optional nome_che As String = "", Optional seq_che As String = "", _
    Optional seq_doc As String = "", Optional num_doc As String = "", _
    Optional cod_loc As String = "", Optional sOrderBy As String = "venc_che", Optional sLocal As String = "", _
    Optional bVenc_che As Boolean = True) As String
    Dim strSQL As String
    
    Dim sTab As String
    
    Dim sCod As String
    Dim sNome As String
    Dim sCPF As String
    
    If tipo_ter = "" Then
        tipo_ter = "C"
    End If
    If tipo_ter = "C" Then
        sTab = "tab_clientes"
        sCod = "cod_cli"
        sNome = "Nome_cli"
        sCPF = "Cpf_cli"
    ElseIf tipo_ter = "R" Then
        sTab = "tab_representantes"
        sCod = "cod_rep"
        sNome = "Nome_rep"
        sCPF = "Cpf_Rep"
    ElseIf tipo_ter = "U" Then 'funcionário
        sTab = "tab_funcionarios"
        sCod = "cod_fun"
        sNome = "Nome_fun"
        sCPF = "cpf_fun"
    End If
    
    
    
        
    
    
        strSQL = "select tl.nome_loc,esp_doc,td.seq_doc,historico_doc,td.valor_Doc as valor_doc,td.parc_doc,td.num_doc,td.data_doc, tc." & sNome & " as nome, tx.*, "
        strSQL = strSQL & "if(status_che='1','Aberto',if(status_che='2','Baixado','DEVOLVIDO')) as status,concat(parc_che,'/',parc_doc) as parcela "
        strSQL = strSQL & " from " & sTab & " as tc, tab_doc1 as td,tab_cheques1 as tx,tab_localche as tl  "
        strSQL = strSQL & " where tc." & sCod & "=td.cod_ter and td.seq_doc=tx.seq_doc "
        strSQL = strSQL & " and tx.tipo_che='R' "
        strSQL = strSQL & " and tl.cod_loc=tx.cod_loc "
        'strSQL = strSQL & " and tipo_ter='" & tipo_ter & "' "
        
    If status_che <> "" Then
        strSQL = strSQL & " and status_che='" & Left(status_che, 1) & "' "
    End If
    If dataF <> "" Then
        If bVenc_che = True Then
            strSQL = strSQL & " and tx.venc_che <= " & F_GravaData(dataF)
        Else
            strSQL = strSQL & " and data_doc <= " & F_GravaData(dataF)
        End If
    End If
    If dataI <> "" Then
        If bVenc_che = True Then
            strSQL = strSQL & " and tx.venc_che >= " & F_GravaData(dataI)
        Else
            strSQL = strSQL & " and data_doc >= " & F_GravaData(dataI)
        End If
    End If
    
    
    
    If cod_cli <> "" Then
            strSQL = strSQL & " and td.cod_ter='" & cod_cli & "'"
    End If
    
    If cpf_che <> "" Then
        strSQL = strSQL & " and cpf_che ='" & F_FormataCPF(cpf_che) & "' "
    End If
    If nome_che <> "" Then
        strSQL = strSQL & " and nome_che like '" & nome_che & "%'"
        
    End If
    If num_che <> "" Then
        strSQL = strSQL & " and num_che='" & num_che & "' "
    End If
    If num_doc <> "" Then
        strSQL = strSQL & " and num_doc='" & num_doc & "' "
    End If
    If seq_che <> "" Then
        strSQL = strSQL & " and tx.seq_che='" & seq_che & "' "
    End If
    If seq_doc <> "" Then
        strSQL = strSQL & " and tx.seq_doc='" & seq_doc & "' "
    End If
    If cod_loc <> "" Then
        strSQL = strSQL & " and tx.cod_loc='" & cod_loc & "'  "
    End If
    If sLocal <> "" Then
        strSQL = strSQL & " and tl.cod_loc ='" & sLocal & "' "
    End If
    If sOrderBy <> "" Then
        strSQL = strSQL & " order by " & sOrderBy
    End If
    
    F_SelectChRec = strSQL
End Function
Function F_BuscaChequeTer(busca As String) As cheque
    frmHistoricoChe.txtbxCPF.Text = busca
    frmHistoricoChe.Tag = "x"
    frmHistoricoChe.Show 1
    If frmHistoricoChe.Tag <> "" Then
        With frmHistoricoChe.msflxgrd
            .Col = 0: F_BuscaChequeTer.num_che = .Text
            .Col = 1: F_BuscaChequeTer.cod_bco = Left(.Text, 3)
            .Col = 2: F_BuscaChequeTer.age_che = .Text
            .Col = 3: F_BuscaChequeTer.cco_che = .Text
            .Col = 4: F_BuscaChequeTer.Valor_che = .Text
            .Col = 5: F_BuscaChequeTer.Venc_che = .Text
            .Col = 6: F_BuscaChequeTer.nome_che = .Text
            .Col = 8: F_BuscaChequeTer.cpf_che = .Text
            .Col = 9: F_BuscaChequeTer.seq_che = .Text
        End With
    End If
    Unload frmHistoricoChe
End Function

Function F_DadosCheque(bVen As Boolean, Optional ven As String = "", Optional nome As String = "", Optional Num As String = "", _
            Optional cco As String = "", Optional age As String = "", Optional bco As String = "", _
            Optional Valor As String = "0,00", Optional sCPF As String = "", Optional JAM As Currency = 0) As cheque
    
    
    With frmDadosChe
        .txtbxAgencia.Text = age
        .txtbxContaCorrente.Text = cco
        .txtbxNome_Che.Text = nome
        .txtbxNum_Che.Text = Num
        .txtbxValor_Che.Text = Valor
        .txtbxVencimento.Text = ven
        .txtbxCPF.Text = sCPF
        .lblValor.caption = FormatVal(Valor)
        .lblJam.caption = FormatVal(JAM)
        If bco <> "" Then
            P_LocalizaOperacao bco, .cmbxBanco
        End If
        If bVen = False Then
            .txtbxVencimento.Enabled = False
        End If
        On Error Resume Next
        
        .Show 1
        If .Tag <> "" Then
            F_DadosCheque.age_che = .txtbxAgencia.Text
            F_DadosCheque.cco_che = .txtbxContaCorrente.Text
            F_DadosCheque.nome_che = .txtbxNome_Che.Text
            F_DadosCheque.num_che = .txtbxNum_Che.Text
            F_DadosCheque.Valor_che = .txtbxValor_Che.Text
            F_DadosCheque.Venc_che = .txtbxVencimento.Text
            F_DadosCheque.cod_bco = Left(.cmbxBanco.Text, 3)
            F_DadosCheque.cpf_che = .txtbxCPF.Text
            
        End If
    End With
        Unload frmDadosChe
End Function
Function F_DadosChePag(Optional ven As String = "", Optional Num As String = "", _
            Optional cod_con As String = "", Optional sVal As String = "", Optional lLock As Boolean = False) As chePag
    
    With frmDadosChPag
        
            .txtbxVencimento.Text = ven
            .txtbxNum_Che.Text = Num
            .txtbxValor_Che.Text = sVal
        
        If cod_con <> "" Then
            P_LocalizaOperacao cod_con, .cmbxContas
        End If
        
        If lLock = True Then
            .txtbxVencimento.Enabled = False
            .txtbxValor_Che.Enabled = False
            .cmbxContas.Enabled = False
        End If
        On Error Resume Next
        
        .Show 1
        
            F_DadosChePag.cod_con = Left(.cmbxContas.Text, 3)
            F_DadosChePag.num_che = .txtbxNum_Che.Text
            F_DadosChePag.Valor_che = CCur(.txtbxValor_Che.Text)
            F_DadosChePag.Venc_che = .txtbxVencimento.Text
        
    End With
        Unload frmDadosChPag
        
    
    
End Function


Function F_DevolveChe(seq_che As String) As Boolean
    frmDevChe.Tag = seq_che
    frmDevChe.Show 1
    If frmDevChe.Tag <> "" Then
        F_DevolveChe = True
    Else
        F_DevolveChe = False
    End If
    
    
End Function
Sub P_HistCheCPF(cpf_cli As String)
    frmHistoricoChe.Tag = cpf_cli
    frmHistoricoChe.txtbxCPF.Text = cpf_cli
    frmHistoricoChe.Show 1
    
    
    
End Sub
Function F_HistCheCPF(CPF As String) As cheque

    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro
    strSQL = "select distinct cco_che,cod_bco,age_che,cpf_che,nome_che from tab_cheques1 where cpf_che='" & CPF & "' "
    strSQL = strSQL & " order by seq_che "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 1 Then
            F_HistCheCPF.age_che = !age_che
            F_HistCheCPF.cco_che = !cco_che
            F_HistCheCPF.cod_bco = !cod_bco
            
            F_HistCheCPF.cpf_che = !cpf_che
            F_HistCheCPF.nome_che = !nome_che
            
        End If
        If .RecordCount > 1 Then
            Do While Not .EOF
                strSQL = !cod_bco & Chr(9)
                strSQL = strSQL & F_BuscaUmCampo("nome_bco", "tab_bancos", "cod_bco='" & !cod_bco & "'") & Chr(9)
                strSQL = strSQL & !age_che & Chr(9)
                strSQL = strSQL & !cco_che & Chr(9)
                strSQL = strSQL & !nome_che
                frmBuscaCheCPF.msflxgrd.AddItem strSQL
                
                .MoveNext
            Loop
            On Error Resume Next
            frmBuscaCheCPF.Show 1
            frmBuscaCheCPF.msflxgrd.Col = 0
            
            F_HistCheCPF.cod_bco = frmBuscaCheCPF.msflxgrd.Text
            frmBuscaCheCPF.msflxgrd.Col = 2
            F_HistCheCPF.age_che = frmBuscaCheCPF.msflxgrd.Text
            frmBuscaCheCPF.msflxgrd.Col = 3
            F_HistCheCPF.cco_che = frmBuscaCheCPF.msflxgrd.Text
            
            frmBuscaCheCPF.msflxgrd.Col = 4
            F_HistCheCPF.nome_che = frmBuscaCheCPF.msflxgrd.Text
            Unload frmBuscaCheCPF
            
        End If
        .Close
    End With
    
Exit Function
trata_Erro:
P_MostraErro

End Function

