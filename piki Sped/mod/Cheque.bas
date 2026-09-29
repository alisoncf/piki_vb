Attribute VB_Name = "Cheques"
Option Explicit
Public Enum tipoVencimentoCheque
    Venc_che = 1
    data_doc = 2
    baixa_che = 3
End Enum
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

Function F_GravaCh(seq_doc As String, num_che As String, cod_bco As String, _
    age_che As String, cco_che As String, nome_che As String, _
    Valor_che As String, Venc_che As String, cod_ter As String, _
    tipo_che As String, num_ven As String, seq_itd As String, Parc_Che As String, seq_cai As String, cpf_che As String, _
    Optional seq_che As String = "", Optional dep_che As String = "") As String
    If cod_ter = "XXX" Or cod_ter = "" Then
        cod_ter = "0"
    End If
    
    
            
    Dim strSQL As String
    strSQL = "insert into " & TAB_CHEQUES & "(Seq_Doc,num_che,cod_bco,age_che,cco_che,"
    strSQL = strSQL & "valor_che,venc_che,cod_ter,tipo_Che,num_ven,seq_che,Parc_Che,nome_che,cpf_che,status_Che,dep_che,hora_che)Values("
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
    
    If seq_che = "" Then
        F_GravaCh = MaxKeySQL(TAB_CHEQUES, "seq_che", "", "000000#")
    Else
        F_GravaCh = seq_che
    End If
    strSQL = strSQL & "'" & F_GravaCh & "', "
    strSQL = strSQL & Parc_Che & ", "
    strSQL = strSQL & "'" & nome_che & "',"
    strSQL = strSQL & "'" & cpf_che & "',1,"
    If dep_che = "" Then
        strSQL = strSQL & "NULL, "
    Else
        strSQL = strSQL & F_GravaData(dep_che) & ", "
    End If
    strSQL = strSQL & "'" & Left(Time, 5) & "')"
    
    conn.Execute strSQL
    
End Function

Function F_GravaChePag(seq_doc As String, num_che As String, cod_con As String, _
    Valor_che As String, Venc_che As String, cod_ter As String, _
    Parc_Che As String, Optional seq_che As String = "", Optional dep_che As String = "") As String

    On Error GoTo trata_erro
            
    Dim strSQL As String
    strSQL = "insert into tab_cheques(cod_con,Seq_Doc,num_che,cod_bco,age_che,cco_che,nome_Che,"
    strSQL = strSQL & "valor_che,venc_che,cod_ter,tipo_Che,num_ven,seq_che,Parc_Che,status_che,dep_che,hora_che)Values("
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
    
    If seq_che = "" Then
    F_GravaChePag = MaxKeySQL("Tab_cheques", "seq_che", "", "000000#")
    Else
        F_GravaChePag = seq_che
    End If
    strSQL = strSQL & "'" & F_GravaChePag & "', "
    strSQL = strSQL & Parc_Che & ",1,"
    If dep_che = "" Then
        strSQL = strSQL & "NULL, "
    Else
        strSQL = strSQL & "'" & dep_che & "', "
    End If
    strSQL = strSQL & "'" & Time & "') "
    conn.Execute strSQL
    Exit Function
trata_erro:
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
        sTab = TAB_FORNECEDORES
        sCod = "cod_for"
        sNome = "Nome_for"
        sCPF = "Cpf_for"
    ElseIf tipo_ter = "R" Then
        sTab = "tab_representantes"
        sCod = "cod_rep"
        sNome = "Nome_rep"
        sCPF = "Cpf_Rep"
    ElseIf tipo_ter = "U" Then 'funcionário
        sTab = TAB_FUNCIONARIOS
        sCod = "cod_fun"
        sNome = "Nome_fun"
        sCPF = "cpf_fun"
    End If


    strSQL = "select tc." & sNome & " as nome,tcc.nome_con, tx.*, td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc,  "
        strSQL = strSQL & " if(status_che='1','Aberto',if(status_che='2','Baixado','DEVOLVIDO')) as status "
        strSQL = strSQL & " from " & sTab & "  as tc,tab_contas as tcc,  " & TAB_DOC & "  as td,tab_cheques as tx  "
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
    Optional bVenc_che As tipoVencimentoCheque = Venc_che) As String
    Dim strSQL As String
    
    Dim sTab As String
    
    Dim sCod As String
    Dim sNome As String
    Dim sCPF As String
    
    If tipo_ter = "" Then
        tipo_ter = "C"
    End If
    If tipo_ter = "C" Then
        sTab = TAB_CLIENTES
        sCod = "cod_cli"
        sNome = "Nome_cli"
        sCPF = "Cpf_cli"
    ElseIf tipo_ter = "R" Then
        sTab = "tab_representantes"
        sCod = "cod_rep"
        sNome = "Nome_rep"
        sCPF = "Cpf_Rep"
    ElseIf tipo_ter = "U" Then 'funcionário
        sTab = TAB_FUNCIONARIOS
        sCod = "cod_fun"
        sNome = "Nome_fun"
        sCPF = "cpf_fun"
    End If
    
    
    
        
    
    
        strSQL = "select tl.nome_loc,esp_doc,td.seq_doc,historico_doc,td.valor_Doc as valor_doc,td.parcelas_doc,td.num_doc,td.data_doc, tc." & sNome & " as nome, tx.*, "
        strSQL = strSQL & "if(status_che='1','Aberto',if(status_che='2','Baixado','DEVOLVIDO')) as status,concat(parc_che,'/',parcelas_doc) as parcela "
        strSQL = strSQL & " from " & sTab & " as tc, " & TAB_DOC & " as td," & TAB_CHEQUES & " as tx left join " & TAB_LOCALCHE & " as tl on tx.cod_loc=tl.cod_loc "
        strSQL = strSQL & " where tc." & sCod & "=td.cod_ter and td.seq_doc=tx.seq_doc "
        strSQL = strSQL & " and tx.tipo_che='R' "
        'strSQL = strSQL & " and tl.cod_loc=tx.cod_loc "
        strSQL = strSQL & " and tipo_ter='" & tipo_ter & "' "
        
    If status_che <> "" Then
        strSQL = strSQL & " and status_che='" & Left(status_che, 1) & "' "
    End If
    If dataF <> "" Then
        If bVenc_che = Venc_che Then
            strSQL = strSQL & " and tx.venc_che <= " & F_GravaData(dataF)
        ElseIf bVenc_che = data_doc Then
            strSQL = strSQL & " and data_doc <= " & F_GravaData(dataF)
        Else
            strSQL = strSQL & " and dep_che <= " & F_GravaData(dataF)
        End If
    End If
    If dataI <> "" Then
        If bVenc_che = Venc_che Then
            strSQL = strSQL & " and tx.venc_che >= " & F_GravaData(dataI)
        ElseIf bVenc_che = data_doc Then
            strSQL = strSQL & " and data_doc >= " & F_GravaData(dataI)
        Else
            strSQL = strSQL & " and dep_che >= " & F_GravaData(dataI)
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
    frmHistoricoChe.txtbxCpf.Text = busca
    frmHistoricoChe.tag = "x"
    frmHistoricoChe.Show 1
    If frmHistoricoChe.tag <> "" Then
        With frmHistoricoChe.msflxgrd
            .col = 0: F_BuscaChequeTer.num_che = .Text
            .col = 1: F_BuscaChequeTer.cod_bco = Left(.Text, 3)
            .col = 2: F_BuscaChequeTer.age_che = .Text
            .col = 3: F_BuscaChequeTer.cco_che = .Text
            .col = 4: F_BuscaChequeTer.Valor_che = .Text
            .col = 5: F_BuscaChequeTer.Venc_che = .Text
            .col = 6: F_BuscaChequeTer.nome_che = .Text
            .col = 8: F_BuscaChequeTer.cpf_che = .Text
            .col = 9: F_BuscaChequeTer.seq_che = .Text
        End With
    End If
    Unload frmHistoricoChe
End Function





Function F_DadosCheque(bVen As Boolean, Optional ven As String = "", Optional nome As String = "", Optional Num As String = "", _
            Optional cco As String = "", Optional age As String = "", Optional bco As String = "", _
            Optional valor As String = "0,00", Optional sCPF As String = "", Optional cod_loc As String) As cheque
    
    
    With frmDadosChe
        .txtbxAgencia.Text = age
        .txtbxContaCorrente.Text = cco
        .txtbxNome_Che.Text = nome
        .txtbxNum_Che.Text = Num
        .txtbxValor_Che.Text = valor
        .txtbxVencimento.Text = ven
        .txtbxCpf.Text = sCPF
        .txtbxLocal.tag = cod_loc
        .txtbxLocal.Text = F_BuscaUmCampo("nome_loc", TAB_LOCALCHE, "cod_loc='" & cod_loc & "'")
        If bco <> "" Then
            .txtbxBanco.tag = bco
            .txtbxBanco.Text = bco
        End If
        If bVen = False Then
            .txtbxVencimento.Enabled = False
        End If
        On Error Resume Next
        
        .Show 1
        If .tag <> "" Then
            F_DadosCheque.age_che = .txtbxAgencia.Text
            F_DadosCheque.cco_che = .txtbxContaCorrente.Text
            F_DadosCheque.nome_che = .txtbxNome_Che.Text
            F_DadosCheque.num_che = .txtbxNum_Che.Text
            F_DadosCheque.Valor_che = .txtbxValor_Che.Text
            F_DadosCheque.Venc_che = .txtbxVencimento.Text
            F_DadosCheque.cod_bco = .txtbxBanco.tag
            F_DadosCheque.cpf_che = .txtbxCpf.Text
            F_DadosCheque.cod_loc = .txtbxLocal.tag
            F_DadosCheque.nome_loc = .txtbxLocal.Text
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


Function F_HistCheCPF(cpf As String) As cheque

    Dim stSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_erro
    
    strSQL = "select distinct cco_che,cod_bco,age_che,cpf_che,nome_che from " & TAB_CHEQUES & " where cpf_che='" & cpf & "' "
    strSQL = strSQL & " order by seq_che "
    abreObjetoAdo rPesq, strSQL
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
            frmBuscaCheCPF.msflxgrd.col = 0
            
            F_HistCheCPF.cod_bco = frmBuscaCheCPF.msflxgrd.Text
            frmBuscaCheCPF.msflxgrd.col = 2
            F_HistCheCPF.age_che = frmBuscaCheCPF.msflxgrd.Text
            frmBuscaCheCPF.msflxgrd.col = 3
            F_HistCheCPF.cco_che = frmBuscaCheCPF.msflxgrd.Text
            
            frmBuscaCheCPF.msflxgrd.col = 4
            F_HistCheCPF.nome_che = frmBuscaCheCPF.msflxgrd.Text
            Unload frmBuscaCheCPF
            
        End If
    End With
    
Exit Function
trata_erro:
P_MostraErro

End Function

Function F_DevolveChe(seq_che As String) As Boolean
    frmDevChe.tag = seq_che
    frmDevChe.Show 1
    If frmDevChe.tag <> "" Then
        F_DevolveChe = True
    Else
        F_DevolveChe = False
    End If
    
    
End Function

