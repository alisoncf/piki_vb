Attribute VB_Name = "sendMail"
Type Mail
    id As String
    nome As String
    email As String
    sch As String
    sendusing As String
    smtpserver As String
    smtpserverport As String
    smtpauthenticate As String
    sendusername As String
    sendpassword As String
    smtpusessl As String
    subject As String
    body As String
    
End Type
Function BuscaMailNome(nome As String) As Mail
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim m As Mail
    On Error GoTo trata_Erro
    strSQL = "select * from sendmail where nome='" & nome & "' "
    
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        m.body = !body
        m.email = !email
        m.id = !id
        m.sch = TiraNull(!sch)
        m.sendpassword = TiraNull(!sendpassword)
        m.sendusername = TiraNull(!sendusername)
        m.sendusing = TiraNull(!sendusing)
        m.smtpauthenticate = TiraNull(!smtpauthenticate)
        m.smtpserver = TiraNull(!smtpserver)
        m.smtpserverport = TiraNull(!smtpserverport)
        m.smtpusessl = TiraNull(!smtpusessl)
        m.subject = TiraNull(!subject)
        
        .Close
        BuscaMailNome = m
    End With
trata_Erro:
    
    
End Function
Function GravaMail(m As Mail) As Boolean
    On Error GoTo trata_Erro:
    Dim strSQL As String
    If m.id = "" Then
        strSQL = "insert into sendmail(nome,sch,email ,sendusing ,"
        strSQL = strSQL & " smtpserver, smtpserverport,  smtpauthenticate ,"
        strSQL = strSQL & " sendusername, sendpassword ,smtpusessl ,subject ,body )values("
        strSQL = strSQL & "'" & m.nome & "',"
        strSQL = strSQL & "'" & m.sch & "',"
        strSQL = strSQL & "'" & m.email & "',"
        strSQL = strSQL & "'" & m.sendusing & "',"
        strSQL = strSQL & "'" & m.smtpserver & "',"
        strSQL = strSQL & m.smtpserverport & ","
        strSQL = strSQL & "'" & m.smtpauthenticate & "',"
        strSQL = strSQL & "'" & m.sendusername & "',"
        strSQL = strSQL & "'" & m.sendpassword & "',"
        strSQL = strSQL & "'" & m.smtpusessl & "',"
        strSQL = strSQL & "'" & m.subject & "',"
        strSQL = strSQL & "'" & m.body & "')"
        conn.Execute strSQL
    Else
        strSQL = "replace into sendmail(id,nome,sch,email ,sendusing ,"
        strSQL = strSQL & " smtpserver, smtpserverport,  smtpauthenticate ,"
        strSQL = strSQL & " sendusername, sendpassword ,smtpusessl ,subject ,body )values("
        strSQL = strSQL & "" & m.id & ","
        strSQL = strSQL & "'" & m.nome & "',"
        strSQL = strSQL & "'" & m.sch & "',"
        strSQL = strSQL & "'" & m.email & "',"
        strSQL = strSQL & "'" & m.sendusing & "',"
        strSQL = strSQL & "'" & m.smtpserver & "',"
        strSQL = strSQL & m.smtpserverport & ","
        strSQL = strSQL & "'" & m.smtpauthenticate & "',"
        strSQL = strSQL & "'" & m.sendusername & "',"
        strSQL = strSQL & "'" & m.sendpassword & "',"
        strSQL = strSQL & "'" & m.smtpusessl & "',"
        strSQL = strSQL & "'" & m.subject & "',"
        strSQL = strSQL & "'" & m.body & "')"
        conn.Execute strSQL
    End If
    GravaMail = True

trata_Erro:
    
End Function
Sub SendMailGmail(sendusername As String, sendpassword As String, asendTo As String, subject As String, _
    msg As String, anexos() As String, Optional smtpserver As String = "smtp.gmail.com", _
    Optional smtpserverport As Integer = 465, Optional sender As String = "")
    On Error GoTo trata_Erro
    Set objMsg = CreateObject("CDO.Message")
    Set config = CreateObject("CDO.Configuration")
    
    
    sch = "http://schemas.microsoft.com/cdo/configuration/"
    config.Fields.item(sch & "sendusing") = 2
    config.Fields.item(sch & "smtpserver") = "smtp.gmail.com"
    config.Fields.item(sch & "smtpserverport") = "465"
    config.Fields.item(sch & "smtpauthenticate") = 1

    config.Fields.item(sch & "sendusername") = sendusername
    config.Fields.item(sch & "sendpassword") = sendpassword
    config.Fields.item(sch & "smtpusessl") = 1
   
   
    
    objMsg.To = asendTo
  
  
  
    objMsg.from = sendusername
    
    Dim i As Integer
    
    P_Status ""
    If IsObject(anexos) > 0 Then
        For i = LBound(anexos) To UBound(anexos)
            If anexos(i) <> "" Then
                If Dir(anexos(i)) = "" Then
                    P_Status "Aguardando arquivo..."
        
        
                    Sleep 5000
                    P_Status ""
                End If
                objMsg.AddAttachment anexos(i)
            End If
        Next i
    End If
    objMsg.subject = subject
    objMsg.HTMLBody = msg
    If sender = "" Then
        objMsg.sender = strFantasia_Emp
    End If
    config.Fields.Update
    
    Set objMsg.Configuration = config
    objMsg.sEnd
    Exit Sub
trata_Erro:
P_MostraErro
End Sub
