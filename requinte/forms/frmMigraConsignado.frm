VERSION 5.00
Begin VB.Form frmMigraConsignado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Plus"
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   735
      Left            =   8640
      TabIndex        =   8
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Migra apenas o arquivo"
      Height          =   735
      Left            =   3480
      TabIndex        =   7
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Migra Dados dos Consignados"
      Height          =   855
      Left            =   3480
      TabIndex        =   6
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton cmndbtnMigraTodos 
      Caption         =   "Migra Todos"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   3135
   End
   Begin VB.TextBox txtbxBancoMySQL 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Text            =   "req_dbf"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ok"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\desenv\requinte\db\temp"
      Top             =   5640
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   5355
      Left            =   0
      Pattern         =   "*.dbf"
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Banco mySQL"
      Height          =   195
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frmMigraConsignado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bOld As database
Dim bNew As database
Dim b As database
Private Type campo
    nome As String
    tipo As String
    tamanho As String
    
    
    
    
End Type







Private Sub cmndbtnMigraTodos_Click()
    Dim iCount As Integer
    
    For iCount = 0 To File1.ListCount - 1
        File1.ListIndex = iCount
        P_Status "Migrando arquivo " & File1.FileName & File1.ListIndex & " de " & File1.ListCount
        migraArquivo
                        
    
    Next
    Me.caption = "Migrado"
    
End Sub

Private Sub Command1_Click()
    
    Dim rPesq As New adodb.Recordset
    Dim r2 As New adodb.Recordset
    Dim strSQL As String
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    
    strSQL = "delete from tab_consignacoes"
    Conn.Execute strSQL
    
    
    strSQL = "delete from tab_estcons"
    Conn.Execute strSQL
        
        
        
    
    If Not F_AbreBancoODBC(b, txtbxBancoMySQL.Text) Then
        MsgBox "Erro.", vbExclamation
        Exit Sub
    End If
    
    strSQL = "show tables"
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        Do While Not .EOF
            strSQL = "show fields from " & rPesq.Fields(0).Value
            abreObjetoAdo r2, strSQL
            'Do While Not .EOF
                
                s1 = ""
                s2 = ""
                s3 = ""
                If r2.RecordCount > 3 Then
                    r2.MoveNext
                    s1 = LCase(r2.Fields(0).Value)
                    r2.MoveNext
                    s2 = LCase(r2.Fields(0).Value)
                    r2.MoveNext
                    s3 = LCase(r2.Fields(0).Value)
                End If
                If s1 = "codproe" And s2 = "qteeste" And s3 = "ultcome" Then
                    'MsgBox .Fields(0).Value
                    
                        migraCons .Fields(0).Value
                    
                    Me.caption = "migrando " & .Fields(0).Value
                End If
                r2.Close
             '   .MoveNext
            'Loop
            
            DoEvents
            .MoveNext
        Loop
        .Close
    End With
    
    
    
End Sub
Sub migraCons(tabela As String)
    Dim sSEQ As String
    Dim rPesq As New adodb.Recordset
    Dim strSQL As String
    Dim sCodVen As String
    
    If Not F_AbreBancoODBC(b, txtbxBancoMySQL.Text) Then
        MsgBox "Erro.", vbExclamation
        Exit Sub
    End If
    
    strSQL = "select codivend from vended where filivend='" & tabela & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            Exit Sub
        Else
        
            sCodVen = Format(rPesq!codivend, "0000#")
        End If
    End With
    
    
    strSQL = "select distinct sum(qteeste) as total," & tabela & ".* from " & tabela & " group by codproe having sum(qteeste) > 0"
    abreObjetoAdo rPesq, strSQL
    
    strSQL = "delete from tab_consignacoes where cod_rep='" & sCodVen & "' "
    Conn.Execute strSQL
    strSQL = "delete from tab_estcons where cod_rep='" & sCodVen & "' "
    Conn.Execute strSQL
    
    sSEQ = MaxKeySQL("tab_consignacoes", "seq_con", "", "000000#")  'buscando o sequencial q servirá para todos os produtos da transf
    
    
    With rPesq
        Do While Not .EOF
'            MsgBox !codproe
            If Consignado.F_GravaConsign(sSEQ, Format(!codproe, "0000000#"), !qteeste, sCodVen, F_BuscaUmCampo("pv_pro", "tab_produtos", "cod_pro=" & !codproe)) = "" Then
                MsgBox "Erro na transferência no item " & .AbsolutePosition, vbCritical
                
            Else
                'If Not Consignado.F_DeletaConsignAberta(seq_con) Then
                
                'End If
            End If
            
            strSQL = !codproe
            
          
            .MoveNext
        Loop
        'DoEvents
    End With
End Sub

Private Sub Command2_Click()
    buscaArquivos
End Sub

Sub buscaArquivos()
    On Error Resume Next
    File1.Path = Text1.Text
End Sub

Private Sub Command3_Click()
    migraCons Mid(File1.FileName, 1, Len(File1.FileName) - 4)
    
    
    
    
End Sub



Private Sub Command5_Click()
    
    
    cmndbtnMigraTodos_Click
    
    frmMigra.tag = "x"
    
    frmMigra.Show 1
    Command1_Click
    
    MsgBox "Tudo foi migrado. Terminou em " & Now
    
    
    
    
End Sub

Private Sub File1_DblClick()
    migraArquivo
'    rPesq.Close
End Sub
Sub migraArquivo()
    Dim rPesq As adodb.Recordset
    Dim dbconn As adodb.Connection
    Set dbconn = openDBFconn(Text1.Text)
    Dim bTest As database
    Dim rp As New adodb.Recordset
    Dim iCount As Double
    Dim strSQL As String
    Dim iPos As Double
    Dim i As Double
    Set rPesq = dbconn.Execute("select * from [" & File1.FileName & "#DBF]")
    
    'bTest.Connect = "[" & File1.FileName & "#DBF]"
    'set rp = bTest.OpenRecordset
    Dim sNomeTb As String
    strSQL = ""
    
    With rPesq
        '.Supports(
        '.MoveFirst
        sNomeTb = Trim(Left(File1.FileName, Len(File1.FileName) - 4))
      If InStr(1, sNomeTb, " ") > 0 Then
     '   MsgBox ""
        sNomeTb = F_SoLetra(sNomeTb)
      End If
      '  sNomeTb = F_SoLetra(sNomeTb)
        
        
        ReDim cmp(.Fields.Count) As campo
        strSQL = "create table " & sNomeTb & "(pk integer(7) auto_increment,"
        For iCount = 0 To .Fields.Count - 1
            
            
            strSQL = strSQL & .Fields(iCount).Name
            If .Fields(iCount).Type = adDate Then
                strSQL = strSQL & " date not null,"
            ElseIf .Fields(iCount).Type = 202 Then
                strSQL = strSQL & " char(" & .Fields(iCount).DefinedSize & ") not null,"
            ElseIf .Fields(iCount).Type = 2 Then
                strSQL = strSQL & " char(" & .Fields(iCount).DefinedSize & ") not null,"
            ElseIf .Fields(iCount).Type = 5 Then
                strSQL = strSQL & " decimal(14,2) not null,"
            
            
            Else
                'MsgBox "o tipo " & .Fields(iCount).Type
            End If
            
            
            
        Next
        strSQL = strSQL & "primary key(pk));"
        
        bNew.Execute "drop table if exists " & sNomeTb
        bNew.Execute strSQL
        
        
        
        Do While Not .EOF
            strSQL = "insert into " & sNomeTb & " values ('',"
            
            For iCount = 0 To .Fields.Count - 1
                If .Fields(iCount).Type = adDate Then
                  '  On Error Resume Next
                    If IsMissing(.Fields(iCount).Value) Then
                        strSQL = strSQL & "'',"
                    Else
                        strSQL = strSQL & F_GravaData(.Fields(iCount).Value)
                    End If
                
                ElseIf .Fields(iCount).Type = 5 Then
                  '  On Error Resume Next
                    If IsMissing(.Fields(iCount).Value) Then
                        strSQL = strSQL & "0,"
                    Else
                        strSQL = strSQL & F_GravaDec(.Fields(iCount).Value)
                    End If
                Else
                   
                   
                   
                   ' On Error Resume Next
                    If IsNull(.Fields(iCount).Value) Then
                   
                        strSQL = strSQL & "''"
                    Else
                        strSQL = strSQL & "'" & F_AjustaTexto(.Fields(iCount).Value) & "'"
                    End If
                End If
                
                If iCount = .Fields.Count - 1 Then
                    strSQL = strSQL & ")"
                Else
                    strSQL = strSQL & ","
                End If
            
            
            '    P_Status "Migrando " & iPos + 1
                
                
                
            Next
            .MoveNext
            i = i + 1
            Me.caption = i
            
            bNew.Execute strSQL
            
            
            DoEvents
                    Loop
        .Close
    End With
    
    'MsgBox "Migrado!", vbInformation
    P_Status ""
End Sub

Private Sub Form_Load()
    buscaArquivos
    If Not F_AbreBancoODBC(bNew, "req_dbf") Then
        MsgBox ""
    End If
End Sub




Function openDBFconn(Path)
  Dim Conn As adodb.Connection
  Set Conn = CreateObject("adodb.connection")
  Conn.Open "provider=microsoft.jet.oledb.4.0;" & _
                   "data source=" & Path & ";" & _
                   "extended properties=""dbase iv;"";"
  Set openDBFconn = Conn
End Function

Private Sub cmd1_click()
'abir conexão para ficheiros DBF em c:\[nome do directorio]
Dim dbconn As Connection
Set dbconn = openDBFconn("c:\testes")
'criar um novo ficheiro DBF chamado persons.DBF
dbconn.Execute "create table persons (name char(50), city char(50), phone char(20), zip decimal(5))"
End Sub

Private Sub cmd2_click()
'inserir alguns registos na tabela persons
Dim dbconn As Connection
Set dbconn = openDBFconn("c:\testes")
dbconn.Execute "insert into persons values('alex p. nor', 'mexico','458962146','14589')"
End Sub

Private Sub cmd3_click()
'abir  recordset da tabela  persons
Dim dbconn As Connection
Set dbconn = openDBFconn("c:\testes")
Dim persons As New adodb.Recordset
'Set persons = dbconn.Execute("select * from [persons#DBF]")
End Sub


