VERSION 5.00
Begin VB.Form frmMigracaoBDUnico 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Passo 1.1 - Histórico DOCS"
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdSetarDocItens 
      Caption         =   "Passo 1.1 -  Setar os ids dos docs no itens"
      Height          =   855
      Left            =   2400
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdHist 
      Caption         =   "Passo 1.1 - Histórico DOCS"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtbxBD 
      Height          =   405
      Left            =   360
      TabIndex        =   1
      Text            =   "requintebd"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdPasso1 
      Caption         =   "Passo 1 - Migrar Docs (20 minutos)"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome do novo BD"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "frmMigracaoBDUnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdHist_Click()
     Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim Arquivo As String
    Dim sCommand As String
    
        
    
    strSQL = "use " & txtbxBD.Text
    Conn.Execute strSQL
    
    
    strSQL = "CREATE TABLE dochist SELECT * FROM requinte.tab_dochist;"
    Conn.Execute strSQL
    strSQL = "ALTER TABLE `dochist`"
    strSQL = strSQL & " ADD COLUMN `id` INT NOT NULL AUTO_INCREMENT FIRST,"
    strSQL = strSQL & " ADD PRIMARY KEY (`id`);"
    
    strSQL = "ALTER TABLE `dochist`   ADD COLUMN `id_doc` INT references doc(id);"
    Conn.Execute strSQL
    
    strSQL = "ALTER TABLE `dochist`    CHANGE COLUMN `seq_doc` `seq_doc` CHAR(12) NOT NULL;"
    Conn.Execute strSQL

    strSQL = "select id,seq_doc from doc"
    abreObjetoAdo rPesq, strSQL
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = Conn
    Conn.BeginTrans
    With rPesq
        Do While Not .EOF
            If .RecordCount Mod 100 = 0 Then
                Me.caption = .AbsolutePosition & "/" & .RecordCount
            End If
        
        
            strSQL = "update dochist set id_doc=" & !id & " where seq_doc=" & !seq_Doc
            Conn.Execute strSQL
            .MoveNext
        Loop
    End With
    Conn.CommitTrans
    
    
    
    
End Sub

Private Sub cmdPasso1_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim Arquivo As String
    Dim sCommand As String
    
    strSQL = "drop database if exists " & txtbxBD.Text
    Conn.Execute strSQL
    
    strSQL = "create database " & txtbxBD.Text
    Conn.Execute strSQL
        
    
    strSQL = "use " & txtbxBD.Text
    Conn.Execute strSQL
    
    strSQL = "CREATE TABLE doc SELECT * FROM requinte.tab_doc;"
    Conn.Execute strSQL
    strSQL = "ALTER TABLE `doc` engine=myisam;"
    Conn.Execute strSQL
    
    strSQL = "ALTER TABLE `doc` "
    strSQL = strSQL & "    ADD COLUMN `id` INT NOT NULL AUTO_INCREMENT FIRST,"
    strSQL = strSQL & "ADD PRIMARY KEY (`id`);"
    Conn.Execute strSQL
    
    strSQL = "ALTER TABLE `doc` "
    strSQL = strSQL & "ADD COLUMN `cod_emp` char(2) NOT NULL DEFAULT '01' AFTER `cod_pla`;"
    Conn.Execute strSQL
    strSQL = "ALTER TABLE `doc` CHANGE COLUMN `seq_doc` `seq_doc` CHAR(12) NOT NULL;"
    Conn.Execute strSQL
    
    strSQL = "CREATE TABLE itdoc SELECT * FROM requinte.tab_itdoc;"
    Conn.Execute strSQL
    strSQL = "ALTER TABLE `itdoc` engine=myisam;"
    Conn.Execute strSQL
    
    strSQL = "ALTER TABLE `itdoc`     ADD COLUMN `id` INT NOT NULL AUTO_INCREMENT FIRST,  ADD PRIMARY KEY (`id`);"
    Conn.Execute strSQL
    
    

    strSQL = "ALTER TABLE `itdoc`     ADD COLUMN `id_doc` INT references doc(id);"
    Conn.Execute strSQL
    
    

    Arquivo = "d:/temp/itdoc.sql"

    strSQL = "SELECT CONCAT('update itdoc set id_doc=',id, ' where seq_doc=',seq_doc,';') FROM doc ORDER BY id DESC "
    strSQL = strSQL & " into outfile '" & Arquivo & "';"
    abreObjetoAdo rPesq, strSQL
    
'    strSQL = "select id,seq_doc from doc"
'    abreObjetoAdo rPesq, strSQL
'    Set cmd = New ADODB.Command
'    Set cmd.ActiveConnection = Conn
'    Conn.BeginTrans
'    With rPesq
'        Do While Not .EOF
'            If .AbsolutePosition Mod 10 = 0 Then
'                Me.caption = .AbsolutePosition & "/" & .RecordCount
'            End If
'
'
'            strSQL = "update itdoc set id_doc=" & !id & " where seq_doc=" & !seq_doc
'            Conn.Execute strSQL
'            .MoveNext
'        Loop
'    End With
'    Conn.CommitTrans
'
    
         
    
    
    sCommand = "mysql -u root -p=masterkey requintebd < d:\Temp\itdoc.sql"
    
    Shell "cmd.exe &&" & sCommand
    
   
    
End Sub



'Sub RunBatchFile()
'    Set Shell = New Shell
'    Shell.Run "cmd.exe /c C:\Caminho\Para\Arquivo.bat", vbHide
'End Sub

Private Sub shell_OnExecuted(ByVal ProcessID As Long)
    MsgBox "O arquivo em lote foi executado com sucesso!"
End Sub

Private Sub shell_OnTerminated(ByVal ProcessID As Long, ByVal ExitCode As Long)
    MsgBox "A execução do arquivo em lote foi concluída."
End Sub
Private Sub cmdSetarDocItens_Click()
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim Arquivo As String
    Dim sCommand As String
    
        
    
    strSQL = "use " & txtbxBD.Text
    Conn.Execute strSQL
    
    Arquivo = "c:/temp/itdoc.sql"

    strSQL = "SELECT CONCAT('update itdoc set id_doc=',id, ' where seq_doc=',seq_doc,';') FROM doc ORDER BY id DESC "
    strSQL = strSQL & " into outfile '" & Arquivo & "';"
    abreObjetoAdo rPesq, strSQL
    
          
    
    sCommand = "mysql -u root -p=masterkey requintebd < d:\Temp\itdoc.sql"
    
    
    Shell "cmd.exe /K " & sCommand, vbMaximizedFocus
    
End Sub

