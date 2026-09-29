VERSION 5.00
Begin VB.Form frmAjuste 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "Num Cup para SaidaNF"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtNota 
      Height          =   285
      Left            =   6600
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Produtos das Notas d Saida"
      Height          =   615
      Left            =   6600
      TabIndex        =   12
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   615
      Left            =   3120
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Saidas, ItSaidas"
      Height          =   615
      Left            =   840
      TabIndex        =   10
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Classificação"
      Height          =   615
      Left            =   7680
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ItVndCnc"
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ajuste Venfin"
      Height          =   855
      Left            =   7440
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   855
      Left            =   5160
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Busca Cupom Sem Venda"
      Height          =   615
      Left            =   4680
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1335
      Left            =   4680
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "busca NF sem complemento"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Nota:"
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   3960
      Width           =   375
   End
End
Attribute VB_Name = "frmAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "show tables"
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "delete from " & .Fields(0).Value
'            conn.execute strSQL
            .MoveNext
        Loop
    End With
End Sub

Private Sub Command10_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iCount As Integer
    strSQL = "select * from tab_saidas"
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_saidas set seq_sai=" & .AbsolutePosition + 1
            conn.Execute strSQL
            Command10.caption = .AbsolutePosition + 1 & " de " & .RecordCount
            DoEvents
            .MoveNext
        Loop
        .Close
    End With
    
    
    
    
    
    
    
End Sub

Private Sub Command11_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCodPro As String
    strSQL = "select cod_pro from tab_produtos "
    rPesq.Open strSQL, conn_client
    
    With rPesq
        Do While Not .EOF
           
            strSQL = "update tab_itcup set cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " where cod_pro='" & Left(!cod_pro, 7) & "'"
            conn.Execute strSQL
           
            Me.caption = .AbsolutePosition + 1 & " de " & .RecordCount
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Command12_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim r2 As New ADODB.Recordset
    strSQL = "select seq_sai,num_ven,picms_sai from tab_saidasnf "
    strSQL = strSQL & " where seq_sai<" & txtNota.Text
    strSQL = strSQL & " and num_ven<>''"
    strSQL = strSQL & " order by seq_sai "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "delete from tab_itsainf where seq_sai=" & !Seq_sai
            conn.Execute strSQL
            
            strSQL = "select ti.* from tab_itvendas as ti where num_ven='" & !num_ven & "' "
            r.Open strSQL
            With r2
                Do While Not .EOF
                    strSQL = "insert into tab_itsainf(seq_sai,seq_its,cod_pro,qtd_its,unit_its,icms_its)values("
                    strSQL = strSQL & rPesq!Seq_sai & ","
                    strSQL = strSQL & .AbsolutePosition + 1 & ","
                    strSQL = strSQL & "'" & !cod_pro & "',"
                    strSQL = strSQL & F_GravaDec(!qtd_ite) & ","
                    strSQL = strSQL & F_GravaDec(!unit_ite) & ","
                    strSQL = strSQL & F_GravaDec(rPesq!picms_sai) & ")"
                    conn.Execute strSQL
                    
                    .MoveNext
                Loop
                .Close
            End With
            
            Me.caption = .AbsolutePosition + 1
            
            .MoveNext
        Loop
        .Close
    End With
    MsgBox "ok"
End Sub

Private Sub Command13_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCup As String
    strSQL = "select seq_sai,num_ven from tab_saidasnf"
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            sCup = F_BuscaUmCampo("seq_cup", "tab_cupons", "num_ven='" & !num_ven & "'")
            strSQL = "update tab_saidasnf set seq_cup='" & sCup & "' "
            strSQL = strSQL & " where seq_sai='" & !Seq_sai & "'"
            conn.Execute strSQL
            P_Status .AbsolutePosition + 1 & "/" & .RecordCount
            .MoveNext
        Loop
        .Close
    End With
P_Status ""
    
End Sub

Private Sub Command2_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_compras order by num_not "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            If F_BuscaUmCampo("seq_not", "tab_cmp", "num_not='" & !num_not & "'") = "" Then
                Text1.Text = Text1.Text & !num_not & Chr(13) + Chr(10)
            End If
            .MoveNext
        Loop
    End With
End Sub

Private Sub Command3_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_itdoc1 "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            If F_BuscaUmCampo("cod_pro", "tab_codbarras", "cod_pro='" & !cod_pro & "'") = "" Then
                strSQL = "insert into tab_codbarras(cod_pro,codbarra_pro)values('" & !cod_pro & "',"
                strSQL = strSQL & "'" & !cod_pro & "') "
                conn.Execute strSQL
            End If
            .MoveNext
        Loop
        .Close
    End With
    
End Sub

Private Sub Command4_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select cod_pro,codfab_pro from tab_produtos where codfab_pro like '00%'"
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "Update tab_produtos set codfab_pro=" & !codfab_pro
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Command5_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    strSQL = "select  * from tab_saidas where emissao_sai >='2006-01-01' order by num_ven "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            If F_BuscaUmCampo("situacao_ven", "tab_vendas", "num_ven='" & !num_ven & "'") = "3" Then
                cTotal = cTotal + !valor_sai
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    MsgBox cTotal
End Sub

Private Sub Command6_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sIn As String
    strSQL = "select seq_doc from tab_docback "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            sIn = sIn & "'" & !seq_doc & "'"
            .MoveNext
            
            If Not .EOF Then
                sIn = sIn & ","
            
                
            End If
        Loop
        sIn = "(" & sIn & ")"
    End With
    strSQL = "select * from tab_doc1 where seq_doc not in" & sIn
    rPesq.Open strSQL, conn_client
    Text1.Text = ""
    With rPesq
        Do While Not .EOF
            Text1.Text = Text1.Text & !seq_doc & "-" & !Data_Doc & "-" & !Valor_Doc & Chr(13)
            .MoveNext
        
        Loop
        MsgBox "Foram encontrados " & .RecordCount & " registros"
        .Close
        
    End With
End Sub

Private Sub Command8_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_itvndcncl"
    rPesq.Open strSQL, conn_client
    
    With rPesq
        Do While Not .EOF
            
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub

Private Sub Command9_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_class "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            
            strSQL = "Update tab_financeiro set cod_pla='" & !cod_pla & "' "
            strSQL = strSQL & " where seq_fin='" & !seq_fin & "' "
            conn.Execute strSQL
            .MoveNext
            Me.caption = .AbsolutePosition + 1 & " de " & .RecordCount
            Me.Refresh
        Loop
        
        .Close
    End With
End Sub

