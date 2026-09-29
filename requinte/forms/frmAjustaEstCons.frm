VERSION 5.00
Begin VB.Form frmAjustaEstCons 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Ajustar Consignações"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ajustar Estoque"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "requinte_09082011"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banco Backup"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1065
   End
End
Attribute VB_Name = "frmAjustaEstCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bancoOld As New ADODB.Connection
Dim sSQL As String
Private Sub Command1_Click()
    bancoOld.ConnectionString = Text1.Text
    bancoOld.Open
End Sub

Private Sub Command2_Click()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    If MsgBox("Tem certeza disso?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    
    strSQL = "select * from tab_estcons where cod_rep=000263"
    
    
    rPesq.Open strSQL, bancoOld
    
    strSQL = "delete from tab_estcons where cod_rep=263"
    Banco.Execute strSQL
    
    
    
    
    With rPesq
        Do While Not .EOF
            i = i + 1
            strSQL = "insert into tab_estcons(cod_rep,cod_pro,estoque_pro)values("
            strSQL = strSQL & !cod_Rep & ","
            strSQL = strSQL & "'" & !cod_pro & "',"
            strSQL = strSQL & F_GravaDec(!estoque_pro) & ")"
            Banco.Execute strSQL
            
            .MoveNext
        Loop
        .Close
    End With
    
End Sub

Private Sub Command3_Click()
        Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    If MsgBox("Tem certeza disso?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    
    strSQL = "select * from tab_consignacoes where cod_rep=000263 order by cod_pro,seq_con"
    
    
    rPesq.Open strSQL, bancoOld
    
    strSQL = "delete from tab_consignacoes where cod_rep=263 "
    Conn.Execute strSQL
    
    
    Open App.path & "\sql.txt" For Output As #1
    
    With rPesq
        Do While Not .EOF
            i = i + 1
            On Error Resume Next
            strSQL = "insert into tab_Consignacoes(seq_con,cod_pro,data_con,qtd_con,cod_rep,cod_usu,reman_con,"
            strSQL = strSQL & "unit_con,seq_cab,retif_con,reservado_con)values("
            strSQL = strSQL & "'" & !seq_con & "',"
            strSQL = strSQL & "'" & !cod_pro & "',"
            
            
            strSQL = strSQL & F_GravaDataTempo(!data_con) & ","
            strSQL = strSQL & F_GravaDec(!qtd_con) & ","
            strSQL = strSQL & "'" & !cod_Rep & "',"
            strSQL = strSQL & "'" & !cod_usu & "',"
            strSQL = strSQL & F_GravaDec(!reman_con) & ","
            strSQL = strSQL & F_GravaDec(!unit_con) & ","
            strSQL = strSQL & "'" & !seq_cab & "',"
            strSQL = strSQL & "'" & !retif_con & "',"
            strSQL = strSQL & F_GravaDec(!reservado_con) & ");"
            Conn.Execute strSQL
            'sSQL = sSQL + strSQL & Chr(13) + Chr(10)
            Print #1, strSQL
            
            
            .MoveNext
        Loop
        Close #1
        .Close
    End With

End Sub
