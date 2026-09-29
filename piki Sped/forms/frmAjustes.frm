VERSION 5.00
Begin VB.Form frmAjustes 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Devolvido 2"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Devolvido"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmAjustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "update itvendas set devolvido_ite=0"
    conn.Execute strSQL
    
    strSQL = "select num_ven,cod_pro, qtd_dev from devolucoes as d inner join itdev as i "
    strSQL = strSQL & " on d.seq_dev=i.seq_dev "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            
            strSQL = "update itvendas set devolvido_ite=" & F_GravaDec(!qtd_dev)
            strSQL = strSQL & " where num_ven='" & !num_ven & "' "
            strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
            conn.Execute strSQL
            .MoveNext
        Loop
    End With
    
    
End Sub

Private Sub Command2_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
  
    strSQL = "update vendas set devolvido_ven=0"
    conn.Execute strSQL
    
    strSQL = "select * from " & TAB_ITVENDAS & " where devolvido_ite>0 order by num_ven desc "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            
            strSQL = "update vendas set devolvido_ven=devolvido_Ven+" & F_GravaDec((!final_ite / !qtd_ite) * !devolvido_ite)
            strSQL = strSQL & " where num_ven='" & !num_ven & "' "
          '  strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
            conn.Execute strSQL
            
            P_Status "aguarde... ajustando... " & .AbsolutePosition & "/" & .RecordCount
            If .AbsolutePosition Mod 50 = 0 Then
                Me.Refresh
            End If
            
            .MoveNext
        Loop
    End With
    P_Status
End Sub
