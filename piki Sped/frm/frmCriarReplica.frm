VERSION 5.00
Begin VB.Form frmCriarReplica 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Criar Réplica"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Limpar Dados Anteriores a 1 Ano"
      Height          =   2535
      Left            =   4680
      TabIndex        =   13
      Top             =   240
      Width           =   3135
      Begin VB.CommandButton cmndbtnExclDadosAntigos 
         Caption         =   "Excluir Dados Antigos"
         Height          =   615
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   1440
      ScaleHeight     =   1545
      ScaleWidth      =   5385
      TabIndex        =   5
      Top             =   3840
      Width           =   5415
      Begin VB.CommandButton cmndbnCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registro:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblRegistro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registro:"
         Height          =   195
         Left            =   1200
         TabIndex        =   10
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblReplicando 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tabelas"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replicando:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblQtdTabelas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tabelas"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   120
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tabelas:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.CommandButton cmndbtnReplicar 
      Caption         =   "Replicar"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtbxBancoReplica 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Market_rep"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Para o Banco de Dados:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblNomeBanco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Market"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Banco de Dados"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1890
   End
End
Attribute VB_Name = "frmCriarReplica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnExclDadosAntigos_Click()
    If MsgBox("Os dados não serão mais recuperados, se não houver réplica, você pode ficar sem os dados definitivamente" & Chr(13) & _
        "Deseja Continuar?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    excluidados
    
    
End Sub
Sub excluidados()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sDataIni As String
    strSQL = "select now();"
    rPesq.Open strSQL, conn_client
    With rPesq
        If Day(Now) <> Day(.Fields(0).Value) Then
            MsgBox "A data deste computador está diferente da data do servidor!", vbExclamation
            Exit Sub
        End If
        sDataIni = DateAdd("M", -12, Date)
        
        .Close
        If MsgBox("Todos os dados anteriores a " & sDataIni & " Serão removidos. Confirma?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End With
    strSQL = "select * from " & TAB_VENDAS & " where data_ven<=" & F_GravaDataTempo(sDataIni)
    strSQL = strSQL & " and situacao_ven>1 "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "delete from " & TAB_ITVENDAS & " where num_ven='" & !num_ven & "' "
            conn.Execute strSQL
            strSQL = "delete from  " & TAB_VENFEC & "  where num_ven='" & !num_ven & "' "
            conn.Execute strSQL
            strSQL = "delete from " & TAB_VENDAS & " where num_ven='" & !num_ven & "' "
            conn.Execute strSQL
            strSQL = "delete from tab_itvndcncl where num_ven='" & !num_ven & "' "
            conn.Execute strSQL
            lblRegistro.caption = "Excluindo venda " & .AbsolutePosition + 1 & "/" & .RecordCount
            DoEvents
            .MoveNext
        Loop
    End With
    
    strSQL = "delete from tab_caixa where data_Cai<=" & F_GravaData(sDataIni)
    conn.Execute strSQL
    
    strSQL = "select * from tab_saidas where emissao_sai<=" & F_GravaData(sDataIni)
    
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "delete from tab_itsai nota_sai='" & !nota_sai & "' "
            strSQL = strSQL & " and cod_Cai='" & !Cod_cai & "'"
            conn.Execute strSQL
            lblRegistro.caption = "Excluindo saida " & .AbsolutePosition + 1 & "/" & .RecordCount
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "delete from tab_saidas where emissao_sai<=" & F_GravaData(sDataIni)
    conn.Execute strSQL
    
    
    strSQL = "flush tables;"
    conn.Execute strSQL
    strSQL = "REPAIR table " & TAB_VENDAS & "," & TAB_ITVENDAS & ",tab_saidas,tab_itsaidas"
    conn.Execute strSQL
End Sub
Private Sub cmndbtnReplicar_Click()
    replica
End Sub
Sub replica()
    Dim strSQL As String
    Dim r1 As New ADODB.Recordset
    Dim r2 As New ADODB.Recordset
    Dim r3 As New ADODB.Recordset
    Dim b As New ADODB.Connection
    Dim sValues As String
    Dim sType As String
    Dim sNomeTab As String
    Dim i As Integer
    On Error GoTo trata_Erro
    If Not F_AbreBancoODBCAdo(b, txtbxBancoReplica.Text) Then
        Exit Sub
    End If
    
    strSQL = "show tables;"
    r1.Open strSQL
    With r1
        
            lblQtdTabelas.caption = r1.RecordCount
        
        Do While Not .EOF
            On Error Resume Next
            
            
            sNomeTab = r1.Fields(0).Value
            strSQL = "select * from " & sNomeTab
            
            r2.Open strSQL
            With r2
                Do While Not .EOF
                    'strSQL = "select * from " & r1.Fields(0).Value
                   ' Set r3 = b.OpenRecordset(strSQL, dbOpenSnapshot)
                    'If r3.RecordCount = 0 Then
                    '    r3.Close
                    'End If
                    sValues = ""
                    For i = 0 To .Fields.Count - 1
                        If .Fields(i).Type = 18 Then
                            sValues = sValues & "'" & .Fields(i).Value & "'"
                        ElseIf .Fields(i).Type = 20 Then
                            sValues = sValues & F_GravaDec(.Fields(i).Value)
                        ElseIf .Fields(i).Type = 23 Then
                            sValues = sValues & F_GravaDataTempo(.Fields(i).Value)
                        ElseIf .Fields(i).Type = 8 Then
                            sValues = sValues & F_GravaData(.Fields(i).Value)
                        
                        Else
                            MsgBox "mudar"
                        End If
                        If i = .Fields.Count - 1 Then
                            sValues = sValues & ""
                        Else
                            sValues = sValues & ","
                        End If
                    Next
                    Err.Clear
                    strSQL = "insert into " & sNomeTab & " values(" & sValues & ")"
                    b.Execute strSQL
                    If Err.Number = 0 Then
                    
                    End If
                    lblRegistro.caption = .AbsolutePosition + 1 & " de " & .RecordCount
                    lblReplicando.caption = sNomeTab & " (" & r1.AbsolutePosition + 1 & " de " & r1.RecordCount & ")"
                    DoEvents
                    .MoveNext
                Loop
                .Close
            End With
            
            
            
            .MoveNext
        Loop
        
        .Close
    End With
    
    
    b.Close
    
    Exit Sub
trata_Erro:
    MostraErro
End Sub
