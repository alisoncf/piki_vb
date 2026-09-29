VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmForcaAbertos 
   Caption         =   "Buscar Produtos Reservados"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Problemas com Comissões de Consignados"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "It Doc sem Doc"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   3600
      Width           =   4695
   End
   Begin VB.CommandButton cmndbtnProblCarneTerceiro 
      Caption         =   "Carnes de Terceiros com Clientes"
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   2880
      Width           =   4695
   End
   Begin VB.CommandButton cmndbtnEstRes 
      Caption         =   "Ajusta Est Res"
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   1455
      Left            =   6480
      TabIndex        =   4
      Top             =   960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin VB.CommandButton cmndbtnBuscaErroCons 
      Caption         =   "Buscar Erros de QTD do Estoque Consign"
      Height          =   735
      Left            =   6480
      TabIndex        =   3
      Top             =   240
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar Todos Documentos em Aberto"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Deletar Vendas:"
      Height          =   1335
      Left            =   360
      TabIndex        =   9
      Top             =   3120
      Width           =   3375
      Begin VB.CommandButton cmnDelete 
         Caption         =   "Deletar"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxData 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optnFechadas 
         Caption         =   "Fechadas"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optnAbertas 
         Caption         =   "Abertas"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até:"
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   285
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código do Representante:"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   1860
   End
End
Attribute VB_Name = "frmForcaAbertos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscaErroCons_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    
End Sub

Private Sub cmndbtnEstRes_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select cod_pro,sum(reman_con) as total from tab_consignacoes "
    strSQL = strSQL & " group by cod_pro order by cod_pro "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_produtos set r_pro=" & F_GravaDec(!total)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    MsgBox "Pronto! "
    
    
End Sub

Private Sub cmndbtnProblCarneTerceiro_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select td.seq_doc from tab_cheques as tc,tab_doc as td  "
    strSQL = strSQL & " where tc.seq_doc=td.seq_doc  and historico_doc like 'BAIXA DOCS%'"
    
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        
        Do While Not .EOF
            strSQL = "update tab_doc set tipo_ter='R' where seq_doc='" & !seq_Doc & "'"
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub

Private Sub cmnDelete_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTotal As Integer
    If MsgBox("Confirma exclusão das Vendas?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    If txtbxData.Text = "" Then
        MsgBox "Informe a data limite para exclusão!", vbExclamation
        txtbxData.SetFocus
        Exit Sub
    End If
    
    
    strSQL = "select num_ven from tab_vendas "
    
    If optnAbertas.Value Then
        strSQL = strSQL & " where situacao_ven=1"
    Else
        strSQL = strSQL & " where situacao_ven=2"
    End If
    strSQL = strSQL & " and data_ven<=" & F_GravaData(Date)
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        sTotal = .RecordCount
        Do While Not .EOF
            strSQL = "delete from tab_itvendas where num_ven='" & !num_ven & "'"
            Conn.Execute strSQL
            
            strSQL = "delete from tab_vendas where num_ven='" & !num_ven & "'"
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    Conn.Execute strSQL
    
    MsgBox sTotal & " Registros excluídos!", vbInformation
    
    
    
    
End Sub

Private Sub Command1_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iAce As Integer
    Dim iIteAce As Integer
    Dim iDev As Integer
    Dim iIteDev As Integer
    Dim iRes As Integer
    
    strSQL = "Tem certeza que deseja cancelar todos (acertos e devoluções) em aberto do consignado "
    strSQL = strSQL & Text1.Text & "?"
    
    If MsgBox(strSQL, vbYesNo + vbCritical) = vbYes Then
        strSQL = "select * from tab_acertosabertos where cod_rep='" & Text1.Text & "' "
        abreObjetoAdo rPesq, strSQL
        With rPesq
            Do While Not .EOF
                strSQL = "delete from tab_acertosabertos where seq_ace='" & !seq_ace & "' "
                Conn.Execute strSQL
                strSQL = "delete from tab_itacertosabertos where seq_ace='" & !seq_ace & "' "
                Conn.Execute strSQL, iIteAce
                
                iAce = .RecordCount
                
                
                .MoveNext
            Loop
            .Close
        End With
        
        strSQL = "select * from tab_devconsignabertas where cod_rep='" & Text1.Text & "' "
        abreObjetoAdo rPesq, strSQL
        With rPesq
            Do While Not .EOF
                strSQL = "delete from tab_devconsignabertas where seq_dev='" & !seq_dev & "' "
                Conn.Execute strSQL
                 '   strSQL = "delete from tab_itdevconsignabertas where seq_dev='" & !seq_dev & "' "
                'conn.execute strSQL
                iIteDev = Banco.RecordsAffected
                iDev = .RecordCount
                
                
                .MoveNext
            Loop
            .Close
        End With
        strSQL = "update tab_consignacoes set reservado_con=0 where cod_rep='" & Text1.Text & "' "
        Conn.Execute strSQL, iRes
        
        
    End If
    
    strSQL = iAce & " Excluídos " & Chr(13)
    strSQL = strSQL & iIteAce & " Itens de Acerto Excluídos" & Chr(13)
    strSQL = strSQL & iDev & " Devoluções Excluídas " & Chr(13)
    strSQL = strSQL & iIteDev & "Itens de  Devoluções Excluídos " & Chr(13)
    strSQL = strSQL & iRes & " Estoques Reservados Retornados " & Chr(13)
    MsgBox strSQL, vbInformation
End Sub

Private Sub Command2_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select seq_doc from tab_itdoc where seq_doc not in " & F_NotIn("seq_doc", "tab_doc", "")
    abreObjetoAdo rPesq, strSQL
    With rPesq
        msflxgrd.rows = 1
        Do While Not .EOF
            msflxgrd.AddItem !seq_Doc
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub

Private Sub Text1_LostFocus()
    On Error Resume Next
     Text1.Text = Format(Text1.Text, "0000#")
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub
