VERSION 5.00
Begin VB.Form frmForcaAbertos 
   Caption         =   "Buscar Produtos Reservados"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
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
Private Sub Command1_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim iAce As Integer
    Dim iIteAce As Integer
    Dim iDev As Integer
    Dim iIteDev As Integer
    Dim iRes As Integer
    
    strSQL = "Tem certeza que deseja cancelar todos (acertos e devoluções) em aberto do consignado "
    strSQL = strSQL & Text1.Text & "?"
    
    If MsgBox(strSQL, vbYesNo + vbCritical) = vbYes Then
        strSQL = "select * from tab_acertosabertos where cod_rep='" & Text1.Text & "' "
        Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
            Do While Not .EOF
                strSQL = "delete from tab_acertosabertos where seq_ace='" & !seq_ace & "' "
                Banco.Execute strSQL
                strSQL = "delete from tab_itacertosabertos where seq_ace='" & !seq_ace & "' "
                Banco.Execute strSQL
                iIteAce = Banco.RecordsAffected
                iAce = .RecordCount
                
                
                .MoveNext
            Loop
            .Close
        End With
        
        strSQL = "select * from tab_devconsignabertas where cod_rep='" & Text1.Text & "' "
        Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
            Do While Not .EOF
                strSQL = "delete from tab_devconsignabertas where seq_dev='" & !seq_ace & "' "
                Banco.Execute strSQL
                strSQL = "delete from tab_itdevconsignabertas where seq_dev='" & !seq_ace & "' "
                Banco.Execute strSQL
                iIteDev = Banco.RecordsAffected
                iDev = .RecordCount
                
                
                .MoveNext
            Loop
            .Close
        End With
        strSQL = "update tab_consignacoes set reservado_con=0 where cod_rep='" & Text1.Text & "' "
        Banco.Execute strSQL
        iRes = Banco.RecordsAffected
        
    End If
    
    strSQL = iAce & " Excluídos " & Chr(13)
    strSQL = strSQL & iIteAce & " Itens de Acerto Excluídos" & Chr(13)
    strSQL = strSQL & iDev & " Devoluções Excluídas " & Chr(13)
    strSQL = strSQL & iIteDev & "Itens de  Devoluções Excluídos " & Chr(13)
    strSQL = strSQL & iRes & " Estoques Reservados Retornados " & Chr(13)
    MsgBox strSQL, vbInformation
End Sub

Private Sub Text1_LostFocus()
    On Error Resume Next
     Text1.Text = Format(Text1.Text, "0000#")
End Sub
