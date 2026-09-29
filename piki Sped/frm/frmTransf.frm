VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTransf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferência"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnMatriz 
      Caption         =   "Matriz P/ &Lojas"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton cmndbtnConcluir 
      Caption         =   "&Concluir"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmndbtnTransferir 
      Caption         =   "&Transferir"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton cmndbtnGeraLog 
      Caption         =   "&Gerar"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmTransf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnConcluir_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sExpressao As String
    Dim t As String
    t = Time
    
    
    
    strSQL = "select * from tab_transferencias where status_tra='0' "
    strSQL = strSQL & " and cod_loj<>'" & strCod_Loj & "' order by seq_tra"
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Não foram encontrados registros para serem transferidos!", vbExclamation
            Exit Sub
        End If
        Do While Not .EOF
            sExpressao = F_BuscaUmCampo("expressao_tra", "tab_transferencias", _
                "seq_tra=" & !seq_tra)
            sExpressao = F_ColocaAspa(sExpressao)
            
            banco.Execute sExpressao
            
            strSQL = "update tab_transferencias set status_tra=1 where seq_tra='" & !seq_tra & "' "
            banco.Execute strSQL
            
            If .AbsolutePosition Mod 100 = 0 Then
                P_Status "Registro " & .AbsolutePosition & " de " & .RecordCount
            End If
            .MoveNext
        Loop
        MsgBox "Hora inicial " & t & " - Hora final " & Time
        .Close
    End With
End Sub

Private Sub cmndbtnGeraLog_Click()
    P_LogToText
End Sub

Private Sub cmndbtnMatriz_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim rPesq2 As Recordset
    strSQL = "select cod_loj from tab_lojas order by cod_loj "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = "select * from tab_transferencias where cod_loj = '" & !cod_loj & "' "
            strSQL = strSQL & " and (tab_tra='TAB_COMPRAS' or tab_tra='TAB_ITCOMPRAS' "
            strSQL = strSQL & "or tab_tra='TAB_CMP' or tab_tra='TAB_ITCMP'  "
            strSQL = strSQL & "or tab_tra='TAB_CMP' or tab_tra='TAB_ITCMP') "
            
            .MoveNext
        Loop
    End With
End Sub

Private Sub cmndbtnTransferir_Click()
    CommonDialog1.Filter = "*.stt"
    CommonDialog1.FileName = App.Path & "\*.stt"
    CommonDialog1.ShowOpen
    
        'MsgBox "Não foi informado um arquivo!", vbExclamation
        'Exit Sub
    
    P_TextToTransf CommonDialog1.FileName
End Sub

