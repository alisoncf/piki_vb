VERSION 5.00
Begin VB.Form frmGerImpNotaSai 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aguardando Nota Fiscal de Saída"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aguardando Nota Fiscal de Saída"
      Height          =   195
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmGerImpNotaSai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_saidasnf where situacao_sai=1 order by seq_sai desc limit 1 "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            If MsgBox("Impressão de Nota Fiscal de Saída. Tecle OK quando a impressora estiver pronta", vbOKCancel) = vbOK Then
                p_imprNota_SaiLPT !seq_sai
            
                .MoveNext
            Else
                .Close
                Unload Me
            End If
        Loop
        .Close
    End With
End Sub
