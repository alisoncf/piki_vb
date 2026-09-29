VERSION 5.00
Begin VB.Form frmContAvanc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contagem Avançada"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "frmContAvanc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   4440
      Picture         =   "frmContAvanc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Fecha formulário."
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnEfetiva 
      Caption         =   "&Confirmar"
      Height          =   495
      Left            =   2640
      Picture         =   "frmContAvanc.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "cria um novo inventário."
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkVendas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "&Abater vendas:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblAte 
         Caption         =   "&Até:"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDe 
         Caption         =   "&De:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opções Avançadas:"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton optnProd 
         Caption         =   "Ze&rar apenas os produtos da contagem"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   3135
      End
      Begin VB.OptionButton optnZera 
         Caption         =   "&Zerar todos os produtos"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label lblContagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contagem"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   720
   End
End
Attribute VB_Name = "frmContAvanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkVendas_Click()
    Dim strSQL As String
    strSQL = "Se esta opção for escolhida, a soma da quantidades vendidas no período "
    strSQL = strSQL & " de cada produto na venda será abatida nos produtos "
    
    MsgBox strSQL, vbInformation
End Sub

Private Sub cmndbtnEfetiva_Click()
    Dim strSQL As String
    If lblContagem.Tag = "" Then
        MsgBox "Não há uma contagem válida!", vbExclamation
        Exit Sub
    End If
    strSQL = "Efetivar a Contagem implica em alterar o estoque de todos os Produtos "
    strSQL = strSQL & " nela contidos para suas quantidades contadas, respectivamente."
    strSQL = strSQL & Chr(13) & "Deseja Realmente Efetivar a Contagem: " & lblContagem.Tag
    strSQL = strSQL & " ?"
    If MsgBox(strSQL, vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
        efetivaCont
        
    End If
    P_AjustaEstoqueReservado
End Sub
Sub efetivaCont()
    Dim strSQL
    Dim rPesq As Recordset
    strSQL = "select * from tab_itCont where seq_cnt='" & lblContagem.Tag & "' "
    strSQL = strSQL & " order by cod_pro "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        P_Status "Efetivando contagem"
        If optnZera.Value = True Then
            strSQL = "update tab_composicao set e2_pro=0,r_pro=0 "
            banco.Execute strSQL
        End If
        Do While Not .EOF
            
            strSQL = "update tab_composicao set e2_pro=" & F_GravaDec(!qtd_itc) & ","
            strSQL = strSQL & "r_pro=0 "
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            If .RecordCount Mod 13 = 0 Then
                P_Status .AbsolutePosition + 1 & " de " & .RecordCount
            End If
            banco.Execute strSQL
            .MoveNext
        Loop
        If chkVendas.Value = 1 Then
            abate_venda
        End If
        strSQL = "update tab_contagens set status_cnt='S' where seq_cnt='"
        strSQL = strSQL & lblContagem.Tag & "'"
        banco.Execute strSQL
        P_Status ""
        MsgBox "Contagem Efetivada!", vbExclamation
        .Close
        cmndbtnEfetiva.Enabled = False
    End With
End Sub
Sub abate_venda()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select sum(ti.qtd_ite) as total from tab_vendas as tv,tab_itvendas as ti, "
    strSQL = strSQL & "tab_itcont as tic where tic.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tic.seq_cnt='" & lblContagem.Tag & "' "
    strSQL = strSQL & " and dataefet_ven>=" & F_PesqData(txtbxDe.Text)
    strSQL = strSQL & " and dataefet_ven<=" & F_PesqData(txtbxAte.Text)
    strSQL = strSQL & " and situacao_ven  =4 "
    strSQL = strSQL & " group by ti.cod_pro order by ti.cod_Pro  "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_composicao set e2_pro=e2_pro-" & F_GravaDec(!Total)
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            banco.Execute strSQL
            .MoveNext
        Loop
        
        .Close
    End With
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub optnZera_Click()
    Dim strSQL As String
    strSQL = "Se esta opção for escolhida, o estoque de todos os produtos"
    strSQL = strSQL & " será zerado, ou seja, os produtos que por ventura não tiverem"
    strSQL = strSQL & " sido contados também serão zerados."
    MsgBox strSQL, vbInformation
End Sub

Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxDe_GotFocus()
        P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub
