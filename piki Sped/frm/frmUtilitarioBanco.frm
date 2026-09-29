VERSION 5.00
Begin VB.Form frmUtilitarioBanco 
   Caption         =   "Utilitário de Banco de Dados"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxData 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "31/12/2008"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnZerarBanco 
      Caption         =   "Zerar Banco"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmUtilitarioBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnZerarBanco_Click()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim sNumVen As String
    sNumVen = F_BuscaUmCampo("max(num_ven)", "" & TAB_VENDAS & "", "situacao_ven=4 and dataefet_ven<=" & F_GravaData(txtbxData.Text))
    If sNumVen = "" Then
        MsgBox "Não encontrado!", vbExclamation
    End If
    
    strSQL = "delete from " & TAB_VENDAS & " where num_ven<='" & sNumVen & "'"
    conn.Execute strSQL
    
    strSQL = "delete from " & TAB_ITVENDAS & " where num_ven<='" & sNumVen & "'"
    conn.Execute strSQL
    
    strSQL = "delete from tab_caixa where num_ven<='" & sNumVen & "' and num_ven<>'' and data_cai<" & F_GravaData(Date)
    conn.Execute strSQL
    
    strSQL = "delete from  " & TAB_VENFEC & "  where num_ven<='" & sNumVen & "' "
    conn.Execute strSQL
    strSQL = "delete from tab_venfin where num_ven<='" & sNumVen & "' and num_ven<>''"
    conn.Execute strSQL
    
    
    
    
End Sub

Private Sub txtbxData_GotFocus()
    ganhaFoco txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub
