VERSION 5.00
Begin VB.Form frmCadCob 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Cobradores/Cobrança"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   480
      MaxLength       =   3
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.CheckBox chkVis 
      Caption         =   "Visível"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   3480
      Picture         =   "frmCadCob.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Salvar"
      Height          =   615
      Left            =   1680
      Picture         =   "frmCadCob.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   480
      MaxLength       =   40
      TabIndex        =   0
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome do Cobrador/Cobrança:"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   2145
   End
End
Attribute VB_Name = "frmCadCob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkVis_Click()
    If chkVis.Value = 1 Then
        chkVis.tag = "S"
    Else
        chkVis.tag = "N"
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If strCadcco_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub grava()
    Dim strSQL As String
    Dim max As String
    On Error GoTo trata_Erro
    max = txtCodigo.Text
    If max = "" Then
        max = MaxKeySQL("tab_contacob", "cod_cob", "", "00#")
    End If
    strSQL = "Insert into tab_contacob(cod_cob,nome_cob,excluido_cob)values("
    strSQL = strSQL & "'" & max & "',"
    strSQL = strSQL & "'" & txtbxNome.Text & "',"
    strSQL = strSQL & "'N')"
    Conn.Execute strSQL
    MsgBox "Gravado com sucesso!", vbInformation
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub altera()

Dim strSQL As String
    Dim max As String
    On Error GoTo trata_Erro
    
    strSQL = "update tab_contacob set "
    
    strSQL = strSQL & "nome_cob='" & txtbxNome.Text & "',"
    strSQL = strSQL & "excluido_cob='" & IIf(chkVis.Value = 1, "N", "S") & "' "
    strSQL = strSQL & " where cod_cob='" & Me.tag & "' "
    Conn.Execute strSQL
    MsgBox "Alterado com sucesso!", vbInformation
    
    Exit Sub
trata_Erro:
    P_MostraErro

End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_contacob where cod_cob='" & Me.tag & "'"
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            txtbxNome.Text = !nome_cob
            txtCodigo.Text = !cod_cob
            txtCodigo.Enabled = False
            If !excluido_cob = "S" Then
                chkVis.Value = 0
            Else
                chkVis.Value = 1
            End If
            chkVis.tag = !excluido_cob
        End If
        Me.caption = "Editanto Cobrador: " & !cod_cob & "-" & !nome_cob
    End With
    
End Sub
Private Sub Form_Activate()
    If bActive Then Exit Sub
    
    If Me.tag <> "" Then
        carrega
    Else
        txtCodigo.Text = MaxKeySQL("tab_contacob", "cod_cob", "", "00#")
    End If
    
End Sub

