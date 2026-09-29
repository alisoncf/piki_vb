VERSION 5.00
Begin VB.Form frmMesclaArea 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mesclar Áreas"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4140
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbxCidade 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnTransf 
      Caption         =   "&Transferir"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.ComboBox cmbxNovo 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
   Begin VB.ComboBox cmbxBairro 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cidade:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Novo Bairro:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro que será Transferido::"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1995
   End
End
Attribute VB_Name = "frmMesclaArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxBairro_Click()
    cmbxBairro.Tag = F_BuscaUmCampo("cod_are", "tab_areas", "cod_cid='" & cmbxCidade.Tag & "' and nome_Are='" & cmbxBairro.Text & "'")
    
End Sub

Private Sub cmbxCidade_Click()
    cmbxCidade.Tag = F_BuscaUmCampo("cod_cid", "tab_Cidades", "nome_cid='" & cmbxCidade.Text & "'")
    P_PreencheArea cmbxBairro, cmbxCidade.Tag
    P_PreencheArea cmbxNovo, cmbxCidade.Tag
End Sub

Private Sub cmbxNovo_Click()
    cmbxNovo.Tag = F_BuscaUmCampo("cod_are", "tab_areas", "cod_cid='" & cmbxCidade.Tag & "' and nome_Are='" & cmbxNovo.Text & "'")
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnTransf_Click()
    Dim strSQL As String
    If cmbxBairro.Text = "" Then
        MsgBox "Informe o bairro que será transferido!", vbCritical
        cmbxBairro.SetFocus
        Exit Sub
    End If
    If cmbxNovo.Text = "" Then
        MsgBox "Informe o novo Bairro!", vbCritical
        cmbxNovo.SetFocus
        Exit Sub
    End If
    If MsgBox("Confirma Transferência do Bairro: " & cmbxBairro.Text & _
        " Para o novo Bairro " & cmbxNovo.Text & "?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = "Update tab_bandidos set bairro_ban='" & cmbxNovo.Tag & "' "
        strSQL = strSQL & " where bairro_ban='" & cmbxBairro.Tag & "' "
        banco.Execute strSQL
        MsgBox banco.RecordsAffected & " Registros modificados!", vbInformation
        strSQL = "delete from tab_areas where cod_are='" & cmbxBairro.Tag & "' "
        banco.Execute strSQL
    
    End If
End Sub

Private Sub Form_Load()
    P_PreencheCidade cmbxCidade, "GO"
    Me.Top = 10
    Me.Left = 10
End Sub
