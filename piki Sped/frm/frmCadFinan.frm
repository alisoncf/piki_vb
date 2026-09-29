VERSION 5.00
Begin VB.Form frmCadFinan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Financeiras"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmCadFinan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3750
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.TextBox txtbxNome 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtbxTel 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   4
         Top             =   960
         Width           =   3495
      End
      Begin VB.ComboBox cmbxBco 
         Height          =   315
         ItemData        =   "frmCadFinan.frx":030A
         Left            =   120
         List            =   "frmCadFinan.frx":031A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome da Financeira:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fone/Fax:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   3735
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   855
         Left            =   120
         Picture         =   "frmCadFinan.frx":0351
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   855
         Left            =   1920
         Picture         =   "frmCadFinan.frx":049B
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCadFinan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxNome.Text = "" Then
        MsgBox "Informe o Nome da Financeira!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    If cmbxBco.ListIndex = -1 Then
        MsgBox "Informe o Tipo da Financeira!", vbExclamation
        cmbxBco.SetFocus
        Exit Sub
    End If
    If Me.tag = "" Then
        P_Grava
    Else
        P_Altera
    
    End If
End Sub
Sub P_Grava()
    Dim strSQL As String
    Dim sCod As String
    sCod = MaxKeySQL("tab_financeiras", "cod_fin", "", "0#")
    strSQL = "insert into tab_financeiras(cod_fin,tel_fin,tipo_fin,nome_fin)Values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxTel.Text & "',"
    strSQL = strSQL & "'" & Left(cmbxBco.Text, 1) & "',"
    strSQL = strSQL & "'" & txtbxNome.Text & "')"
    conn.Execute strSQL
    MsgBox "Financeira Gravada!", vbInformation
    If IsLoad("frmEdiFinan") Then
        frmEdiFinan.adc.Refresh
    End If
    txtbxNome.Text = ""
    txtbxTel.Text = ""
    cmbxBco.ListIndex = -1
    
    txtbxNome.SetFocus
End Sub
Sub P_Altera()
    strSQL = "update tab_financeiras set "
    
    strSQL = strSQL & " tel_Fin='" & txtbxTel.Text & "',"
    strSQL = strSQL & "tipo_fin='" & Left(cmbxBco.Text, 1) & "',"
    strSQL = strSQL & "nome_fin='" & txtbxNome.Text & "' "
    strSQL = strSQL & " where cod_fin='" & Me.tag & "' "
    conn.Execute strSQL
    MsgBox "Financeira Alterada!", vbInformation
    If IsLoad("frmEdiFinan") Then
        RefreshDataPosition frmEdiFinan.adc
    End If
    Unload Me
End Sub

Private Sub txtbxNome_GotFocus()
    ganhaFoco txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    AjustaTexto txtbxNome
End Sub

Private Sub txtbxTel_GotFocus()
    ganhaFoco txtbxTel
End Sub

Private Sub txtbxTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxTel_LostFocus()
    txtbxTel.Text = F_AjustaTelefone(txtbxTel.Text)
    
End Sub
