VERSION 5.00
Begin VB.Form frmIncCon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Contas"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
   Icon            =   "frmIncCon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cmbxBco 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtbxTit 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtbxAge 
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtbxNum 
         Height          =   285
         Left            =   1440
         MaxLength       =   18
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtbxNome 
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Banco:"
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Titular:"
         Height          =   195
         Left            =   870
         TabIndex        =   9
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Agência:"
         Height          =   255
         Left            =   255
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N&umero:"
         Height          =   195
         Left            =   750
         TabIndex        =   3
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome da Conta"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   5175
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   3360
         Picture         =   "frmIncCon.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmIncCon.frx":0475
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmIncCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxAge.Text = "" Then
        MsgBox "Informe a agência!", vbExclamation
        txtbxAge.SetFocus
        Exit Sub
    End If
    If txtbxNome.Text = "" Then
        MsgBox "Informe um nome para a Conta!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    If txtbxNum.Text = "" Then
        MsgBox "Informe o  número da Conta!", vbExclamation
        txtbxNum.SetFocus
        Exit Sub
    End If
    If txtbxTit.Text = "" Then
        MsgBox "Informe o nome do Titular!", vbExclamation
        txtbxTit.SetFocus
        Exit Sub
    End If
    
    If Me.tag = "" Then
        If F_BuscaUmCampo("Nome_Con", "Tab_Contas", "nome_con='" & txtbxNome.Text & "'") <> "" Then
        MsgBox "Já existe uma conta com este nome!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
        If F_BuscaUmCampo("Nome_Con", "Tab_Contas", "nome_con='" & _
            txtbxNome.Text & "' and cod_con <> '" & Me.tag & "'") <> "" Then
        MsgBox "Já existe uma conta com este nome!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
        Grava
    Else
        Altera
    End If
End Sub
Sub Grava()
    Dim strSQL As String
    Dim sCod As String
    
    sCod = F_BuscaUmCampo("Max(Cod_Con)", "Tab_Contas", "")
    If sCod = "" Then
        sCod = "002"
    Else
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    strSQL = "insert into tab_Contas(Cod_Con,Nome_Con,Num_Con,"
    strSQL = strSQL & "Cod_Bco,Age_Con,Tit_Con,status_con)Values("
    strSQL = strSQL & "'" & sCod & "', "
    strSQL = strSQL & "'" & txtbxNome.Text & "',"
    strSQL = strSQL & "'" & txtbxNum.Text & "',"
    strSQL = strSQL & "'" & Left(cmbxBco.Text, 3) & "',"
    strSQL = strSQL & "'" & txtbxAge.Text & "',"
    strSQL = strSQL & "'" & txtbxTit.Text & "','H') "
    Conn.Execute strSQL
    MsgBox "Conta Gravada!", vbExclamation
    txtbxAge.Text = ""
    txtbxNome.Text = ""
    txtbxNum.Text = ""
    cmbxBco.ListIndex = -1
    txtbxNome.SetFocus
End Sub

Sub Altera()
    Dim strSQL As String
    strSQL = "update tab_Contas set "
    
    strSQL = strSQL & " nome_Con = '" & txtbxNome.Text & "',"
    strSQL = strSQL & " num_con = '" & txtbxNum.Text & "',"
    strSQL = strSQL & " cod_bco = '" & Left(cmbxBco.Text, 3) & "',"
    strSQL = strSQL & "age_con = '" & txtbxAge.Text & "',"
    strSQL = strSQL & "tit_con = '" & txtbxTit.Text & "'  "
    strSQL = strSQL & " where cod_Con='" & Me.tag & "'"
    Conn.Execute strSQL
    MsgBox "Conta alterada!", vbInformation
    Unload Me
End Sub
Private Sub Form_Load()
'    Me.Height = 3315
 '   Me.Width = 3945
    P_CentraForm Me
    P_PreencheBanco cmbxBco
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsLoad("frmEdiCon") Then
        frmEdiCon.adc.Refresh
    End If
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub
Private Sub txtbxAge_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxAge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAge_LostFocus()
    P_AjustaTexto txtbxNome
End Sub
Private Sub txtbxNum_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNum_LostFocus()
    P_AjustaTexto txtbxNome
End Sub
Private Sub txtbxTit_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxTit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxTit_LostFocus()
    P_AjustaTexto txtbxNome
End Sub

