VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Acesso ao PIKI EXPRESS"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   ClipControls    =   0   'False
   ForeColor       =   &H80000001&
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcSair 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4080
      ScaleHeight     =   855
      ScaleWidth      =   1455
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
      Begin VB.PictureBox pcCancel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   600
         Picture         =   "frmLogin.frx":0E42
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   630
      End
   End
   Begin VB.PictureBox pcEntrar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1680
      ScaleHeight     =   855
      ScaleWidth      =   2175
      TabIndex        =   6
      Top             =   3840
      Width           =   2175
      Begin VB.PictureBox pcOk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   960
         Picture         =   "frmLogin.frx":3209
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblEntrar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrar"
         Height          =   195
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   6540
      TabIndex        =   5
      Top             =   0
      Width           =   6570
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   0
         Picture         =   "frmLogin.frx":3593
         ScaleHeight     =   690
         ScaleWidth      =   1635
         TabIndex        =   12
         Top             =   0
         Width           =   1635
      End
   End
   Begin VB.TextBox txtbxNome_Usu 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txtbxSenha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   4575
      Left            =   0
      Top             =   720
      Width           =   6570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite seu Nome e Senha"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   6000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuário...:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha.....:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   915
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iTentativas As Byte
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub cmndbtnEntr_Click()
    Dim strSQL As String
    Dim sUsu As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select tu.cod_fun,status_usu,senha_usu from tab_funcionarios as tf inner join tab_usuarios as tu "
    strSQL = strSQL & " on tf.cod_fun=tu.cod_fun "
    strSQL = strSQL & " where (guerra_fun='" & txtbxNome_Usu.Text & "' "
    strSQL = strSQL & " or tu.cod_fun='" & txtbxNome_Usu.Text & "') "
   ' strSQL = strSQL & " and status_usu=1 "
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    If rPesq.RecordCount = 0 Then
        MsgBox "Este usuário não existe!", vbExclamation
        Exit Sub
    ElseIf rPesq.RecordCount > 1 Then
        MsgBox "existem " & rPesq.RecordCount & " usários com este nome", vbExclamation
        Exit Sub
    
    Else
        strCod_Usu = rPesq!cod_fun
        If rPesq!status_usu <> "1" Then
            MsgBox "O usuário não está habilitado para uso!", vbExclamation
            P_Fim
        End If
    End If
    
    
    
    strSQL = F_RetornaPW(rPesq!senha_usu)
    If Me.tag = "" Then
        If UCase(strSQL) = UCase(txtbxSenha.Text) Then
            strNome_Usu = txtbxNome_Usu.Text
            strCod_Car = F_BuscaUmCampo("cod_car", "tab_funcionarios", "cod_fun='" & strCod_Usu & "'")
            Unload Me
        Else
            MsgBox "Senha Incorreta!", vbCritical
            txtbxSenha.SetFocus
            iTentativas = iTentativas + 1
            If iTentativas = 4 Then
                MsgBox "Autenticação Inválida!", vbCritical
                If Not IsLoad("frmVenda") Then
                    P_Fim
                End If
            End If
            Exit Sub
        End If
    Else
        If strSQL = txtbxSenha.Text Then
            Me.tag = strCod_Usu
'            strNome_Usu = txtbxNome_Usu.Text
            Me.Hide
            
        Else
            MsgBox "Senha Incorreta!", vbCritical
            iTentativas = iTentativas + 1
            If iTentativas = 4 Then
                MsgBox "Autenticação Inválida!", vbCritical
                If Me.tag = "" Then
                    P_Fim
                End If
            End If
            Exit Sub
        End If
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnSai_Click()
    If Me.tag = "" Then
        P_Fim
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmndbtnSai_Click
    End If
End Sub

Private Sub Image2_Click()
 
End Sub

Private Sub Label3_Click()
    cmndbtnSai_Click
End Sub

Private Sub lblEntrar_Click()
    cmndbtnEntr_Click
End Sub

Private Sub pcCancel_Click()
    cmndbtnSai_Click
End Sub

Private Sub pcEntrar_Click()
    cmndbtnEntr_Click
End Sub

Private Sub pcOk_Click()
    cmndbtnEntr_Click
End Sub

Private Sub pcSair_Click()
    cmndbtnSai_Click
End Sub

Private Sub txtbxNome_Usu_GotFocus()
    P_SelecTexto txtbxNome_Usu
End Sub

Private Sub txtbxNome_Usu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Usu_LostFocus()
    P_AjustaTexto txtbxNome_Usu
End Sub

Private Sub txtbxSenha_GotFocus()
    P_SelecTexto txtbxSenha
End Sub

Private Sub txtbxSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxSenha_LostFocus
        cmndbtnEntr_Click
    End If
End Sub

Private Sub txtbxSenha_LostFocus()
    P_AjustaTexto txtbxSenha
End Sub


