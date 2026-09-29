VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:Piki Express - BOX - Senha:."
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ClipControls    =   0   'False
   ForeColor       =   &H80000001&
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbxSenha 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox txtbxNome_Usu 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1200
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   4980
      Left            =   0
      Picture         =   "frmLogin.frx":0E42
      ScaleHeight     =   4920
      ScaleWidth      =   2385
      TabIndex        =   8
      Top             =   0
      Width           =   2445
   End
   Begin VB.CommandButton cmndbtnSai 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5160
      Picture         =   "frmLogin.frx":100A6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnEntr 
      Caption         =   "&Confirmar"
      Height          =   615
      Left            =   3720
      Picture         =   "frmLogin.frx":10D70
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha.....:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   2
      Top             =   2040
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuário...:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   1470
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
      Height          =   345
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   6840
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Piki Express "
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   5520
      TabIndex        =   6
      Top             =   4680
      Width           =   900
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
    If Me.Tag = "" Then
        P_Fim
    Else
    
        Unload Me
    End If
End Sub

Private Sub cmndbtnEntr_Click()
    Dim strSQL As String
    Dim sUsu As String
    sUsu = F_BuscaUmCampo("Cod_fun", "Tab_funcionarios", "guerra_fun='" & txtbxNome_Usu.Text & "'")
    sUsu = F_BuscaUmCampo("Cod_fun", "Tab_Usuarios", "cod_fun='" & sUsu & "'")
    If sUsu = "" Then
        MsgBox "Este usuário não existe!", vbExclamation
        Exit Sub
    Else
        strCod_Usu = sUsu
    End If
    strSQL = F_BuscaUmCampo("Senha_Usu", "Tab_Usuarios", " cod_fun='" & sUsu & "'")
    strSenha_Par = strSQL
    strSQL = F_RetornaPW(strSQL)
    If Me.Tag = "" Then
        If UCase(strSQL) = UCase(txtbxSenha.Text) Then

            strNome_Usu = txtbxNome_Usu.Text
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
            Me.Tag = sUsu
            strNome_Usu = txtbxNome_Usu.Text
            Me.Hide
            
        Else
            MsgBox "Senha Incorreta!", vbCritical
            iTentativas = iTentativas + 1
            If iTentativas = 4 Then
                MsgBox "Autenticação Inválida!", vbCritical
                If Me.Tag = "" Then
                    P_Fim
                Else
                    Me.Tag = ""
                    Me.Hide
                End If
            End If
            Exit Sub
        End If
    End If
    
End Sub

Private Sub cmndbtnSai_Click()
    If Me.Tag = "" Then
        P_Fim
    Else
        Me.Tag = ""
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

Private Sub txtbxNome_Usu_GotFocus()
    P_SelecTexto txtbxNome_Usu
End Sub

Private Sub txtbxNome_Usu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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


