VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso ao Piki Express"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000001&
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8040
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
      Left            =   2280
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3120
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
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
   End
   Begin VB.PictureBox pcLeft 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   0
      Picture         =   "frmLogin.frx":0E42
      ScaleHeight     =   1155
      ScaleWidth      =   8010
      TabIndex        =   8
      Top             =   0
      Width           =   8040
   End
   Begin VB.CommandButton cmndbtnSai 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   4800
      Picture         =   "frmLogin.frx":45A2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnEntr 
      Caption         =   "&Entrar"
      Height          =   735
      Left            =   2280
      Picture         =   "frmLogin.frx":526C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1395
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuário:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1260
      TabIndex        =   0
      Top             =   2400
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite seu Nome e Senha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   2685
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   5880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Piki Express "
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   5640
      TabIndex        =   6
      Top             =   7560
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
    If Me.tag = "" Then
        P_Fim
    Else
    
        Unload Me
    End If
End Sub

Private Sub cmndbtnEntr_Click()
    Dim strSQL As String
    Dim sUsu As String
    sUsu = F_BuscaUmCampo("Cod_fun", TAB_FUNCIONARIOS, "guerra_fun='" & txtbxNome_Usu.Text & "'")
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
    If Me.tag = "" Then
        If UCase(strSQL) = UCase(txtbxSenha.Text) Then

            strNome_Usu = txtbxNome_Usu.Text
            strNivel_Usu = F_BuscaUmCampo("nivel_usu", "tab_usuarios", "cod_fun='" & strCod_Usu & "'")
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
        If UCase(strSQL) = UCase(txtbxSenha.Text) Then
            Me.tag = sUsu
            strNome_Usu = txtbxNome_Usu.Text
            strNivel_Usu = F_BuscaUmCampo("nivel_usu", "tab_usuarios", "cod_fun='" & strCod_Usu & "'")
            Me.Hide
            
        Else
            MsgBox "Senha Incorreta!", vbCritical
            iTentativas = iTentativas + 1
            If iTentativas = 4 Then
                MsgBox "Autenticação Inválida!", vbCritical
                If Me.tag = "" Then
                    P_Fim
                Else
                    Me.tag = ""
                    Me.Hide
                End If
            End If
            Exit Sub
        End If
    End If
    
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

Private Sub Form_Load()
    BuscaPic
     
End Sub
Sub BuscaPic()
    On Error GoTo trata_erro
'    pcLeft.Picture = LoadPicture(App.Path & "\pic\logo.jpg")
    Exit Sub
trata_erro:
    pcLeft.Picture = Nothing
End Sub

Private Sub txtbxNome_Usu_GotFocus()
    ganhaFoco txtbxNome_Usu
End Sub

Private Sub txtbxNome_Usu_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        txtbxSenha.SetFocus
    End If
End Sub

Private Sub txtbxNome_Usu_LostFocus()
    AjustaTexto txtbxNome_Usu
End Sub

Private Sub txtbxSenha_GotFocus()
    ganhaFoco txtbxSenha
End Sub

Private Sub txtbxSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxSenha_LostFocus
        cmndbtnEntr_Click
    End If
End Sub

Private Sub txtbxSenha_LostFocus()
    AjustaTexto txtbxSenha
End Sub


