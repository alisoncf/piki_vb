VERSION 5.00
Begin VB.Form frmNovoMail 
   Caption         =   "Nova Configuração de Email"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9300
   Icon            =   "frmNovoMail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBody 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   3960
      Width           =   9015
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   9015
   End
   Begin VB.TextBox txtSMTPUseSSL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8160
      TabIndex        =   21
      Text            =   "1"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtSendPassWord 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   19
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox txtSendUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   4695
   End
   Begin VB.TextBox txtSMTPAuthenticate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8760
      TabIndex        =   15
      Text            =   "1"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtSMTPPort 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8160
      TabIndex        =   13
      Text            =   "465"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtSMTPServer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5880
      TabIndex        =   11
      Text            =   "smtp.gmail.com"
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtSch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Text            =   "http://schemas.microsoft.com/cdo/configuration/"
      Top             =   1800
      Width           =   4695
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   9270
      TabIndex        =   3
      Top             =   5325
      Width           =   9300
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Salvar"
         Height          =   735
         Left            =   6000
         Picture         =   "frmNovoMail.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   7560
         Picture         =   "frmNovoMail.frx":10CF2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.TextBox txtSendUsing 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   2
      Text            =   "2"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   780
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   585
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use SSL"
      Height          =   195
      Left            =   8160
      TabIndex        =   22
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   4920
      TabIndex        =   20
      Top             =   2280
      Width           =   690
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auth"
      Height          =   195
      Left            =   8760
      TabIndex        =   16
      Top             =   1560
      Width           =   330
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      Height          =   195
      Left            =   8160
      TabIndex        =   14
      Top             =   1560
      Width           =   285
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP Server"
      Height          =   195
      Left            =   5880
      TabIndex        =   12
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SendUsing"
      Height          =   195
      Left            =   4920
      TabIndex        =   10
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schema"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmNovoMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m As Mail

Private Sub cmndbtnSair_Click()
     Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If strNivel_Usu < 3 Then
        MsgBox "Não é possível acessar essas informações.", vbExclamation
        Exit Sub
    End If
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub buscaM()
    
    m.body = txtBody.Text
    m.email = txtEmail.Text
    m.id = Me.tag
    m.nome = txtNome.Text
    m.sch = txtSch.Text
    m.sendpassword = txtSendPassWord.Text
    m.sendusername = txtSendUserName.Text
    m.sendusing = txtSendUsing.Text
    m.smtpauthenticate = txtSMTPAuthenticate.Text
    m.smtpserver = txtSMTPServer.Text
    m.smtpserverport = txtSMTPPort.Text
    m.smtpusessl = txtSMTPUseSSL.Text
    m.subject = txtSubject.Text
    
End Sub
Sub grava()
    buscaM
    If SendMail.GravaMail(m) Then
        MsgBox "Email Gravado!", vbInformation
    End If
End Sub
Sub altera()
    buscaM
    If SendMail.GravaMail(m) Then
        MsgBox "Email Gravado!", vbInformation
    End If
    
End Sub
Sub carrega()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select * from sendmail where id=" & Me.tag & " "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            txtNome.Text = !nome
            txtEmail.Text = !email
            txtBody.Text = TiraNull(!body)
            txtSch.Text = TiraNull(!sch)
            txtSendPassWord.Text = TiraNull(!sendpassword)
            txtSendUserName.Text = TiraNull(!sendusername)
            txtSendUsing.Text = TiraNull(!sendusing)
            txtSMTPAuthenticate.Text = TiraNull(!smtpauthenticate)
            txtSMTPPort.Text = TiraNull(!smtpserverport)
            txtSMTPServer.Text = TiraNull(!smtpserver)
            txtSMTPUseSSL.Text = TiraNull(!smtpusessl)
            txtSubject.Text = TiraNull(!subject)
        End If
        .Close
    End With
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        carrega
    End If
End Sub

