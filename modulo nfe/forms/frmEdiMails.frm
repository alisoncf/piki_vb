VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEdiMails 
   Caption         =   "Configuração de Emails"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   Icon            =   "frmEdiMails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   13290
      TabIndex        =   1
      Top             =   5115
      Width           =   13320
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Selecionar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   5160
         Picture         =   "frmEdiMails.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiMails.frx":10D4C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "Alterar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmEdiMails.frx":110D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "E&xcluir"
         Height          =   735
         Left            =   3480
         Picture         =   "frmEdiMails.frx":11220
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   11520
         Picture         =   "frmEdiMails.frx":1136A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMail 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8281
      _Version        =   393216
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmEdiMails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ado As ADODB.Recordset
Private Sub cmndbtnSair_Click()

End Sub

Private Sub cmdAlterar_Click()
    gridMail.col = 0
    frmNovoMail.tag = gridMail.Text
    frmNovoMail.Show 1
    busca
End Sub

Private Sub cmdNovo_Click()
    frmNovoMail.Show 1
    busca
End Sub

Private Sub cmdSair_Click()
    If Me.tag = "" Then
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo trata_Erro
    Me.tag = ado!id
    Me.Hide
trata_Erro:
End Sub

Private Sub Form_Activate()
    busca
    If Me.tag <> "" Then
        cmdSelect.Enabled = True
    Else
        cmdSelect.Enabled = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        busca
    End If
End Sub
Sub busca()
    Dim strSQL As String
    Dim colunas As String
    Dim tam As String
    Dim titulo As String
    strSQL = "select * from sendmail order by id "
    abreObjetoAdo ado, strSQL
    colunas = "id, nome, email, smtpserver, smtpserverport, subject"
    
    titulo = "id, nome, email, SMTP, Porta, Assunto"
    tam = "500,1500,2500,2500,1000,3000"
    AjustaGridData gridMail, 6, 1, 1, titulo, colunas, tam, ado, "", True
    
End Sub
Private Sub Form_Load()
    Set ado = New ADODB.Recordset
End Sub
