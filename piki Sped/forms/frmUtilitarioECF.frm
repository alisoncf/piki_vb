VERSION 5.00
Begin VB.Form frmUtilitarioECF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alíquotas de ICMS da ECF"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Adicionar Nova Alíquota:"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   5895
      Begin VB.CommandButton cmndbtnAddAliq 
         Caption         =   "&Adicionar"
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "I&SS"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&ICMS"
         Height          =   195
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtbxAliq 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "0000"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alíquota:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.TextBox txtbxRetorno 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmndbtnAliq 
      Caption         =   "&Alíquotas"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmUtilitarioECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnAddAliq_Click()
    Dim strSQL As String
    Dim i As Integer
    strSQL = "Uma vez adicionada uma alíquota, ela só pode ser retirada via "
    strSQL = strSQL & "intervenção técnica. Deseja Realmente Continuar?"
    If MsgBox(strSQL, vbCritical + vbYesNo) = vbYes Then
        If strImprFiscal = "D" Then
            i = Daruma_FI_ProgramaAliquota(txtbxAliq.Text, 0)
        Else
            i = Bematech_FI_ProgramaAliquota(txtbxAliq.Text, 0)
        End If
        
        cmndbtnAliq_Click
        
    End If
End Sub

Private Sub cmndbtnAliq_Click()
    Dim c As String
    Dim i As Integer
    c = Space(80)
    If strImprFiscal = "D" Then
        i = Daruma_FI_RetornoAliquotas(c)
        txtbxRetorno.Text = F_QuebraAliq(c)
    Else
        i = Bematech_FI_RetornoAliquotas(c)
        txtbxRetorno.Text = F_QuebraAliq(c)
        
    End If
End Sub
Function F_QuebraAliq(c As String)
    Dim sAux As String
    Dim sBuf As String
    Dim sMsg As String
    Dim iIndice As Integer
    sMsg = "Indice|Alíquota  |"
    sMsg = sMsg & "-----+----------+" & Chr(13) + Chr(10)
    For iCount = 1 To Len(c) - 1
        sBuf = Mid(c, iCount, 1)
        If sBuf = "," Then
            iIndice = iIndice + 1
            sMsg = sMsg & "  " & Format(iIndice, "0#") & "  |  " & Left(sAux, 2) & " %" & Chr(13) + Chr(10)
            sAux = ""
        Else
            
            sAux = sAux & sBuf
            
        End If
    Next
    sMsg = sMsg & sAux
    F_QuebraAliq = sMsg
End Function

Private Sub txtbxAliq_GotFocus()
    ganhaFoco txtbxAliq
End Sub

Private Sub txtbxAliq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxAliq_LostFocus()
    Dim Dec As String
    Dim Uni As String
     
    PerdeFoco txtbxAliq
    Dec = Format(Int(txtbxAliq.Text), "0#")
    Uni = CCur(txtbxAliq.Text) - CCur(Dec)
    Uni = Uni * 100
    Uni = Format(Uni, "0#")
    txtbxAliq.Text = Dec & Uni
End Sub
