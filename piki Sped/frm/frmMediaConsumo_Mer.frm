VERSION 5.00
Begin VB.Form frmMediaConsumo_Mer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Média de Consumo dos Produtos"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmMediaConsumo_Mer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5040
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox cmbxForn 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2520
         Width           =   4575
      End
      Begin VB.CommandButton cmndbtnCalcula 
         Caption         =   "&Calcular"
         Height          =   735
         Left            =   120
         Picture         =   "frmMediaConsumo_Mer.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3720
         Width           =   1695
      End
      Begin VB.ComboBox cmbxMar 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1800
         Width           =   4575
      End
      Begin VB.ComboBox cmbxCat 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Sair"
         Height          =   735
         Left            =   3000
         Picture         =   "frmMediaConsumo_Mer.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtbxDataFin 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Fornecedor:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Marca"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Tag             =   "!"
         Top             =   3000
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "Data &Final"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "&Data Inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   345
         Left            =   120
         Picture         =   "frmMediaConsumo_Mer.frx":06D6
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   15
      End
   End
End
Attribute VB_Name = "frmMediaConsumo_Mer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxGru_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxGru.ListIndex = -1
    End If
End Sub

Private Sub cmbxCat_Click()
    cmbxCat.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxForn_Click()
    cmbxForn.tag = F_BuscaUmCampo("cod_for", TAB_FORNECEDORES, "nome_for='" & cmbxForn.Text & "'")
End Sub

Private Sub cmbxForn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxForn.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxMar_Click()
    cmbxMar.tag = F_BuscaUmCampo("cod_mar", TAB_MARCAS, "nome_mar='" & cmbxMar.Text & "'")
End Sub

Private Sub cmbxMar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxMar.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxSub.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnCalcula_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCat As String
    Dim sFor As String
    Dim sMar As String
    If F_Vazio(txtbxDAtaFIn.Text) Then
        MsgBox "Informe a data final!", vbExclamation
        txtbxDAtaFIn.SetFocus
        Exit Sub
    ElseIf F_Vazio(txtbxDataIni.Text) Then
        MsgBox "Informe a data Inicial!", vbExclamation
        txtbxDataIni.SetFocus
        Exit Sub
    End If
    
    If DateDiff("D", txtbxDataIni.Text, txtbxDAtaFIn.Text) < 14 Then
        MsgBox "A Diferença entre as Datas deve ser maior ou superior a 15", vbExclamation
        txtbxDAtaFIn.SetFocus
        Exit Sub
    End If
    
    Label3.Visible = True
    Image1.Visible = True
    
    P_MediaDiaria txtbxDataIni.Text, txtbxDAtaFIn.Text, cmbxMar.tag, cmbxCat.tag, "", cmbxForn.tag
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_Load()
    P_PreencheCateg cmbxCat
    P_PreencheMarca cmbxMar
    P_PreencheFornecedores cmbxForn
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
    
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
