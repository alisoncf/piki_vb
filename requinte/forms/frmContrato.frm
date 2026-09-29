VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmContrato 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contrato de Cliente:"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Height          =   7320
      Left            =   8445
      ScaleHeight     =   7260
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnEditarClau 
         Caption         =   "Editar Cláusula"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnGravar 
         Caption         =   "&Gravar"
         Height          =   495
         Left            =   120
         Picture         =   "frmContrato.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnNova 
         Caption         =   "Nova Cláusula"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtImprimir 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6480
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   1320
      ScaleHeight     =   2625
      ScaleWidth      =   6945
      TabIndex        =   8
      Top             =   3240
      Width           =   6975
      Begin VB.TextBox txtbxTextoCla 
         Height          =   2055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   6615
      End
      Begin VB.Label lblTitulo_Cla 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   16
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.PictureBox pcContrato 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   8145
      TabIndex        =   1
      Top             =   480
      Width           =   8175
      Begin VB.TextBox txtbxPreAmbulo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   6975
      End
      Begin VB.TextBox txtbxCabCon 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   120
         Width           =   6975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preêmbulo:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Data dtServ 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmContrato.frx":1CFA
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "frmContrato.frx":1D0F
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cláusulas:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1230
   End
   Begin VB.Label lblContrato 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub Command4_Click()

End Sub

Private Sub cmndbtImprimir_Click()
    On Error GoTo trata_Erro
    imprimeContrato adc.Recordset!seq_con
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmndbtnEditarClau_Click()
    If Me.tag = "" Then
        MsgBox "Primeiro você dfeve gravar o contrato!", vbExclamation
        Exit Sub
    End If
    frmClausula.tag = Me.tag
    frmClausula.txtbxCab.text = adc.Recordset!cab_cla
    frmClausula.txtbxCab.tag = adc.Recordset!seq_cla
    frmClausula.txtbxTexto.text = txtbxTextoCla.text
    frmClausula.Show 1
    buscaCla
End Sub

Private Sub cmndbtnGravar_Click()
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub grava()
    Dim strSQL As String
    Dim sSeq As String
    sSeq = MaxKeySQL("tab_contratos", "seq_con", "", "")
    strSQL = "insert into tab_contratos(seq_con,cab_con,preambulo_con)values("
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & txtbxCabCon.text & "',"
    strSQL = strSQL & "'" & txtbxPreAmbulo.text & "')"
    Conn.Execute strSQL
    Me.tag = sSeq
    MsgBox "Contrato Salvo com sucesso!", vbInformation
    
End Sub
Private Sub cmndbtnNova_Click()
    If Me.tag = "" Then
        MsgBox "Primeiro você dfeve gravar o contrato!", vbExclamation
        Exit Sub
    End If
    frmClausula.tag = Me.tag
    frmClausula.Show 1
    
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()
    Unload Me
End Sub

Sub buscaClausula()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
  '  On Error Resume Next
    strSQL = "select * from tab_clausulas where seq_con=" & Me.tag
    strSQL = strSQL & " and seq_cla=" & adc.Recordset!seq_cla
    abreObjetoAdo rPesq, strSQL
    With rPesq
        lblTitulo_Cla.caption = !cab_cla
        txtbxTextoCla.text = !texto_cla
        .Close
    End With
End Sub
Private Sub Form_Activate()
    
    buscaCla
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
End Sub
Sub altera()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_contratos set cab_con='" & txtbxCabCon.text & "',"
    strSQL = strSQL & "preambulo_con='" & txtbxPreAmbulo.text & "' "
    strSQL = strSQL & " where seq_con=" & Me.tag & " "
    Conn.Execute strSQL
    MsgBox "Contrato alterado!", vbInformation
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub carrega()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    'on error resume next
    strSQL = "select * from tab_contratos where seq_con=" & Me.tag
    abreObjetoAdo rPesq, strSQL
    With rPesq
        lblContrato.caption = !cab_con
        txtbxCabCon.text = !cab_con
        txtbxPreAmbulo.text = !preambulo_con
        .Close
    End With
End Sub
Private Sub Form_Load()
    adc.ConnectionString = Conn.ConnectionString
End Sub
Sub buscaCla()
    Dim strSQL As String
    strSQL = "select * from tab_clausulas where seq_con='" & Me.tag & "' "
    adc.RecordSource = strSQL
    adc.Refresh
End Sub

Private Sub txtbxCabCon_GotFocus()
    P_SelecTexto txtbxCabCon
End Sub

Private Sub txtbxCabCon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCabCon_LostFocus()
    P_AjustaTexto txtbxCabCon, -1
End Sub

Private Sub txtbxPreAmbulo_LostFocus()
    P_AjustaTexto txtbxPreAmbulo, -1
End Sub
