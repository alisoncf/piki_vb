VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmImpostos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impostos"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   Icon            =   "frmImpostos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3540
   Begin VB.Frame frImp 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtbxNome_Imp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblNome_Imp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome do Imposto"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1245
      End
   End
   Begin VB.Data dtImp 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmndbtnEdit 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   960
      Picture         =   "frmImpostos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin MSDBGrid.DBGrid dbgrdImp 
      Bindings        =   "frmImpostos.frx":058C
      Height          =   1335
      Left            =   120
      OleObjectBlob   =   "frmImpostos.frx":05A0
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmbxNovo 
      Caption         =   "&Novo"
      Height          =   735
      Left            =   120
      Picture         =   "frmImpostos.frx":0F7B
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   735
      Left            =   2640
      Picture         =   "frmImpostos.frx":1305
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1800
      Picture         =   "frmImpostos.frx":1747
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblImp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 Impostos Encontrados"
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   907
      TabIndex        =   4
      Top             =   2280
      Width           =   1710
   End
End
Attribute VB_Name = "frmImpostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub P_Salva()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sCod As String
    strSQL = "select max(cod_imp) as maximo from tab_impostos"
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If IsNull(rPesq!maximo) Then
        sCod = "01"
    Else
        sCod = Format(rPesq!maximo + 1, "0#")
    End If
    strSQL = "insert into tab_impostos(cod_imp,nome_imp) values ("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxNome_Imp.Text & "') "
    banco.Execute strSQL
    MsgBox "Imposto cadastrado", vbInformation
    dtImp.Refresh
    txtbxNome_Imp.Text = ""
    
End Sub
Sub P_Altera()
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_erro
    strSQL = "update tab_impostos set nome_imp = '" & txtbxNome_Imp.Text & "' "
    strSQL = strSQL & " where cod_imp = '" & Me.Tag & "' "
    banco.Execute strSQL
    dtImp.Refresh
    MsgBox "Imposto alterado!", vbInformation
    Exit Sub
    
trata_erro:
        MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmbxNovo_Click()
    frImp.Enabled = True
    lblNome_Imp.Enabled = True
    cmndbtnSalvar.Enabled = True
    txtbxNome_Imp.Enabled = True
    cmndbtnEdit.Enabled = False
    txtbxNome_Imp.SetFocus
        
End Sub

Private Sub cmndbtnEdit_Click()
    If dtImp.Recordset.RecordCount > 0 Then
        frImp.BorderStyle = 1
        frImp.Enabled = True
        lblNome_Imp.Enabled = True
        lblNome_Imp.Enabled = True
        txtbxNome_Imp.Text = dtImp.Recordset!nome_Imp
        Me.Tag = dtImp.Recordset!cod_Imp
        cmndbtnEdit.Enabled = False
        cmbxNovo.Enabled = False
        cmndbtnSalvar.Enabled = True
        txtbxNome_Imp.Enabled = True
        txtbxNome_Imp.SetFocus
        
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxNome_Imp.Text = "" Then
        MsgBox "Informe o nome do imposto!", vbCritical
        txtbxNome_Imp.SetFocus
        Exit Sub
    End If
    If Me.Tag = "" Then
        If F_BuscaUmCampo("Cod_Imp", "Tab_Impostos", "Nome_Imp='" & txtbxNome_Imp.Text & _
                        "'") <> "" Then
            MsgBox "já existe um imposto com esse Nome", vbInformation
            Exit Sub
        End If
        P_Salva
    Else
        P_Altera
        cmndbtnSalvar.Enabled = False
        cmbxNovo.Enabled = True
        frImp.BorderStyle = 0
    End If
    Form_Activate
End Sub

Private Sub dbgrdImp_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dtImp_Reposition()
    cmndbtnEdit.Enabled = True
    txtbxNome_Imp.Text = ""
    txtbxNome_Imp.Enabled = False
    lblNome_Imp.Enabled = False
    cmbxNovo.Enabled = True
    cmndbtnSalvar.Enabled = False
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    dtImp.Connect = banco.Connect
    strSQL = "Select * from tab_impostos"
    dtImp.RecordSource = strSQL
    dtImp.Refresh
    lblImp.caption = dtImp.Recordset.RecordCount & " impostos encontrados"
    If dtImp.Recordset.RecordCount > 0 Then
        cmndbtnEdit.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
End Sub

Private Sub txtbxNome_Imp_GotFocus()
    P_SelecTexto txtbxNome_Imp
End Sub

Private Sub txtbxNome_Imp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Imp_LostFocus()
    P_AjustaTexto txtbxNome_Imp
End Sub
