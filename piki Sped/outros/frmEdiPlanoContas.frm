VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiPlanoContas 
   Caption         =   "Edição de Plano de Contas"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   Icon            =   "frmEdiPlanoContas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   3990
   Begin VB.Frame frOutros 
      BackColor       =   &H80000004&
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   3975
      Begin VB.CommandButton cmndbtnExcluir 
         Caption         =   "Desab."
         Enabled         =   0   'False
         Height          =   735
         Left            =   2040
         Picture         =   "frmEdiPlanoContas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiPlanoContas.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   1080
         Picture         =   "frmEdiPlanoContas.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   3000
         Picture         =   "frmEdiPlanoContas.frx":06E8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   165
         Width           =   855
      End
   End
   Begin VB.Data dtPla 
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
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid dbgrdPla 
      Bindings        =   "frmEdiPlanoContas.frx":0832
      Height          =   3375
      Left            =   0
      OleObjectBlob   =   "frmEdiPlanoContas.frx":0846
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmEdiPlanoContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    With dtPla.Recordset
    'If ActionOnForm("frmCadPlano", !cod_pla, ActFormTag, ActNoAction) Then
     '   Exit Sub
    'End If
        Dim fr As New frmCadPlano
        
        fr.Tag = !cod_pla
        fr.txtbxNome.Text = !Nome_Pla
        If !Tipo_Pla = "C" Then
            fr.optnCred.Value = True
        Else
            fr.optnDeb.Value = True
        End If
        fr.optnCred.Enabled = False
        fr.optnDeb.Enabled = False
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnNovo_Click()
    Dim fr As New frmCadPlano
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrdPla_DblClick()
    If dtPla.Recordset.RecordCount > 0 Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub dbgrdPla_GotFocus()
    dbgrdPla.MarqueeStyle = 3
End Sub

Private Sub dbgrdPla_HeadClick(ByVal ColIndex As Integer)
    busca dbgrdPla.Columns(ColIndex).DataField
End Sub

Private Sub dbgrdPla_KeyDown(KeyCode As Integer, Shift As Integer)
        
        If KeyCode = vbKeyReturn Then
        dbgrdPla_DblClick
    End If
End Sub

Private Sub dbgrdPla_LostFocus()
    dbgrdPla.MarqueeStyle = 0
End Sub

Private Sub dtPla_Reposition()
    If dtPla.Recordset.RecordCount > 0 Then
        cmndbtnEdit.Enabled = True
        If F_BuscaUmCampo("Cod_Cai", "tab_caixas", " cod_pla='" & _
            dtPla.Recordset!cod_pla & "'") = "" Then
            cmndbtnExcluir.Enabled = True
        End If
    Else
        cmndbtnEdit.Enabled = False
        cmndbtnExcluir.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
        MudaAparencia Me
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dtPla.Connect = banco.Connect
    Me.Height = 5070
    Me.Width = 4110
    busca "Nome_Pla"
End Sub
Sub busca(sOrd As String)
    Dim strSQL As String
    strSQL = "select * from tab_planocontas "
    strSQL = strSQL & " order by " & sOrd
    dtPla.RecordSource = strSQL
    dtPla.Refresh
End Sub
