VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelMemCob 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Memória das Cobranças"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13290
      TabIndex        =   11
      Top             =   0
      Width           =   13320
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   11880
         Picture         =   "frmRelMemCob.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   10440
         Picture         =   "frmRelMemCob.frx":016B
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   615
         Left            =   9000
         Picture         =   "frmRelMemCob.frx":0578
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1335
      End
      Begin VB.CheckBox chkCobrador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apenas Carnês com o Cobrador"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6000
         TabIndex        =   9
         Top             =   360
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkAberto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apenas Parcelas em Aberto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6000
         TabIndex        =   8
         Top             =   120
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxContaCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmndbtnContaCob 
         Height          =   285
         Left            =   2880
         Picture         =   "frmRelMemCob.frx":0902
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   3240
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   4440
         TabIndex        =   6
         Top             =   0
         Width           =   285
      End
      Begin VB.Label lblContaCob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Cobrança:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Data dtCob 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid dbgrdCob 
      Bindings        =   "frmRelMemCob.frx":0C8C
      Height          =   6975
      Left            =   120
      OleObjectBlob   =   "frmRelMemCob.frx":0CA0
      TabIndex        =   0
      Top             =   840
      Width           =   13095
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   8280
      TabIndex        =   14
      Top             =   7800
      Width           =   4815
      Begin VB.OptionButton Option1 
         Caption         =   "Numero"
         Height          =   195
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optnCob 
         Caption         =   "Por Cobrança"
         Height          =   195
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optnData 
         Caption         =   "Por Data"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRelMemCob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAberto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub chkCobrador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    strSQL = "select distinct td.num_doc,th.seq_doc,th.cobout_doc,tipocobin_doc,tipocobout_doc,th.cobin_doc,th.data_doc,th.cod_usu,"
    strSQL = strSQL & " tc.nome_cli from tab_dochist as th, tab_doc as td, tab_itdoc as ti, tab_clientes as tc "
    strSQL = strSQL & " where tc.cod_cli=td.cod_ter"
    strSQL = strSQL & " and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.seq_doc=th.seq_doc "
    
    If chkAberto.Value = 1 Then
        strSQL = strSQL & " and reman_itd>0"
    End If
    If chkCobrador.Value = 1 Then
        strSQL = strSQL & " and ti.cod_cob=cobin_doc "
    End If
    If txtbxContaCob.tag <> "" Then
        strSQL = strSQL & " and th.cobin_doc='" & txtbxContaCob.tag & "' "
    End If
    If txtbxDataFin.Text <> "" Then
        strSQL = strSQL & " and th.data_doc<=" & F_GravaDataTempo(txtbxDataFin.Text & " 23:59:59")
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and th.data_doc>=" & F_GravaDataTempo(txtbxDataIni.Text & " 00:00:00")
    End If
    If optnCob.Value Then
        strSQL = strSQL & " order by data_doc desc,ti.cod_cob     "
    ElseIf optnData.Value Then
        strSQL = strSQL & " order by data_doc desc"
    Else
        strSQL = strSQL & " order by num_doc, data_doc desc"
    End If
    
    dtCob.RecordSource = strSQL
    dtCob.Refresh
    
End Sub

Private Sub cmndbtnContaCob_Click()
    Dim t As ContaCob
    
    t = Cobranca.F_BuscaContaCob("")
    If t.cod_cob <> "" Then
        txtbxContaCob.Text = t.nome_cob
        txtbxContaCob.tag = t.cod_cob
        'If t.cod_cob = "000" Then
        '    txtbxFunc.Enabled = True
        '    cmndbtnFun.Enabled = True
        'E'lse
        '    txtbxFunc.Enabled = False
        '    cmndbtnFun.Enabled = False
        'End If
        txtbxDataIni.SetFocus
    End If


End Sub

Private Sub cmndbtnImprimir_Click()
    P_DbgrdToHtml "Relatório de Memória das Cobranças de " & txtbxDataIni.Text & " A " & txtbxDataFin.Text, _
        dbgrdCob, dtCob, "Cobrança: " & txtbxContaCob.Text
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dtCob.Connect = Banco.Connect
    P_CentraForm Me
End Sub

Private Sub txtbxContaCob_Change()
    txtbxContaCob.tag = ""
End Sub

Private Sub txtbxContaCob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnContaCob_Click
    End If
    
End Sub

Private Sub txtbxdatafin_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxdatafin_LostFocus()
    P_AjustaData txtbxDataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

