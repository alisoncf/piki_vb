VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Contas"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmEdiCon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6810
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiCon.frx":014A
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "frmEdiCon.frx":015F
      TabIndex        =   0
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmndbtnDesab 
         Caption         =   "&Desab. / Hab."
         Height          =   615
         Left            =   3480
         Picture         =   "frmEdiCon.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiCon.frx":1332
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1800
         Picture         =   "frmEdiCon.frx":16BC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   5160
         Picture         =   "frmEdiCon.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1575
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
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "frmEdiCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnDesab_Click()
    Dim strSQL As String
    With dtServ.Recordset
        If MsgBox("Deseja Realmente mudat status da Conta " & _
            !nome_con & "?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
        strSQL = "update tab_contas set status_con='"
        If !Status_Con = "D" Then
            strSQL = strSQL & "H'"
        Else
            strSQL = strSQL & "D'"
        End If
        strSQL = strSQL & " where cod_Con='" & !cod_Con & "'"
        Banco.Execute strSQL
        dtServ.Refresh
    End With
    
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmIncCon
    With dtServ.Recordset
    
        fr.Tag = !cod_Con
        fr.txtbxAge.Text = !Age_con
        fr.txtbxNome.Text = !nome_con
        fr.txtbxNum.Text = !Num_Con
        fr.txtbxTit.Text = !Tit_Con
        P_LocalizaOperacao !cod_bco, fr.cmbxBco
    End With
    fr.Show 0
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmIncCon
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    dtServ.Recordset.Close
    Unload Me
    
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Load()
    'Me.Height = 4500
    'Me.Width = 6975
    Dim strSQL As String
    P_CentraForm Me
    dtServ.Connect = Banco.Connect
    strSQL = "select tc.*,tb.nome_bco from tab_contas as tc, tab_bancos as tb "
    strSQL = strSQL & " where tc.cod_bco=tb.cod_bco order by nome_con"
    dtServ.RecordSource = strSQL
    dtServ.Refresh
End Sub
