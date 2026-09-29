VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiBan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edição de Bancos"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5115
   Icon            =   "frmEdiBan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.Data dtServ 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   4815
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   3240
         Picture         =   "frmEdiBan.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmEdiBan.frx":0858
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiBan.frx":09A2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiBan.frx":0D2C
      Height          =   4695
      Left            =   120
      OleObjectBlob   =   "frmEdiBan.frx":0D41
      TabIndex        =   0
      Top             =   105
      Width           =   4815
   End
End
Attribute VB_Name = "frmEdiBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadBan
    With dtServ.Recordset
        fr.Tag = !cod_bco
        fr.txtbxCod.Text = !cod_bco
        fr.txtbxCod.Enabled = False
        fr.txtbxNOme.Text = !nome_bco
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadBan
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        strSQL = "select * from tab_bancos order by cod_bco"
    Else
        strSQL = "select * from tab_bancos order by nome_bco"
    End If
    
    
    dtServ.RecordSource = strSQL
    dtServ.Refresh
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Load()
    MudaAparencia Me
    'Me.Height = 5205
    'Me.Width = 5145
    P_CentraForm Me
    strSQL = "select * from tab_bancos"
    dtServ.Connect = banco.Connect
    dtServ.RecordSource = strSQL
    dtServ.Refresh
End Sub
