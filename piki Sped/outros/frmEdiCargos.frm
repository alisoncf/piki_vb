VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiCargos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Cargos"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "frmEdiCargos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4845
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   2535
      Width           =   4815
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiCargos.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmEdiCargos.frx":04D4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   3240
         Picture         =   "frmEdiCargos.frx":061E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1455
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
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiCargos.frx":09A8
      Height          =   2535
      Left            =   0
      OleObjectBlob   =   "frmEdiCargos.frx":09BD
      TabIndex        =   4
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmEdiCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadCargos
    With dtServ.Recordset
        If .RecordCount <= 0 Then
            MsgBox "Não há o que ser Editado!", vbExclamation
            Exit Sub
        End If
    
        fr.Tag = !cod_car
        P_LocalizaOperacao !venda_car, fr.cmbxVenda
        fr.txtbxNome.Text = !nome_car
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadCargos
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
    P_CentraForm Me
    MudaAparencia Me
    'Me.Height = 5205
    'Me.Width = 5145
    strSQL = "select * from tab_cargos order by nome_car"
    dtServ.Connect = banco.Connect
    dtServ.RecordSource = strSQL
    dtServ.Refresh
End Sub

