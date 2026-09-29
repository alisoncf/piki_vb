VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiFinan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Financeiras"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "frmEdiFinan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5985
   Begin VB.Data dtServ 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   5895
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   3960
         Picture         =   "frmEdiFinan.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   2040
         Picture         =   "frmEdiFinan.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiFinan.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiFinan.frx":0768
      Height          =   2415
      Left            =   0
      OleObjectBlob   =   "frmEdiFinan.frx":077D
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmEdiFinan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    With dtServ.Recordset
        If .RecordCount <= 0 Then
            MsgBox "Não há o que ser editado!", vbExclamation
            Exit Sub
        End If
        Dim fr As New frmCadFinan
        fr.Tag = !Cod_Fin
        fr.txtbxNome.Text = !nome_fin
        fr.txtbxTel.Text = !tel_fin
        
        P_LocalizaOperacao !tipo_fin, fr.cmbxBco
        fr.Show 0
            
        
    End With
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadFinan
    fr.Show 0
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    P_CentraForm Me
    dtServ.Connect = banco.Connect
    strSQL = "select * from tab_financeiras order by nome_fin "
    dtServ.RecordSource = strSQL
    dtServ.Refresh
End Sub
