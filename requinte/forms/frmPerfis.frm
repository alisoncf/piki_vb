VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPerfis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfis"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmPerfis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4455
   Begin MSDBGrid.DBGrid dbgrdImp 
      Bindings        =   "frmPerfis.frx":0442
      Height          =   1095
      Left            =   120
      OleObjectBlob   =   "frmPerfis.frx":0456
      TabIndex        =   1
      Top             =   2280
      Width           =   4215
   End
   Begin VB.CommandButton cmbxNovo 
      Caption         =   "&Novo"
      Height          =   735
      Left            =   1560
      Picture         =   "frmPerfis.frx":1179
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnSair 
      BackColor       =   &H8000000A&
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   3480
      Picture         =   "frmPerfis.frx":1503
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnEdit 
      Caption         =   "&Editar"
      Height          =   735
      Left            =   2520
      Picture         =   "frmPerfis.frx":1945
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   855
   End
   Begin VB.Data dtImp 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Data dtPer 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSDBGrid.DBGrid dbGrdPer 
      Bindings        =   "frmPerfis.frx":1A8F
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frmPerfis.frx":1AA3
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmPerfis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxNovo_Click()
    frmIncPerfil.Show 0
    
End Sub

Private Sub cmndbtnEdit_Click()
    Dim strSQL As String
    If dbGrdPer.Row = -1 Or dtPer.Recordset.RecordCount = 0 Then
        MsgBox "Não há o que ser editado!", vbExclamation
        Exit Sub
    End If
    With frmIncPerfil
        .Tag = dtPer.Recordset!cod_per
        .txtbxDesc_Per.Text = dtPer.Recordset!Desc_Per
        .txtbxNome_Per.Text = dtPer.Recordset!nome_per
        With dtImp.Recordset
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                 strSQL = !cod_Imp & Chr(9) & !nome_Imp & Chr(9) & !Val_Imp & Chr(9) & !data_imp
                 frmIncPerfil.msflxgrdItens.AddItem strSQL
                .MoveNext
            Loop
        End With
        
        .Show 0
    End With
    
    dtPer.Refresh
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dtPer_Reposition()
    Dim strSQL As String
    strSQL = "select ti.*,tp.* from tab_impostos as ti, tab_perimp as tp "
    strSQL = strSQL & " where ti.cod_imp=tp.cod_imp"
    strSQL = strSQL & " and tp.cod_per='" & dtPer.Recordset!cod_per & "'"
    dtImp.RecordSource = strSQL
    dtImp.Refresh
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    P_CentraForm Me
    dtImp.Connect = banco.Connect
    dtPer.Connect = banco.Connect
    strSQL = "Select * from tab_perfis order by nome_Per "
    dtPer.RecordSource = strSQL
    dtPer.Refresh
End Sub
