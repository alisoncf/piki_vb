VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frmBuscaBan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selecione um Banco"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmBuscaBan.frx":0000
      Height          =   5175
      Left            =   120
      OleObjectBlob   =   "frmBuscaBan.frx":0015
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox txtbxExpr 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   4815
   End
   Begin VB.Data dtServ 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2295
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "frmBuscaBan"
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
        fr.txtbxNome.Text = !nome_bco
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
    On Error Resume Next
    Me.Tag = dtServ.Recordset!cod_bco
    Me.Hide
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub
Sub busca(Optional ordem As Integer = 0)
    On Error GoTo trata_erro
    dtServ.Connect = Banco.Connect
    For iCount = 1 To 2
        strSQL = "select * from tab_bancos "
        If iCount = 1 Then
            strSQL = strSQL & " where cod_bco like '" & txtbxExpr.Text & "%'"
        Else
            strSQL = strSQL & " where nome_bco like '%" & txtbxExpr.Text & "%'"
        End If
        If ordem = 0 Then
            strSQL = strSQL & " order by cod_bco"
        Else
            strSQL = strSQL & " order by nome_bco"
        End If
        dtServ.RecordSource = strSQL
        dtServ.Refresh
        If dtServ.Recordset.RecordCount = 1 Then
            Me.Tag = dtServ.Recordset!cod_bco
            On Error Resume Next
            
            Me.Hide
            Exit Sub
        ElseIf dtServ.Recordset.RecordCount > 1 Then
            Exit For
        End If
    Next
        
    
    Exit Sub
trata_erro:
'    P_MostraErro
End Sub
Private Sub dbgrd_HeadClick(ByVal ColIndex As Integer)
    busca ColIndex
    
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Tag = ""
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    MudaAparencia Me
    'Me.Height = 5205
    'Me.Width = 5145
    P_CentraForm Me
    busca
    dtServ.Connect = Banco.Connect
    
End Sub

Private Sub txtbxExpr_Change()
    busca 0
End Sub
