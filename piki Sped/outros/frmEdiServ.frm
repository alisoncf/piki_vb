VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Serviços:"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "frmEdiServ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8415
   Begin VB.Frame frControle 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEdiServ.frx":08CA
         Height          =   3975
         Left            =   2640
         OleObjectBlob   =   "frmEdiServ.frx":08DF
         TabIndex        =   12
         Top             =   240
         Width           =   5415
      End
      Begin VB.Frame Frame3 
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2415
         Begin VB.TextBox txtbxExpr 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   7
            Top             =   1560
            Width           =   2175
         End
         Begin VB.ComboBox cmbxBusca 
            Height          =   315
            ItemData        =   "frmEdiServ.frx":1607
            Left            =   120
            List            =   "frmEdiServ.frx":1611
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox cmbxOrd 
            Height          =   315
            ItemData        =   "frmEdiServ.frx":1623
            Left            =   120
            List            =   "frmEdiServ.frx":1633
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Buscar Por:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   825
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "E&xpressão:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Ordenar:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   615
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
         Left            =   4920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Frame Frame2 
         Height          =   2295
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   2415
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Novo"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiServ.frx":1656
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   180
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiServ.frx":19E0
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiServ.frx":1B2A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1560
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmEdiServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    With dtServ.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que Editar!", vbExclamation
            Exit Sub
        End If
        Dim fr As New frmIncServic
        fr.Tag = !cod_ser
        fr.txtbxComissao_Ser.Text = FormatVal(!comissao_ser)
        fr.txtbxNome_Ser.Text = !nome_Ser
        fr.txtbxValor_Ser.Text = FormatVal(!valor_ser)
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmIncServic
    fr.Show 0
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtServ_Reposition()
    If dtServ.Recordset.RecordCount > 0 Then
        cmndbtnEdit.Enabled = True
    Else
        cmndbtnEdit.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    dtServ.Connect = banco.Connect
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 0
    Me.Height = 4935
    Me.Width = 8220
    P_Busca
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sOrder As String
    
    If cmbxOrd.ListIndex <= 0 Then
        sOrder = "nome_ser "
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrder = "cod_ser"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrder = "valor_ser"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrder = "valor_ser"
    End If
    
    strSQL = "select * from tab_servicos "
    If cmbxBusca.ListIndex <= 0 Then
        strSQL = strSQL & " where nome_ser like '" & txtbxExpr.Text & "%'"
    Else
        strSQL = strSQL & " where cod_ser like '" & txtbxExpr.Text & "%'"
    End If
    strSQL = strSQL & " order by " & sOrder
    dtServ.RecordSource = strSQL
    dtServ.Refresh
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub
