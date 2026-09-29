VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiMar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Marcas:"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "frmEdiMar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5730
   Begin VB.Frame frControle 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Data dtProd 
         Caption         =   "Edição de Gêneros"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1560
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   3600
         TabIndex        =   1
         Top             =   120
         Width           =   1935
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Novo"
            Height          =   855
            Left            =   120
            Picture         =   "frmEdiMar.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   855
            Left            =   120
            Picture         =   "frmEdiMar.frx":0C54
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   855
            Left            =   120
            Picture         =   "frmEdiMar.frx":0D9E
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2160
            Width           =   1695
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmEdiMar.frx":0EE8
         Height          =   3015
         Left            =   120
         OleObjectBlob   =   "frmEdiMar.frx":0EFD
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmEdiMar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadMar
    If dtProd.Recordset.RecordCount > 0 Then
        fr.Tag = dtProd.Recordset!cod_mar
        fr.txtbxNome.Text = dtProd.Recordset!nome_mar
        fr.Show 0
    Else
        MsgBox "Não há o que ser Editado!", vbExclamation
    End If
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadMar
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub DBGrid1_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.MarqueeStyle = 3
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
'    Me.Height = 3840
'    Me.Width = 3570
    P_CentraForm Me
    dtProd.Connect = banco.Connect
    strSQL = "select * from tab_marcas order by nome_mar "
    dtProd.RecordSource = strSQL
    dtProd.Refresh
End Sub


