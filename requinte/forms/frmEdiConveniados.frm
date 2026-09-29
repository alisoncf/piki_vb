VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiConveniados 
   Caption         =   "Edição de Conveniados"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   Icon            =   "frmEdiConveniados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   10305
   Begin VB.Data dtClientes 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox pcBotoes 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6090
      Left            =   8730
      ScaleHeight     =   6060
      ScaleWidth      =   1545
      TabIndex        =   5
      Top             =   615
      Width           =   1575
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConveniados.frx":0BA2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConveniados.frx":0F2C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConveniados.frx":1076
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "&Histórico"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConveniados.frx":11C0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnConv 
         Caption         =   "Con&veniados"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConveniados.frx":1A8A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDeb 
         Caption         =   "&Débitos"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConveniados.frx":1B3F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3000
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   0
      Width           =   10305
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   315
         Left            =   4080
         Picture         =   "frmEdiConveniados.frx":1C89
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxConvenio 
         Height          =   315
         Left            =   120
         MaxLength       =   40
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   315
         Left            =   4560
         MaxLength       =   40
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Conveniado"
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Convênio:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiConveniados.frx":1D82
      Height          =   5895
      Left            =   0
      OleObjectBlob   =   "frmEdiConveniados.frx":1D9B
      TabIndex        =   12
      Top             =   720
      Width           =   7935
   End
End
Attribute VB_Name = "frmEdiConveniados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Dim c As Cliente
    c = F_BuscaCliente(txtbxConvenio.Text, True)
    If c.cod_cli <> "" Then
        txtbxConvenio.Text = c.nome_Cli
        txtbxConvenio.tag = c.cod_cli
        
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadConveniado
    With dtClientes.Recordset
        If .RecordCount > 0 Then
            fr.tag = dtClientes.Recordset!seq_con
            fr.txtbxConv.tag = !cod_cli
            fr.Show 0
        Else
            MsgBox "Não hã o que ser editado!", vbExclamation
            
        End If
    End With
End Sub

Private Sub cmndbtnLocaliza_Click()
    Dim strSQL As String
    strSQL = "select tc.* from tab_conveniados as tc where tc.nome_con like '%" & txtbxExpr.Text & "%' "
    If txtbxConvenio.tag <> "" Then
        strSQL = strSQL & " and tc.cod_cli='" & txtbxConvenio.tag & "' "
    End If
    
    dtClientes.RecordSource = strSQL
    dtClientes.Refresh
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadConveniado
    fr.Show 0
End Sub

Private Sub Form_Load()
    dtClientes.Connect = Banco.Connect
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbgrd.Width = pcBotoes.Left - 100
    
   
    dbgrd.Height = Me.Height - dbgrd.Top - 500
   
End Sub

Private Sub txtbxConvenio_GotFocus()
    P_SelecTexto txtbxConvenio
End Sub

Private Sub txtbxConvenio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxConvenio_LostFocus()
    P_AjustaTexto txtbxConvenio
End Sub
