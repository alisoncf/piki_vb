VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSelectServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleção de Serviços"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmSelectServ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   2175
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   120
         Picture         =   "frmSelectServ.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   1080
         Picture         =   "frmSelectServ.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   2280
      TabIndex        =   3
      Top             =   0
      Width           =   3855
      Begin VB.Data dtPro 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmSelectServ.frx":11DE
         Height          =   3975
         Left            =   120
         OleObjectBlob   =   "frmSelectServ.frx":11F2
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   3240
      Width           =   2175
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblPreco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   840
         TabIndex        =   15
         Tag             =   "!"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comiss.(%):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label lblComissao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   1200
         TabIndex        =   13
         Tag             =   "!"
         Top             =   600
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2175
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmSelectServ.frx":1BCE
         Left            =   120
         List            =   "frmSelectServ.frx":1BD8
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmSelectServ.frx":1BEA
         Left            =   120
         List            =   "frmSelectServ.frx":1BF4
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSelectServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxBusca_Click()
    Busca
End Sub

Private Sub cmbxDiv_Click()
    lblDiv.caption = F_BuscaUmCampo("Cod_div", "tab_divisoes", "nome_div='" & cmbxDiv.Text & "'")
    Busca
End Sub

Private Sub cmbxDiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxDiv.ListIndex = -1
    End If
End Sub

Private Sub cmbxEsp_Click()
    lblEsp.caption = F_BuscaUmCampo("Cod_Esp", "tab_Especies", "nome_esp='" & cmbxEsp.Text & "'")
    Busca
End Sub

Private Sub cmbxEsp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxEsp.ListIndex = -1
    End If
End Sub


Private Sub cmbxGenero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxGenero.ListIndex = -1
    End If
End Sub

Private Sub cmbxOrd_Click()
    Busca
End Sub
Private Sub cmndbtnEdit_Click()
    Me.Tag = dtPro.Recordset!cod_ser
    Me.Hide
End Sub

Private Sub cmndbtnNew_Click()
    frmIncPro.Show 0
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

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtPro_Reposition()
    If dtPro.Recordset.RecordCount = 0 Then
    Else
        With dtPro.Recordset
            lblPreco.caption = FormatVal(dtPro.Recordset!valor_ser)
            
            lblComissao.caption = FormatVal(dtPro.Recordset!comissao_ser)
            
        End With
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 6285
    Me.Height = 4845
    P_CentraForm Me, 0
    dtPro.Connect = banco.Connect
    
    Busca
    MudaAparencia Me
End Sub

Private Sub txtbxExpr_Change()
    Busca
End Sub

Sub Busca()
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    
    If cmbxBusca.ListIndex = 1 Then
        sCampo = "cod_ser"
    Else
        sCampo = "nome_ser"
    End If
    
    If cmbxOrd.ListIndex = 1 Then
        sOrder = "cod_ser"
    
    Else
        sOrder = "nome_ser,cod_ser "
    End If
    strSQL = "select * from tab_servicos "
    strSQL = strSQL & " where " & sCampo & " like '"
    strSQL = strSQL & txtbxExpr.Text & "%' "
    strSQL = strSQL & " order by " & sOrder
    dtPro.RecordSource = strSQL
    dtPro.Refresh
    If dtPro.Recordset.RecordCount = 0 Then
        MsgBox "Não foram encontrados serviços...", vbExclamation
    End If
End Sub


