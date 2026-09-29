VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frmEdiImo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edição de Imobilizados"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   98
      TabIndex        =   0
      Top             =   4215
      Width           =   7215
      Begin VB.CommandButton cmndbtnExcluir 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   3240
         Picture         =   "frmEdiImo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiImo.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmEdiImo.frx":04D4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   5640
         Picture         =   "frmEdiImo.frx":061E
         Style           =   1  'Graphical
         TabIndex        =   6
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
      Left            =   1935
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2295
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiImo.frx":09A8
      Height          =   4095
      Left            =   105
      OleObjectBlob   =   "frmEdiImo.frx":09BD
      TabIndex        =   2
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   5280
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   405
   End
End
Attribute VB_Name = "frmEdiImo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadImo
    With dtServ.Recordset
        fr.Tag = !cod_imo
        fr.txtbxNome.Text = !desc_IMO
        fr.txtbxQtd.Text = FormatVal(!qtd_imo)
        fr.txtbxUnit.Text = FormatVal(!unit_imo)
        fr.txtbxUni.Text = !unidade_imo
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnExcluir_Click()
    If MsgBox("Confirma exclusão do imobilizado?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from tab_imobilizado where cod_imo='" & dtServ.Recordset!cod_imo & "' "
        Banco.Execute strSQL
        busca
    End If
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadImo
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
        strSQL = "select * from tab_imob order by cod_imo"
    Else
        strSQL = "select * from tab_imob order by desc_imo"
    End If
    
    
    dtServ.RecordSource = strSQL
    dtServ.Refresh
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    busca
    
End Sub
Sub busca()
    strSQL = "select ti.*,ti.unit_imo*ti.qtd_imo as total from tab_imob as ti"
    dtServ.Connect = Banco.Connect
    dtServ.RecordSource = strSQL
    dtServ.Refresh
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    With dtServ.Recordset
        Do While Not .EOF
            cTotal = cTotal + !total
            .MoveNext
        Loop
        
        
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        lblTotal.caption = FormatVal(cTotal)
    End With
    
End Sub
Private Sub Form_Load()
    MudaAparencia Me
    'Me.Height = 5205
    'Me.Width = 5145
    P_CentraForm Me
    
End Sub

