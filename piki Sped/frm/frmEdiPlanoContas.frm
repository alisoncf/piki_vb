VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiPlanoContas 
   Caption         =   "Edição de Plano de Contas"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   Icon            =   "frmEdiPlanoContas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   9840
   Begin VB.PictureBox pcTipoPLa 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   2355
      TabIndex        =   9
      Top             =   6600
      Width           =   2415
      Begin VB.OptionButton optnDeb 
         Caption         =   "&Débito"
         Height          =   195
         Left            =   1170
         TabIndex        =   11
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton optnCred 
         Caption         =   "&Crédito"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   6840
      Width           =   5295
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   9810
      TabIndex        =   1
      Top             =   7320
      Width           =   9840
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Selecionar"
         Height          =   735
         Left            =   1680
         Picture         =   "frmEdiPlanoContas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   8040
         Picture         =   "frmEdiPlanoContas.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   3240
         Picture         =   "frmEdiPlanoContas.frx":05BF
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiPlanoContas.frx":0709
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnExcluir 
         Caption         =   "Desab."
         Enabled         =   0   'False
         Height          =   735
         Left            =   4800
         Picture         =   "frmEdiPlanoContas.frx":0853
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiPlanoContas.frx":099D
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11245
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "cod_pla"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nome_pla"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "tipo_pla"
         Caption         =   "Tipo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5880,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   645,165
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Descrição:"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   6600
      Width           =   975
   End
End
Attribute VB_Name = "frmEdiPlanoContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdSelect_Click()
    If Me.tag <> "" Then
        Me.Hide
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    With dtPla.Recordset
    'If ActionOnForm("frmCadPlano", !cod_pla, ActFormTag, ActNoAction) Then
     '   Exit Sub
    'End If
        Dim fr As New frmCadPlano
        
        fr.tag = !cod_pla
        fr.txtbxNOme.Text = !nome_pla
        If !tipo_pla = "C" Then
            fr.optnCred.Value = True
        Else
            fr.optnDeb.Value = True
        End If
        fr.optnCred.Enabled = False
        fr.optnDeb.Enabled = False
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnNovo_Click()
    Dim fr As New frmCadPlano
    fr.Show 1
    ado.Requery
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub dbgrdPla_DblClick()
    If dtPla.Recordset.RecordCount > 0 Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub dbgrdPla_GotFocus()
    dbgrdPla.MarqueeStyle = 3
End Sub

Private Sub dbgrdPla_HeadClick(ByVal ColIndex As Integer)
    busca dbgrdPla.Columns(ColIndex).DataField
End Sub

Private Sub dbgrdPla_KeyDown(KeyCode As Integer, Shift As Integer)
        
        If KeyCode = vbKeyReturn Then
        dbgrdPla_DblClick
    End If
End Sub

Private Sub dbgrdPla_LostFocus()
    dbgrdPla.MarqueeStyle = 0
End Sub

Private Sub dtPla_Reposition()
    If dtPla.Recordset.RecordCount > 0 Then
        cmndbtnEdit.Enabled = True
        If F_BuscaUmCampo("Cod_Cai", "tab_caixas", " cod_pla='" & _
            dtPla.Recordset!cod_pla & "'") = "" Then
            cmndbtnExcluir.Enabled = True
        End If
    Else
        cmndbtnEdit.Enabled = False
        cmndbtnExcluir.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        busca ""
    Else
        cmdSelect.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    busca "Nome_Pla"
End Sub
Sub busca(sOrd As String)
    Dim strSQL As String
    strSQL = "select * from " & TAB_PLANOCONTAS
    strSQL = strSQL & " where (nome_pla like '%" & txtDesc.Text & "%'    "
    strSQL = strSQL & " or cod_pla like '%" & txtDesc.Text & "%'    )"
    If optnCred.Value = True Then
        strSQL = strSQL & " and tipo_pla='C' "
    Else
        strSQL = strSQL & " and tipo_pla='D' "
    End If
    If sOrd <> "" Then
        strSQL = strSQL & " order by " & sOrd
    End If
    abreObjetoAdo ado, strSQL, dtgrd
End Sub

Private Sub optnCred_Click()
    busca ""
End Sub

Private Sub optnDeb_Click()
    busca ""
End Sub

Private Sub txtDesc_Change()
    busca ""
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        busca ""
    ElseIf KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
End Sub
