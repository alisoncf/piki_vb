VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Contas"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   Icon            =   "frmEdiCon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10860
   Begin VB.CheckBox chkDesab 
      Caption         =   "Desabilitadas"
      Height          =   195
      Left            =   9360
      TabIndex        =   9
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtBusca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   4935
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   10830
      TabIndex        =   3
      Top             =   6180
      Width           =   10860
      Begin VB.CommandButton cmdSelec 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmEdiCon.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   9120
         Picture         =   "frmEdiCon.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmEdiCon.frx":0A12
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiCon.frx":0B5C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnDesab 
         Caption         =   "&Desab. / Hab."
         Height          =   735
         Left            =   5160
         Picture         =   "frmEdiCon.frx":0EE6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frmEdiCon.frx":1030
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9340
      _Version        =   393216
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
      Caption         =   "Contas"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cod_con"
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
         DataField       =   "nome_con"
         Caption         =   "Descrição"
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
         DataField       =   "nome_bco"
         Caption         =   "Banco"
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
      BeginProperty Column03 
         DataField       =   "age_con"
         Caption         =   "Age"
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
      BeginProperty Column04 
         DataField       =   "num_con"
         Caption         =   "Num."
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
      BeginProperty Column05 
         DataField       =   "tit_con"
         Caption         =   "Titular"
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
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4155,024
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Pesquisar"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   735
   End
End
Attribute VB_Name = "frmEdiCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdSelec_Click()
    On Error Resume Next
    Me.tag = ado!cod_con
    Me.Hide
End Sub

Private Sub cmndbtnDesab_Click()
    Dim strSQL As String
    With ado
        If MsgBox("Deseja Realmente mudat status da Conta " & _
            !nome_con & "?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
        strSQL = "update tab_contas set status_con='"
        If !Status_Con = "D" Then
            strSQL = strSQL & "H'"
        Else
            strSQL = strSQL & "D'"
        End If
        strSQL = strSQL & " where cod_Con='" & !cod_con & "'"
        conn.Execute strSQL
        dtServ.Refresh
    End With
    
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmIncCon
    With ado
    
        fr.tag = !cod_con
        fr.txtbxAge.Text = TiraNull(!Age_con)
        fr.txtbxNome.Text = TiraNull(!nome_con)
        fr.txtbxNum.Text = TiraNull(!Num_Con)
        fr.txtbxTit.Text = TiraNull(!Tit_Con)
        P_LocalizaOperacao TiraNull(!cod_bco), fr.cmbxBco
    End With
    fr.Show 1
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmIncCon
    fr.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    If Me.tag = "" Then
    
    ado.Close
    Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
    
End Sub



Private Sub Form_Activate()
    
    If Me.tag <> "" Then
        cmdSelec.Visible = True
    Else
        cmdSelec.Visible = False
    End If
    If Me.tag <> "" Then
        busca
    End If
    
End Sub

Private Sub Form_Load()
    'Me.Height = 4500
    'Me.Width = 6975
    Dim strSQL As String
    Set ado = New ADODB.Recordset
    
    
    P_CentraForm Me
End Sub
Sub busca()
Dim strSQL As String
On Error GoTo trata_Erro
    strSQL = "select tc.*,tb.nome_bco from tab_contas as tc left join tab_bancos as tb "
    strSQL = strSQL & " on tc.cod_bco=tb.cod_bco"
    
    strSQL = strSQL & " where nome_con like '" & txtBusca.Text & "%' "
    strSQL = strSQL & " order by nome_con "
    abreObjetoAdo ado, strSQL, grid
    If Me.tag <> "" And ado.RecordCount = 1 Then
        cmdSelec_Click
    End If
trata_Erro:
End Sub

Private Sub grid_DblClick()
    If Me.tag <> "" Then
        cmdSelec_Click
    Else
        cmndbtnEdit_Click
    End If
    
End Sub

Private Sub grid_GotFocus()
    grid.RecordSelectors = True
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grid_DblClick
    End If
    
End Sub

Private Sub grid_LostFocus()
    grid.RecordSelectors = False
End Sub

Private Sub txtBusca_Change()
    busca
End Sub

Private Sub txtBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        grid.SetFocus
    End If
End Sub
