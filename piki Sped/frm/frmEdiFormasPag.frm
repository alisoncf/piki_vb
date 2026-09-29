VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiFormasPag 
   Caption         =   "Edição de Formas de Pagamento"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   Icon            =   "frmEdiFormasPag.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7470
   ScaleWidth      =   9315
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtExpr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   6120
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   9285
      TabIndex        =   5
      Top             =   6495
      Width           =   9315
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3720
         Picture         =   "frmEdiFormasPag.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiFormasPag.frx":05D1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1920
         Picture         =   "frmEdiFormasPag.frx":095B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   7560
         Picture         =   "frmEdiFormasPag.frx":0AA5
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiFormasPag.frx":0FF3
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9975
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
      Caption         =   "Edição de Formas de Pagamento"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "cod_for"
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
         DataField       =   "nome_for"
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
         DataField       =   "desconto_for"
         Caption         =   "Desconto"
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
         DataField       =   "tipo_cai"
         Caption         =   "Tipo Caixa"
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
         DataField       =   "parcelas_for"
         Caption         =   "Parcelas"
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
         DataField       =   "avista_for"
         Caption         =   "À Vista"
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
      BeginProperty Column06 
         DataField       =   "nome_tip"
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
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2520
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1920,189
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   5880
      Width           =   735
   End
End
Attribute VB_Name = "frmEdiFormasPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents adc As ADODB.Recordset
Attribute adc.VB_VarHelpID = -1
Private Sub cmdSelecionar_Click()
    If Me.tag <> "" Then
        Me.Hide
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    If strCadformaspag_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    
    With adc
        If .RecordCount = 0 Then
            MsgBox "Não há Tipo de Pagamento para editar!", vbExclamation
            Exit Sub
        End If
        frmCadFormasPag.tag = adc!cod_for
        
        frmCadFormasPag.Show 1
        RefreshAdoPosition adc
        
    End With
    busca
End Sub

Private Sub cmndbtnNew_Click()
    
    If strCadformaspag_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
        
    End If
    frmCadFormasPag.Show 1
    RefreshAdoPosition adc
    
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtgrd_DblClick()
    If Me.tag <> "" Then
        cmdSelecionar_Click
    Else
        cmndbtnEdit_Click
    End If
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    
    Set adc = New ADODB.Recordset
    
    
End Sub

Sub busca()
    Dim strSQL As String
    strSQL = "select tf.*, tt.cod_tip,tt.nome_tip from tab_formaspag as tf LEFT join tab_tipopag as tt "
    strSQL = strSQL & " on tt.cod_tip=tf.cod_tip where 1> 0 "
    If txtCod.Text <> "" Then
        strSQL = strSQL & " and cod_for='" & txtCod.Text & "' "
    Else
        strSQL = strSQL & " and nome_for like '%" & txtExpr.Text & "%'"
    End If
    
    
    
    
    strSQL = strSQL & " order by NOME_for"
    abreObjetoAdo adc, strSQL, dtgrd
    
    If Me.tag <> "" Then
        If adc.RecordCount = 1 Then
            Me.Hide
        End If
    End If
    
    
End Sub

Private Sub Option1_Click()
    busca
End Sub

Private Sub optnCheque_Click()
    busca
End Sub

Private Sub optnDup_Click()
    busca
End Sub

Private Sub txtCod_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        busca
     End If
End Sub

Private Sub txtExpr_Change()
     
        busca
     
End Sub
