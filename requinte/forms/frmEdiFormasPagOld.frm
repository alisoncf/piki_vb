VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiFormasPagOld 
   Caption         =   "Edição de Formas de Pagamento"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   Icon            =   "frmEdiFormasPagOld.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7470
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDesab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Carregar Desabilitadas"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9720
      TabIndex        =   12
      Top             =   6120
      Width           =   1935
   End
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
      ScaleWidth      =   11940
      TabIndex        =   5
      Top             =   6495
      Width           =   11970
      Begin VB.CommandButton cmdDesab 
         Caption         =   "Desabilitar"
         Height          =   735
         Left            =   7560
         Picture         =   "frmEdiFormasPagOld.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdGerenciais 
         Caption         =   "Carregar formas gerenciais"
         Height          =   735
         Left            =   5520
         TabIndex        =   10
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3720
         Picture         =   "frmEdiFormasPagOld.frx":090C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiFormasPagOld.frx":0C4E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1920
         Picture         =   "frmEdiFormasPagOld.frx":0FD8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   9960
         Picture         =   "frmEdiFormasPagOld.frx":1122
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiFormasPagOld.frx":1670
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9975
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         DataField       =   "parcelas_cai"
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
Attribute VB_Name = "frmEdiFormasPagOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents adc As ADODB.Recordset
Attribute adc.VB_VarHelpID = -1
Public SoAvista As Boolean

Private Sub chkDesab_Click()
    busca
End Sub

Private Sub cmdDesab_Click()
    If strAltformaspag_usu <> "S" Or strNivel_Usu < "2" Then
        MsgBox "Você não tem permissão para essa operação!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja realmente desabilitar essa forma de pagamento para novas vendas?", vbYesNo + vbQuestion) = vbYes Then
        If FormasPag.DesabilitaForma(adc!cod_For, IIf(adc!visivel_for = "S", False, True)) Then
            adc.Requery
        End If
    End If
End Sub

Private Sub cmdGerenciais_Click()
    Dim sUsu As String
    If strAltformaspag_usu <> "S" Then
        
        sUsu = F_senhaReturnUsu()
        If sUsu = "" Then
            Exit Sub
        End If
        If F_BuscaUmCampo("altformaspag_usu", "tab_defusu", "cod_usu='" & sUsu & "'") <> "S" Then
            MsgBox "Este usuário não possui permissão para esse acesso!", vbExclamation
            Exit Sub
        End If
        
    End If
    
    
    busca True
    
End Sub

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
        Dim fr As New frmIncFormasPag
        fr.tag = !cod_For
        fr.txtbxDesc.Text = !nome_for
        fr.txtbxInterv.Text = !intervalo_for
        'fr.txtbxJuros.Text = FormatVal(!juros_for)
        fr.txtbxParc.Text = !parcelas_for
        'If !Entrada_tip = "S" Then
        '    fr.cmbxEntr.ListIndex = 0
        'Else
        '    fr.cmbxEntr.ListIndex = 1
        'End If
        If Left(!tipo_cai, 1) = "C" Then
            fr.optnChe.Value = True
        Else
            fr.optnDup.Value = True
        End If
        fr.Show 0
        
    End With
    busca
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmIncFormasPag
    fr.Show 0
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
    If strAltformaspag_usu <> "S" Then
        chkDesab.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    
    Set adc = New ADODB.Recordset
    
    
End Sub

Sub busca(Optional bGerencial As Boolean = False)
    Dim strSQL As String
    strSQL = "select tf.*, tt.nome_tip from tab_formaspag as tf, tab_tipopag as tt "
    strSQL = strSQL & " where tt.cod_tip=tf.cod_tip "
    If txtCod.Text <> "" Then
        strSQL = strSQL & " and cod_for='" & txtCod.Text & "' "
    Else
        strSQL = strSQL & " and nome_for like '%" & txtExpr.Text & "%'"
    End If
    
    If Not bGerencial Then
        strSQL = strSQL & " and gerencial_for='N' "
    Else
        'strSQL = strSQL & " and gerencial_for='S' "
    End If
    
    If SoAvista Then
        strSQL = strSQL & " and tt.cod_tip='01' "
    End If
    
    If chkDesab.Value = 0 Then
        strSQL = strSQL & " and visivel_for='S' "
    End If
    
    strSQL = strSQL & " order by cod_for"
    abreObjetoAdo adc, strSQL, dtgrd
    
    If Me.tag <> "" Then
        If adc.RecordCount = 1 Then
            On Error Resume Next
            Me.tag = "x"
            
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

Private Sub txtExpr_GotFocus()
    P_SelecTexto txtExpr
End Sub

Private Sub txtExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dtgrd.SetFocus
    End If
End Sub
