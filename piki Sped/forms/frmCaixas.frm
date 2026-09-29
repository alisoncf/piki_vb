VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCaixas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caixas Cadastrados"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid gridcx 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Caixas"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cod_cai"
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
         DataField       =   "nome_cai"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2580,095
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   7560
      Picture         =   "frmCaixas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnEditar 
      Caption         =   "&Editar"
      Height          =   735
      Left            =   1440
      Picture         =   "frmCaixas.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnNovo 
      Caption         =   "&Novo Caixa"
      Height          =   735
      Left            =   120
      Picture         =   "frmCaixas.frx":0255
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid gridIp 
      Height          =   3975
      Left            =   6360
      TabIndex        =   4
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Computadores Conectados"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "ip_cai"
         Caption         =   "Ip"
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
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Excluir"
      Height          =   255
      Left            =   8040
      TabIndex        =   5
      Top             =   4200
      Width           =   735
   End
End
Attribute VB_Name = "frmCaixas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents adoCx As ADODB.Recordset
Attribute adoCx.VB_VarHelpID = -1
Dim WithEvents adoIp As ADODB.Recordset
Attribute adoIp.VB_VarHelpID = -1





Private Sub adoCx_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    buscaIp
End Sub
Sub buscaIp()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from tab_caixasip where cod_cai='" & adoCx!Cod_Cai & "' "
    abreObjetoAdo adoIp, strSQL, gridIp
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    If strExclcai_usu <> "S" Then
        MsgBox "Você não ter permissão para excluir caixa!", vbExclamation
        Exit Sub
    End If
    strSQL = "delete from tab_caixasip where cod_cai='" & adoIp!Cod_Cai & "' and ip_cai='" & adoIp!ip_cai & "'"
    Conn.Execute strSQL
    adoIp.Requery
    
    
End Sub

Private Sub cmndbtnEditar_Click()
    On Error Resume Next
    With adc.Recordset
        frmCadCai.tag = !Cod_Cai
        frmCadCai.caption = !Cod_Cai & "-" & !nome_cai
        frmCadCai.txtbxCod.Text = !Cod_Cai
        frmCadCai.txtbxConta.Text = !cod_con
        frmCadCai.txtbxIP.Text = !ip_cai
        frmCadCai.txtbxNome.Text = !nome_cai
        frmCadCai.txtbxCod.Enabled = False
        frmCadCai.Show 1
    End With
    
End Sub

Private Sub cmndbtnNovo_Click()
    frmCadCaixa.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
    
    If strExclcai_usu <> "S" Then
        cmdDel.Visible = False
    End If
    
End Sub

Private Sub Form_Load()
    Set adoCx = New ADODB.Recordset
    Set adoIp = New ADODB.Recordset
    
       
    
    
    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select * from tab_caixas order by nome_cai"
    abreObjetoAdo adoCx, strSQL, gridcx
        
    
    
    
End Sub

Private Sub gridIp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmdDel_Click
    End If
End Sub
