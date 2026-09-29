VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCrediario 
   Caption         =   "Crediário"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10560
   Icon            =   "frmCrediario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   10560
   Begin MSDataGridLib.DataGrid grid 
      Height          =   6855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   12091
      _Version        =   393216
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
      Caption         =   "Pendências para Crediário"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "num_ven"
         Caption         =   "Número"
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
         DataField       =   "nomecli_ven"
         Caption         =   "Nome do Cliente na Venda"
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
         DataField       =   "nome_cli"
         Caption         =   "Cliente cadastrado"
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
         DataField       =   "valor_ven"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "##0,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "guerra_fun"
         Caption         =   "Vendendor"
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
         DataField       =   "data_cre"
         Caption         =   "Horário"
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
   Begin VB.Timer tmr 
      Interval        =   3000
      Left            =   2640
      Top             =   4200
   End
   Begin VB.PictureBox pcBotoes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   10560
      TabIndex        =   5
      Top             =   7305
      Width           =   10560
      Begin VB.CommandButton cmndbtnCliente 
         Caption         =   "F8 &Novo Cliente"
         Height          =   825
         Left            =   4200
         Picture         =   "frmCrediario.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   825
         Left            =   120
         Picture         =   "frmCrediario.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   825
         Left            =   10200
         Picture         =   "frmCrediario.frx":1D1E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnFinalizarVenda 
         Caption         =   "&F7 Finalizar Pendência"
         Height          =   825
         Left            =   2160
         Picture         =   "frmCrediario.frx":1E89
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar"
         Height          =   825
         Left            =   6240
         Picture         =   "frmCrediario.frx":2213
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmCrediario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim iPos As Integer
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    'iPos = ado.AbsolutePosition
End Sub

Private Sub cmndbtnCancel_Click()
    On Error GoTo trata_Erro
    'Permissao_usu.cred_cancelar_crediario = True
    'If Permissao_usu.cred_cancelar_crediario = False Then
    '    MsgBox "Você não tem permissão para este acesso!", vbExclamation
    '    Exit Sub
    'End If
    If MsgBox("Confirma cancelamento do crediário?", vbYesNo + vbCritical) = vbNo Then
        
        Exit Sub
    End If
    If ado!status_cre <> 1 Then
        MsgBox "Este crediário não pode ser cancelado!", vbExclamation
        Exit Sub
    End If
    strSQL = "delete from tab_crediarioven where num_ven='" & ado!num_ven & "' "
    Conn.Execute strSQL
    strSQL = "update tab_vendas set situacao_ven='1',crediario_ven='0' where num_ven='" & ado!num_ven & "' "
    Conn.Execute strSQL
    MsgBox "Crediário cancelado, a venda foi reaberta!", vbExclamation
    
    busca
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmndbtnCliente_Click()
    Dim fr As New frmCadCli
    fr.Show 0
    
End Sub


Private Sub cmndbtnFinalizarVenda_Click()
    On Error GoTo trata_Erro
     
    
    With ado
    
        frmPendenciaCrediario.caption = "Pendências de Crediário - Venda: " & !num_ven
        frmPendenciaCrediario.txtbxNomeCli.tag = !cod_cli
        frmPendenciaCrediario.tag = !num_ven
        
        frmPendenciaCrediario.Show 0
        
        frmPendenciaCrediario.SetFocus
        Do While frmPendenciaCrediario.Visible
            DoEvents
        Loop
        
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnImpr_Click()
    ImprVenda ado!num_ven
    
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dtVendas_Reposition()
    
End Sub

Private Sub Form_Activate()
    bActive = True
    Form_Resize
    tmr_Timer
End Sub

Private Sub Form_Load()
    bActive = False
    Me.top = 10
    Set ado = New ADODB.Recordset
    iPos = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'dbgrd.Height = pcBotoes.Top - dbgrd.Top - 100
    'dbgrd.Width = Me.Width - dbgrd.Left - 100
    PosicionaGrade grid, Me, pcBotoes
    posicionaADireita cmndbtnSair, pcBotoes
    
    
End Sub

Private Sub tmr_Timer()
     busca
End Sub
Sub busca()
    Dim strSQL As String
    
    
    
    Dim i As Integer
    If bActive = False Then Exit Sub
    
    On Error Resume Next
        iPos = ado.AbsolutePosition
    On Error GoTo trata_Erro
    strSQL = "select tv.cod_cli,tc.nome_cli,tv.valor_ven,data_ven,tv.nomecli_ven,data_cre,tp.*,tf.guerra_fun from "
    strSQL = strSQL & TAB_VENDAS & " as tv inner join " & TAB_CLIENTES & " as tc on tc.cod_cli=tv.cod_cli "
    strSQL = strSQL & " inner join tab_Crediarioven as tp on tp.num_ven=tv.num_ven "
    strSQL = strSQL & " inner join tab_funcionarios as tf on tf.cod_fun=tv.abrevenda_usu "
    strSQL = strSQL & " and status_cre='1' "
    strSQL = strSQL & " order by data_cre desc "
    
    
    
    
    abreObjetoAdo ado, strSQL, grid
    
    
    If ado.RecordCount > 0 Then
        ado.Move iPos - 1
        
    End If
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
    
End Sub


