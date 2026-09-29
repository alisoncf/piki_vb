VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReemissao 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vendas na Data:"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13335
   Icon            =   "frmReemissao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   9000
      ScaleHeight     =   3705
      ScaleWidth      =   4305
      TabIndex        =   7
      Top             =   960
      Width           =   4335
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Emitir"
         Height          =   660
         Left            =   120
         Picture         =   "frmReemissao.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtCPF 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label lblICMS 
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
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   345
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   13335
      TabIndex        =   1
      Top             =   0
      Width           =   13335
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   660
         Left            =   11160
         Picture         =   "frmReemissao.frx":07CC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin VB.ComboBox cmbxCaixa 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtbxData 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   390
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmReemissao.frx":08C1
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6800
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "dataefet_ven"
         Caption         =   "Data"
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
      BeginProperty Column02 
         DataField       =   "nomecli_ven"
         Caption         =   "Cliente"
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
         DataField       =   "fechavenda_usu"
         Caption         =   "vendedor"
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
         DataField       =   "id"
         Caption         =   "Id.Cupom"
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
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3330,142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   945,071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid gridProd 
      Bindings        =   "frmReemissao.frx":08D3
      Height          =   2175
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   3836
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cod_Pro"
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
         DataField       =   "desc_ite"
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
         DataField       =   "qtd_ite"
         Caption         =   "Qtd"
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
         DataField       =   "unit_ite"
         Caption         =   "Unit"
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
         DataField       =   "final_ite"
         Caption         =   "Total"
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
         DataField       =   "icms"
         Caption         =   "ICMS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
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
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3330,142
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   945,071
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReemissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim WithEvents adoProd As ADODB.Recordset
Attribute adoProd.VB_VarHelpID = -1
Private Sub chkTodas_Click()
    P_Busca
End Sub

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    buscaProd
End Sub

Private Sub cmbxCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    ElseIf KeyCode = vbKeyDelete Then
        cmbxCaixa.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    P_Busca
End Sub

Private Sub cmndbtnImpr_Click()
    Dim sID As String
    Dim CFOP As String
    If MsgBox("Confirma reemissão da venda? ", vbYesNo + vbQuestion) = vbYes Then
        If F_EmiteCupomVenda(ado!num_ven, 12, txtCliente.Text, txtCPF.Text, ado!valor_ven) Then
            P_Status ""
            MsgBox "Cupom Impresso!", vbInformation
            
            'P_GravaSaida  ado!num_ven, strNum_Cpm, IIf(pc.BackColor = vbBlue, True, False)
            
            sID = SpedCupom.gravasped_cupom(ado!num_ven, ado!Cod_cai, Date, Time, SpedModeloCUPOM, SpedSituacaoRegular, _
                strNum_Cpm, ado!valor_ven, "0", "0", txtCPF.Text, txtCliente.Text)
                If sID <> "" Then
                    With adoProd
                        Do While Not .EOF
                            If !icms_pro = "FF" Then
                                CFOP = spedCFOPPreferencialCUPOMRetida
                            Else
                                CFOP = spedCFOPPreferencialCUPOMTributada
                            End If
                             SpedCupom.gravasped_cupom_itens sID, !cod_pro, !qtd_ite, "0", !unidade_pro, !unit_ite, _
                             !cst_icms, CFOP, !perc_imp, "0", "0", !icms, !desconto_ite
                            
                            .MoveNext
                        Loop
                    End With
                End If
            'P_BaixaSaidaEF ado!num_ven
        End If
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVis_Click()
    ImprVenda ado!num_ven
    
End Sub
Function F_JustificaTexto(sText As String, tamanho As Integer) As String
    If Len(sText) >= tamanho Then
        sText = Left(sText, tamanho)
    Else
        sText = sText + Space(tamanho - Len(sText))
    End If
    F_JustificaTexto = sText
End Function



Private Sub Form_Activate()
    
    P_Busca
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        pc.BackColor = vbGreen
        P_Busca
        cmndbtnImpr.SetFocus
        
    ElseIf KeyCode = vbKeyF9 Then
        pc.BackColor = vbYellow
        P_Busca
        cmndbtnImpr.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set ado = New ADODB.Recordset
    Set adoProd = New ADODB.Recordset
    txtbxData.Text = Date
    
    
    P_PreencheCaixas cmbxCaixa
    
        P_LocalizaOperacao strCod_Cai, cmbxCaixa
        
    If strExclcai_usu <> "S" Then
        cmbxCaixa.Enabled = False
    End If
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    On Error GoTo trata_Erro
  
    
    strSQL = "select tv.*, c.id from " & TAB_VENDAS & " as tv left join sped_cupom as c on c.num_ven=tv.num_ven "
    
    strSQL = strSQL & " where Dataefet_ven=" & F_GravaData(txtbxData.Text)
    If cmbxCaixa.Text <> "" Then
        strSQL = strSQL & " and tv.cod_cai='" & Left(cmbxCaixa.Text, 2) & "'  "
    End If
    
        
    strSQL = strSQL & " order by num_ven desc "
    abreObjetoAdo ado, strSQL, dtgrd
    Exit Sub
trata_Erro:
MostraErro
    
End Sub
Sub buscaProd()
    Dim strSQL As String
    
    On Error GoTo trata_Erro
    gridProd.Visible = False
    
    txtCliente.Text = ado!nomecli_ven
    txtCPF.Text = F_BuscaUmCampo("cnpj_cli", TAB_CLIENTES, "cod_cli=" & ado!cod_cli & "")
    lblCliente.Visible = True
    txtCliente.Visible = True
    lblCPF.Visible = True
    txtCPF.Visible = True
    
    
    strSQL = "select tp.unidade_pro,perc_imp,ti.*,(final_ite*(perc_imp/100)) as icms from " & TAB_ITVENDAS & " as ti inner join " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " on tp.cod_pro=ti.cod_pro  left join  " & TAB_IMPOSTOS & " as imp on tp.cod_imp=imp.cod_imp "
    strSQL = strSQL & " where num_ven='" & ado!num_ven & "' "
    abreObjetoAdo adoProd, strSQL, gridProd
    
    gridProd.Visible = True
    Exit Sub
trata_Erro:
    gridProd.Visible = False
End Sub
    
Private Sub Text2_Change()

End Sub

Private Sub txtbxData_GotFocus()
    ganhaFoco txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        P_Busca
    End If
    If KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
    P_Busca
End Sub

