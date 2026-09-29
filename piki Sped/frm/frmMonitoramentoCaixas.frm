VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMonitoramentoCaixas 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Monitoramento dos Caixas"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   12345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TabStrip tbstrp 
      Height          =   615
      Left            =   240
      TabIndex        =   31
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   11400
      Top             =   6120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   240
      ScaleHeight     =   6825
      ScaleWidth      =   11385
      TabIndex        =   1
      Top             =   600
      Width           =   11415
      Begin VB.CommandButton cmdSuspende 
         Caption         =   "Suspender Caixa"
         Height          =   375
         Left            =   5640
         TabIndex        =   34
         Top             =   3120
         Width           =   1575
      End
      Begin VB.PictureBox pcTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1395
         Left            =   6720
         Picture         =   "frmMonitoramentoCaixas.frx":0000
         ScaleHeight     =   1395
         ScaleWidth      =   4050
         TabIndex        =   32
         Top             =   5040
         Width           =   4050
         Begin VB.Label lblFinal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   29.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   660
            Left            =   1560
            TabIndex        =   33
            Top             =   360
            Width           =   1050
         End
      End
      Begin VB.CommandButton cmndbtnVisualizaItCnc 
         Caption         =   "Visualizar"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   4800
         Width           =   1095
      End
      Begin VB.PictureBox pcModo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnBlock 
         Caption         =   "Bloquear Caixa"
         Height          =   375
         Left            =   5640
         TabIndex        =   26
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Itens Cancelados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   30
         Top             =   4485
         Width           =   2115
      End
      Begin VB.Label lblItensCanc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2940
         TabIndex        =   29
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Modo:"
         Height          =   255
         Left            =   5640
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblstatusCai 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa Aberto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F8AC67&
         Height          =   360
         Left            =   360
         TabIndex        =   24
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   4560
         TabIndex        =   23
         Top             =   3630
         Width           =   795
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Left            =   3015
         TabIndex        =   22
         Top             =   3960
         Width           =   390
      End
      Begin VB.Label lblECF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Left            =   1440
         TabIndex        =   21
         Top             =   3960
         Width           =   525
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saídas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Left            =   360
         TabIndex        =   20
         Top             =   3960
         Width           =   810
      End
      Begin VB.Line Line2 
         X1              =   5640
         X2              =   8400
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblTotalMov 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7620
         TabIndex        =   19
         Top             =   2115
         Width           =   735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5640
         TabIndex        =   18
         Top             =   2160
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3840
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label lblTotalVendas 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2835
         TabIndex        =   17
         Top             =   3315
         Width           =   765
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   16
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblSai 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7620
         TabIndex        =   15
         Top             =   1515
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saídas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5640
         TabIndex        =   14
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblEntr 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7620
         TabIndex        =   13
         Top             =   1035
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entradas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5640
         TabIndex        =   12
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label lblFinan 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2880
         TabIndex        =   11
         Top             =   2715
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas Financ.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   10
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblAp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2835
         TabIndex        =   9
         Top             =   1875
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas a Prazo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   1740
      End
      Begin VB.Label lblNomeFun 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3840
         TabIndex        =   7
         Top             =   75
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1800
         TabIndex        =   6
         Top             =   120
         Width           =   1290
      End
      Begin VB.Label lblTur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   75
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   645
      End
      Begin VB.Label lblAv 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2835
         TabIndex        =   3
         Top             =   1035
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas a Vista:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   1665
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "F7 - Fechar Caixa"
      Height          =   195
      Left            =   9960
      MouseIcon       =   "frmMonitoramentoCaixas.frx":1273E
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   7800
      Width           =   1245
   End
End
Attribute VB_Name = "frmMonitoramentoCaixas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sCod_Cai As String
Private Sub cmndbtnBlock_Click()
    If cmndbtnBlock.tag = "" Then
        blockCaixa True
    Else
        blockCaixa False
    End If
End Sub
Sub blockCaixa(opcao As Boolean)
    Dim strSQL As String
    On Error Resume Next
    strSQL = "update tab_caixas set bloqueado_cai='" & IIf(opcao = True, "S", "N") & "'"
    strSQL = strSQL & " where cod_cai='" & sCod_Cai & "' "
    conn.Execute strSQL
    
End Sub

Private Sub cmndbtnVisualizaItCnc_Click()
On Error Resume Next
 P_LocalizaOperacao strCod_Cai, frmRelItensCancelados.cmbxCai
    frmRelItensCancelados.cmbxTurno.Text = strCod_Tur
    frmRelItensCancelados.Show 1
End Sub

Private Sub Form_Activate()
    Busca
    buscaCAixa
End Sub
Sub Busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    tbstrp.Tabs.Clear
    strSQL = "select * from tab_caixas order by cod_cai "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
        
            tbstrp.Tabs.Add , , !Cod_cai & "-" & !nome_cai
            .MoveNext
        Loop
        .Close
        
    End With
End Sub
Sub buscaCAixa()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    sCod_Cai = Left(tbstrp.Tabs(tbstrp.SelectedItem.Index), 2)
    strSQL = "select * from tab_caixas where cod_cai='" & sCod_Cai & "'  "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        lblTur.caption = !cod_tur
        If !status_cai = "A" Then
            lblstatusCai.caption = "Caixa Aberto"
        Else
            lblstatusCai.caption = "Caixa Fechado"
        End If
        strSQL = "select tf.guerra_fun from " & TAB_TURNCAI & " as tt, "
        strSQL = strSQL & " tab_Funcionarios as tf where cod_cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and cod_Tur='" & lblTur.caption & "' "
        strSQL = strSQL & " and data_tur=" & F_GravaData(Date)
        strSQL = strSQL & " and tf.cod_fun=tt.login_usu "
        abreObjetoAdo rPesq, strSQL
        If rPesq.RecordCount > 0 Then
            lblNomeFun.caption = rPesq!guerra_fun
        Else
            lblNomeFun.caption = "Não Há funcionário logado!"
        End If
        strSQL = "select sum(tf.avista_ven) as av,sum(tf.aprazo_ven) as ap,"
        strSQL = strSQL & "sum(tf.finan_ven) as finan from " & TAB_VENDAS & " as tv, " & TAB_VENFEC & "  as tf "
        strSQL = strSQL & " where tf.num_ven=tv.num_ven and tf.cod_cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and tf.cod_tur='" & lblTur.caption & "'"
        strSQL = strSQL & " and tv.dataefet_ven=" & F_GravaData(Date)
        abreObjetoAdo rPesq, strSQL
        With rPesq
            If .RecordCount > 0 Then
                lblAv.caption = FormatVal(!av)
                lblAp.caption = FormatVal(!ap)
                lblFinan.caption = FormatVal(!finan)
                lblTotalVendas.caption = FormatVal(!av + !ap + !finan)
            End If
            .Close
        End With
        strSQL = "select sum(valor_cai) as total from " & TAB_CAIXA & "  where cod_Cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and cod_tur='" & lblTur & "' "
        strSQL = strSQL & " and num_ven='' and tipo_cai='C' "
        strSQL = strSQL & " and data_cai=" & F_GravaData(Date)
        rPesq.Open strSQL, conn_client
        lblEntr.caption = FormatVal(rPesq!total)
        strSQL = "select sum(valor_cai) as total from " & TAB_CAIXA & "  where cod_Cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and cod_tur='" & lblTur & "'"
        strSQL = strSQL & " and num_ven='' and tipo_cai='D' "
        strSQL = strSQL & " and data_cai=" & F_GravaData(Date)
        abreObjetoAdo rPesq, strSQL
        lblSai.caption = FormatVal(rPesq!total)
        lblTotalMov.caption = FormatVal(CCur(lblEntr.caption) - CCur(lblSai.caption))
        lblFinal.caption = CCur(lblTotalVendas.caption) + CCur(lblTotalMov.caption)
        rPesq.Close
        strSQL = "select sum(vNF) as total from  nfe as n inner join nfe_icmstot as ni on n.id=ni.id_nfe "
        strSQL = strSQL & " inner join " & TAB_VENDAS & " as tv on tv.num_ven=n.num_ven "
        strSQL = strSQL & " where status=2 "
        strSQL = strSQL & " and tv.cod_cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and cod_tur='" & lblTur & "' "
        strSQL = strSQL & " and dataefet_ven=" & F_GravaData(Date)
        rPesq.Open strSQL, conn_client
        lblECF.caption = FormatVal(rPesq!total)
        If CCur(lblTotalVendas.caption) > 0 Then
        
            lblPercent.caption = Round(rPesq!total / CCur(lblTotalVendas.caption) * 100, 2) & " %"
        Else
            lblPercent.caption = "0,00 %"
        End If
        .Close
    End With
    
    
    verModoCai
End Sub
Sub buscaCancel()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
End Sub
Sub verModoCai()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select modo_cai,bloqueado_cai from tab_caixas where cod_cai='" & sCod_Cai & "' "
    rPesq.Open strSQL, conn_client
    If rPesq!modo_cai = "C" Then
        pcModo.BackColor = vbGreen
    ElseIf rPesq!modo_cai = "B" Then
        pcModo.BackColor = vbYellow
    Else
        pcModo.BackColor = vbWhite
    End If
    If rPesq!bloqueado_cai = "S" Then
        cmndbtnBlock.caption = "&Liberar Caixa"
        cmndbtnBlock.tag = "B"
    Else
        cmndbtnBlock.caption = "&Bloquear Caixa"
        cmndbtnBlock.tag = ""
    End If
    rPesq.Close
    Exit Sub
trata_Erro:
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        Label2_Click
    End If
End Sub

Private Sub Form_Load()
    tbstrp.Tabs.Clear
End Sub

Private Sub Label2_Click()
    If strFechacai_usu = "S" Then
        frmMovDiarioCx.Show 1
    Else
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
    End If
End Sub

Private Sub tbstrp_Click()
    buscaCAixa
End Sub

Private Sub Timer1_Timer()
    buscaCAixa
End Sub
