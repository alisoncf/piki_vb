VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMonitoramentoCaixas 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Monitoramento dos Caixas"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   8400
      Top             =   5040
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4545
      ScaleWidth      =   8625
      TabIndex        =   1
      Top             =   120
      Width           =   8655
      Begin VB.PictureBox pcModo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnBlock 
         Caption         =   "Bloquear Caixa"
         Height          =   375
         Left            =   5640
         TabIndex        =   27
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Modo:"
         Height          =   255
         Left            =   5640
         TabIndex        =   28
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label lblFinal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   6240
         TabIndex        =   24
         Top             =   3457
         Width           =   1065
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final (A+B):"
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
         Top             =   3637
         Width           =   1560
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
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7680
         TabIndex        =   19
         Top             =   2115
         Width           =   675
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
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   17
         Top             =   3315
         Width           =   675
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
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7680
         TabIndex        =   15
         Top             =   1515
         Width           =   675
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
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7680
         TabIndex        =   13
         Top             =   1035
         Width           =   675
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
         Alignment       =   1  'Right Justify
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
         Left            =   3000
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
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   10
         Top             =   2760
         Width           =   2010
      End
      Begin VB.Label lblAp 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   9
         Top             =   1875
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas a Prazo:"
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
         TabIndex        =   8
         Top             =   1920
         Width           =   2010
      End
      Begin VB.Label lblNomeFun 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Funcionário"
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
         Left            =   3600
         TabIndex        =   7
         Top             =   75
         Width           =   3540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário:"
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
         Left            =   1800
         TabIndex        =   6
         Top             =   120
         Width           =   1590
      End
      Begin VB.Label lblTur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Left            =   1320
         TabIndex        =   5
         Top             =   75
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
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
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblAv 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   3
         Top             =   1035
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas a Vista:"
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
         TabIndex        =   2
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   1095
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   3270
         Width           =   4095
      End
   End
   Begin MSComctlLib.TabStrip tbstrp 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   1296
      MultiRow        =   -1  'True
      Placement       =   1
      Separators      =   -1  'True
      TabMinWidth     =   3528
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "F7 - Fechar Caixa"
      Height          =   195
      Left            =   7680
      MouseIcon       =   "frmMonitoramentoCaixas.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   5520
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
    Conn.Execute strSQL
    
End Sub
Private Sub Form_Activate()
    busca
    buscaCAixa
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    tbstrp.Tabs.Clear
    strSQL = "select * from tab_caixas order by cod_cai "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
        
            tbstrp.Tabs.Add , , !cod_cai & "-" & !nome_cai
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
        strSQL = "select tf.guerra_fun from tab_turncai as tt, "
        strSQL = strSQL & " tab_Funcionarios as tf where cod_cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and cod_Tur='" & lblTur.caption & "' "
        strSQL = strSQL & " and data_tur=" & F_GravaData(Date)
        strSQL = strSQL & " and tf.cod_fun=tt.login_usu "
        FechaObjeto rPesq
        abreObjetoAdo rPesq, strSQL
        
        If rPesq.RecordCount > 0 Then
            lblNomeFun.caption = rPesq!guerra_fun
        Else
            lblNomeFun.caption = "Não Há funcionário logado!"
        End If
        strSQL = "select sum(tf.valor_ven) as total,nome_for,esp_cai  "
        strSQL = strSQL & " from tab_vendas as tv,tab_venfec as tf, tab_ "
        strSQL = strSQL & " where tf.num_ven=tv.num_ven and tf.cod_cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and tf.cod_tur='" & lblTur.caption & "'"
        strSQL = strSQL & " and tv.dataefet_ven=" & F_GravaData(Date)
        FechaObjeto rPesq
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
        strSQL = "select sum(valor_cai) as total from tab_caixa where cod_Cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and cod_tur='" & lblTur & "' "
        strSQL = strSQL & " and num_ven='' and tipo_cai='C' "
        strSQL = strSQL & " and data_cai=" & F_GravaData(Date)
        strSQL = strSQL & " and excluido_cai=false "
        abreObjetoAdo rPesq, strSQL
        lblEntr.caption = FormatVal(rPesq!total)
        strSQL = "select sum(valor_cai) as total from tab_caixa where cod_Cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and cod_tur='" & lblTur & "'"
        strSQL = strSQL & " and num_ven='' and tipo_cai='D' "
        strSQL = strSQL & " and data_cai=" & F_GravaData(Date)
        strSQL = strSQL & " and excluido_cai=false "
        abreObjetoAdo rPesq, strSQL
        lblSai.caption = FormatVal(rPesq!total)
        lblTotalMov.caption = FormatVal(CCur(lblEntr.caption) - CCur(lblSai.caption))
        lblFinal.caption = CCur(lblTotalVendas.caption) + CCur(lblTotalMov.caption)
        rPesq.Close
        strSQL = "select sum(valor_cup) as total from tab_cupons "
        strSQL = strSQL & " where cod_cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and cod_tur='" & lblTur & "' "
        strSQL = strSQL & " and emissao_Sai=" & F_GravaData(Date)
        abreObjetoAdo rPesq, strSQL
        lblECF.caption = FormatVal(rPesq!total)
        If CCur(lblTotalVendas.caption) > 0 Then
        
            lblPercEnt.caption = Round(rPesq!total / CCur(lblTotalVendas.caption) * 100, 2) & " %"
        Else
            lblPercEnt.caption = "0,00 %"
        End If
    End With
    verModoCai
End Sub
Sub verModoCai()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select modo_cai,bloqueado_cai from tab_caixas where cod_cai='" & sCod_Cai & "' "
    abreObjetoAdo rPesq, strSQL
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
