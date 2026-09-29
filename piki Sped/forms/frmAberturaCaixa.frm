VERSION 5.00
Begin VB.Form frmAberturaCaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abertura de Caixa"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "frmAberturaCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   28
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   7065
      TabIndex        =   3
      Top             =   240
      Width           =   7095
      Begin VB.Label lblDataAnterior 
         AutoSize        =   -1  'True
         Caption         =   "00/00/0000"
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
         Left            =   4710
         TabIndex        =   26
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblCaixa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   4680
         TabIndex        =   25
         Top             =   2280
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abertura de Caixa"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         TabIndex        =   24
         Top             =   0
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fundo de Caixa:"
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
         Left            =   30
         TabIndex        =   23
         Top             =   2400
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         Caption         =   "00/00/0000"
         Height          =   195
         Left            =   4830
         TabIndex        =   22
         Top             =   45
         Width           =   870
      End
      Begin VB.Label lblFundo 
         AutoSize        =   -1  'True
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
         Left            =   2310
         TabIndex        =   21
         Top             =   2400
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Em:"
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
         Left            =   4110
         TabIndex        =   20
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fechado Por:"
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
         Left            =   30
         TabIndex        =   19
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label lblNomeFec 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Funcionário"
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
         Left            =   1470
         TabIndex        =   18
         Top             =   1320
         Width           =   1800
      End
      Begin VB.Label lblValorAbertura 
         AutoSize        =   -1  'True
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
         Left            =   2310
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor de Abertura:"
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
         Left            =   30
         TabIndex        =   16
         Top             =   1680
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   1110
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "0 - Status do Turno"
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
         Left            =   1950
         TabIndex        =   14
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label lblNomeAbre 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Funcionário"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aberto Por:"
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
         Left            =   0
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblTurno 
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   870
         TabIndex        =   11
         Top             =   600
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
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
         Left            =   30
         TabIndex        =   10
         Top             =   600
         Width           =   570
      End
      Begin VB.Label lblHoraAbert 
         AutoSize        =   -1  'True
         Caption         =   "00:00"
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
         Left            =   4710
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Às:"
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
         Left            =   4110
         TabIndex        =   8
         Top             =   960
         Width           =   270
      End
      Begin VB.Label lblHoraFec 
         AutoSize        =   -1  'True
         Caption         =   "00:00"
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
         Left            =   4710
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Às:"
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
         Left            =   4110
         TabIndex        =   6
         Top             =   1320
         Width           =   270
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor de Fechamento:"
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
         Left            =   30
         TabIndex        =   5
         Top             =   2040
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lblValorFechamento 
         AutoSize        =   -1  'True
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
         Left            =   2310
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   390
      End
   End
   Begin VB.ComboBox cmbxTurno 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmndbtnConf 
      Caption         =   "Confirmar"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   3480
      Width           =   345
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Turno:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   3480
      Width           =   675
   End
End
Attribute VB_Name = "frmAberturaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub cmbxTurno_Click()
    busca
End Sub

Private Sub cmndbtnConf_Click()
    Dim strSQL As String
    If cmbxTurno.Text = "" Then
        MsgBox "Informe o Turno", vbCritical
        Exit Sub
    End If
    If strCod_Cai = "" Then
        MsgBox "Caixa Inválido!", vbExclamation
        Exit Sub
    End If
    
    If Left(lblStatus.caption, 1) = "A" Then
        
        If CDate(lblDataAnterior.caption) = Date Then
            MsgBox "O " & lblCaixa.caption & " já se encontra em aberto!" & _
            Chr(13) & "Não é possível o mesmo caixa no mesmo turno duas vezes no dia!", vbCritical, "Abertura de Caixa"
            Exit Sub
        Else
            If MsgBox("O movimento do turno: " & lblTurno.caption & " do caixa  " & _
                lblCaixa.caption & " do dia " & lblDataAnterior.caption & _
                " está aberto, deseja continuar?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        End If
         
    End If
     If Left(lblStatus.caption, 1) = "D" Then
        MsgBox "O " & lblCaixa.caption & " se encontra desabilitado!", vbExclamation
        Exit Sub
    End If
    If Left(lblStatus.caption, 1) = "I" Then
        MsgBox "Caixa inexistente!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma abertura do " & lblCaixa & "? ", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
'    If F_BuscaUmCampo("cod_tur", TAB_TURNCAI, "data_tur=" & _
'        F_GravaData(Date) & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' and cod_Cai='" & Left(lblCaixa.caption, 2) & "'") <> "" Then
'        MsgBox "Já existe movimentação para este turno Hoje!", vbCritical
'        Exit Sub
'    End If
    strSQL = F_senhaReturnUsu()
    If strSQL = "" Then
        Exit Sub
    End If
    If F_BuscaUmCampo("fechacai_usu", "" & TAB_DEFUSU & "", "cod_usu='" & strSQL & "'") <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    AbreCaixa
End Sub
Sub AbreCaixa()
    Dim strSQL As String
    strSQL = "replace into " & TAB_TURNCAI & " (cod_tur,cod_cai,data_tur,login_usu,abertura_tur,"
    strSQL = strSQL & "valorabertura_tur,status_tur)values("
    strSQL = strSQL & "'" & Left(cmbxTurno.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(lblCaixa.caption, 2) & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & Hour(Now) & "',"
    strSQL = strSQL & F_GravaDec(lblFundo.caption) & ","
    strSQL = strSQL & "'A') "
    conn.Execute strSQL
    
    strSQL = "update tab_caixas set status_Cai='A',"
    strSQL = strSQL & "cod_tur='" & Left(cmbxTurno.Text, 1) & "', "
    strSQL = strSQL & " ip_cai='" & strIP_Usu & "' "
    strSQL = strSQL & " where cod_cai='" & Left(lblCaixa.caption, 2) & "' "
    conn.Execute strSQL
    strCod_Tur = Left(cmbxTurno.Text, 1)
    
    strSQL = F_GravaCai("C", Date, lblFundo.caption, "FUNDO DE CAIXA REF " & lblData.caption, _
        "", "TRA", "", 0, 0, Left(lblCaixa.caption, "01"), strCod_Tur, "", "")
    If strSQL = "" Then
        'MsgBox "Erro no lançamento de fundo de caixa!", vbExclamation
    End If
    
    strSQL = "update " & TAB_CAIXAS & " set status_cai='A',cod_tur='" & strCod_Tur & "'"
    strSQL = strSQL & " where cod_cai='" & strCod_Cai & "'"
    conn.Execute strSQL
    
    
    MsgBox "Caixa aberto com sucesso!", vbInformation
    
    
    
    busca
End Sub
Private Sub Form_Activate()
     P_PreencheTurnos cmbxTurno
     busca
     
End Sub

Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sStatus As String
    lblData.caption = Date
    lblCaixa.caption = strCod_Cai & " - " & F_BuscaUmCampo("nome_cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
    sStatus = F_BuscaUmCampo("status_Cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
    
    strSQL = "select  tt.*,max(data_tur) as maximo,cod_tur from  " & TAB_TURNCAI & " as tt "
    strSQL = strSQL & " where cod_cai='" & Left(lblCaixa.caption, 2) & "' "
    strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    strSQL = strSQL & " group by data_tur order by maximo desc,cod_tur desc"
    abreObjetoAdo rPesq, strSQL
    
    
    If rPesq.RecordCount > 0 Then
        'strSQL = "select * from tab_fechacaixa where data_cai=" & F_PesqData(rPesq!maximo)
        'strSQL = strSQL & " and cod_cai='" & Left(lblCaixa.caption, 2) & "' "
'        rpesq.open  strsql, conn
        
    End If
    With rPesq
        If .RecordCount > 0 Then
            lblTurno.caption = !cod_tur
            lblDataAnterior.caption = !data_tur
            lblFundo.caption = FormatVal(!fundo_tur)
            lblValorAbertura.caption = FormatVal(!valorfechamento_tur)
            lblValorFechamento.caption = FormatVal(!valorfechamento_tur)
            lblNomeFec.caption = !logout_usu & " - " & F_BuscaUmCampo("guerra_fun", TAB_FUNCIONARIOS, "cod_fun='" & !logout_usu & "'")
            lblNomeAbre.caption = !login_usu & " - " & F_BuscaUmCampo("guerra_fun", TAB_FUNCIONARIOS, "cod_fun='" & !login_usu & "'")
            lblHoraAbert.caption = !abertura_tur

            lblHoraFec.caption = Format(!fechamento_tur)
            sStatus = !status_tur
            If sStatus = "A" Then
                lblStatus.caption = "A - Aberto"
            ElseIf sStatus = "F" Then
                lblStatus.caption = "F - Fechado"
            ElseIf sStatus = "D" Then
                lblStatus.caption = "D - Desabilitado"
            Else
                lblStatus.caption = "I - Inexistente"
            End If
        Else
            lblTurno.caption = F_BuscaUmCampo("cod_tur", "tab_caixas", "cod_cai='" & Left(lblCaixa.caption, 2) & "'")
            lblDataAnterior.caption = ""
            lblFundo.caption = "0,00"
            lblValorAbertura.caption = FormatVal(!valorfechamento_tur)
            lblValorFechamento.caption = FormatVal(!valorfechamento_tur)
            lblNomeFec.caption = ""
            lblNomeAbre.caption = ""
            lblHoraAbert.caption = ""

            lblHoraFec.caption = ""
            
            
                lblStatus.caption = "F - Fechado"
            End If
    
    
        
        lblStatus = lblStatus.caption
        .Close
    End With
End Sub


Private Sub Form_Load()
    txtData.Text = Date
End Sub
