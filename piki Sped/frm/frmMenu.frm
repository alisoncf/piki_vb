VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Menu de Opções"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14400
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   14400
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   14370
      TabIndex        =   19
      Top             =   8895
      Width           =   14400
      Begin VB.Label lblTool 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   10545
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   14370
      TabIndex        =   17
      Top             =   9150
      Width           =   14400
      Begin VB.Label lblNomeEmp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   11745
      End
   End
   Begin VB.PictureBox pcControle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   0
      ScaleHeight     =   8295
      ScaleWidth      =   14175
      TabIndex        =   16
      Top             =   0
      Width           =   14175
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6750
         Left            =   4320
         Picture         =   "frmMenu.frx":3132
         ScaleHeight     =   6750
         ScaleWidth      =   9750
         TabIndex        =   15
         Top             =   240
         Width           =   9750
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DrawWidth       =   2
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   0
            Left            =   240
            Picture         =   "frmMenu.frx":C87C
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   0
            Top             =   480
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   1
            Left            =   240
            Picture         =   "frmMenu.frx":1489D
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   1
            Top             =   1680
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   2
            Left            =   240
            Picture         =   "frmMenu.frx":1CC4B
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   2
            Top             =   2880
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   3
            Left            =   240
            Picture         =   "frmMenu.frx":24273
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   3
            Top             =   4080
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   4
            Left            =   240
            Picture         =   "frmMenu.frx":2BED3
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   4
            Top             =   5280
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   14
            Left            =   6480
            Picture         =   "frmMenu.frx":33F18
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   14
            Top             =   5280
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   13
            Left            =   6480
            Picture         =   "frmMenu.frx":3B4A9
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   13
            Top             =   4080
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   12
            Left            =   6480
            Picture         =   "frmMenu.frx":42E9C
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   12
            Top             =   2880
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   11
            Left            =   6480
            Picture         =   "frmMenu.frx":4B3CD
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   11
            Top             =   1680
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   10
            Left            =   6480
            Picture         =   "frmMenu.frx":52B83
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   10
            Top             =   480
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   9
            Left            =   3360
            Picture         =   "frmMenu.frx":5A35C
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   9
            Top             =   5280
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   8
            Left            =   3360
            Picture         =   "frmMenu.frx":61E03
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   8
            Top             =   4080
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   7
            Left            =   3360
            Picture         =   "frmMenu.frx":69367
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   7
            Top             =   2880
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   6
            Left            =   3360
            Picture         =   "frmMenu.frx":70E8C
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   6
            Top             =   1680
            Width           =   3000
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1050
            Index           =   5
            Left            =   3360
            Picture         =   "frmMenu.frx":786C5
            ScaleHeight     =   1050
            ScaleWidth      =   3000
            TabIndex        =   5
            Top             =   480
            Width           =   3000
         End
      End
      Begin VB.PictureBox pcLogo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1680
         Left            =   -120
         Picture         =   "frmMenu.frx":8019D
         ScaleHeight     =   1680
         ScaleWidth      =   5520
         TabIndex        =   34
         Top             =   6240
         Width           =   5520
      End
      Begin VB.Timer tmr 
         Interval        =   1000
         Left            =   480
         Top             =   1080
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4500
         Left            =   -360
         Picture         =   "frmMenu.frx":88BBB
         ScaleHeight     =   4500
         ScaleWidth      =   4500
         TabIndex        =   21
         Top             =   -480
         Width           =   4500
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2010
            Picture         =   "frmMenu.frx":8D874
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   29
            Top             =   2760
            Width           =   240
         End
         Begin VB.Label lblDIa 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1920
            TabIndex        =   33
            Top             =   3120
            Width           =   465
         End
         Begin VB.Label lblNumDia 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   825
            Left            =   1950
            TabIndex        =   32
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label lblMes 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1905
            TabIndex        =   31
            Top             =   3840
            Width           =   465
         End
         Begin VB.Label lblAno 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2007"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1800
            TabIndex        =   30
            Top             =   4440
            Width           =   675
         End
         Begin VB.Label lblTurno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   540
            Left            =   2160
            TabIndex        =   28
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   13.5
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   600
            TabIndex        =   27
            Top             =   5040
            Width           =   150
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Turno:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1440
            TabIndex        =   26
            Top             =   2400
            Width           =   585
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   2160
            TabIndex        =   25
            Top             =   1920
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caixa:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   24
            Top             =   2040
            Width           =   570
         End
         Begin VB.Label lblUsu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   23
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuário:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Top             =   1320
            Width           =   750
         End
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
  



Private Sub cmndbtnAbertura_Click()
    frmAberturaCaixa.Show 1
End Sub

Private Sub cmndbtnCaixa_Click()
    frmSaidas.Show 0
    frmSaidas.SetFocus
End Sub

Private Sub cmndbtnCancelaCupom_Click()
    cancela_Cupom
End Sub

Private Sub cmndbtnCompras_Click()
    Dim strSQL As String
    If strVisrelest_usu = "S" Then
        frmEdiVendas.Show 0
        frmEdiVendas.SetFocus
    Else
        
        MsgBox "Você não tem permissão para este acesso!", vbCritical
    End If
End Sub

Private Sub cmndbtnECF_Click()
    'frmImpressoraECF.Show 1
    frmECF.Show 0, mdiFrm
End Sub

Private Sub cmndbtnFechamento_Click()
       If strFechacai_usu = "S" Then
         frmMovDiarioCx.Show 1
        'frmFechaCaixa.Show 1
    Else
        MsgBox "Você não tem permissão para fechar o caixa!", vbExclamation
    End If
End Sub

Private Sub cmndbtnLeituraX_Click()
    ecf_LeituraX
End Sub

Private Sub cmndbtnMapa_Click()
    'If bUsaTEF = False Then
    '    MsgBox "Não há tef!", vbExclamation
   '
    'End If
    P_AdmTEF
End Sub

Private Sub cmndbtnMemoriaFiscal_Click()
        Dim sDataIni As String
    Dim sDataFin As String
    Dim s As String
    s = F_DataInicialFinal()
    
        
    ecf_MemoriaFiscal Left(s, 10), Right(s, 10)
End Sub

Private Sub cmndbtnNovaVenda_Click()
        If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim strSQL As String
    
    
    If F_BuscaUmCampo("status_tur", "tab_turncai", "cod_Cai='" & strCod_Cai & "' and cod_tur='" & strCod_Tur & "' and data_tur=" & F_GravaData(Date)) = "F" Then
        'MsgBox "Este turno já se encontra Fechado!, para fazer vendas você deve abrir o turno", vbCritical
        'Exit Sub
    End If
    
    
    strSQL = F_GeraNovaVenda("00000", strCod_Usu)
    
    
    
    frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.tag = strSQL
    On Error Resume Next
        frmVenda.Show 0, mdiFrm

End Sub

Private Sub cmndbtnReducaoZ_Click()
    Dim strSQL As String
    Dim i As Integer
    strSQL = "A Redução Z é a leitura de Fechamento que deve ser feita no final do Dia." & Chr(13)
    strSQL = strSQL & "Uma vez feita a redução Z, até às 23:59:00hs "
    strSQL = strSQL & "a impressora estará bloqueada." & Chr(13)
    strSQL = strSQL & "Deseja realmente emitir a redução Z?"
    If MsgBox(strSQL, vbCritical + vbYesNo) = vbYes Then
         ecf_ReducaoZ
    End If
End Sub

Private Sub cmndbtnSair_Click()
    P_Fim
End Sub

Private Sub cmndbtnUsu_Click()
    frmOptUsu.Show 0
End Sub

Private Sub Command1_Click()
    Dim Buffer As String
    Buffer = "777-xxx"
    CriaArquivo Buffer, DiretorioTEFreq
End Sub

Private Sub Form_Activate()
    lblCaixa.caption = strCod_Cai
    lblTurno.caption = strCod_Tur
    mdiFrm.lblUsuario.caption = strNome_Usu
    lblData.caption = Date
    lblNomeEmp.caption = strNome_Emp
    'P_CentraObejct pcControle, Me
    mdiFrm.lblDIa.caption = StrConv(WeekdayName(Weekday(Date)), vbProperCase)
    mdiFrm.lblNumDia.caption = Day(Date)
    mdiFrm.lblMes.caption = StrConv(MonthName(Month(Date)), vbProperCase)
    mdiFrm.lblAno.caption = Year(Date)
    
    lblDIa.caption = StrConv(WeekdayName(Weekday(Date)), vbProperCase)
    lblNumDia.caption = Day(Date)
    lblMes.caption = StrConv(MonthName(Month(Date)), vbProperCase)
    lblAno.caption = Year(Date)
    
End Sub
Sub P_foco(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).BorderStyle = 0
    '    pcBotao(iCount).BackColor = vbWhite
    Next
    pcBotao(Indice).BorderStyle = 1
    'pcBotao(iCount).BackColor = vbWhite
End Sub
Sub P_focoMouse(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).BorderStyle = 0
        
    Next
    pcBotao(Indice).BorderStyle = 1
'    pcBotao(Indice).SetFocus
End Sub

Private Sub Form_Resize()
    'P_CentraObejct pcControle, Me
End Sub

'Private Sub lblBotao_Click(Index As Integer)
'    pcBotao_Click Index
'End Sub

Private Sub pcBotao_Click(Index As Integer)
    If Index = 0 Then
        cmndbtnAbertura_Click
    ElseIf Index = 1 Then
        cmndbtnFechamento_Click
    ElseIf Index = 2 Then
        cmndbtnCaixa_Click
    ElseIf Index = 3 Then
        cmndbtnECF_Click
    ElseIf Index = 4 Then
        cmndbtnCompras_Click
    ElseIf Index = 5 Then
        cmndbtnNovaVenda_Click
    ElseIf Index = 6 Then
        cmndbtnUsu_Click
    ElseIf Index = 7 Then
        cmndbtnCancelaCupom_Click
    ElseIf Index = 8 Then
        cmndbtnLeituraX_Click
    ElseIf Index = 9 Then
        cmndbtnMemoriaFiscal_Click
    ElseIf Index = 10 Then
        cmndbtnReducaoZ_Click
    ElseIf Index = 11 Then
        cmndbtnMapa_Click
    ElseIf Index = 12 Then
        cmndbtnEfetVen_Click
    ElseIf Index = 13 Then
        frmRecebConta.Show 1
    ElseIf Index = 14 Then
        cmndbtnSair_Click
    End If
End Sub

Private Sub pcBotao_GotFocus(Index As Integer)
    
     P_foco Index
    If Index = 0 Then
        lblTool.caption = "Início das atividades..."
    ElseIf Index = 1 Then
        lblTool.caption = "Término das atividades..."
    ElseIf Index = 2 Then
        lblTool.caption = "Abre lançamentos de caixa"
    ElseIf Index = 3 Then
        lblTool.caption = "Abre controle da Impressora Fiscal"
    ElseIf Index = 4 Then
        lblTool.caption = "Abre lista das Vendas ainda não concluídas..."
    ElseIf Index = 5 Then
        lblTool.caption = "Inicia uma nova venda"
    ElseIf Index = 6 Then
        lblTool.caption = "Opções de Usuário/Mudança de Senha"
    ElseIf Index = 7 Then
        lblTool.caption = "Cancela o Último Cupom Fiscal/Última Venda Impressos"
    ElseIf Index = 8 Then
        lblTool.caption = "Emite uma Leitura X (Início do Dia)"
    ElseIf Index = 9 Then
        lblTool.caption = "Emite Relatório de Memória Fiscal a partir de uma data"
    ElseIf Index = 10 Then
        lblTool.caption = "Emite uma Redução Z (Final do Dia - Encerra Todas as Atividades)"
    ElseIf Index = 11 Then
        lblTool.caption = "Configurações, Relatórios TEF, Cancelamento de Transação"
    ElseIf Index = 12 Then
        lblTool.caption = "Lançamento de dados pendentes nas Vendas a Finalizar"
    ElseIf Index = 13 Then
        lblTool.caption = "Lançamento de Recebimento Avulso Multi-fácil"
    ElseIf Index = 14 Then
        lblTool.caption = "Sair do Sistema"
    
    End If
End Sub

Private Sub pcBotao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If Index = 0 Then
            pcBotao(pcBotao.count - 1).SetFocus
        Else
            pcBotao(Index - 1).SetFocus
        End If
        
    ElseIf KeyCode = vbKeyDown Then
        If Index = pcBotao.count - 1 Then
            pcBotao(0).SetFocus
        Else
            pcBotao(Index + 1).SetFocus
        End If
    ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcBotao_Click Index
    ElseIf KeyCode = vbKeyLeft Then
        If Index < 5 Then
            If Index = 0 Then
                pcBotao(14).SetFocus
            Else
                pcBotao(Index + 9).SetFocus
            End If
        Else
            pcBotao(Index - 5).SetFocus
        End If
    ElseIf KeyCode = vbKeyRight Then
        If Index > 9 Then
            If Index = 14 Then
                pcBotao(0).SetFocus
            Else
                pcBotao(Index - 9).SetFocus
            End If
        Else
            
            pcBotao(Index + 5).SetFocus
            
        End If
    End If
End Sub

Private Sub pcBotao_LostFocus(Index As Integer)
    pcBotao(Index).BorderStyle = 0
    lblTool.caption = "..."
End Sub

    
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmndbtnAbertura_Click
    ElseIf KeyCode = vbKeyF3 Then
        cmndbtnFechamento_Click
    ElseIf KeyCode = vbKeyF4 Then
        cmndbtnCaixa_Click
    ElseIf KeyCode = vbKeyF5 Then
        cmndbtnECF_Click
    ElseIf KeyCode = vbKeyF6 Then
        cmndbtnCompras_Click
    ElseIf KeyCode = vbKeyF7 Then
        cmndbtnNovaVenda_Click
    ElseIf KeyCode = vbKeyF8 Then
        cmndbtnUsu_Click
    ElseIf KeyCode = vbKeyF10 Then
        frmRecebConta.Show 1
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnSair_Click
    End If
End Sub

Private Sub Form_Load()
    strIP_Usu = Winsock1.LocalIP
    strCod_Cai = F_BuscaUmCampo("cod_cai", "tab_caixas", "ip_cai='" & strIP_Usu & "'")
    strImprFiscal = F_BuscaUmCampo("impressora_cai", "tab_caixas", "ip_cai='" & strIP_Usu & "'")
    If strImprFiscal = "" Then
        strImprFiscal = "B"
    End If
    If strCod_Cai = "" Then
        strCod_Cai = "01"
    End If
    strCod_Tur = F_BuscaUmCampo("cod_tur", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
   
    strSQL = F_UsuLogadoCaixa(strCod_Cai)
    If strSQL = "" Then
'        MsgBox "Não existe usuário logado neste caixa!", vbExclamation
        
    ElseIf strSQL <> strNome_Usu Then
        strSQL = "update tab_turncai set login_usu = '" & strCod_Usu & "' "
        strSQL = strSQL & " where cod_cai='" & strCod_Cai & "' "
        strSQL = strSQL & " and cod_Tur='" & strCod_Tur & "' "
        strSQL = strSQL & " and data_tur=" & F_GravaData(Date)
        banco.Execute strSQL
    End If
    Dim i As Integer
    'i = ecf_VerificaImpressoraLigada()
    'If i <> 1 Then
     
     
     '   MsgBox "A impressora fiscal não foi localizada!", vbExclamation
     
     
    'End If
    ecf_verificaImpressora
    BuscaPic
End Sub
Sub BuscaPic()
    On Error Resume Next
    pcLogo.Picture = LoadPicture(App.path & "\pic\logo.jpg")
End Sub
Private Sub cmndbtnEfetVen_Click()
    frmEfetivaVen.Show 1
End Sub

'Private Sub pcIco_Click(Index As Integer)
'    pcBotao_Click Index
'End Sub
'
'Private Sub pcIco_GotFocus(Index As Integer)
'    trocatabkey keycode
'End Sub

Private Sub tmr_Timer()
    lblData.caption = Now
End Sub
