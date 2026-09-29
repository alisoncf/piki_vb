VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLanParcelasVen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lançamento de Cheque:"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      Caption         =   "Dados da Venda:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   10335
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   195
         Left            =   4080
         TabIndex        =   35
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00000000000000"
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
         Left            =   4560
         TabIndex        =   34
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblValorVen 
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
         Left            =   2760
         TabIndex        =   31
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2280
         TabIndex        =   30
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblNomeCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda a Consumidor"
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
         Left            =   7680
         TabIndex        =   29
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   6840
         TabIndex        =   28
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblDataVen 
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
         Left            =   6600
         TabIndex        =   27
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   6120
         TabIndex        =   26
         Top             =   0
         Width           =   390
      End
      Begin VB.Label lblNumVen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
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
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C6DADD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   10335
      TabIndex        =   0
      Top             =   1440
      Width           =   10335
      Begin VB.TextBox txtbxValor_Che 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8400
         MaxLength       =   14
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtbxVencimento 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox pcCheque 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   7065
         TabIndex        =   54
         Top             =   240
         Width           =   7095
         Begin VB.TextBox txtbxContaCorrente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5760
            MaxLength       =   12
            TabIndex        =   13
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtbxNum_Che 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   9
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtbxNome_Che 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   4
            Top             =   240
            Width           =   4815
         End
         Begin VB.TextBox txtbxAgencia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4680
            MaxLength       =   8
            TabIndex        =   11
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtbxCPFChe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   2
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtbxBanco 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   6
            Top             =   840
            Width           =   2895
         End
         Begin VB.CommandButton cmndbtnBuscaBanco 
            Height          =   285
            Left            =   3000
            Picture         =   "frmLanParcelaVen.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C&onta C.:"
            Height          =   195
            Left            =   5760
            TabIndex        =   12
            Top             =   600
            Width           =   660
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Titular:"
            Height          =   195
            Left            =   2040
            TabIndex        =   3
            Top             =   0
            Width           =   480
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Cheque Nº:"
            Height          =   195
            Left            =   3360
            TabIndex        =   8
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Agência:"
            Height          =   195
            Left            =   4680
            TabIndex        =   10
            Top             =   600
            Width           =   630
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&CPF:"
            Height          =   195
            Left            =   120
            TabIndex        =   1
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.CheckBox chkCh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cheque:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   0
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   2535
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Dê <Enter> ou duplo click para editar a linha do cheque"
         Top             =   1560
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4471
         _Version        =   393216
         AllowUserResizing=   3
         Appearance      =   0
      End
      Begin VB.CommandButton cmndbtnEdi 
         Height          =   375
         Left            =   8880
         Picture         =   "frmLanParcelaVen.frx":00F9
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Editar o cheque ou  parcela que você selecionou na grade abaixo..."
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnAdd 
         Height          =   375
         Left            =   8400
         Picture         =   "frmLanParcelaVen.frx":0AFB
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Adiciona o Cheque que você digitou.."
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmbxSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   8760
         Picture         =   "frmLanParcelaVen.frx":1085
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnConfirmar 
         Caption         =   "Con&firmar"
         Height          =   615
         Left            =   7200
         Picture         =   "frmLanParcelaVen.frx":11F0
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Va&lor:"
         Height          =   195
         Left            =   7320
         TabIndex        =   16
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   7320
         TabIndex        =   14
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Lançado:"
         Height          =   195
         Left            =   480
         TabIndex        =   33
         Top             =   4560
         Width           =   1080
      End
      Begin VB.Label lblTotal 
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
         Left            =   1800
         TabIndex        =   32
         Top             =   4560
         Width           =   390
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   36
      Top             =   600
      Width           =   10335
      Begin VB.CommandButton cmndbtnDistribuir 
         Caption         =   "Distribuir"
         Height          =   255
         Left            =   7440
         TabIndex        =   52
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDescForma 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   42
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtbxParcelas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   40
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbxInterv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   39
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtbxDataBase 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   38
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxValorDoc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   37
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxFormaPag 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnFormasPag 
         Height          =   285
         Left            =   2280
         Picture         =   "frmLanParcelaVen.frx":133A
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   2760
         TabIndex        =   50
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pagamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label lblTipoCai 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIN"
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
         Left            =   2280
         TabIndex        =   48
         Top             =   120
         Width           =   345
      End
      Begin VB.Label lblParcelas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas:"
         Height          =   195
         Left            =   3240
         TabIndex        =   47
         Top             =   120
         Width           =   660
      End
      Begin VB.Label lblInternv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intervalo."
         Height          =   195
         Left            =   4080
         TabIndex        =   46
         Top             =   120
         Width           =   660
      End
      Begin VB.Label lblInterMes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mês:"
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
         Height          =   285
         Left            =   4560
         TabIndex        =   45
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblDataBase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base:"
         Height          =   195
         Left            =   6120
         TabIndex        =   44
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Prazo:"
         Height          =   195
         Left            =   5040
         TabIndex        =   43
         Top             =   120
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmLanParcelasVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim c As cheque
Private Sub cmbxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub chkCh_Click()
    If chkCh.Value = 1 Then
        pcCheque.Visible = True
    Else
        pcCheque.Visible = False
    End If
End Sub

Private Sub cmbxSair_Click()
    bActive = False
    Me.Hide
    
    
    
End Sub

Private Sub cmndbtnAdd_Click()
    
    If chkCh.Value = 1 Then
        If txtbxNum_Che.Text = "" Then
            MsgBox "Informe o numero do cheque!", vbExclamation
            Exit Sub
        End If
    End If
    
    If cmndbtnAdd.tag = "" Then
    Else
        
        msflxgrd.row = cmndbtnAdd.tag
        msflxgrd.col = 1:
        If chkCh.Value = 1 Then
            msflxgrd.Text = "Cheque"
        Else
            msflxgrd.Text = "Duplicata"
        End If
        msflxgrd.col = 2:
        
        msflxgrd.Text = txtbxVencimento.Text
        msflxgrd.col = 3:
        msflxgrd.Text = txtbxValor_Che.Text
        msflxgrd.col = 4:
        msflxgrd.Text = txtbxNum_Che.Text
        msflxgrd.col = 5:
        msflxgrd.Text = txtbxAgencia.Text
        
        msflxgrd.col = 6:
        msflxgrd.Text = txtbxContaCorrente.Text
        msflxgrd.col = 7:
        msflxgrd.Text = txtbxBanco.tag
        msflxgrd.col = 8:
        msflxgrd.Text = txtbxCPFChe.Text
        msflxgrd.col = 9:
        msflxgrd.Text = txtbxNome_Che.Text
        cmndbtnAdd.tag = ""
        c.age_che = txtbxAgencia.Text
        c.num_che = txtbxNum_Che.Text
        c.cod_bco = txtbxBanco.tag
        c.cco_che = txtbxContaCorrente.Text
        soma
        msflxgrd.SetFocus
    End If
End Sub

Sub soma()
    Dim cTotal As Currency
    Dim iCount As Integer
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 3
            cTotal = cTotal + CCur(.Text)
            
        Next
    End With
    lblTotal.caption = FormatVal(cTotal)
End Sub
Private Sub cmndbtnConfirmar_Click1()
    Dim sSEQ_DOC As String
    Dim sSeq_che As String
    Dim strSQL As String
    Dim iCount As Integer
    Dim sNum As String
    Dim sAge As String
    Dim sCCO As String
    Dim sBco As String
    Dim sValor As String
    Dim sVen As String
    Dim sCPF As String
    Dim sTit As String
    If msflxgrd.rows = 1 Then
        MsgBox "Não há cheques lançados!", vbExclamation
        Exit Sub
    End If
    If CCur(lblValorVen.caption) > CCur(lblTotal.caption) + 1 Or CCur(lblValorVen.caption) < CCur(lblTotal.caption) - 1 Then
        MsgBox "O Valor lançado não está correto!", vbExclamation
        Exit Sub
    End If
    
    sSEQ_DOC = F_GravaDoc("VENDA " & Me.tag & lblNomeCli.caption, Me.tag, lblValorVen.caption, _
        Date, lblNomeCli.tag, "R", Me.tag, msflxgrd.rows, "F", "C", txtbxClass.tag)
    If sSEQ_DOC <> "" Then
        For iCount = 1 To msflxgrd.rows - 1
            msflxgrd.col = 0: sNum = msflxgrd.Text
            msflxgrd.col = 1: sAge = msflxgrd.Text
            msflxgrd.col = 2: sCCO = msflxgrd.Text
            msflxgrd.col = 3: sBco = msflxgrd.Text
            msflxgrd.col = 5: sVen = msflxgrd.Text
            msflxgrd.col = 6: sValor = msflxgrd.Text
            msflxgrd.col = 7: sCPF = msflxgrd.Text
            msflxgrd.col = 8: sTit = msflxgrd.Text
            sSeq_che = F_GravaCh(sSEQ_DOC, sNum, sBco, sAge, sCCO, sTit, sValor, sVen, lblNomeCli.tag, "R", _
                Me.tag, "", CStr(iCount), "", sCPF)
        Next
        strSQL = "update tab_vendas set situacao_ven=4 "
        strSQL = strSQL & " where num_ven='" & Me.tag & "' "
        Conn.Execute strSQL
    End If
    APIInserirDoc sSEQ_DOC, True
    MsgBox "Cheques Gravados!", vbInformation
   Unload Me
End Sub

Private Sub cmndbtnBuscaBanco_Click()
    Dim b As Bancos.TBanco
    b = F_BuscaBanco(txtbxBanco.Text)
    If b.cod_bco <> "" Then
        txtbxBanco.Text = b.nome_bco
        txtbxBanco.tag = b.cod_bco
    End If
    
End Sub


Private Sub cmndbtnConfirmar_Click()
    Dim sSEQ_DOC As String
Dim strSQL As String
    Dim sTip As String, sVen As String, sAge As String, sConta As String, sBanco As String
    Dim sCPF As String, sNome As String, sVal As String, sChe As String
    If Not F_confMsflxCH(msflxgrd, 2, 4, 5, 6, 7, 8, 9, 1) Then
        Exit Sub
    End If
    
    If CCur(lblValorVen.caption) <> CCur(lblTotal) Then
        MsgBox "O valor distribuídos foi de " & lblTotal.caption & " enquanto o total deve ser " & lblValorVen.caption & "!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirma Fechamento da Venda?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    sSEQ_DOC = F_GravaDoc("VENDA " & Me.tag & " " & lblNomeCli.caption, Me.tag, lblValorVen.caption, _
             Date, lblNomeCli.tag, "R", Me.tag, txtbxParcelas.Text, "V", "C", "000")
                
            If sSEQ_DOC <> "" Then
                With msflxgrd
                    
                    For iCount = 1 To .rows - 1
                        .row = iCount
                        .col = 1: sTip = .Text
                        .col = 2: sVen = .Text
                        .col = 3: sVal = .Text
                        If Left(sTip, 1) = "D" Then
                            strSQL = F_GravaItDoc(sSEQ_DOC, sVal, sVen, sVen, CStr(iCount), a_receber_avulso)
                            If strSQL = "" Then
                                MsgBox " Erro na gravação da parcela " & iCount
                            End If
                        Else
                            .col = 4: sChe = .Text
                            .col = 5: sAge = .Text
                            .col = 6: sConta = .Text
                            .col = 7: sBanco = .Text
                            .col = 8: sCPF = .Text
                            .col = 9: sNome = .Text
                            strSQL = F_GravaCh(sSEQ_DOC, sChe, sBanco, sAge, sConta, sNome, sVal, sVen, lblNomeCli.tag, _
                                    "R", Me.tag, "", CStr(iCount), "", sCPF)
                                    
                            If strSQL = "" Then
                                MsgBox " Erro na gravação da parcela " & iCount
                            End If
                            
                        End If
                            
                    Next
                      
                End With
                
                strSQL = F_GravaCai("C", Date, lblValorVen.caption, _
                    "VENDA " & Me.tag & " " & lblNomeCli.caption, Me.tag, lblTipoCai.caption, "", 0, 0, _
                    strCod_Cai, strCod_Tur, "000", "S")
                
                strSQL = "UPDATE TAB_VENDAS set situacao_ven='4' "
'                strSQL = strSQL & "fechavenda_usu='" & strCod_Usu & "' "
'                strSQL = strSQL & "dataefet_ven=" & F_GravaData(Date) & " "
                strSQL = strSQL & " where num_ven='" & Me.tag & "'"
                Conn.Execute strSQL
                APIInserirDoc sSEQ_DOC, True
                MsgBox "Venda Efetivada com sucesso!", vbInformation
                
                ImprVenda Me.tag
                Unload Me
            End If
End Sub

Private Sub cmndbtnDistribuir_Click()
    distribuiparcelas
End Sub


Private Sub cmndbtnEdi_Click()
    With msflxgrd
        If .row < 1 Then
            MsgBox "Escolha uma parcela para editar!", vbExclamation
            Exit Sub
            Exit Sub
            
        End If
                
        cmndbtnAdd.tag = msflxgrd.row
        .col = 1:
        If Left(.Text, 1) = "C" Then
            'chkCh.Value = 1
        Else
            'chkCh.Value = 0
        End If
        
        .col = 2:
        If .Text <> "" Then
            txtbxVencimento.Text = .Text
        End If
        .col = 3:
        If .Text <> "" Then
            txtbxValor_Che.Text = .Text
        End If
        .col = 4:
        If .Text <> "" Then
            txtbxNum_Che.Text = .Text
        End If
        .col = 5:
        If .Text <> "" Then
            txtbxAgencia.Text = .Text
        End If
        .col = 6:
        If .Text <> "" Then
            txtbxContaCorrente.Text = .Text
        End If
        .col = 7:
        If .Text <> "" Then
            txtbxBanco.tag = .Text
            txtbxBanco.Text = .Text
        End If
        
        .col = 8:
        If .Text <> "" Then
            txtbxCPFChe.Text = .Text
        End If
        .col = 9:
        If .Text <> "" Then
            txtbxNome_Che.Text = .Text
        End If
        .col = 1
        If Left(.Text, 1) = "C" Then
            If txtbxCPFChe.Text = "" Then
                txtbxCPFChe.Text = lblcpf.caption
                txtbxNome_Che.Text = lblNomeCli.caption
            End If
           On Error Resume Next
            txtbxCPFChe.SetFocus
        Else
            txtbxVencimento.SetFocus
        End If
    End With
End Sub

Private Sub cmndbtnExcl_Click()
    On Error GoTo trata_Erro
    With msflxgrd
        If .rows = 2 Then
            msflxgrd.rows = 1
        Else
            msflxgrd.RemoveItem msflxgrd.row
        End If
    End With
    Exit Sub
trata_Erro:
    
End Sub


Private Sub cmndbtnFormasPag_Click()
    Dim f As FormasPag.FormasPag
    If txtbxFormaPag.tag <> "" Then
        f = F_SelectFormasPag(txtbxFormaPag.tag)
    Else
        f = F_SelectFormasPag(txtbxFormaPag.Text)
    End If
    If f.cod_For <> "" Then
        txtbxFormaPag.Text = f.nome_for
        txtbxFormaPag.tag = f.cod_For
        txtbxDescForma.Text = f.desconto_for
        lblTipoCai.caption = f.tipo_cai
        lblTipoCai.tag = f.cod_tip
        txtbxParcelas.Text = f.parcelas_for
        If f.tipo_cai = "CHV" Or f.tipo_cai = "CHP" Then
            chkCh.Value = 1
        Else
            chkCh.Value = 0
        End If
        distribuiparcelas
    End If
        
        
        
         
        
        
        
End Sub
Sub ajustaGrid()
    With msflxgrd
        .rows = 1
        .Cols = 10
        .col = 0: .Text = "..."
        .col = 1: .Text = "Tipo"
        .col = 2: .Text = "Vencimento"
        .col = 3: .Text = "Valor"
        .col = 4: .Text = "Cheque"
        .col = 5: .Text = "Agência"
        .col = 6: .Text = "Conta"
        .col = 7: .Text = "banco"
        .col = 8: .Text = "CPF"
        .col = 9: .Text = "Titular"
        .ColWidth(0) = 300
        .ColWidth(1) = 800
        .ColWidth(2) = 1000
        .ColWidth(7) = 2000
        .ColWidth(8) = 2000
        .ColWidth(9) = 3000
        
    End With
End Sub

Sub distribuiparcelas()
    msflxgrd.rows = 1
    Dim iCount As Integer
    Dim cTotal As Currency
    Dim sTipo As String
    Dim cValor As Currency
    Dim sVlrParcela As String
    
    If lblTipoCai.tag = "01" Then
        msflxgrd.rows = 1
        Exit Sub
    End If
    
    
    cValor = CCur(lblValorVen.caption)
    
    If txtbxDataBase.Text = "" Then
        txtbxDataBase.Text = Format(Date, "dd/mm/yyyy")
    End If
    If (lblTipoCai.caption = "CHP" Or lblTipoCai.caption = "CHV") Then
        sTipo = "Cheque"
    Else
        sTipo = "Duplicata"
    End If
    For iCount = 1 To Int(txtbxParcelas.Text)
        strSQL = iCount & Chr(9) & sTipo & Chr(9)
        strSQL = strSQL & DateAdd("M", iCount, txtbxDataBase.Text) & Chr(9)
        
        
        If iCount = Int(txtbxParcelas) Then
            sVlrParcela = FormatVal(CCur(cValor) - cTotal)
        Else
            sVlrParcela = FormatVal(CCur(cValor) / Int(txtbxParcelas.Text))
            cTotal = cTotal + CCur(sVlrParcela)
        End If
        strSQL = strSQL & sVlrParcela & Chr(9)
        msflxgrd.AddItem strSQL
            
    Next
    
    soma
    
End Sub


    Private Sub cmndbtnPlanoContas_Click()
    Dim p As PlanoContas
    
    p = F_BuscaPlanoContas(Left(lblTipo.caption, 1), txtbxClass.Text)
    If p.cod_pla <> "" Then
        txtbxClass.tag = p.cod_pla
        txtbxClass.Text = p.nome_pla
        txtbxParc.SetFocus
    End If
End Sub

Private Sub txtbxClass_GotFocus()
    
End Sub

Private Sub txtbxClass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnPlanoContas_Click
    End If
End Sub


Private Sub Form_Activate()
    If bActive Then Exit Sub
    
    carrega
    bActive = True
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sCPF As String
    strSQL = F_BuscaUmCampo("tipocli_ven", "tab_vendas", "num_ven='" & Me.tag & "'")
    
    If strSQL = "U" Then
        
        sTab = "tab_funcionarios"
        sCod = "cod_fun"
        sNome = "nome_fun"
        sCPF = "cpf_fun"
    Else
        sTab = "tab_clientes"
        sCod = "cod_cli"
        sNome = "nome_cli"
        sCPF = "cpf_cli"
        
    End If
    
    strSQL = "select tc.limite_cli,esp_cai,tc." & sCPF & " as cpf,tt.nome_tip,sum(tvf.valor_ven) as valor,tv.*,tf.*,"
    strSQL = strSQL & " tfu.guerra_fun from tab_vendas as tv, "
    strSQL = strSQL & sTab & " as tc, tab_formaspag as tf,tab_funcionarios as tfu,tab_tipopag as tt,tab_venfec as tvf "
    strSQL = strSQL & " where tf.cod_for=tv.cod_for and tc." & sCod & "=tv.cod_cli "
    strSQL = strSQL & " and tt.cod_tip=tf.cod_tip "
    strSQL = strSQL & " and tv.num_ven=tvf.num_ven "
    strSQL = strSQL & " and tv.num_ven='" & Me.tag & "' "
    strSQL = strSQL & " and tv.abrevenda_usu=tfu.cod_fun "
    strSQL = strSQL & " and esp_cai in('DUP','CHP') "
    strSQL = strSQL & " group by tv.num_ven "
    abreObjetoAdo rPesq, strSQL
    
    
    rPesq.Close
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblcpf.caption = !cpf
            lblNumVen.caption = !num_ven
            lblDataVen.caption = !data_ven
            lblNomeCli.caption = !nomecli_ven
            lblValorVen.caption = !valor
            lblcpf.caption = !cpf
            lblNomeCli.tag = !cod_cli
            txtbxFormaPag.tag = !cod_For
            txtbxFormaPag.Text = !nome_for
            lblTipoCai.caption = !esp_cai
            txtbxParcelas.Text = !parcelas_for
            txtbxDataBase.Text = !dataefet_ven
            distribuiparcelas
            If lblTipoCai.caption = "CHP" Or lblTipoCai.caption = "CHV" Then
                chkCh.Value = 1
            End If
            
            
            chkCh.Enabled = False
        End If
        .Close
    End With
End Sub
Private Sub Form_Load()
    ajustaGrid
End Sub

Private Sub msflxgrd_DblClick()
    cmndbtnEdi_Click
End Sub

Private Sub msflxGrd_GotFocus()
    msflxgrd.GridLines = flexGridInset
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    msflxgrd_DblClick
End Sub

Private Sub msflxGrd_LostFocus()
    msflxgrd.GridLines = flexGridFlat
End Sub

Private Sub Picture1_GotFocus()
    'txtbxCPF.SetFocus
End Sub



Private Sub txtbxAgencia_GotFocus()
    P_SelecTexto txtbxAgencia
End Sub

Private Sub txtbxAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxAgencia_LostFocus()
    P_AjustaTexto txtbxAgencia
End Sub


Private Sub txtbxCNPJ_LostFocus()
    
End Sub

Private Sub txtbxBanco_GotFocus()
    P_SelecTexto txtbxBanco
End Sub

Private Sub txtbxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaBanco
    End If
End Sub
Sub buscaBanco()
    Dim b As TBanco
    b = F_SelectBanco(txtbxBanco.Text)
    If b.cod_bco <> "" Then
        txtbxBanco.Text = b.nome_bco
        txtbxBanco.tag = b.cod_bco
        txtbxNum_Che.SetFocus
    End If
    
End Sub
Private Sub txtbxContaCorrente_GotFocus()
    P_SelecTexto txtbxContaCorrente
End Sub

Private Sub txtbxContaCorrente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxContaCorrente_LostFocus()
    P_AjustaTexto txtbxContaCorrente
End Sub


Sub buscaCH()
    Dim c As cheque
    'On Error Resume Next
    txtbxCPFChe_LostFocus
    If txtbxCPFChe.Text = "" Then Exit Sub
    
    c = F_HistCheCPF(txtbxCPFChe.Text)
    
    If c.cod_bco <> "" Then
        txtbxNome_Che.Text = c.nome_che
        txtbxAgencia.Text = c.age_che
        txtbxContaCorrente.Text = c.cco_che
        
        txtbxBanco.Text = F_BuscaUmCampo("nome_bco", "tab_Bancos", "cod_bco='" & c.cod_bco & "'")
        txtbxBanco.tag = c.cod_bco
        txtbxNum_Che.SetFocus
    Else
        txtbxNome_Che.SetFocus
    End If
    
    
    
End Sub

Private Sub txtbxCPFChe_GotFocus()
    P_SelecTexto txtbxCPFChe
End Sub

Private Sub txtbxCPFChe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCH
    End If
End Sub

Private Sub txtbxCPFChe_LostFocus()
   On Error Resume Next
    txtbxCPFChe.Text = F_SoNumero(txtbxCPFChe.Text)
    If Len(txtbxCPFChe.Text) = 0 Then
       
    ElseIf Len(txtbxCPFChe.Text) >= 14 Then
        'lblCPF.caption = "&CNPJ:"
        P_FormataCGC txtbxCPFChe
        If Not F_CGC(txtbxCPFChe.Text) Then
            MsgBox "CNPJ Inválido!", vbExclamation
            txtbxCPF.SetFocus
            Exit Sub
        End If
    Else
        'lblCPF.caption = "&CPF:"
        P_FormataCPF txtbxCPFChe
        If Not F_CPF(txtbxCPFChe.Text) Then
            MsgBox "CPF Inválido!", vbExclamation
            txtbxCPF.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtbxFormaPag_Change()
    txtbxFormaPag.tag = ""
End Sub

Private Sub txtbxNome_Che_GotFocus()
    P_SelecTexto txtbxNome_Che
End Sub

Private Sub txtbxNome_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Che_LostFocus()
    P_AjustaTexto txtbxNome_Che
End Sub

Private Sub txtbxNum_Che_GotFocus()
    P_SelecTexto txtbxNum_Che
End Sub

Private Sub txtbxNum_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNum_Che_LostFocus()
    P_AjustaTexto txtbxNum_Che
End Sub

Private Sub txtbxValor_Che_GotFocus()
    P_SelecTexto txtbxValor_Che
End Sub

Private Sub txtbxValor_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Che_LostFocus()
    P_ValorLostFocus txtbxValor_Che
End Sub

Private Sub txtbxVencimento_GotFocus()
    P_SelecTexto txtbxVencimento
End Sub

Private Sub txtbxVencimento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxVencimento_LostFocus()
    P_AjustaData txtbxVencimento
End Sub

