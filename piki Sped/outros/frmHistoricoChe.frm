VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHistoricoChe 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Histórico de Cheques:"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   3675
      TabIndex        =   6
      Top             =   2640
      Width           =   3735
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
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   390
      End
      Begin VB.Label lblDev 
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
         Left            =   2640
         TabIndex        =   13
         Top             =   0
         Width           =   390
      End
      Begin VB.Label lblBaixado 
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
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   390
      End
      Begin VB.Label lblAbert 
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
         Left            =   840
         TabIndex        =   11
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Devolvido:"
         Height          =   195
         Left            =   1800
         TabIndex        =   9
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baixado:"
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Em Aberto:"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Buscar"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtbxCPF 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   9120
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4048
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   345
   End
End
Attribute VB_Name = "frmHistoricoChe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnBuscar_Click()
    If txtbxCPF.Text = "" Then
        MsgBox "Informe o CPF/CNPJ", vbExclamation
        txtbxCPF.SetFocus
        Exit Sub
    End If
    
    busca
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
    MudaAparencia Me
    
End Sub

Private Sub Form_Load()
    Me.Top = 600
    'Me.Left = mdiFrm.pcMenu.Width
    With msflxgrd
        .rows = 1
        .FixedCols = 0
        .Cols = 8
        .Col = 0: .Text = "Nº"
        .Col = 1: .Text = "Banco"
        .Col = 2: .Text = "Agência"
        .Col = 3: .Text = "Conta C/"
        .Col = 4: .Text = "Valor"
        .Col = 5: .Text = "Vencim."
        .Col = 6: .Text = "Titular"
        .Col = 7: .Text = "Status"
        .ColWidth(1) = 1500
        .ColWidth(6) = 2500
    End With
End Sub
Sub busca()
    Dim stSQL As String
    Dim rPesq As Recordset
    Dim cTotalAb As Currency
    Dim cTotalDev As Currency
    Dim cTotal As Currency
    Dim cTotalBai As Currency
    strSQL = "select tc.*,tb.nome_bco from tab_cheques1 as tc, tab_bancos as tb where cpf_che='" & txtbxCPF.Text & "' "
    strSQL = strSQL & " and tc.cod_bco=tb.cod_bco "
    strSQL = strSQL & " order by seq_che desc"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
    msflxgrd.Visible = False
    msflxgrd.rows = 1
        Do While Not .EOF
            strSQL = !num_che & Chr(9)
            strSQL = strSQL & !nome_bco & Chr(9)
            strSQL = strSQL & !age_che & Chr(9)
            strSQL = strSQL & !cco_che & Chr(9)
            strSQL = strSQL & FormatVal(!Valor_che) & Chr(9)
            strSQL = strSQL & !Venc_che & Chr(9)
            
            strSQL = strSQL & !nome_che
            
            
            msflxgrd.AddItem strSQL
            msflxgrd.Col = 7
            msflxgrd.Row = msflxgrd.rows - 1
            If !status_Che = 1 Then
                msflxgrd.Text = "Em Aberto"
                ColoreLinhaGrid msflxgrd, msflxgrd.Row, vbYellow
                cTotalAb = cTotalAb + !Valor_che
            ElseIf !status_Che = 3 Then
                msflxgrd.Text = "Devolvido"
                ColoreLinhaGrid msflxgrd, msflxgrd.Row, vbRed
                cTotalDev = cTotalDev + !Valor_che
            Else
                msflxgrd.Text = "Baixado"
                ColoreLinhaGrid msflxgrd, msflxgrd.Row, vbGreen
                cTotalBai = cTotalBai + !Valor_che
            End If
            cTotal = cTotal + !Valor_che
            
            
            .MoveNext
        Loop
        .Close
    End With
    msflxgrd.Visible = True
    lblAbert.caption = FormatVal(cTotalAb)
    lblBaixado.caption = FormatVal(cTotalBai)
    lblDev.caption = FormatVal(cTotalDev)
    lblTOtal.caption = FormatVal(cTotal)
End Sub

Private Sub txtbxCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCPF_LostFocus()
    txtbxCPF.Text = F_SoNumero(txtbxCPF.Text)
    If Len(txtbxCPF.Text) >= 14 Then
        P_FormataCGC txtbxCPF
    Else
        P_FormataCPF txtbxCPF
    End If
End Sub
