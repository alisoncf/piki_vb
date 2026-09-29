VERSION 5.00
Begin VB.Form frmDevChe 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Devolução de Cheque"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbxData 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtbxMotivo 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   4575
   End
   Begin VB.CommandButton cmbxSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   3360
      Picture         =   "frmDevChe.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnConfirmar 
      Caption         =   "Con&firmar"
      Height          =   615
      Left            =   1800
      Picture         =   "frmDevChe.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6DADD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2655
      ScaleWidth      =   4815
      TabIndex        =   7
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtbxValor_Che 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         MaxLength       =   14
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtbxContaCorrente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cmbxBanco 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtbxVencimento 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtbxNum_Che 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtbxNome_Che 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtbxAgencia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtbxCPF 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1920
         Width           =   3495
      End
      Begin VB.Label lblCodCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         Height          =   195
         Left            =   720
         TabIndex        =   25
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Va&lor:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   600
         TabIndex        =   23
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&onta C.:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Banco:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2520
         TabIndex        =   20
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Titular:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   525
         TabIndex        =   19
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cheque Nº:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Agência:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2880
         TabIndex        =   17
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblCPF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CPF:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   660
         TabIndex        =   16
         Top             =   1920
         Width           =   345
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Situação do Cheque:"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Motivo da Devolução:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
End
Attribute VB_Name = "frmDevChe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select tc.*,td.cod_ter from tab_cheques1 as tc,tab_doc1 as td where seq_che='" & Me.Tag & "' "
    strSQL = strSQL & " and td.seq_doc=tc.seq_doc "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            txtbxAgencia.Text = !age_che
            txtbxContaCorrente.Text = !cco_che
            txtbxCPF.Text = !cpf_che
            txtbxData.Text = Date
            txtbxMotivo.Text = ""
            txtbxNome_Che.Text = !nome_che
            txtbxNum_Che.Text = !num_che
            txtbxValor_Che.Text = FormatVal(!Valor_che)
            lblCodCli.Tag = !cod_ter
            lblCodCli.caption = !cod_ter
            txtbxVencimento.Text = !Venc_che
            P_LocalizaOperacao !cod_bco, cmbxBanco
            
            
        End If
        .Close
    End With
End Sub

Private Sub cmbxSair_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub cmndbtnConfirmar_Click()
    If F_Vazio(txtbxMotivo.Text) Then
        MsgBox "Informe o motivo da Devolução!", vbExclamation
        txtbxMotivo.SetFocus
        Exit Sub
    End If
    
    
    If F_Vazio(txtbxData.Text) Then
        MsgBox "Informe a data!", vbExclamation
        txtbxData.SetFocus
        Exit Sub
        
    End If
    If MsgBox("Confirma devolução do cheque " & txtbxNum_Che.Text & "?", vbYesNo + vbQuestion) = vbYes Then
        grava
    End If
End Sub
Sub grava()
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim sSeq As String
    sSeq = MaxKeySQL("tab_chequedev", "seq_Dev", "seq_che='" & Me.Tag & "'", "")
    
    strSQL = "insert into tab_chequedev(seq_Che,seq_dev,data_dev,motivo_dev)values("
    strSQL = strSQL & "'" & Me.Tag & "',"
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & F_GravaData(txtbxData.Text) & ","
    strSQL = strSQL & "'" & txtbxMotivo.Text & "') "
    Banco.Execute strSQL
    strSQL = "update tab_cheques1 set status_che=3 where seq_che='" & Me.Tag & "' "
    Banco.Execute strSQL
    Me.Tag = sSeq
    
    strSQL = "update tab_clientes set status_cli=2 "
    strSQL = strSQL & " where cod_cli='" & lblCodCli.Tag & "' "
    Banco.Execute strSQL
    
    Me.Hide
    Exit Sub
trata_erro:
    P_MostraErro
    
End Sub
Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    P_PreencheBanco cmbxBanco
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub

Private Sub txtbxMotivo_GotFocus()
    P_SelecTexto txtbxMotivo
End Sub

Private Sub txtbxMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMotivo_LostFocus()
    P_AjustaTexto txtbxMotivo
End Sub
