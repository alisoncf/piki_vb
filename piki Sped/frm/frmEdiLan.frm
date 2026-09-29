VERSION 5.00
Begin VB.Form frmEdiLan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Lançamentos"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "frmEdiLan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5655
      Begin VB.ComboBox cmbxPlanoConta 
         Height          =   315
         ItemData        =   "frmEdiLan.frx":030A
         Left            =   120
         List            =   "frmEdiLan.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2400
         Width           =   5415
      End
      Begin VB.TextBox txtbxNome 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5415
      End
      Begin VB.TextBox txtbxHist 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   5415
      End
      Begin VB.TextBox txtbxVenc 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtbxValor 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Centro de Custo:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Histórico:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Vencimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "V&alor:"
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   1560
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   5655
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   3840
         Picture         =   "frmEdiLan.frx":0337
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   2280
         Picture         =   "frmEdiLan.frx":0481
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   4560
      TabIndex        =   27
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblTipo 
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
      Left            =   5040
      TabIndex        =   26
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      Height          =   195
      Left            =   3120
      TabIndex        =   25
      Top             =   360
      Width           =   405
   End
   Begin VB.Label lblValor 
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
      Left            =   3600
      TabIndex        =   24
      Top             =   360
      Width           =   390
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parc"
      Height          =   195
      Left            =   1560
      TabIndex        =   23
      Top             =   360
      Width           =   330
   End
   Begin VB.Label lblParc 
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
      Left            =   2040
      TabIndex        =   22
      Top             =   360
      Width           =   750
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seq:"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   330
   End
   Begin VB.Label lblSeq_Itd 
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
      Left            =   600
      TabIndex        =   20
      Top             =   360
      Width           =   750
   End
   Begin VB.Label lblData 
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
      Left            =   3600
      TabIndex        =   19
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      Height          =   195
      Left            =   3120
      TabIndex        =   18
      Top             =   120
      Width           =   390
   End
   Begin VB.Label lblNUM_Doc 
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
      Left            =   1920
      TabIndex        =   17
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº"
      Height          =   195
      Left            =   1560
      TabIndex        =   16
      Top             =   120
      Width           =   180
   End
   Begin VB.Label lblSeq_Doc 
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
      Left            =   600
      TabIndex        =   15
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "frmEdiLan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxPlanoConta_Click()
    cmbxPlanoConta.tag = F_BuscaUmCampo("cod_pla", "tab_planocontas", "nome_pla='" & cmbxPlanoConta.Text & _
        "' and tipo_pla='" & IIf(Left(lblTipo.caption, 1) = "R", "C", "D") & "'")
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxValor.Text = "0,00" Then
        MsgBox "Este valor é inválido!", vbCritical
        txtbxValor.SetFocus
        Exit Sub
    End If
    If txtbxHist.Text = "" Then
'        MsgBox "Informe o Histórico!", vbCritical
'        txtbxHist.SetFocus
'        Exit Sub
        
    End If
    If txtbxNome.tag = "" Then
        MsgBox "Dê enter no campo 'NOME' para escolher um terceiro!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
        
    End If
    If MsgBox("Confirma alteração da Parcela?", vbYesNo + vbQuestion) = vbYes Then
        altera
    End If
    
End Sub
Sub altera()
    Dim strSQL As String
    On Error GoTo trata_erro
    
    strSQL = "update tab_doc1 set historico_doc='" & txtbxHist.Text & "', "
    strSQL = strSQL & "cod_Ter='" & txtbxNome.tag & "' "
    If cmbxPlanoConta.tag <> "" Then
        strSQL = strSQL & " ,cod_pla='" & cmbxPlanoConta.tag & "' "
    End If
    strSQL = strSQL & " where seq_doc='" & lblSeq_Doc.caption & "' "
    conn.Execute strSQL
    strSQL = "update tab_itdoc1 set venc_itd=" & F_GravaData(txtbxVenc.Text) & ","
    strSQL = strSQL & "reman_itd=" & F_GravaDec(txtbxValor.Text)
    strSQL = strSQL & " where seq_itd='" & lblSeq_Itd & "'"
    conn.Execute strSQL
    MsgBox "Documento Alterado!", vbInformation
    Unload Me
    
    
    
    
    
    
    'replicação
    
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub txtbxHist_GotFocus()
    ganhaFoco txtbxHist
End Sub

Private Sub txtbxHist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxHist_LostFocus()
    AjustaTexto txtbxHist
End Sub

Private Sub txtbxNome_GotFocus()
    ganhaFoco txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    If KeyCode = vbKeyReturn Then
    
        If Left(lblTipo.caption, 1) = "R" Then
            frmSelectCli.txtbxNome.tag = "C"
        Else
            frmSelectCli.txtbxNome.tag = "F"
        End If
        frmSelectCli.txtbxNome.Text = txtbxNome.Text
    
        On Error Resume Next
    
        frmSelectCli.Show 1
        If frmSelectCli.adc.Recordset!cod <> "" Then
            strSQL = "Mudar o Terceiro de uma parcela alterará"
            strSQL = strSQL & " o terceiro de todas as parcelas" & Chr(13)
            strSQL = strSQL & "referentes a este Documento!" & Chr(13)
            strSQL = strSQL & "Deseja Continunar?"
            If MsgBox(strSQL, vbYesNo + vbCritical) = vbYes Then
                txtbxNome.Text = frmSelectCli.adc.Recordset!nome
                txtbxNome.tag = frmSelectCli.adc.Recordset!cod
                
        
            End If
            
            
        End If
        Unload frmSelectCli

    End If
    
End Sub

Private Sub txtbxNome_LostFocus()
    AjustaTexto txtbxNome
End Sub

Private Sub txtbxValor_GotFocus()
    ganhaFoco txtbxValor
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    PerdeFoco txtbxValor
End Sub

Private Sub txtbxVenc_GotFocus()
    ganhaFoco txtbxVenc
End Sub

Private Sub txtbxVenc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxVenc_LostFocus()
    P_AjustaData txtbxVenc
End Sub
