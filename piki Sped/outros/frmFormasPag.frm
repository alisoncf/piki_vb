VERSION 5.00
Begin VB.Form frmIncFormasPag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Formas de Pagamento"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5490
   Icon            =   "frmFormasPag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fr 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5235
      Begin VB.OptionButton optnChe 
         Caption         =   "&Cheque"
         Height          =   195
         Left            =   2160
         TabIndex        =   13
         Top             =   2760
         Width           =   1215
      End
      Begin VB.OptionButton optnDup 
         Caption         =   "&Duplicata"
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   2760
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnConfirmar 
         Caption         =   "&Confirmar"
         Height          =   615
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtbxInterv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Text            =   "0"
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox cmbxEntr 
         Height          =   315
         ItemData        =   "frmFormasPag.frx":030A
         Left            =   1080
         List            =   "frmFormasPag.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtbxJuros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Text            =   "0"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtbxParc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtbxDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Formas de Pagamento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   3975
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   5040
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dias"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Intervalo:"
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Entrada:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tx.Jur(%):"
         Height          =   195
         Left            =   270
         TabIndex        =   5
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Parcelas:"
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D&escrição:"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   720
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmIncFormasPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxEntr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnConfirmar_Click()
    If txtbxDesc.Text = "" Then
        MsgBox "informe a Descrição da forma de pagamento!", vbExclamation
        txtbxDesc.SetFocus
        Exit Sub
    End If
    If txtbxParc.Text = "0" Then
        MsgBox "Informe uma quantidade de Parcelas Válida!", vbExclamation
        txtbxParc.SetFocus
        Exit Sub
    End If
    If CCur(txtbxJuros.Text) = 0 And CCur(txtbxParc.Text) > 1 Then
        If MsgBox("O Valor dos Juros está correto ?", vbYesNo + vbExclamation) = vbNo Then
            txtbxJuros.SetFocus
            Exit Sub
            End If
    End If
    If cmbxEntr.ListIndex = -1 Then
        MsgBox "Informe se há entrada ou não!", vbExclamation
        cmbxEntr.SetFocus
        Exit Sub
    End If
    If Me.Tag = "" Then
        P_Grava
    Else
        P_Altera
    End If
End Sub
Sub P_Grava()
    Dim strSQL As String
    Dim sCod As String
    sCod = MaxKeySQL("tab_tipoven", "cod_tip", "", "0#")
    strSQL = "insert into tab_tipoven(cod_tip,parcelas_tip,juros_tip,descricao_tip,"
    strSQL = strSQL & "entrada_tip,intervalo_tip,tipo_tip)Values("
    strSQL = strSQL & "'" & sCod & "', "
    strSQL = strSQL & Str(txtbxParc.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxJuros.Text) & ","
    strSQL = strSQL & "'" & txtbxDesc.Text & "', "
    strSQL = strSQL & "'" & Left(cmbxEntr.Text, 1) & "',"
    strSQL = strSQL & Str(txtbxInterv.Text) & ", "
    strSQL = strSQL & "'" & IIf(optnDup.Value, "D", "C") & "') "
    banco.Execute strSQL
    MsgBox "Plano de Pagamento gravado!", vbInformation
    txtbxDesc.Text = ""
    txtbxInterv.Text = "0"
    txtbxJuros.Text = "0,00"
    txtbxDesc.SetFocus
    If IsLoad("frmEdiFormasPag") Then
        frmEdiFormasPag.dtProd.Refresh
    End If
End Sub
Sub P_Altera()
    Dim strSQL As String
    
    strSQL = "Update tab_tipoven set "
    strSQL = strSQL & "parcelas_tip = " & Str(txtbxParc.Text) & ","
    strSQL = strSQL & "juros_tip =" & F_GravaDec(txtbxJuros.Text) & ","
    strSQL = strSQL & "descricao_tip = '" & txtbxDesc.Text & "', "
    strSQL = strSQL & "entrada_tip = '" & Left(cmbxEntr.Text, 1) & "',"
    strSQL = strSQL & "intervalo_tip =" & Str(txtbxInterv.Text) & ", "
    strSQL = strSQL & "tipo_tip='" & IIf(optnDup.Value, "D", "C") & "' "
    strSQL = strSQL & " where cod_tip='" & Me.Tag & "' "
    banco.Execute strSQL
    MsgBox "Plano de Pagamento Alterado!", vbInformation
    If IsLoad("frmEdiFormasPag") Then
        frmEdiFormasPag.dtProd.Refresh
    End If
    Unload Me
End Sub
Private Sub cmndbtnSair_Click()
    If IsLoad(frmEdiFormasPag.Name) Then
        frmEdiFormasPag.dtProd.Refresh
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    
    P_CentraForm Me, 0
End Sub

Private Sub txtbxDesc_GotFocus()
    P_SelecTexto txtbxDesc
End Sub

Private Sub txtbxDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDesc_LostFocus()
    P_AjustaTexto txtbxDesc
End Sub

Private Sub txtbxInterv_GotFocus()
    P_SelecTexto txtbxInterv
End Sub

Private Sub txtbxInterv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If

End Sub

Private Sub txtbxInterv_LostFocus()
    P_ValorLostFocus txtbxInterv, "###,###,###"
End Sub

Private Sub txtbxJuros_GotFocus()
    P_SelecTexto txtbxJuros
End Sub

Private Sub txtbxJuros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If

End Sub

Private Sub txtbxJuros_LostFocus()
    P_ValorLostFocus txtbxJuros, "###,##0.0000"
End Sub

Private Sub txtbxParc_GotFocus()
    P_SelecTexto txtbxParc
End Sub

Private Sub txtbxParc_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If

End Sub

Private Sub txtbxParc_LostFocus()
    P_ValorLostFocus txtbxParc, "###"
End Sub
