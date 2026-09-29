VERSION 5.00
Begin VB.Form frmLanGuias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Guias"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   Icon            =   "frmLanGuias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3690
   Begin VB.TextBox txtbxNum_Gui 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   2640
      Picture         =   "frmLanGuias.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnConfirmar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   1680
      Picture         =   "frmLanGuias.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtbxJuros_Gui 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   9
      Text            =   "0,00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtbxVenc_GUi 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtbxValor_Gui 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   5
      Text            =   "0,00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cmbxCod_Imp 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "&Número:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Juros:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "V&encimento:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Valor:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Imposto:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   600
   End
End
Attribute VB_Name = "frmLanGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub P_Grava()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sCod As String
    On Error GoTo trata_erro
    sCod = F_BuscaUmCampo("Max(COD_GUI)", "tab_Guias", "")
    If sCod = "" Then
        sCod = "00000001"
    Else
        sCod = Format(Int(sCod) + 1, "0000000#")
    End If
    
    strSQL = "insert into tab_guias(COD_CLI,COD_GUI,COD_IMP,NUM_GUI,Valor_Gui,Venc_Gui,"
    strSQL = strSQL & "juros_gui,status_GUI)values("
    strSQL = strSQL & "'" & strCOD_CLI & "',"
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & Left(cmbxCod_Imp.Text, 2) & "', "
    strSQL = strSQL & "'" & txtbxNum_Gui.Text & "', "
    strSQL = strSQL & f_GravaDec(txtbxValor_Gui.Text) & ","
    strSQL = strSQL & F_GravaData(txtbxVenc_GUi.Text) & ","
    strSQL = strSQL & "0.00,1)" 'em aberto
    banco.Execute strSQL
    MsgBox "Guia lançada com sucesso!", vbInformation
    txtbxNum_Gui.Text = ""
    txtbxValor_Gui.Text = "0,00"
    txtbxVenc_GUi.Text = ""
    cmbxCod_Imp.SetFocus
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub cmndbtnConfirmar_Click()
    If Trim(txtbxNum_Gui.Text) = "" Then
        MsgBox "Informe o número da Guia", vbExclamation
        txtbxNum_Gui.SetFocus
    ElseIf txtbxValor_Gui.Text = "0,00" Then
        MsgBox "Informe o Valor da Guia", vbExclamation
        txtbxValor_Gui.SetFocus
    ElseIf txtbxVenc_GUi.Text = "" Then
        MsgBox "Informe o Vencimento da Guia", vbExclamation
        txtbxVenc_GUi.SetFocus
    ElseIf cmbxCod_Imp.Text = "" Then
        MsgBox "Informe o Imposto referente à Guia", vbExclamation
        cmbxCod_Imp.SetFocus
    Else
        If Me.Tag = "" Then
            If F_BuscaUmCampo("Cod_Gui", "Tab_Guias", "num_gui='" & txtbxNum_Gui.Text & _
                        "' and cod_imp='" & Left(cmbxCod_Imp.Text, 2) & "' ") <> "" Then
                MsgBox "Já existe uma guia com este número para este imposto!", vbExclamation
                Exit Sub
            End If
            P_Grava
        Else
            If F_BuscaUmCampo("Cod_Gui", "Tab_Guias", "num_gui='" & txtbxNum_Gui.Text & _
                        "' and cod_imp='" & Left(cmbxCod_Imp.Text, 2) & _
                        "' and cod_gui <> '" & Me.Tag & "'") <> "" Then
                MsgBox "Já existe uma guia com este número para este imposto!", vbExclamation
                Exit Sub
            End If
            P_Altera
        End If
    End If
End Sub
Sub P_Altera()
    Dim strSQL As String
    Dim iPos As Integer
    On Error GoTo trata_erro
    strSQL = "update tab_guias set "
    strSQL = strSQL & "cod_imp='" & Left(cmbxCod_Imp.Text, 2) & "', "
    strSQL = strSQL & "num_gui='" & txtbxNum_Gui.Text & "', "
    strSQL = strSQL & "valor_gui=" & f_GravaDec(txtbxValor_Gui.Text) & ","
    strSQL = strSQL & "Venc_gui=" & F_GravaData(txtbxVenc_GUi.Text) & ","
    strSQL = strSQL & "juros_gui=" & f_GravaDec(txtbxJuros_Gui.Text)
    strSQL = strSQL & " where cod_gui='" & Me.Tag & "'"
    banco.Execute strSQL
    MsgBox "Guia Alterada com sucesso!", vbInformation
    
    If frmEdiGui.Visible = True Then
        Unload Me
        iPos = frmEdiGui.dtGui.Recordset.AbsolutePosition
        frmEdiGui.dtGui.Refresh
        frmEdiGui.dtGui.Recordset.Move iPos
    End If
    
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    mdiFrm.mnIncGuias.Enabled = False
    mdiFrm.mnGuias.Enabled = False
    P_CentraForm Me
    P_PreencheImpostos cmbxCod_Imp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiFrm.mnIncGuias.Enabled = True
    mdiFrm.mnGuias.Enabled = True
End Sub

Private Sub txtbxJuros_Gui_GotFocus()
    P_SelecTexto txtbxJuros_Gui
End Sub

Private Sub txtbxJuros_Gui_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxJuros_Gui_LostFocus()
    P_ValorLostFocus txtbxJuros_Gui
End Sub

Private Sub txtbxNum_Gui_GotFocus()
    P_SelecTexto txtbxNum_Gui
End Sub

Private Sub txtbxNum_Gui_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Gui_GotFocus()
    P_SelecTexto txtbxValor_Gui
End Sub

Private Sub txtbxValor_Gui_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Gui_LostFocus()
    P_ValorLostFocus txtbxValor_Gui
End Sub

Private Sub txtbxVenc_GUi_GotFocus()
    P_SelecTexto txtbxVenc_GUi
End Sub

Private Sub txtbxVenc_GUi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxVenc_GUi_LostFocus()
    P_AjustaData txtbxVenc_GUi
End Sub
