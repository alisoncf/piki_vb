VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLanDoc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lançamento de Documentos"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8880
   DrawStyle       =   5  'Transparent
   Icon            =   "frmLanDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbxPlanoConta 
      Height          =   315
      ItemData        =   "frmLanDoc.frx":030A
      Left            =   120
      List            =   "frmLanDoc.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   6360
      Width           =   4215
   End
   Begin VB.TextBox txtbxHist 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   13
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Frame frVal 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   8655
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   5160
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtbxValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   30
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbxEspecie 
         Height          =   315
         ItemData        =   "frmLanDoc.frx":0337
         Left            =   960
         List            =   "frmLanDoc.frx":0341
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtbxData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbxItem 
         Height          =   315
         ItemData        =   "frmLanDoc.frx":0364
         Left            =   120
         List            =   "frmLanDoc.frx":036B
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cmbxOpc 
         Height          =   315
         ItemData        =   "frmLanDoc.frx":0372
         Left            =   2160
         List            =   "frmLanDoc.frx":037C
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   2295
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4048
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         BorderStyle     =   0
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2880
         TabIndex        =   29
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3960
         TabIndex        =   31
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Item:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "E&spécie:"
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame frCab 
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   8655
      Begin VB.CommandButton cmndbtnDistr 
         Caption         =   "Distribuir"
         Height          =   255
         Left            =   4800
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtbxInterv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   50
         Text            =   "30"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtbxParc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   20
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtbxVenIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxValor_Doc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbxDias 
         Height          =   315
         ItemData        =   "frmLanDoc.frx":039F
         Left            =   3360
         List            =   "frmLanDoc.frx":03A9
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   345
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         Caption         =   "Considerar Data Base"
         Height          =   195
         Left            =   6360
         TabIndex        =   49
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "&Valor:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Tag             =   "!"
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Intervalo:"
         Height          =   195
         Left            =   3000
         TabIndex        =   21
         Tag             =   "!"
         Top             =   120
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Parc.:"
         Height          =   195
         Left            =   2520
         TabIndex        =   19
         Tag             =   "!"
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data &Base:"
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Tag             =   "!"
         Top             =   120
         Width           =   795
      End
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "F7 - &Confirmar"
      Height          =   735
      Left            =   5040
      Picture         =   "frmLanDoc.frx":03BE
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "ESC - &Fechar"
      Height          =   735
      Left            =   6960
      Picture         =   "frmLanDoc.frx":0508
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6000
      Width           =   1815
   End
   Begin VB.PictureBox pcAtalhos 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C6AE86&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   8850
      TabIndex        =   42
      Top             =   6870
      Width           =   8880
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito"
         Height          =   195
         Left            =   7920
         TabIndex        =   52
         Top             =   0
         Width           =   495
      End
      Begin VB.Label lblBuscaPro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Del> Excluir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1440
         TabIndex        =   48
         Top             =   0
         Width           =   1020
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F7 - Concluir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9000
         TabIndex        =   47
         Top             =   45
         Width           =   1020
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<F5> Navegar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3960
         TabIndex        =   46
         Top             =   0
         Width           =   1065
      End
      Begin VB.Label lblMsg2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<+> Adicionar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2640
         TabIndex        =   45
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Sair"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   10440
         TabIndex        =   44
         Top             =   45
         Width           =   735
      End
      Begin VB.Label lblTef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Enter> Editar "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   43
         Top             =   0
         Width           =   1170
      End
   End
   Begin VB.TextBox txtbxNum_Doc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtbxRef 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame frNome_cli 
      Caption         =   "&Nome do Cliente"
      Height          =   855
      Left            =   2880
      TabIndex        =   6
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         Picture         =   "frmLanDoc.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxNOme 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label lblCPFCap 
         AutoSize        =   -1  'True
         Caption         =   "CPF/CNPJ:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00000000000"
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
         Left            =   1080
         TabIndex        =   40
         Top             =   600
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Natureza"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      Begin VB.OptionButton optnPag 
         Caption         =   "A &Pagar"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optnReceb 
         Caption         =   "A &Receber"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo:"
      Height          =   855
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Width           =   1335
      Begin VB.OptionButton optnCh 
         Caption         =   "C&heque:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optnDup 
         Caption         =   "D&uplicata"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Centro de Custo:"
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   6120
      Width           =   1185
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Histórico:"
      Height          =   195
      Left            =   3000
      TabIndex        =   12
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento:"
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Top             =   960
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Referência:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   825
   End
End
Attribute VB_Name = "frmLanDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bCheque As Boolean
Function F_Soma() As Currency
    Dim cVal As Currency
    Dim iCount As Integer
    Dim iRow As Integer
    iRow = msflxGrd.Row
    With msflxGrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 1
            cVal = cVal + CCur(.Text)
        Next
        
    End With
    F_Soma = cVal
    msflxGrd.Row = iRow
End Function

Private Sub cmbxDias_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOpc_LostFocus()
    cmbxOpc.Visible = False
    
End Sub

Private Sub cmndbtnAdd_Click()
    Dim p As parcela
    Dim strSQL As String
    If F_Soma = CCur(txtbxValor_Doc.Text) Then
        MsgBox "Você não pode adicionar parcelas quando o Valor total está fechado!", vbExclamation
        Exit Sub
    End If
    p = F_DadosPar
    If p.valor <> 0 Then
        strSQL = p.vencimento & Chr(9) & FormatVal(p.valor)
        msflxGrd.AddItem strSQL
    Else
    End If
End Sub

Private Sub cmbxPlanoConta_Click()
    cmbxPlanoConta.Tag = F_BuscaUmCampo("cod_pla", "tab_planocontas", "nome_pla='" & _
        cmbxPlanoConta.Text & "' and tipo_pla='" & IIf(Left(lblTipo.caption, 1) = "P", "D", "C") & "'")
End Sub

Private Sub cmndbtnBuscar_Click()
    If Left(lblTipo.caption, 1) = "R" Then
        frmSelectCli.txtbxNome.Tag = "C"
    Else
        frmSelectCli.txtbxNome.Tag = "F"
    End If
    frmSelectCli.txtbxNome.Text = txtbxNome.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    If frmSelectCli.Tag <> "" Then
        txtbxNome.Text = frmSelectCli.dtBusca.Recordset!nome
        txtbxNome.Tag = frmSelectCli.dtBusca.Recordset!cod
        lblCPF.caption = frmSelectCli.dtBusca.Recordset!CNPJ
        txtbxNum_Doc.SetFocus
    End If
    Unload frmSelectCli
    
End Sub

Private Sub cmndbtnCancelar_Click()
    If MsgBox("Deseja Realmente Cancelar a Distribuição das parcelas? ", _
                vbYesNo + vbCritical) = vbYes Then
        msflxGrd.rows = 1
        frCab.Enabled = True
        frVal.Enabled = False
    End If
End Sub

Private Sub cmndbtnDistr_Click()
    Dim iCount As Integer
    Dim strSQL As String
    Dim iDias As Integer
    Dim cTotal As Currency
    If optnCH.Value Then bCheque = True Else bCheque = False
    If cmbxDias.ListIndex = 1 And CCur(txtbxInterv.Text) > 2 Then
        MsgBox "Intervalo muito grande!", vbExclamation
        cmbxDias.SetFocus
        Exit Sub
    End If
    
    If txtbxNome.Tag = "" Then
        MsgBox "Escolha o Terceiro primeiramente!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Confirma distribuição das parcelas?", vbYesNo + vbQuestion) = vbYes Then
        'iDias = Int(txtbxInterv.Text)
        msflxGrd.rows = 1
        If CCur(txtbxValor_Doc.Text) = 0 Then
            MsgBox "Informe um valor válido!", vbExclamation
            txtbxValor_Doc.SetFocus
            Exit Sub
        End If
        
        For iCount = 1 To Int(txtbxParc.Text)
            If cmbxDias.ListIndex = 0 Then
                iDias = Int(txtbxInterv) * (iCount - IIf(chk.Value = 1, 1, 0))
                strSQL = DateAdd("D", iDias, txtbxVenIni.Text) & Chr(9)
            Else
                iDias = Int(txtbxInterv) * iCount
                strSQL = DateAdd("M", iDias, txtbxVenIni.Text) & Chr(9)
            End If
            strSQL = strSQL & FormatVal(CCur(txtbxValor_Doc.Text) / CCur(txtbxParc.Text)) & Chr(9)
            If bCheque Then
                strSQL = strSQL & "2 - Cheque pré"
            Else
                strSQL = strSQL & "1 - Duplicata"
            End If
            msflxGrd.AddItem strSQL
            
            cTotal = cTotal + Round((CCur(txtbxValor_Doc.Text) / CCur(txtbxParc.Text)), 2)
            
        Next
            If cTotal <> CCur(txtbxValor_Doc.Text) Then
                msflxGrd.Row = msflxGrd.rows - 1
                msflxGrd.Col = 1
                msflxGrd.Text = FormatVal(CCur(msflxGrd.Text) + CCur(txtbxValor_Doc.Text) - cTotal)
            End If
    End If
    frCab.Enabled = False
    frVal.Enabled = True
End Sub

Private Sub cmndbtnEdit_Click()
    Dim sVen As String
    Dim sVal As String
    If msflxGrd.rows = 1 Then
        MsgBox "Não há o que ser Editado!", vbExclamation
    Else
        With msflxGrd
            .Col = 0
            sVen = .Text
            .Col = 1
            sVal = .Text
            .Col = 2
            If Left(cmbxEspecie.Text, 1) = "1" Then
                Dim p As parcela
                p = F_DadosPar(sVen, sVal)
                
                If p.valor <> 0 Then
                    If (F_Soma - CCur(sVal) + p.valor) > CCur(txtbxValor_Doc.Text) Then
                        MsgBox "O Valor excede o valor total!", vbExclamation
                        Exit Sub
                    End If
                    .Col = 0
                    .Text = p.vencimento
                    .Col = 1
                    .Text = FormatVal(p.valor)
                    .Col = 2
                    .Text = cmbxEspecie.Text
                    .Col = 3
                    .Text = ""
                    .Col = 4
                    .Text = ""
                    .Col = 5
                    .Text = ""
                    .Col = 6
                    .Text = ""
                    .Col = 7
                    .Text = ""
                    .Col = 8
                    .Text = ""
                    .Col = 9
                    .Text = ""
                End If
            End If
        End With
    End If
End Sub

Private Sub cmndbtnRem_Click()
    If msflxGrd.rows = 1 Then
        MsgBox "Não há o que ser excluído!", vbExclamation
    Else
        If MsgBox("Deseja excluir Item? ", vbYesNo + vbCritical) = vbYes Then
            If msflxGrd.rows > 2 Then
                msflxGrd.RemoveItem msflxGrd.Row
            Else
                msflxGrd.rows = 1
            End If
        
            
        End If
        
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Function F_confMsflx() As Boolean
     Dim iCount As Integer
     Dim sTipo As String
     Dim Retorno As Boolean
     Dim strSQL As String
     Retorno = True
     strSQL = "os seguintes dados estão faltando: " & Chr(13) & Chr(13)
     With msflxGrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            If optnPag.Value Then
                .Col = 2
                sTipo = Left(.Text, 1)
                If sTipo = "2" Then
                    .Col = 3
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 4
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        Retorno = False
                    End If
                    
                    
                    
                    
                    
                End If
            
            Else
                .Col = 2
                sTipo = Left(.Text, 1)
                If sTipo = "2" Then
                    .Col = 3
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 4
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 5
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " agência" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 6
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " conta" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 9
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " nome do titular" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 10
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " CPF do titular" & Chr(13)
                        Retorno = False
                    End If
                End If
            End If
        Next
     End With
     F_confMsflx = Retorno
     If Retorno = False Then
        MsgBox strSQL, vbExclamation
        
     End If
End Function
Private Sub cmndbtnSalvar_Click()
    Dim sSeq_Doc As String
    Dim sSeq_Itd As String
    Dim sSeq_che As String
    Dim iCount As Integer
    Dim sVen As String, sVal As String, sEsp As String, sNum_Che As String, sBco As String
    Dim sAge As String, sCCO As String, sNome_che As String
    Dim sCPF As String
    If Not F_confMsflx Then Exit Sub
    If msflxGrd.rows = 1 Then
        MsgBox "Os valores não estão distribuídos!", vbExclamation
        Exit Sub
    End If
    If cmbxPlanoConta.Tag = "" Then
        MsgBox "Não é possível gravar o documento por que não foi encontrado um centro de custo", vbExclamation
        cmbxPlanoConta.SetFocus
        Exit Sub
    End If
    If F_Soma <> CCur(txtbxValor_Doc.Text) Then
        MsgBox "Os valores não estão corretamente distribuídos!", vbExclamation
        Exit Sub
    ElseIf F_Vazio(txtbxNum_Doc.Text) Then
        MsgBox "Informe o número do Documento!", vbExclamation
        txtbxNum_Doc.SetFocus
        Exit Sub
    End If
    sSeq_Doc = F_GravaDoc(txtbxHist.Text, txtbxNum_Doc.Text, txtbxValor_Doc.Text, _
        CStr(txtbxRef.Text), txtbxNome.Tag, Left(lblTipo.caption, 1), "", _
            msflxGrd.rows - 1, cmbxPlanoConta.Tag)
    If sSeq_Doc <> "" Then
        With msflxGrd
            For iCount = 1 To .rows - 1
                .Row = iCount
                msflxGrd.Col = 0
                sVen = msflxGrd.Text
                msflxGrd.Col = 1
                sVal = msflxGrd.Text
                msflxGrd.Col = 2
                sEsp = Left(msflxGrd.Text, 1)
                msflxGrd.Col = 3
                sNum_Che = msflxGrd.Text
                msflxGrd.Col = 4
                sBco = msflxGrd.Text
                msflxGrd.Col = 5
                sAge = msflxGrd.Text
                msflxGrd.Col = 6
                sCCO = msflxGrd.Text
                msflxGrd.Col = 9
                sNome_che = msflxGrd.Text
                msflxGrd.Col = 10
                sCPF = msflxGrd.Text
                If sEsp = "1" Then
                    sSeq_Itd = F_GravaItDoc(sSeq_Doc, sVal, sVen, iCount)
                ElseIf sEsp = "2" Then
                    If optnReceb.Value = True Then 'cheque a receber
                        sSeq_che = F_GravaCh(sSeq_Doc, sNum_Che, sBco, sAge, sCCO, _
                        sNome_che, sVal, sVen, txtbxNome.Tag, Left(lblTipo.caption, 1), _
                        "", "", CStr(iCount), "", sCPF, Time)
                    Else 'cheque a pagar
                        sSeq_che = F_GravaChePag(sSeq_Doc, sNum_Che, sBco, sVal, sVen, _
                        txtbxNome.Tag, CStr(iCount), Time)
                    End If
                End If
                
            Next
        End With
    End If
    
    MsgBox "Documento Gravado!", vbInformation
    msflxGrd.rows = 1
    txtbxValor_Doc.Text = "0,00"
    txtbxParc.Text = "1"
    frCab.Enabled = True
    frVal.Enabled = False
    txtbxNome.SetFocus
    
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF5 Then
        msflxGrd.SetFocus
    End If
End Sub

Private Sub Form_Load()
bCheque = True
txtbxRef.Text = Date
cmbxDias.ListIndex = 0
optnReceb_Click
P_CentraForm Me, 600
    msflxGrd.rows = 1
    msflxGrd.Cols = 11
    msflxGrd.Col = 0
    msflxGrd.Text = "Vencimento"
    msflxGrd.Col = 1
    msflxGrd.Text = "Valor"
    msflxGrd.Col = 2
    msflxGrd.Text = "Espécie"
    msflxGrd.Col = 3
    msflxGrd.Text = "Cheque Nº"
    msflxGrd.Col = 4
    msflxGrd.Text = "Ban/Con"
    msflxGrd.Col = 5
    msflxGrd.Text = "Agencia"
    msflxGrd.Col = 6
    msflxGrd.Text = "Conta C."
    msflxGrd.Col = 7
    msflxGrd.Text = "Valor"
    msflxGrd.Col = 8
    msflxGrd.Text = "Vencimento"
    msflxGrd.Col = 9
    msflxGrd.Text = "Titular"
    msflxGrd.Col = 10
    msflxGrd.Text = "CPF"
    msflxGrd.ColAlignment(2) = 1
    txtbxVenIni.Text = Date
    msflxGrd.ColWidth(0) = 1000
    
    msflxGrd.ColWidth(1) = 1000
    msflxGrd.ColWidth(2) = 1500
End Sub

Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
End Sub


Private Sub msflxGrd_GotFocus()
    msflxGrd.BorderStyle = flexBorderSingle
    
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
        With msflxGrd
            If .Col = 2 Then
                If KeyCode = vbKeySpace Then
                    cmbxOpc.Left = .Left + .CellLeft
                    cmbxOpc.Top = .Top + .CellTop
                    cmbxOpc.Width = .CellWidth
                    cmbxOpc.Text = msflxGrd.Text
                    msflxGrd.SetFocus
                    cmbxOpc.Visible = True
                
                End If
                
            End If
            
        End With
        If KeyCode = vbKeyReturn Then
            msflxgrd_DblClick
        ElseIf KeyCode = vbKey1 Or KeyCode = 97 Then
            msflxGrd.Col = 2
            msflxGrd.Text = "1 - Duplicata"
            msflxgrd_DblClick
        ElseIf KeyCode = vbKey2 Or KeyCode = 98 Then
            msflxGrd.Col = 2
            msflxGrd.Text = "2 - Cheque Pré"
            msflxgrd_DblClick
        ElseIf KeyCode = vbKeyAdd Or KeyCode = 187 Then
            cmndbtnAdd_Click
        ElseIf KeyCode = vbKeyF5 Then
            msflxGrd.SetFocus
        ElseIf KeyCode = vbKeyDelete Then
            cmndbtnRem_Click
        End If
        
    
    
    Exit Sub
    
End Sub

Private Sub msflxGrd_LostFocus()
    msflxGrd.BorderStyle = flexBorderNone
End Sub

Private Sub optnCH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnDup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnPag_Click()
        lblTipo.caption = "P - Contas a Pagar"
    txtbxNome.Tag = ""
    
    txtbxNome.Text = ""
    frNome_cli.caption = "&Nome do Fornecedor"
    P_PreenchePlanoContas cmbxPlanoConta, False
End Sub

Private Sub optnPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnReceb_Click()
    lblTipo.caption = "R - Contas a Receber"
    txtbxNome.Tag = ""
    txtbxNome.Text = ""
    frNome_cli.caption = "&Nome do Cliente"
    P_PreenchePlanoContas cmbxPlanoConta, True
End Sub

Private Sub optnReceb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
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

Private Sub txtbxHist_GotFocus()
    P_SelecTexto txtbxHist
End Sub


Private Sub txtbxHist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxHist_LostFocus()
    P_AjustaTexto txtbxHist
End Sub

Private Sub txtbxInterv_GotFocus()
    P_SelecTexto txtbxInterv
End Sub

Private Sub txtbxInterv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        txtbxNome_LostFocus
        SendKeys Chr(9)
    End If

End Sub

Private Sub txtbxInterv_LostFocus()
    P_ValorLostFocus txtbxInterv, "###"
    If Int(txtbxInterv.Text) < 1 Then
        MsgBox "Informe o intervalo corretamente!", vbExclamation
        txtbxInterv.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub

Private Sub txtbxNum_Doc_GotFocus()
    P_SelecTexto txtbxNum_Doc
End Sub

Private Sub txtbxNum_Doc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNum_Doc_LostFocus()
    P_AjustaTexto txtbxNum_Doc
    buscaDoc
End Sub
Sub buscaDoc()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro
    If txtbxNum_Doc.Text <> "" Then
        strSQL = "select * from tab_doc1 where num_doc='" & txtbxNum_Doc.Text & "' "
        strSQL = strSQL & " and cod_Ter='" & txtbxNome.Tag & "' "
        Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
            If .RecordCount > 0 Then
                strSQL = "Já existe (m) " & .RecordCount & " documento (s) para este "
                
                strSQL = strSQL & IIf(optnPag.Value, "Fornecedor", "Cliente") & " com este Número!"
                strSQL = strSQL & Chr(13) & "Deseja continuar?"
                If MsgBox(strSQL, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                    txtbxNum_Doc.Text = ""
                    txtbxNum_Doc.SetFocus
                End If
                
            End If
            .Close
        End With
    
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Private Sub txtbxParc_GotFocus()
    P_SelecTexto txtbxParc
End Sub

Private Sub txtbxParc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        SendKeys Chr(9)
    End If

End Sub

Private Sub txtbxParc_LostFocus()
    P_ValorLostFocus txtbxParc, "###"
    If Int(txtbxParc.Text) < 1 Then
        MsgBox "informe um valor decente", vbExclamation
        txtbxParc.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtbxRef_GotFocus()
    P_SelecTexto txtbxRef
End Sub

Private Sub txtbxRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxRef_LostFocus()
    P_AjustaData txtbxRef
End Sub

Private Sub txtbxValor_GotFocus()
    P_SelecTexto txtbxValor
End Sub

Private Sub txtbxValor_Doc_GotFocus()
    P_SelecTexto txtbxValor_Doc
End Sub

Private Sub txtbxValor_Doc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Doc_LostFocus()
    P_ValorLostFocus txtbxValor_Doc
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    P_ValorLostFocus txtbxValor
End Sub

Private Sub txtbxVenIni_GotFocus()
    P_SelecTexto txtbxVenIni
End Sub

Private Sub txtbxVenIni_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyF2 Then
        strSQL = InputBox("Quantos dias deseja adicionar?", "Soma de Dias", "0")
        strSQL = F_SoNumero(strSQL)
        txtbxVenIni.Text = DateAdd("D", CDbl(strSQL), txtbxVenIni.Text)
        
    End If
End Sub

Private Sub txtbxVenIni_LostFocus()
    P_AjustaData txtbxVenIni
End Sub
Private Sub msflxgrd_DblClick()
    cmbxItem_GotFocus
    cmbxItem.Text = Format(msflxGrd.Row, "00#")
    If optnReceb.Value = True Then
        cmbxEspecie_Click
    End If
    msflxGrd.Col = 2
    'cmndbtnEdit_Click
    
End Sub

Private Sub cmbxEspecie_Click()
    Dim c As cheque
    Dim sVal As String
    Dim sNum As String
    Dim sVen As String
    Dim sNome As String
    Dim sBan As String
    Dim sCCO As String
    Dim sAge As String
    Dim cpp As chePag
    Dim sCPF As String
    If msflxGrd.Row = -1 Or msflxGrd.rows <= 1 Then
        Exit Sub
    End If
    
    If cmbxEspecie.ListIndex = 0 Then
        With msflxGrd
            .Col = 0
            If .Text <> txtbxVenIni.Text Then
                
            End If
        End With
    ElseIf cmbxEspecie.ListIndex = 0 Then
        cmndbtnEdit_Click
    Else
        If cmbxEspecie.ListIndex = 1 Then
            
            
        
            
            msflxGrd.Col = 0
            sVen = msflxGrd.Text
            msflxGrd.Col = 1
            sVal = msflxGrd.Text
            msflxGrd.Col = 3
            sNum = msflxGrd.Text
            msflxGrd.Col = 4
            sBan = msflxGrd.Text
            msflxGrd.Col = 5
            sAge = msflxGrd.Text
            msflxGrd.Col = 6
            sCCO = msflxGrd.Text
            msflxGrd.Col = 7
            sNome = msflxGrd.Text
            msflxGrd.Col = 10
            sCPF = msflxGrd.Text
            If sNum = "" Then
                sNum = txtbxNum_Doc.Text
            End If
            If sCPF = "" Then
                sCPF = lblCPF.caption
            End If
            
            If optnReceb.Value = True Then
                c = F_DadosCheque(True, sVen, txtbxNome.Text, sNum, sCCO, sAge, sBan, sVal, sCPF)
            Else
                cpp = F_DadosChePag(sVen, sNum, sCCO, sVal)
            End If
            If c.num_che = "" And cpp.num_che = "" Then
                MsgBox "Não foi informado um Cheque Válido!", vbExclamation
                Exit Sub
            Else
                msflxGrd.Col = 2
                msflxGrd.Text = cmbxEspecie.Text
            End If
            
        End If
        
        With msflxGrd
           If c.num_che <> "" Then
                .Col = 0
                .Text = c.Venc_che
                .Col = 1
                .Text = FormatVal(c.Valor_che)
                
                
                .Col = 3
                .Text = c.num_che
                .Col = 4
                .Text = c.cod_bco
                .Col = 5
                .Text = c.age_che
                .Col = 6
                .Text = c.cco_che
                .Col = 7
                .Text = c.Valor_che
                .Col = 8
                .Text = c.Venc_che
                .Col = 9
                .Text = c.nome_che
                .Col = 10
                .Text = c.cpf_che
            End If
            If cpp.num_che <> "" Then
                .Col = 0
                .Text = cpp.Venc_che
                .Col = 1
                .Text = FormatVal(cpp.Valor_che)
                
                
                .Col = 3
                .Text = cpp.num_che
                .Col = 4
                .Text = cpp.cod_con
                
            End If
        End With
    End If
    
End Sub
Private Sub cmbxItem_Click()
    msflxGrd.Row = Int(cmbxItem.Text)
    msflxGrd.Col = 2
    cmbxEspecie.ListIndex = Int(Left(msflxGrd.Text, 1)) - 1
    msflxGrd.Col = 0
    txtbxData.Text = msflxGrd.Text
    msflxGrd.Col = 1
    txtbxValor.Text = msflxGrd.Text
End Sub

Private Sub cmbxItem_GotFocus()
    cmbxItem.Clear
    For iCount = 1 To msflxGrd.rows - 1
        cmbxItem.AddItem Format(iCount, "0##")
    Next
    
End Sub

Private Sub cmbxNat_Click()
    If cmbxNat.ListIndex = 1 Then
        Dim c As cheque
        c = F_DadosCheque(False, Date, lblNome_Che.caption, lblNum_Che.caption)
        If c.num_che <> "" Then
            lblAge_Che.caption = c.age_che
            lblBco.caption = c.cod_bco
            lblCco_Che.caption = c.cco_che
            lblNome_Che.caption = c.nome_che
            lblNum_Che.caption = c.num_che
            lblValor_Che.caption = c.Valor_che
            lblVenc_Che.caption = c.Venc_che
            frCheque.Visible = True
        End If
    End If
End Sub


