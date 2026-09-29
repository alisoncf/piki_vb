VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLanDoc 
   Caption         =   "Lançamento de Documentos"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   DrawStyle       =   5  'Transparent
   Icon            =   "frmLanDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9660
   Begin VB.ComboBox cmbxPlanoConta 
      BackColor       =   &H000080FF&
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
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3120
      MaxLength       =   50
      TabIndex        =   13
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Frame frVal 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   9255
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
         TabIndex        =   50
         Top             =   1920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   2295
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
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
      Width           =   9255
      Begin VB.CommandButton cmndbtnDistr 
         Caption         =   "Distribuir"
         Height          =   255
         Left            =   5160
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtbxInterv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   49
         Text            =   "30"
         Top             =   360
         Width           =   615
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
         Left            =   3600
         List            =   "frmLanDoc.frx":03A9
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   345
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         Caption         =   "Considerar Data Base"
         Height          =   195
         Left            =   6720
         TabIndex        =   48
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
      Left            =   5640
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
      Left            =   7560
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
      ScaleWidth      =   9630
      TabIndex        =   41
      Top             =   7185
      Width           =   9660
      Begin VB.Label lblTipo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito"
         Height          =   195
         Left            =   8760
         TabIndex        =   51
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
         TabIndex        =   47
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
         Left            =   5520
         TabIndex        =   46
         Top             =   0
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   0
         Width           =   1170
      End
   End
   Begin VB.TextBox txtbxNum_Doc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtbxRef 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   240
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
      Width           =   6495
      Begin VB.TextBox txtbxNOme 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5175
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         Picture         =   "frmLanDoc.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblCPFCap 
         AutoSize        =   -1  'True
         Caption         =   "CPF/CNPJ:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
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
         TabIndex        =   39
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
      Left            =   3120
      TabIndex        =   12
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento:"
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   960
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Referência:"
      Height          =   195
      Left            =   240
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
    iRow = msflxgrd.Row
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .col = 1
            cVal = cVal + CCur(.Text)
        Next
        
    End With
    F_Soma = cVal
    msflxgrd.Row = iRow
End Function

Private Sub cmbxDias_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
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
        msflxgrd.AddItem strSQL
    Else
    End If
End Sub

Private Sub cmbxPlanoConta_Click()
    cmbxPlanoConta.tag = F_BuscaUmCampo("cod_pla", "tab_planocontas", "nome_pla='" & _
        cmbxPlanoConta.Text & "' and tipo_pla='" & IIf(Left(lblTipo.caption, 1) = "P", "D", "C") & "'")
End Sub

Private Sub cmndbtnBuscar_Click()
    If optnPag.Value Then
        buscaFor
    Else
        buscaCli
    End If
    
   
    
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_BuscaFornecedor(txtbxNome.Text)
    If f.cod_for <> "" Then
        txtbxNome.Text = f.nome_for
        txtbxNome.tag = f.cod_for
        lblCPF.caption = f.cnpj_for
        txtbxRef.SetFocus
    End If
    
End Sub
Sub buscaCli()
    Dim c As Clientes.tCliente
    
    c = Clientes.F_BuscaCliente("C", txtbxNome.Text)
    If c.cod_cli <> "" Then
        txtbxNome.Text = c.nome_Cli
        txtbxNome.tag = c.cod_cli
        lblCPF.caption = c.cnpj_Cli
        txtbxRef.SetFocus
    End If
    
End Sub
Private Sub cmndbtnCancelar_Click()
    If MsgBox("Deseja Realmente Cancelar a Distribuição das parcelas? ", _
                vbYesNo + vbCritical) = vbYes Then
        msflxgrd.rows = 1
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
    
    If txtbxNome.tag = "" Then
        MsgBox "Escolha o Terceiro primeiramente!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Confirma distribuição das parcelas?", vbYesNo + vbQuestion) = vbYes Then
        'iDias = Int(txtbxInterv.Text)
        msflxgrd.rows = 1
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
            msflxgrd.AddItem strSQL
            
            cTotal = cTotal + Round((CCur(txtbxValor_Doc.Text) / CCur(txtbxParc.Text)), 2)
            
        Next
            If cTotal <> CCur(txtbxValor_Doc.Text) Then
                msflxgrd.Row = msflxgrd.rows - 1
                msflxgrd.col = 1
                msflxgrd.Text = FormatVal(CCur(msflxgrd.Text) + CCur(txtbxValor_Doc.Text) - cTotal)
            End If
    End If
    frCab.Enabled = False
    frVal.Enabled = True
End Sub

Private Sub cmndbtnEdit_Click()
    Dim sVen As String
    Dim sVal As String
    If msflxgrd.rows = 1 Then
        MsgBox "Não há o que ser Editado!", vbExclamation
    Else
        With msflxgrd
            .col = 0
            sVen = .Text
            .col = 1
            sVal = .Text
            .col = 2
            If Left(cmbxEspecie.Text, 1) = "1" Then
                Dim p As parcela
                p = F_DadosPar(sVen, sVal)
                
                If p.valor <> 0 Then
                    If (F_Soma - CCur(sVal) + p.valor) > CCur(txtbxValor_Doc.Text) Then
                        MsgBox "O Valor excede o valor total!", vbExclamation
                        Exit Sub
                    End If
                    .col = 0
                    .Text = p.vencimento
                    .col = 1
                    .Text = FormatVal(p.valor)
                    .col = 2
                    .Text = cmbxEspecie.Text
                    .col = 3
                    .Text = ""
                    .col = 4
                    .Text = ""
                    .col = 5
                    .Text = ""
                    .col = 6
                    .Text = ""
                    .col = 7
                    .Text = ""
                    .col = 8
                    .Text = ""
                    .col = 9
                    .Text = ""
                End If
            End If
        End With
    End If
End Sub

Private Sub cmndbtnRem_Click()
    If msflxgrd.rows = 1 Then
        MsgBox "Não há o que ser excluído!", vbExclamation
    Else
        If MsgBox("Deseja excluir Item? ", vbYesNo + vbCritical) = vbYes Then
            If msflxgrd.rows > 2 Then
                msflxgrd.RemoveItem msflxgrd.Row
            Else
                msflxgrd.rows = 1
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
     With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            If optnPag.Value Then
                .col = 2
                sTipo = Left(.Text, 1)
                If sTipo = "2" Then
                    .col = 3
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        Retorno = False
                    End If
                    .col = 4
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        Retorno = False
                    End If
                    
                    
                    
                    
                    
                End If
            
            Else
                .col = 2
                sTipo = Left(.Text, 1)
                If sTipo = "2" Then
                    .col = 3
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        Retorno = False
                    End If
                    .col = 4
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        Retorno = False
                    End If
                    .col = 5
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " agência" & Chr(13)
                        Retorno = False
                    End If
                    .col = 6
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " conta" & Chr(13)
                        Retorno = False
                    End If
                    .col = 9
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " nome do titular" & Chr(13)
                        Retorno = False
                    End If
                    .col = 10
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
    If msflxgrd.rows = 1 Then
        MsgBox "Os valores não estão distribuídos!", vbExclamation
        Exit Sub
    End If
    If cmbxPlanoConta.tag = "" Then
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
        CStr(txtbxRef.Text), txtbxNome.tag, Left(lblTipo.caption, 1), "", _
            msflxgrd.rows - 1, cmbxPlanoConta.tag)
    If sSeq_Doc <> "" Then
        With msflxgrd
            For iCount = 1 To .rows - 1
                .Row = iCount
                msflxgrd.col = 0
                sVen = msflxgrd.Text
                msflxgrd.col = 1
                sVal = msflxgrd.Text
                msflxgrd.col = 2
                sEsp = Left(msflxgrd.Text, 1)
                msflxgrd.col = 3
                sNum_Che = msflxgrd.Text
                msflxgrd.col = 4
                sBco = msflxgrd.Text
                msflxgrd.col = 5
                sAge = msflxgrd.Text
                msflxgrd.col = 6
                sCCO = msflxgrd.Text
                msflxgrd.col = 9
                sNome_che = msflxgrd.Text
                msflxgrd.col = 10
                sCPF = msflxgrd.Text
                If sEsp = "1" Then
                    sSeq_Itd = F_GravaItDoc(sSeq_Doc, sVal, sVen, iCount)
                ElseIf sEsp = "2" Then
                    If optnReceb.Value = True Then 'cheque a receber
                        sSeq_che = F_GravaCh(sSeq_Doc, sNum_Che, sBco, sAge, sCCO, _
                        sNome_che, sVal, sVen, txtbxNome.tag, Left(lblTipo.caption, 1), _
                        "", "", CStr(iCount), "", sCPF, Time)
                    Else 'cheque a pagar
                        sSeq_che = F_GravaChePag(sSeq_Doc, sNum_Che, sBco, sVal, sVen, _
                        txtbxNome.tag, CStr(iCount), Time)
                    End If
                End If
                
            Next
        End With
    End If
    
    MsgBox "Documento Gravado!", vbInformation
    msflxgrd.rows = 1
    txtbxValor_Doc.Text = "0,00"
    txtbxParc.Text = "1"
    frCab.Enabled = True
    frVal.Enabled = False
    txtbxNome.SetFocus
    
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF5 Then
        msflxgrd.SetFocus
    End If
End Sub

Private Sub Form_Load()
bCheque = True
txtbxRef.Text = Date
cmbxDias.ListIndex = 0
optnReceb_Click
P_CentraForm Me, 600
    msflxgrd.rows = 1
    msflxgrd.Cols = 11
    msflxgrd.col = 0
    msflxgrd.Text = "Vencimento"
    msflxgrd.col = 1
    msflxgrd.Text = "Valor"
    msflxgrd.col = 2
    msflxgrd.Text = "Espécie"
    msflxgrd.col = 3
    msflxgrd.Text = "Cheque Nº"
    msflxgrd.col = 4
    msflxgrd.Text = "Ban/Con"
    msflxgrd.col = 5
    msflxgrd.Text = "Agencia"
    msflxgrd.col = 6
    msflxgrd.Text = "Conta C."
    msflxgrd.col = 7
    msflxgrd.Text = "Valor"
    msflxgrd.col = 8
    msflxgrd.Text = "Vencimento"
    msflxgrd.col = 9
    msflxgrd.Text = "Titular"
    msflxgrd.col = 10
    msflxgrd.Text = "CPF"
    msflxgrd.ColAlignment(2) = 1
    txtbxVenIni.Text = Date
    msflxgrd.ColWidth(0) = 1000
    
    msflxgrd.ColWidth(1) = 1000
    msflxgrd.ColWidth(2) = 1500
End Sub

Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
End Sub


Private Sub msflxGrd_GotFocus()
    msflxgrd.BorderStyle = flexBorderSingle
    
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
        With msflxgrd
            If .col = 2 Then
                If KeyCode = vbKeySpace Then
                    cmbxOpc.Left = .Left + .CellLeft
                    cmbxOpc.Top = .Top + .CellTop
                    cmbxOpc.Width = .CellWidth
                    cmbxOpc.Text = msflxgrd.Text
                    msflxgrd.SetFocus
                    cmbxOpc.Visible = True
                
                End If
                
            End If
            
        End With
        If KeyCode = vbKeyReturn Then
            msflxgrd_DblClick
        ElseIf KeyCode = vbKey1 Or KeyCode = 97 Then
            msflxgrd.col = 2
            msflxgrd.Text = "1 - Duplicata"
            msflxgrd_DblClick
        ElseIf KeyCode = vbKey2 Or KeyCode = 98 Then
            msflxgrd.col = 2
            msflxgrd.Text = "2 - Cheque Pré"
            msflxgrd_DblClick
        ElseIf KeyCode = vbKeyAdd Or KeyCode = 187 Then
            cmndbtnAdd_Click
        ElseIf KeyCode = vbKeyF5 Then
            msflxgrd.SetFocus
        ElseIf KeyCode = vbKeyDelete Then
            cmndbtnRem_Click
        End If
        
    
    
    Exit Sub
    
End Sub

Private Sub msflxGrd_LostFocus()
    msflxgrd.BorderStyle = flexBorderNone
End Sub

Private Sub optnCH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub optnDup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub optnPag_Click()
        lblTipo.caption = "P - Contas a Pagar"
    txtbxNome.tag = ""
    
    txtbxNome.Text = ""
    frNome_cli.caption = "&Nome do Fornecedor"
    P_PreenchePlanoContas cmbxPlanoConta, False
End Sub

Private Sub optnPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub optnReceb_Click()
    lblTipo.caption = "R - Contas a Receber"
    txtbxNome.tag = ""
    txtbxNome.Text = ""
    frNome_cli.caption = "&Nome do Cliente"
    P_PreenchePlanoContas cmbxPlanoConta, True
End Sub

Private Sub optnReceb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxData_GotFocus()
    ganhaFoco txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
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

Private Sub txtbxInterv_GotFocus()
    ganhaFoco txtbxInterv
End Sub

Private Sub txtbxInterv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        txtbxNome_LostFocus
        TrocaTabKey KeyCode
    End If

End Sub

Private Sub txtbxInterv_LostFocus()
    PerdeFoco txtbxInterv, "###"
    If Int(txtbxInterv.Text) < 1 Then
        MsgBox "Informe o intervalo corretamente!", vbExclamation
        txtbxInterv.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtbxNome_GotFocus()
    ganhaFoco txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    AjustaTexto txtbxNome
End Sub

Private Sub txtbxNum_Doc_GotFocus()
    ganhaFoco txtbxNum_Doc
End Sub

Private Sub txtbxNum_Doc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNum_Doc_LostFocus()
    AjustaTexto txtbxNum_Doc
    buscaDoc
End Sub
Sub buscaDoc()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_erro
    If txtbxNum_Doc.Text <> "" Then
        strSQL = "select * from tab_doc1 where num_doc='" & txtbxNum_Doc.Text & "' "
        strSQL = strSQL & " and cod_Ter='" & txtbxNome.tag & "' "
        rPesq.Open strSQL, conn_client
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
trata_erro:
    MostraErro
End Sub
Private Sub txtbxParc_GotFocus()
    ganhaFoco txtbxParc
End Sub

Private Sub txtbxParc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        TrocaTabKey KeyCode
    End If

End Sub

Private Sub txtbxParc_LostFocus()
    PerdeFoco txtbxParc, "###"
    If Int(txtbxParc.Text) < 1 Then
        MsgBox "informe um valor decente", vbExclamation
        txtbxParc.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtbxRef_GotFocus()
    ganhaFoco txtbxRef
End Sub

Private Sub txtbxRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxRef_LostFocus()
    P_AjustaData txtbxRef
End Sub

Private Sub txtbxValor_GotFocus()
    ganhaFoco txtbxValor
End Sub

Private Sub txtbxValor_Doc_GotFocus()
    ganhaFoco txtbxValor_Doc
End Sub

Private Sub txtbxValor_Doc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValor_Doc_LostFocus()
    PerdeFoco txtbxValor_Doc
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    PerdeFoco txtbxValor
End Sub

Private Sub txtbxVenIni_GotFocus()
    ganhaFoco txtbxVenIni
End Sub

Private Sub txtbxVenIni_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        txtbxNome_LostFocus
        TrocaTabKey KeyCode
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
    cmbxItem.Text = Format(msflxgrd.Row, "00#")
    If optnReceb.Value = True Then
        cmbxEspecie_Click
    End If
    msflxgrd.col = 2
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
    If msflxgrd.Row = -1 Or msflxgrd.rows <= 1 Then
        Exit Sub
    End If
    
    If cmbxEspecie.ListIndex = 0 Then
        With msflxgrd
            .col = 0
            If .Text <> txtbxVenIni.Text Then
                
            End If
        End With
    ElseIf cmbxEspecie.ListIndex = 0 Then
        cmndbtnEdit_Click
    Else
        If cmbxEspecie.ListIndex = 1 Then
            
            
        
            
            msflxgrd.col = 0
            sVen = msflxgrd.Text
            msflxgrd.col = 1
            sVal = msflxgrd.Text
            msflxgrd.col = 3
            sNum = msflxgrd.Text
            msflxgrd.col = 4
            sBan = msflxgrd.Text
            msflxgrd.col = 5
            sAge = msflxgrd.Text
            msflxgrd.col = 6
            sCCO = msflxgrd.Text
            msflxgrd.col = 7
            sNome = msflxgrd.Text
            msflxgrd.col = 10
            sCPF = msflxgrd.Text
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
                msflxgrd.col = 2
                msflxgrd.Text = cmbxEspecie.Text
            End If
            
        End If
        
        With msflxgrd
           If c.num_che <> "" Then
                .col = 0
                .Text = c.Venc_che
                .col = 1
                .Text = FormatVal(c.Valor_che)
                
                
                .col = 3
                .Text = c.num_che
                .col = 4
                .Text = c.cod_bco
                .col = 5
                .Text = c.age_che
                .col = 6
                .Text = c.cco_che
                .col = 7
                .Text = c.Valor_che
                .col = 8
                .Text = c.Venc_che
                .col = 9
                .Text = c.nome_che
                .col = 10
                .Text = c.cpf_che
            End If
            If cpp.num_che <> "" Then
                .col = 0
                .Text = cpp.Venc_che
                .col = 1
                .Text = FormatVal(cpp.Valor_che)
                
                
                .col = 3
                .Text = cpp.num_che
                .col = 4
                .Text = cpp.cod_con
                
            End If
        End With
    End If
    
End Sub
Private Sub cmbxItem_Click()
    msflxgrd.Row = Int(cmbxItem.Text)
    msflxgrd.col = 2
    cmbxEspecie.ListIndex = Int(Left(msflxgrd.Text, 1)) - 1
    msflxgrd.col = 0
    txtbxData.Text = msflxgrd.Text
    msflxgrd.col = 1
    txtbxValor.Text = msflxgrd.Text
End Sub

Private Sub cmbxItem_GotFocus()
    cmbxItem.Clear
    For iCount = 1 To msflxgrd.rows - 1
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


