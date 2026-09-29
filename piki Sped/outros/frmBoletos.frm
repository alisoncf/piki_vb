VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBoletos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão de Boletos"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   Icon            =   "frmBoletos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9015
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.TextBox txtbxDataProc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   51
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtbxDataDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtbxLocal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Text            =   "Pagamento em qualquer Agência até o Vencimento"
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox txtbxEspDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtbxAceite 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtbxInstr 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "frmBoletos.frx":0442
         Top             =   3600
         Width           =   4215
      End
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   8535
         Begin VB.TextBox txtbxBoleto 
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtbxValor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   9
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxParc 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtbxVenc 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmndbtnAdd 
            Height          =   405
            Left            =   120
            Picture         =   "frmBoletos.frx":046B
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1440
            Width           =   695
         End
         Begin MSFlexGridLib.MSFlexGrid msflxGrd 
            Height          =   1575
            Left            =   2520
            TabIndex        =   13
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
            ForeColorFixed  =   -2147483642
            BackColorBkg    =   16777215
            GridLinesFixed  =   1
            Appearance      =   0
         End
         Begin VB.CommandButton cmndbtnEdit 
            Height          =   405
            Left            =   865
            Picture         =   "frmBoletos.frx":05B5
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1440
            Width           =   695
         End
         Begin VB.CommandButton cmndbtnCancel 
            Height          =   405
            Left            =   1605
            Picture         =   "frmBoletos.frx":06FF
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1440
            Width           =   695
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "&Nº do Boleto:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1320
            TabIndex        =   8
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Parcela:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Vencimento:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1320
            TabIndex        =   4
            Top             =   120
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   4560
         TabIndex        =   26
         Top             =   2160
         Width           =   4095
         Begin VB.TextBox txtbxSacado 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            TabIndex        =   28
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txtbxCnpj 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            TabIndex        =   34
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox txtbxEnd 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            TabIndex        =   30
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox txtbxBairro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            TabIndex        =   36
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox txtbxCidade 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            TabIndex        =   32
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox txtbxUf 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3600
            TabIndex        =   40
            Text            =   "GO"
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox txtbxCep 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            TabIndex        =   38
            Text            =   "75.115-430"
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nome:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblCnpj 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CNPJ:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Bairro:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cidade:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "UF:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3120
            TabIndex        =   39
            Top             =   2040
            Width           =   450
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CEP:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   37
            Top             =   2040
            Width           =   570
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "End:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   570
         End
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Proc.:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         TabIndex        =   52
         Top             =   2760
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local de Pagamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Doc"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Número Doc:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   4440
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Esp. Docum."
         Height          =   195
         Left            =   1200
         TabIndex        =   20
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aceite:"
         Height          =   195
         Left            =   2280
         TabIndex        =   22
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Instruções:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   780
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Buscar:"
      Height          =   1095
      Left            =   120
      TabIndex        =   41
      Top             =   4800
      Width           =   2655
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   285
         Left            =   2040
         Picture         =   "frmBoletos.frx":0849
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   600
         Width           =   285
      End
      Begin VB.TextBox txtbxBusca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optnSeq 
         Caption         =   "Seqüencial"
         Height          =   195
         Left            =   1320
         TabIndex        =   43
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optnDoc 
         Caption         =   "Documento"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   2880
      TabIndex        =   46
      Top             =   4800
      Width           =   6015
      Begin VB.CommandButton cmndbtnConfig 
         Caption         =   "&Preferências"
         Height          =   735
         Left            =   120
         Picture         =   "frmBoletos.frx":0993
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir Boletos"
         Height          =   735
         Left            =   1560
         Picture         =   "frmBoletos.frx":0ADD
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "Salvar Boletos"
         Height          =   735
         Left            =   3000
         Picture         =   "frmBoletos.frx":0C27
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "Sair"
         Height          =   735
         Left            =   4440
         Picture         =   "frmBoletos.frx":0D71
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmBoletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnAdd_Click()
    If cmndbtnAdd.Tag = "" Then
        MsgBox "Escolha a parcela que deseja editar!", vbExclamation
        Exit Sub
    Else
        msflxGrd.Row = cmndbtnAdd.Tag
        msflxGrd.Col = 0
        msflxGrd.Text = txtbxBoleto.Text
        txtbxBoleto.Text = ""
        cmndbtnCancel_Click
        msflxGrd.SetFocus
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select td.*,ti.*,tc.* from tab_clientes as tc,"
    strSQL = strSQL & " tab_doc1 as td, tab_itdoc1 as ti "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.cod_ter=tc.cod_cli "
    If optnDoc.Value Then
        strSQL = strSQL & " and num_doc='" & txtbxBusca.Text & "'"
    Else
        strSQL = strSQL & " and td.seq_doc='" & txtbxBusca.Text & "'"
    End If
    strSQL = strSQL & " order by seq_itd "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Documento não encontrado!", vbCritical
            .Close
            Exit Sub
        End If
        msflxGrd.rows = 1
        txtbxSacado.Text = !nome_Cli
        txtbxEnd.Text = !endereco_cli
        txtbxBairro.Text = !bairro_cli
        txtbxCidade.Text = !cidade_cli
        txtbxUf.Text = !uf_cli
        txtbxCep.Text = !cep_cli
        txtbxCNPJ.Text = !cpf_cli
        txtbxNumDoc.Text = !Num_Doc
        txtbxDataDoc.Text = !Data_Doc
        
        Do While Not .EOF
            strSQL = Chr(9) & Format(!Parc_Itd, "0#") & "/" & Format(!Parc_Doc, "0#") & Chr(9)
            strSQL = strSQL & !Venc_Itd & Chr(9)
            strSQL = strSQL & FormatVal(!reman_itd) & Chr(9) & !seq_itd
            msflxGrd.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub cmndbtnCancel_Click()
    txtbxParc.Text = ""
    txtbxVenc.Text = ""
    txtbxBoleto.Text = ""
    txtbxValor.Text = ""
    cmndbtnAdd.Tag = ""
End Sub

Private Sub cmndbtnConfig_Click()
    frmBolConfig.Show 1
    
End Sub

Private Sub cmndbtnEdit_Click()
    With msflxGrd
        .Col = 0
        txtbxBoleto.Text = .Text
        .Col = 1
        txtbxParc.Text = .Text
        .Col = 2
        txtbxVenc.Text = .Text
        .Col = 3
        txtbxValor.Text = .Text
        cmndbtnAdd.Tag = .Row
        txtbxBoleto.SetFocus
        
    End With
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime Printer
    Printer.EndDoc
    
End Sub
Sub imprime(obj As Object)
    Dim x As Single
    Dim i As Single
    Dim sAltura As Single
    Dim sLargura As Single
    Dim iCount As Integer
    Dim rPesq As Recordset
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "select * from tab_bolconfig"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    
    
     'A ESCALA SEMPRE TEM QUE SER PASSADA EM VBTWIPS, PORTANTO É NECESSÁRIO TRANSFORMAR
    sAltura = 102 '10 cm
    sLargura = 170 '17 cm
    Printer.ScaleMode = vbMillimeters
    Printer.PaperSize = 256
    sAltura = Printer.ScaleY(sAltura, vbMillimeters, vbTwips)
    sLargura = Printer.ScaleX(sLargura, vbMillimeters, vbTwips)
    obj.FontName = "arial"
    obj.fontsize = 10
    Printer.Height = sAltura
    Printer.Width = sLargura
    Printer.ScaleHeight = 80.76855
    Printer.ScaleWidth = 163.2304
    For iCount = 0 To 170
        For j = 0 To 90
            
        Next
    Next
    
    With msflxGrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            'P_PrintXY obj, txtbxLocal.Text, 5, 5
            
            .Col = 1 'PARCELA
            P_PrintXY obj, .Text, rPesq!parcx_bol, rPesq!parcy_bol
            .Col = 2 'VENCIMENTO
            P_ImprX obj, .Text, rPesq!vencx_bol, True, True
            P_PrintXY obj, txtbxDataDoc.Text, rPesq!emissx_bol, rPesq!emissy_bol
            P_PrintXY obj, txtbxNumDoc.Text, rPesq!numdocx_bol, rPesq!numdocY_boL
            P_PrintXY obj, txtbxEspDoc.Text, rPesq!espDOCx_BOL, rPesq!espDOCy_BOL
            P_PrintXY obj, txtbxAceite.Text, rPesq!aceitex_bol, rPesq!aceiteY_bol
            .Col = 3 'VALOR
            Printer.CurrentY = rPesq!valory_bol
            P_ImprX obj, .Text, rPesq!valorx_bol, True, True
            P_PrintXY obj, txtbxInstr.Text, rPesq!instrx_bol, rPesq!instry_bol
            P_PrintXY obj, txtbxSacado.Text, rPesq!sacadox_bol, rPesq!sacadoy_bol
            P_PrintXY obj, txtbxEnd.Text & " - " & txtbxBairro.Text & " - " & txtbxCidade.Text & " /" & txtbxUf.Text, rPesq!endx_bol, rPesq!endy_bol
            P_PrintXY obj, "Cep: " & txtbxCep.Text & " - " & lblCnpj.caption & ": " & txtbxCNPJ.Text, 5, 72
            Printer.NewPage
        Next
    End With
    Exit Sub
trata_erro:
    Printer.KillDoc
    P_MostraErro
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()

    Dim iCount As Integer
    If msflxGrd.rows = 1 Then
        MsgBox "Não há o que salvar!", vbExclamation
        Exit Sub
    End If
    With msflxGrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 0
            If .Text = "" Then
                MsgBox "É necessário informar o número de cada boleto!", vbExclamation
                msflxGrd.SetFocus
                Exit Sub
            End If
        Next
    End With
    If F_Vazio(txtbxSacado.Text) Then
        MsgBox "Não foi informado o nome do Sacado!", vbCritical
        txtbxSacado.SetFocus
        Exit Sub
    End If
    grava
End Sub
Sub grava()
    Dim strSQL As String
    Dim sSeq As String
    Dim iCount As Integer
    On Error GoTo trata_erro
    With msflxGrd
        For iCount = 1 To .rows - 1
            .Row = iCount:
            sSeq = MaxKeySQL("tab_boletos", "seq_bol", "", "0000#")
            strSQL = "INSERT INTO tab_boletos (seq_bol, num_bol, parc_bol,  venc_bol, valor_bol, seq_itd, data_bol) VALUES ( "
            strSQL = strSQL & "'" & sSeq & "',"
            .Col = 0
            strSQL = strSQL & "'" & .Text & "',"
            .Col = 1
            strSQL = strSQL & "'" & .Text & "',"
            .Col = 2
            strSQL = strSQL & "'" & .Text & "',"
            .Col = 3
            strSQL = strSQL & F_GravaDec(.Text) & ","
            .Col = 4
            strSQL = strSQL & "'" & .Text & "',"
            strSQL = strSQL & F_GravaDataTempo(Now) & ") "
            
            Banco.Execute strSQL
        Next
    End With
        MsgBox "Boleto(s) Gravado(s)", vbInformation
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    txtbxInstr.Text = F_BuscaUmCampo("instr_bol", "tab_bolconfig", "")
    txtbxLocal.Text = "Pagamento em qualquer Agência até o Vencimento"
    txtbxDataDoc.Text = Date
    txtbxDataProc.Text = Date
    With msflxGrd
        .Cols = 5
        .rows = 1
        .Col = 0: .Text = "Boleto"
        .Col = 1: .Text = "Parc"
        .Col = 2: .Text = "Vencimento"
        .Col = 3: .Text = "Valor"
        .Col = 4: .Text = "Lançamento"
        
        .ColWidth(0) = 1800
        .ColWidth(4) = 1000
        
    End With
End Sub

Private Sub msflxgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim ob As Object
        
        msflxgrd_DblClick
    End If
End Sub

Private Sub txtbxAceite_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxBoleto_GotFocus()
    P_SelecTexto txtbxBoleto
End Sub

Private Sub txtbxBoleto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnAdd_Click
        
    End If
End Sub

Private Sub txtbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
    
End Sub

Private Sub txtbxCEP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataProc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEspDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxParc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSacado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxVenc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub
