VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFecharVenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Concluir Venda:"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   ControlBox      =   0   'False
   Icon            =   "frmFecharVenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10365
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H0071CACE&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   5985
      TabIndex        =   36
      Top             =   5640
      Width           =   6015
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   195
         Left            =   0
         TabIndex        =   40
         Top             =   405
         Width           =   345
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   39
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblNome_Cli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome de Cliente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   38
         Top             =   75
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   0
         TabIndex        =   37
         Top             =   120
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00B0B0B0&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   9825
      TabIndex        =   23
      Top             =   0
      Width           =   9855
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr Venda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblvlrVenda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00 "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   34
         Tag             =   "!"
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1440
         TabIndex        =   33
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblDesconto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00 "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1440
         TabIndex        =   32
         Tag             =   "!"
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr Final:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   31
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00 "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2760
         TabIndex        =   30
         Tag             =   "!"
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Vista:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   29
         Top             =   120
         Width           =   675
      End
      Begin VB.Label lblAvIsta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00 "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3840
         TabIndex        =   28
         Tag             =   "!"
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Prazo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6480
         TabIndex        =   27
         Top             =   120
         Width           =   750
      End
      Begin VB.Label lblAPrazo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00 "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   6480
         TabIndex        =   26
         Tag             =   "!"
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblFinan 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00 "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5160
         TabIndex        =   25
         Tag             =   "!"
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financeira:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5160
         TabIndex        =   24
         Top             =   120
         Width           =   990
      End
   End
   Begin VB.Frame frAp 
      Enabled         =   0   'False
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9855
      Begin VB.Frame frForm 
         Caption         =   "&Pagamento a Prazo:"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9615
         Begin VB.OptionButton optnCheque 
            Caption         =   "&2 - Cheque"
            Height          =   195
            Left            =   6000
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton otpnCart 
            Caption         =   "&1 - Carteira"
            Height          =   195
            Left            =   4680
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cmbxFormas_Pag 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame frVal 
         Enabled         =   0   'False
         Height          =   3855
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   9615
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   8040
            Picture         =   "frmFecharVenda.frx":08CA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   22
            Top             =   3420
            Width           =   240
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   6600
            Picture         =   "frmFecharVenda.frx":0A14
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   21
            Top             =   3420
            Width           =   240
         End
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   5160
            Picture         =   "frmFecharVenda.frx":0B5E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   20
            Top             =   3420
            Width           =   240
         End
         Begin VB.ComboBox cmbxEspecie 
            Height          =   315
            ItemData        =   "frmFecharVenda.frx":0CA8
            Left            =   840
            List            =   "frmFecharVenda.frx":0CB2
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtbxValor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2520
            TabIndex        =   9
            Text            =   "0,00"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtbxData 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cmbxItem 
            Height          =   315
            ItemData        =   "frmFecharVenda.frx":0CD4
            Left            =   120
            List            =   "frmFecharVenda.frx":0CDB
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   375
            Left            =   6480
            TabIndex        =   14
            Top             =   3360
            Width           =   1335
         End
         Begin VB.CommandButton cmndbtnAdd 
            Caption         =   "&Adicionar"
            Height          =   375
            Left            =   7920
            TabIndex        =   15
            Top             =   3360
            Width           =   1455
         End
         Begin VB.CommandButton cmndbtnRem 
            Caption         =   "E&xcluir"
            Height          =   375
            Left            =   5040
            TabIndex        =   13
            Top             =   3360
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrd 
            Height          =   2535
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4471
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Valor:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2520
            TabIndex        =   8
            Top             =   120
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&Data:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3480
            TabIndex        =   10
            Top             =   120
            Width           =   390
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "&Item:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "E&spécie:"
            Height          =   255
            Left            =   840
            TabIndex        =   6
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Cancelar (F12)"
      Height          =   855
      Left            =   8280
      Picture         =   "frmFecharVenda.frx":0CE2
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Confirmar Venda (F7)"
      Height          =   855
      Left            =   6240
      Picture         =   "frmFecharVenda.frx":0E2C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5640
      Width           =   1695
   End
End
Attribute VB_Name = "frmFecharVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bActive As Boolean
Dim bBusca As Boolean
Public operacao As String 'venda ou troca
Function F_Soma()
    Dim iCount As Integer
    Dim cTotal As Currency
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 1
            cTotal = cTotal + CCur(.Text)
        Next
        
    End With
    F_Soma = cTotal
End Function
Private Sub cmbxEspecie_Click()
    Dim c As cheque
    Dim sDate As String
    Dim sNum As String
    Dim sBan As String
    Dim sValor As String
    Dim sAge As String
    Dim sCCO As String
    Dim sCPF As String
    
    If msflxgrd.Row = -1 Or msflxgrd.rows <= 1 Then
        
        Exit Sub
    End If
    
    
    If cmbxEspecie.ListIndex = 1 Then
        With msflxgrd
            
            msflxgrd.Col = 2
            msflxgrd.Text = cmbxEspecie.Text
            .Col = 0
            sDate = .Text
            .Col = 1
            sValor = .Text
            .Col = 3
            sNum = .Text
            .Col = 4
            sBan = .Text
            .Col = 5
            sAge = .Text
            .Col = 6
            sCCO = .Text
            .Col = 10
            sCPF = .Text
        End With
        c = F_DadosCheque(True, sDate, lblNome_cli.caption, sNum, sCCO, sAge, sBan, sValor, sCPF)
        If c.num_che = "" Then
            cmbxEspecie.ListIndex = 0
            Exit Sub
        End If
        
    End If
    
    With msflxgrd
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
            .Col = 2
            
        Else

            .Col = 2
            .Text = cmbxEspecie.Text
        End If
    End With
    
    
End Sub


Private Sub cmbxFormas_Pag_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim cAux As Currency
    Dim iCount As Integer
    Dim iAux As Integer
    Dim iDias As Integer
    Dim cTotal As Currency
    Dim cValor As Currency
    Dim iParcelas As Integer
    Dim cEntrada As Currency
    Dim cApr As Currency
    If bActive = False Then Exit Sub
    If cmbxFormas_Pag.ListIndex = -1 Then Exit Sub
    
    strSQL = "select * from tab_tipoven where cod_tip='" & Left(cmbxFormas_Pag.Text, 2) & "' "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    msflxgrd.rows = 1
    lblAprazo.caption = lblAprazo.Tag
    cApr = CCur(lblAprazo.caption)
    
    With rPesq
        
        If !juros_tip > 1 Then
            lblAprazo.caption = FormatVal(cApr * (1 + (!juros_tip / 100)))
            
        End If
        If !Entrada_tip = "S" Then
            iAux = 2
            cAux = (CCur(lblAprazo.caption) / !parcelas_tip)
            cAux = Round(cAux, 2)
            strSQL = Date & Chr(9) & FormatVal(cAux) & Chr(9) & "1 - Dinheiro"
            msflxgrd.AddItem strSQL
            cEntrada = Round(CCur(lblAprazo.caption) / !parcelas_tip, 2)
        Else
            
            
            iAux = 1
            cEntrada = 0
        End If
        iDias = !Intervalo_tip
        For iCount = iAux To !parcelas_tip
            cValor = Round(CCur((lblAprazo.caption) / !parcelas_tip), 2) '- cEntrada
            strSQL = DateAdd("D", iDias, Date) & Chr(9)
            strSQL = strSQL & FormatVal(cValor) & Chr(9) & IIf(optnCheque.Value, "2 - Cheque-Pré", "1 - Carteira")
            
            msflxgrd.AddItem strSQL
            iDias = iDias + !Intervalo_tip
            cTotal = cTotal + Round(cValor, 2)
        Next
            If cTotal + Round(cEntrada, 2) <> CCur(lblAprazo.caption) Then
                msflxgrd.Row = msflxgrd.rows - 1
                msflxgrd.Col = 1
                msflxgrd.Text = FormatVal(CCur(msflxgrd.Text) + CCur(lblAprazo.caption) - (cTotal + cEntrada))
            End If
           
        .Close
    End With
    If optnCheque.Value Then
        msflxgrd_DblClick
        
    End If
End Sub

Private Sub cmbxFormas_Pag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Command3_Click
    End If
End Sub

Private Sub cmbxItem_Click()
    On Error Resume Next
    msflxgrd.Row = Int(cmbxItem.Text)
    msflxgrd.Col = 2
    cmbxEspecie.ListIndex = Int(Left(msflxgrd.Text, 1)) - 1
    msflxgrd.Col = 0
    txtbxData.Text = msflxgrd.Text
    msflxgrd.Col = 1
    txtbxValor.Text = msflxgrd.Text

End Sub

Private Sub cmbxItem_GotFocus()
    Dim iCount As Integer
    cmbxItem.Clear
    For iCount = 1 To msflxgrd.rows - 1
        cmbxItem.AddItem Format(iCount, "0##")
    Next
    
End Sub

'Private Sub cmbxNat_Click()
'    If cmbxNat.ListIndex = 1 Then
'        Dim c As cheque
'        MSFlxGrd.Col = 0
'        c = F_DadosCheque(False, MSFlxGrd.Text, lblNome_Che.caption, lblNum_Che.caption)
'        If c.num_che <> "" Then
'            lblAge_Che.caption = c.age_che
'            lblBco.caption = c.cod_bco
'            lblCco_Che.caption = c.cco_che
'            lblNome_Che.caption = c.Nome_Che
'            lblNum_Che.caption = c.num_che
'            lblValor_Che.caption = c.Valor_che
'            lblVenc_Che.caption = c.Venc_che
'            frCheque.Visible = True
'        End If
'    End If
'End Sub

Private Sub cmndbtnEdit_Click()
    Dim sVen As String
    Dim sVal As String
    Dim sNum_Che As String
    Dim sAge As String
    Dim sBco As String
    Dim sTit As String
    Dim sCCO As String
    Dim iRow As Integer
    Dim p As parcela
    Dim c As cheque
    If msflxgrd.rows = 1 Then
        MsgBox "Não há o que ser Editado!", vbExclamation
    Else
        With msflxgrd
            .Col = 0
            sVen = .Text
            .Col = 1
            sVal = .Text
            .Col = 2
            If Left(.Text, 1) = "1" Then
                
                iRow = .Row
                p = F_DadosPar(sVen, sVal)
                
                If p.valor <> 0 Then
                    If (F_Soma - CCur(sVal) + p.valor) > CCur(lblAprazo.caption) Then
                        MsgBox "O Valor excede o valor total!", vbExclamation
                        Exit Sub
                    End If
                    
                    .Row = iRow
                    
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
            
                
            Else
                .Col = 0
                sVen = .Text
                .Col = 1
                sVal = FormatVal(.Text)
                .Col = 3
                sNum_Che = .Text
                .Col = 4
                sBco = .Text
                .Col = 5
                sAge = .Text
                .Col = 6
                sCCO = .Text
                
                .Col = 9
                sTit = .Text
                
                If sTit = "" Then
                    sTit = lblNome_cli.caption
                End If
                iRow = .Row
                c = F_DadosCheque(True, sVen, sTit, sNum_Che, sCCO, sAge, sBco, sVal)
                
                If c.Valor_che <> 0 Then
                    If (F_Soma - CCur(sVal) + p.valor) > CCur(lblAprazo.caption) Then
                        MsgBox "O Valor excede o valor total!", vbExclamation
                        Exit Sub
                    End If
                    
                    .Row = iRow
                    
                    .Col = 0: .Text = c.Venc_che
                    .Col = 1: .Text = FormatVal(c.Valor_che)
                    .Col = 2: .Text = cmbxEspecie.Text
                    .Col = 3: .Text = c.num_che
                    .Col = 4: .Text = c.cod_bco
                    .Col = 5: .Text = c.age_che
                    .Col = 6: .Text = c.cco_che
                    .Col = 7: .Text = FormatVal(c.Valor_che)
                    .Col = 8: .Text = c.Venc_che
                    .Col = 9: .Text = c.nome_che
                    
                
                
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
Private Sub cmndbtnAdd_Click()
    Dim p As parcela
    Dim strSQL As String
    If F_Soma = CCur(lblAprazo.caption) Then
        MsgBox "Você não pode adicionar parcelas quando o Valor total está fechado!", vbExclamation
        Exit Sub
    End If
    p = F_DadosPar
    If p.valor <> 0 Then
        strSQL = p.vencimento & Chr(9) & FormatVal(p.valor) & Chr(9) & "1 - Carteira"
        msflxgrd.AddItem strSQL
    Else
    End If
End Sub

Private Sub Command3_Click()
    Dim strSQL As String
    
    If CCur(lblAprazo.caption) > 0 Then
        If msflxgrd.rows = 1 Then
            MsgBox "Os valores A Prazo não estão definidos!", vbExclamation
            Exit Sub
        End If
    End If
    
    
        If msflxgrd.rows > 1 Then
            Dim sSeq_Doc As String
            Dim sSeq_Itd As String
            Dim sSeq_che As String
            Dim iCount As Integer
            Dim sVen As String, sVal As String, sEsp As String, sNum_Che As String, sBco As String
            Dim sAge As String, sCCO As String, sNome_che As String
            Dim sCPF As String
            
            If F_Soma <> CCur(lblAprazo.caption) Then
                MsgBox "Os valores não estão corretamente distribuídos!", vbExclamation
                Exit Sub
           
            End If
            sSeq_Doc = F_GravaDoc(operacao & " " & Me.Tag & "-" & lblNome_cli.caption, Me.Tag, _
                lblAprazo.caption, CStr(Date), lblNome_cli.Tag, "R", Me.Tag, msflxgrd.rows - 1, strPlaVendasEfetuadas)
            If sSeq_Doc <> "" Then
                With msflxgrd
                    For iCount = 1 To .rows - 1
                        .Row = iCount
                        msflxgrd.Col = 0
                        sVen = msflxgrd.Text
                        msflxgrd.Col = 1
                        sVal = msflxgrd.Text
                        msflxgrd.Col = 2
                        sEsp = Left(msflxgrd.Text, 1)
                        msflxgrd.Col = 3
                        sNum_Che = msflxgrd.Text
                        msflxgrd.Col = 4
                        sBco = msflxgrd.Text
                        msflxgrd.Col = 5
                        sAge = msflxgrd.Text
                        msflxgrd.Col = 6
                        sCCO = msflxgrd.Text
                        msflxgrd.Col = 9
                        sNome_che = msflxgrd.Text
                        msflxgrd.Col = 10
                        sCPF = msflxgrd.Text
                        If sEsp = "1" Or sEsp = "" Then
                            sSeq_Itd = F_GravaItDoc(sSeq_Doc, sVal, sVen, iCount)
                        ElseIf sEsp = "2" Then
                            sSeq_che = F_GravaCh(sSeq_Doc, sNum_Che, sBco, sAge, sCCO, _
                            sNome_che, sVal, sVen, lblNome_cli.Tag, "R", _
                            "", "", str(iCount), "", sCPF, Time)
                        End If
                        
                    Next
                End With
            End If
            
        

        msflxgrd.rows = 1
            
            
        
        
            
        P_BaixaVenda Me.Tag
        strSQL = "update tab_vendas set situacao_ven='4', "
        strSQL = strSQL & "dataefet_ven=" & F_GravaData(Date) & ","
        
        strSQL = strSQL & "cod_tip='" & Left(cmbxFormas_Pag.Text, 2) & "' "
        
        strSQL = strSQL & " where num_ven='" & Me.Tag & "' "
        banco.Execute strSQL

        'gavar caixa
        strSQL = F_GravaCai("C", Date, lblAprazo.caption, "Referente a  " & operacao & _
             " " & Me.Tag, Me.Tag, IIf(optnCheque.Value, "CHP", "CAR"), "01", 0, 0, strCod_Cai, strCod_Tur)
            If strSQL = "" Then
                MsgBox "Erro!", vbExclamation
                
            End If
 
 

            
            Me.Hide
            Exit Sub
        ElseIf LCase(operacao) = "troca" Then
            Me.Hide
            Exit Sub
        End If
    
    Unload Me
End Sub


Private Sub Command4_Click()
    Me.Tag = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    Dim sTipo_Ven As String
    If lblNome_cli.Tag = "00000" Then 'venda a consumidor
        'optnPraz.Enabled = False
        'opntPerson.Enabled = False
    End If
    MudaAparencia Me
    
    
    If cmbxFormas_Pag.Tag <> "" Then
        
        P_LocalizaOperacao cmbxFormas_Pag.Tag, cmbxFormas_Pag
        bActive = True
        
        sTipo_Ven = F_BuscaUmCampo("tipo_ven", "tab_venfec", "num_ven='" & Me.Tag & "'")
        If sTipo_Ven = "D" Then
            otpnCart.Value = True
        Else
            optnCheque.Value = True
        End If
        cmbxFormas_Pag_Click
    End If
    bActive = True
    optnCheque.Enabled = False
    otpnCart.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        Command3_Click
    End If
    If KeyCode = vbKeyF12 Then
        Command4_Click
    End If
    
End Sub

Private Sub Form_Load()
    bActive = False
    msflxgrd.rows = 1
    msflxgrd.Cols = 11
    msflxgrd.Col = 0
    msflxgrd.Text = "Vencimento"
    msflxgrd.Col = 1
    msflxgrd.Text = "Valor"
    msflxgrd.Col = 2
    msflxgrd.Text = "Espécie"
    msflxgrd.Col = 3
    msflxgrd.Text = "Cheque Nº"
    msflxgrd.Col = 4
    msflxgrd.Text = "Banco"
    msflxgrd.Col = 5
    msflxgrd.Text = "Agencia"
    msflxgrd.Col = 6
    msflxgrd.Text = "Conta C."
    msflxgrd.Col = 7
    msflxgrd.Text = "Valor"
    msflxgrd.Col = 8
    msflxgrd.Text = "Vencimento"
    msflxgrd.Col = 9
    msflxgrd.Text = "Titular"
    msflxgrd.ColAlignment(2) = 1
    msflxgrd.Col = 10
    msflxgrd.Text = "CPF"
    
    
    
    P_PreencheFormasPag cmbxFormas_Pag, ""
    
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub msflxgrd_DblClick()
    On Error Resume Next
    cmbxItem_GotFocus
    cmbxItem.Text = Format(msflxgrd.Row, "00#")
    cmbxEspecie.SetFocus
End Sub

Private Sub opntPerson_Click()

End Sub



Sub Valores()
'    If optnAv.Value = True Then
'        lblDesconto.caption = txtbxDesconto.Text
'        lblAvIsta.caption = FormatVal(CCur(lblvlrVenda.caption) - CCur(lblDesconto.caption))
'        lblTotal.caption = lblAvIsta.caption
'        lblAPrazo.caption = "0,00"
'    ElseIf optnPraz.Value = True Then
'        lblDesconto.caption = txtbxDesconto.Text
'        lblAPrazo.caption = FormatVal(CCur(lblvlrVenda.caption) - CCur(lblDesconto.caption))
'        lblTotal.caption = lblAPrazo.caption
'        lblAvIsta.caption = "0,00"
'    Else
'
'    End If
End Sub




Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        msflxgrd_DblClick
    End If
    If KeyCode = vbKeyReturn Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub msflxgrd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub optnCheque_Click()
    Dim iCount As Integer
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 2
            .Text = "2 - Cheque Pré"
            
        Next
        cmbxEspecie.ListIndex = 1
    End With
End Sub

Private Sub otpnCart_Click()
    Dim iCount As Integer
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Col = 2
            .Text = "1 - Carteira"
            
        Next
        cmbxEspecie.ListIndex = 0
    End With
End Sub

