VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmClass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Classificação"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "frmClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtbxValorBaixa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   1665
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnLancar 
         Caption         =   ">> &Lançar"
         Height          =   405
         Left            =   4800
         TabIndex        =   5
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "<< E&xcluir"
         Height          =   405
         Left            =   4800
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox cmbxConta 
         BackColor       =   &H00008000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmClass.frx":014A
         Left            =   120
         List            =   "frmClass.frx":0154
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "!"
         Top             =   1650
         Width           =   3375
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1931
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         FillStyle       =   1
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label lblDistr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distribuidos:"
         Height          =   195
         Left            =   4800
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor a Baixar:"
         Height          =   195
         Left            =   3600
         TabIndex        =   3
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&onta:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   465
      End
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar Baixa"
      Height          =   735
      Left            =   6000
      Picture         =   "frmClass.frx":0172
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Cancela&r"
      Height          =   735
      Left            =   6000
      Picture         =   "frmClass.frx":02BC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   6000
      TabIndex        =   12
      Top             =   1680
      Width           =   405
   End
End
Attribute VB_Name = "frmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function F_Restante() As Currency
    F_Restante = CCur(lblTOtal.caption) - CCur(lblDistr.caption)
End Function

Private Sub cmndbtnConfirma_Click()
    'ReDim class(msflxgrd.Rows - 1)
    Dim strSQL As String
    Dim iCount As Integer
    Dim c As Currency
    c = F_Restante
    If c > 0 Then
        MsgBox "O Valor não foi completamente distribuído!", vbExclamation
        Exit Sub
    ElseIf c < 0 Then
        MsgBox "O Valor distribuído excede o total!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "delete from tab_class where seq_fin ='" & Me.tag & "' "
    conn.Execute strSQL
    strSQL = ""
    For iCount = 1 To msflxgrd.rows - 1
        With msflxgrd
            strSQL = ""
            .Row = iCount
            strSQL = strSQL & "insert into tab_class(seq_fin,cod_pla,valor_cla,cod_usu)values("
            strSQL = strSQL & "'" & Me.tag & "', "
            .col = 0
            strSQL = strSQL & "'" & F_PlanoConta(False, .Text, cmndbtnConfirma.tag) & "', "
            .col = 1
            strSQL = strSQL & F_GravaDec(.Text) & ","
            strSQL = strSQL & "'" & strCod_Usu & "') "
            conn.Execute strSQL
        End With
    Next
    Me.Hide
End Sub

Private Sub cmndbtnExcl_Click()
    Dim sVal As Currency
    If msflxgrd.rows = 1 Then
        MsgBox "Não há o que ser excluído!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma exlcusão do Lançamento?", vbYesNo + vbCritical) = vbYes Then
        msflxgrd.col = 1
        sVal = CCur(msflxgrd.Text)
        lblDistr.caption = FormatVal(CCur(lblDistr.caption) - sVal)
        If msflxgrd.rows = 2 Then
            msflxgrd.rows = 1
        Else
            msflxgrd.RemoveItem msflxgrd.Row
        End If
        
    End If
End Sub

Private Sub cmndbtnLancar_Click()
    Dim strSQL As String
    If cmbxConta.Text = "" Then
        MsgBox "Informe a Conta!", vbExclamation
        cmbxConta.SetFocus
        Exit Sub
    End If
    If txtbxValorBaixa.Text = "0,00" Then
        MsgBox "Informe um valor válido!", vbExclamation
        txtbxValorBaixa.SetFocus
        Exit Sub
    End If
    If F_InFlxgrd(msflxgrd, cmbxConta.Text, 0) <> -1 Then
        MsgBox "Já existe um lançamento para esta Conta!", vbExclamation
        Exit Sub
    End If
    strSQL = cmbxConta.Text & Chr(9)
    strSQL = strSQL & txtbxValorBaixa.Text
    msflxgrd.AddItem strSQL
    lblDistr.caption = FormatVal(CCur(lblDistr.caption) + CCur(txtbxValorBaixa.Text))
    
End Sub

Private Sub cmndbtnSair_Click()
    If F_Restante < 0 Then
        MsgBox "O Valor distribuído excede o total!", vbExclamation
        Exit Sub
    End If
    Me.tag = ""
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    P_PreenchePlanoContas cmbxConta, IIf(cmndbtnConfirma.tag <> "D", True, False)
    With msflxgrd
        .FixedCols = 0
        .FixedRows = 1
        .rows = 1
        .Cols = 2
        .ColWidth(0) = 2500
        .Text = "Classificação"
        .ColWidth(1) = 800
        .Text = "Valor"
        
    End With
    strSQL = "select * from tab_class where seq_fin='" & Me.tag & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = F_PlanoConta(True, !cod_pla, cmndbtnConfirma.tag) & _
            Chr(9) & FormatVal(!valor_cla)
            msflxgrd.AddItem strSQL
            cTotal = cTotal + !valor_cla
            .MoveNext
        Loop
        lblDistr.caption = FormatVal(cTotal)
        .Close
    End With
    'MudaAparencia Me
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmndbtnExcl_Click
    End If
End Sub

Private Sub txtbxValorBaixa_GotFocus()
    txtbxValorBaixa.Text = F_Restante
    ganhaFoco txtbxValorBaixa
End Sub

Private Sub txtbxValorBaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValorBaixa_LostFocus()
    PerdeFoco txtbxValorBaixa
End Sub
