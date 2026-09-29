VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmValoresForn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de Estoques das Fornadas:"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmValoresForn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frControle 
      Caption         =   "Distribuição das Fornadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Data dtProd 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtbxQtd_Itf 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnOk 
         Caption         =   "&Adicionar"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtbxForn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin MSDBGrid.DBGrid dbgrdProd 
         Bindings        =   "frmValoresForn.frx":08CA
         Height          =   1695
         Left            =   120
         OleObjectBlob   =   "frmValoresForn.frx":08DF
         TabIndex        =   7
         Top             =   840
         Width           =   7335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fornada:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "F&echar"
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lançado:"
      Height          =   195
      Left            =   4440
      TabIndex        =   18
      Top             =   2760
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Lançar:"
      Height          =   195
      Left            =   3000
      TabIndex        =   17
      Top             =   2760
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forn.:"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   3000
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Série"
      Height          =   195
      Left            =   1680
      TabIndex        =   15
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota:"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   390
   End
   Begin VB.Label lblDistr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   195
      Left            =   5160
      TabIndex        =   13
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label lblQtd_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   195
      Left            =   3720
      TabIndex        =   12
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label lblFornecedor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      Height          =   195
      Left            =   720
      TabIndex        =   11
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label lblSerie 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Série"
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label lblNota 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota:"
      Height          =   195
      Left            =   720
      TabIndex        =   9
      Top             =   2760
      Width           =   390
   End
End
Attribute VB_Name = "frmValoresForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnConfirma_Click()
    Unload Me
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    
    With dtProd.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que excluir!", vbExclamation
            Exit Sub
        End If
        If MsgBox("confirma exclusão do lançamento?", vbYesNo + vbCritical) = vbYes Then
            strSQL = "Delete from tab_itforn "
            strSQL = strSQL & " where num_not='" & lblNota.caption & "' "
            strSQL = strSQL & " and cod_for = '" & lblFornecedor.Tag & "' "
            strSQL = strSQL & " and serie_not='" & lblSerie.caption & "' "
            strSQL = strSQL & " and cod_pro = '" & Me.Tag & "' "
            strSQL = strSQL & " and fornada_itf = '" & !fornada_itf & "' "
            banco.Execute strSQL
            Busca
        End If
        
    End With
End Sub

Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    If CCur(txtbxQtd_Itf.Text) + CCur(lblDistr.caption) > CCur(lblQtd_Ite) Then
        MsgBox "Valor está maior que o permitido!", vbExclamation
        txtbxQtd_Itf.SetFocus
        Exit Sub
    End If
    If F_Vazio(txtbxForn.Text) Then
        MsgBox "Informe a Fornada!", vbExclamation
        txtbxForn.SetFocus
        Exit Sub
    End If
    If txtbxQtd_Itf.Text = "0,00" Then
        MsgBox "Informe uma quantidade válida!", vbExclamation
        txtbxQtd_Itf.SetFocus
        Exit Sub
    End If
    If F_BuscaUmCampo("Fornada_Itf", "tab_itforn", "num_not='" & lblNota.caption & "'" _
        & " and serie_not='" & lblSerie.caption & "' and cod_for ='" & _
        lblFornecedor.Tag & "' and cod_pro='" & Me.Tag & "' and fornada_itf ='" & txtbxForn.Text & "'") <> "" Then
        MsgBox "Já existe esta fornada para este produto!", vbExclamation
        txtbxForn.SetFocus
        Exit Sub
    End If
    
    
    If MsgBox("Confirma Inclusão?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = "insert into tab_itforn(cod_for,num_not,serie_not,"
        strSQL = strSQL & "cod_pro,fornada_itf,qtd_itf)values("
        strSQL = strSQL & "'" & lblFornecedor.Tag & "',"
        strSQL = strSQL & "'" & lblNota.caption & "',"
        strSQL = strSQL & "'" & lblSerie.caption & "',"
        strSQL = strSQL & "'" & Me.Tag & "',"
        strSQL = strSQL & "'" & txtbxForn.Text & "',"
        strSQL = strSQL & f_GravaDec(txtbxQtd_Itf.Text) & ") "
        banco.Execute strSQL
        
    End If
    Busca
    txtbxQtd_Itf.Text = "0,00"
    txtbxForn.SetFocus
    
End Sub
Sub Busca()
    Dim strSQL As String
    Dim cTotal As Currency
    strSQL = "select * from tab_itforn "
    strSQL = strSQL & " where num_not='" & lblNota.caption & "' "
    strSQL = strSQL & " and cod_for = '" & lblFornecedor.Tag & "' "
    strSQL = strSQL & " and serie_not='" & lblSerie.caption & "' "
    strSQL = strSQL & " and cod_pro = '" & Me.Tag & "' "
    
    
    
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    With dtProd.Recordset
        Do While Not .EOF
            
            cTotal = cTotal + !qtd_itf
            .MoveNext
        Loop
    End With
    lblDistr.caption = FormatVal(cTotal)
    
End Sub

Private Sub dbgrdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    Busca
End Sub

Private Sub Form_Load()
    dtProd.Connect = banco.Connect
End Sub

Private Sub txtbxForn_GotFocus()
    P_SelecTexto txtbxForn
End Sub

Private Sub txtbxForn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxForn_LostFocus()
    P_AjustaTexto txtbxForn
End Sub

Private Sub txtbxQtd_Itf_GotFocus()
    P_SelecTexto txtbxQtd_Itf
End Sub

Private Sub txtbxQtd_Itf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_Itf_LostFocus()
    P_ValorLostFocus txtbxQtd_Itf
End Sub
