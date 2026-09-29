VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiMovFin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimento de Financeiras"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   Icon            =   "frmEdiMovFin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9150
   Begin VB.Frame frControle 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.CheckBox chkAberto 
         Caption         =   "Apenas Parc. em Aberto"
         Height          =   195
         Left            =   2400
         TabIndex        =   16
         Top             =   5280
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Frame Frame5 
         Caption         =   "&Ordenar Por:"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
         Begin VB.ComboBox cmbxOrder 
            Height          =   315
            ItemData        =   "frmEdiMovFin.frx":0442
            Left            =   120
            List            =   "frmEdiMovFin.frx":0452
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "&Financeira:"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2175
         Begin VB.ComboBox cmbxFin 
            Height          =   315
            ItemData        =   "frmEdiMovFin.frx":0478
            Left            =   120
            List            =   "frmEdiMovFin.frx":0482
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2175
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxdataFin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1080
            TabIndex        =   6
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.Data dtNot 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   3495
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   2175
         Begin VB.CommandButton cmndbtnBuscar 
            Caption         =   "&Pesquisar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiMovFin.frx":04B6
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   160
            Width           =   1980
         End
         Begin VB.CommandButton cmndbtnBaixar 
            Caption         =   "&Baixar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiMovFin.frx":0840
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   795
            Width           =   1980
         End
         Begin VB.CommandButton cmndbtnImpr 
            Caption         =   "&Imprimir"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiMovFin.frx":0BCA
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1440
            Width           =   1980
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiMovFin.frx":0F54
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2085
            Width           =   1980
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiMovFin.frx":109E
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2745
            Width           =   1980
         End
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEdiMovFin.frx":11E8
         Height          =   4935
         Left            =   2400
         OleObjectBlob   =   "frmEdiMovFin.frx":11FC
         TabIndex        =   17
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total: 0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7680
         TabIndex        =   18
         Top             =   5280
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmEdiMovFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAberto_Click()
    If chkAberto.Value = 1 Then
        cmndbtnBaixar.Enabled = True
    Else
        cmndbtnBaixar.Enabled = False
    End If
End Sub

Private Sub cmbxFin_Click()
    cmbxFin.Tag = F_BuscaUmCampo("cod_fin", "tab_financeiras", "nome_fin='" & _
    cmbxFin.Text & "'")
End Sub

Private Sub cmbxFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxFin.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBaixar_Click()
    Dim iCount As Integer
    Dim strSQL As String
    If IsLoad("frmListaDoc") Then
        If frmListaDoc.Tag = Me.Name Then
            If MsgBox("Confirma baixa de todas as parcelas listadas no valor de " _
                & frmListaDoc.lblTotal.caption & "?", vbYesNo + vbQuestion) = vbYes Then
                With frmListaDoc.msflxgrd
                    For iCount = 0 To .Rows - 1
                        .Row = iCount
                        .Col = 0
                        strSQL = "update tab_venfin set baixa_ven=" & F_GravaData(Date)
                        strSQL = strSQL & ",status_ven=2, "
                        strSQL = strSQL & "cod_usu='" & strCod_Usu & "'  "
                        strSQL = strSQL & " where num_ven='" & .Text & "' "
                        banco.Execute strSQL
                        
                    Next
                Unload frmListaDoc
                dtNot.Refresh
                End With
            End If
        End If
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    strSQL = "select tf.nome_fin,tvf.*,tc.nome_cli from tab_financeiras as tf, "
    strSQL = strSQL & " tab_venfin as tvf,tab_clientes as tc where tvf.cod_fin=tf.cod_fin "
    strSQL = strSQL & " and tc.cod_cli=tvf.cod_cli "
    strSQL = strSQL & " and tvf.data_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tvf.data_ven<=" & F_GravaData(txtbxdataFin.Text)
    If chkAberto.Value = 1 Then
        strSQL = strSQL & " and tvf.status_ven=1 "
    End If
    If cmbxFin.Tag <> "" Then
        strSQL = strSQL & " and tvf.cod_fin='" & cmbxFin.Tag & "' "
    End If
    If cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by nome_fin, data_ven "
    ElseIf cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by  data_ven,nome_fin,num_ven "
    ElseIf cmbxOrder.ListIndex = 2 Then
        strSQL = strSQL & " order by  nome_cli,data_ven,nome_fin,num_ven "
    Else
        strSQL = strSQL & " order by  num_ven "
    End If
    dtNot.RecordSource = strSQL
    dtNot.Refresh
    If dtNot.Recordset.RecordCount > 0 Then
        cmndbtnBaixar.Enabled = True
        cmndbtnEdit.Enabled = True
        cmndbtnImpr.Enabled = True
    Else
        cmndbtnBaixar.Enabled = False
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
    End If
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    With dtNot.Recordset
        Do While Not .EOF
            cTotal = cTotal + !valor_ven
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    lblTotal.caption = "Total " & FormatVal(cTotal)
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_DblClick()
    With dtNot.Recordset
        If !status_ven <> 1 Then
            Exit Sub
        End If
        If IsLoad("frmListaDoc") Then
            If frmListaDoc.Tag <> Me.Name Then
                MsgBox "A lista está sendo usada por outro formulário!", vbExclamation
                Exit Sub
            End If
            If F_InFlxgrd(frmListaDoc.msflxgrd, !Num_Ven, 0) = -1 Then
                frmListaDoc.msflxgrd.AddItem !Num_Ven & Chr(9) & !valor_ven
            End If
            frmListaDoc.soma
        Else
            frmListaDoc.Tag = Me.Name
            frmListaDoc.caption = "Documentos de Financeiras"
            frmListaDoc.msflxgrd.Cols = 2
            frmListaDoc.msflxgrd.Rows = 0
            frmListaDoc.msflxgrd.AddItem !Num_Ven & Chr(9) & !valor_ven
            frmListaDoc.Show 0, mdiFrm
            frmListaDoc.soma
        End If
    
    End With
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Load()
    dtNot.Connect = banco.Connect
    P_CentraForm Me
    P_PreencheFinanceiras cmbxFin
    cmbxOrder.ListIndex = 0
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDAtaFIn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
