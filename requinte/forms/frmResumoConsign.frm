VERSION 5.00
Begin VB.Form frmResumoConsign 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumo das Consignações"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnFecharMov 
      Caption         =   "Fechar Movimento"
      Height          =   975
      Left            =   5880
      Picture         =   "frmResumoConsign.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Buscar"
      Height          =   685
      Left            =   240
      Picture         =   "frmResumoConsign.frx":04A8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   685
      Left            =   4080
      Picture         =   "frmResumoConsign.frx":0832
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   685
      Left            =   8040
      Picture         =   "frmResumoConsign.frx":0C3F
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnGraf 
      Caption         =   "&Gráfico"
      Height          =   685
      Left            =   6000
      Picture         =   "frmResumoConsign.frx":0DAA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnExpor 
      Caption         =   "&Exportar"
      Height          =   685
      Left            =   2160
      Picture         =   "frmResumoConsign.frx":0EF4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox txtbxData 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   240
      ScaleHeight     =   5145
      ScaleWidth      =   4065
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Width           =   1995
      End
      Begin VB.Label lblTotalAcerto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   2100
         TabIndex        =   25
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dinheiro(DIN):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lblAcertoDin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   23
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Duplicata(DUP):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lblAcertoDup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   21
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cartão de Crédito(CRT):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lblAcertoCrt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   19
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lblAcertoCHV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   18
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheque Pré-Datado(CHP):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheques a Vista (CHV):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label lblAcertoCHP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   15
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblAcertoCrd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   14
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cartão de Débito(CRD):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label lblDevolucoes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   12
         Top             =   3720
         Width           =   1155
      End
      Begin VB.Label lblConsign 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   11
         Top             =   3240
         Width           =   1155
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Devoluções:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   1995
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transferências:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acertos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   345
   End
End
Attribute VB_Name = "frmResumoConsign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    P_Status "Buscando Dados..."
    busca
    P_Status ""
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnFecharMov_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    If MsgBox("Deseja realmente fechar o movimento de consignado de " & txtbxData.Text & "?", vbInformation + vbQuestion) = vbOK Then
        Exit Sub
    End If
    strSQL = "select * from tab_fechaconsign where data_fec=" & F_GravaDec(txtbxData.Text)
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then
            .Close
            Exit Sub
        End If
        Do While Not .EOF
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub

Private Sub Form_Load()
    txtbxData.Text = Date
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    lblAcertoCHP.caption = "0,00"
    lblAcertoDin.caption = "0,00"
    lblAcertoCHV.caption = "0,00"
    lblAcertoCrd.caption = "0,00"
    lblAcertoCrt.caption = "0,00"
    lblAcertoDup.caption = "0,00"
    
    strSQL = " select sum(ta.valor_ace) as valor,tp.tipo_cai "
    strSQL = strSQL & " from tab_acertos as ta, tab_formaspag as tp  "
    strSQL = strSQL & " where ta.cod_for=tp.cod_for "
    strSQL = strSQL & " and data_Ace=" & F_GravaData(txtbxData.Text)
    
    
    strSQL = strSQL & " group by tp.tipo_cai "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            If !tipo_cai = "DIN" Then
                lblAcertoDin.caption = FormatVal(!valor)
            ElseIf !tipo_cai = "DUP" Then
                lblAcertoDup.caption = FormatVal(!valor)
            ElseIf !tipo_cai = "CRT" Then
                lblAcertoCrt.caption = FormatVal(!valor)
            ElseIf !tipo_cai = "CRD" Then
                lblAcertoCrd.caption = FormatVal(!valor)
            ElseIf !tipo_cai = "CHP" Then
                lblAcertoCHP.caption = FormatVal(!valor)
            ElseIf !tipo_cai = "CHV" Then
                lblAcertoCHV.caption = FormatVal(!valor)
            
            Else
                MsgBox ""
            End If
            
            cTotal = cTotal + !valor
            .MoveNext
        Loop
        .Close
    End With
    lblTotalAcerto.caption = FormatVal(cTotal)
    
    strSQL = "select sum(qtd_con*unit_con) as total from tab_consignacoes "
    strSQL = strSQL & " where left(data_con,10)=" & F_GravaData(txtbxData.Text)
    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            lblConsign.caption = FormatVal(rPesq!total)
        Else
            lblConsign.caption = "0,00"
        End If
        .Close
    End With
    
    strSQL = "select sum(ti.qtd_dev*tc.unit_con) as total from tab_devconsign as td, "
    strSQL = strSQL & "tab_itdevconsign as ti,tab_consignacoes as tc "
    strSQL = strSQL & " where ti.seq_dev=td.seq_dev"
    strSQL = strSQL & " and ti.seq_con=tc.seq_con "
    strSQL = strSQL & " and left(data_con,10)=" & F_GravaData(txtbxData.Text)
    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            lblDevolucoes.caption = FormatVal(rPesq!total)
        Else
            lblDevolucoes.caption = "0,00"
        End If
        .Close
    End With
    
End Sub
Sub busca1()
    Dim strSQL As String
    Dim rPesq As Recordset
    
    strSQL = "select sum(tf.valor_fin)as total from tab_financeiro as tf, tab_acertofin as taf, tab_acertos as ta "
    strSQL = strSQL & " where tf.seq_fin=taf.seq_fin  "
    strSQL = strSQL & " and ta.seq_ace=taf.seq_ace "
    strSQL = strSQL & " and data_ace=" & F_GravaData(txtbxData.Text)
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            lblDin.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    strSQL = "select sum(tvf.valor_fin)as total from tab_financeiras as tf, tab_venfin as tvf , tab_acertos as ta "
    strSQL = strSQL & " where tf.cod_fin=tvf.cod_fin"
    strSQL = strSQL & " and ta.seq_ace=tvf.num_ven"
    strSQL = strSQL & " and tf.operacao_fin='C' "
    strSQL = strSQL & " and data_ace=" & F_GravaData(txtbxData.Text)
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            lblCC.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    strSQL = "select sum(tvf.valor_fin)as total from tab_financeiras as tf, tab_venfin as tvf , tab_acertos as ta "
    strSQL = strSQL & " where tf.cod_fin=tvf.cod_fin"
    strSQL = strSQL & " and ta.seq_ace=tvf.num_ven"
    strSQL = strSQL & " and tf.operacao_fin='D' "
    strSQL = strSQL & " and data_ace=" & F_GravaData(txtbxData.Text)
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            lblCD.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    strSQL = "select sum(valor_che)as total from tab_doc as td, tab_cheques as tc , tab_acertos as ta "
    strSQL = strSQL & " where td.seq_doc=tc.seq_doc "
    
    strSQL = strSQL & " and ta.seq_ace=td.num_ven"
    strSQL = strSQL & " and td.esp_doc='A'"
    strSQL = strSQL & " and data_ace=" & F_GravaData(txtbxData.Text)
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            lblChv.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    
    strSQL = "select sum(valor_itd)as total from tab_doc as td, tab_itdoc as tc , tab_acertos as ta "
    strSQL = strSQL & " where td.seq_doc=tc.seq_doc "
    
    strSQL = strSQL & " and ta.seq_ace=td.num_ven"
    strSQL = strSQL & " and td.esp_doc='A'"
    strSQL = strSQL & " and data_ace=" & F_GravaData(txtbxData.Text)
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            lblAprazo.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    lblTotal.caption = FormatVal(CCur(lblDin.caption) + CCur(lblChv.caption) + CCur(lblCC.caption) + CCur(lblCD.caption) + CCur(lblAprazo.caption))
    
End Sub

Private Sub lblCC_Click()
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub
