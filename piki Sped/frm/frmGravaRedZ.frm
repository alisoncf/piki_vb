VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGravaRedZ 
   Caption         =   "Gravar Redução Z"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10650
   Icon            =   "frmGravaRedZ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGrid 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Text            =   "0,00"
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2535
      Left            =   5160
      TabIndex        =   17
      Top             =   360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4471
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   10620
      TabIndex        =   13
      Top             =   3315
      Width           =   10650
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   9120
         Picture         =   "frmGravaRedZ.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Salvar"
         Height          =   735
         Left            =   7560
         Picture         =   "frmGravaRedZ.frx":0E58
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.TextBox txtVlBrt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   14
      Text            =   "0,00"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtGT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Text            =   "0,00"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtCOO 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtCRZ 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtCRO 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtCaixa 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Venda Bruta"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grande Total Final"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contador de Cupom Fiscal"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   1860
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contador de Reduções Z:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1845
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contador de Reinício de Operação:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   390
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   435
   End
End
Attribute VB_Name = "frmGravaRedZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmdNovo_Click()
    If txtCaixa.Text = "" Then
        MsgBox "Informe o caixa!", vbExclamation
        txtCaixa.SetFocus
        Exit Sub
    End If
    If txtData.Text = "" Then
        MsgBox "Informe a data!", vbExclamation
        txtData.SetFocus
        Exit Sub
    End If
    If txtCRO.Text = "" Then
        MsgBox "Informe o CRO!", vbExclamation
        txtCRO.SetFocus
        Exit Sub
    End If
    If txtCRZ.Text = "" Then
        MsgBox "Informe o CRZ!", vbExclamation
        txtCRZ.SetFocus
        Exit Sub
    End If
    If txtCOO.Text = "" Then
        MsgBox "Informe o COO!", vbExclamation
        txtCOO.SetFocus
        Exit Sub
    End If
    If txtGT.Text = "" Then
        MsgBox "Informe o GT!", vbExclamation
        txtGT.SetFocus
        Exit Sub
    End If
    If txtVlBrt.Text = "" Then
        MsgBox "Informe o VLBRT!", vbExclamation
        txtVlBrt.SetFocus
        Exit Sub
    End If
    
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub gravaTot()
    Dim i As Integer
    Dim s1 As String
    Dim s2 As String
    strSQL = "delete from sped_redz_tot where id_red=" & Me.tag
    conn.Execute strSQL
    With grid
        For i = 1 To .rows - 1
            .Row = i
            .col = 0
            s1 = .Text
            .col = 2
            s2 = .Text
            If GravaSpedRedZTot(Me.tag, s2, s1) Then
            End If
            
        Next
    End With
End Sub
Sub grava()
    If GravaSpedRedZ(txtCaixa.Text, txtData.Text, txtCRO.Text, txtCRZ.Text, txtCOO.Text, txtGT.Text, txtVlBrt.Text) Then
        Me.tag = MaxKeySQL("sped_redz", "id", "caixa='" & txtCaixa.Text & "'", "")
        gravaTot
        
        MsgBox "Gravado com sucesso", vbInformation
        txtData.SetFocus
    Else
        MsgBox "Erro na gravação!", vbExclamation
    
    End If

End Sub
Sub altera()
    If AlteraSpedRedZ(Me.tag, txtCaixa.Text, txtData.Text, txtCRO.Text, txtCRZ.Text, txtCOO.Text, txtGT.Text, txtVlBrt.Text) Then
        gravaTot
        MsgBox "Alterado com sucesso", vbInformation
        Unload Me
    Else
        MsgBox "Erro na gravação!", vbExclamation
    
    End If

End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    bActive = True
    
    If Me.tag <> "" Then
        carrega
    End If
    
    
End Sub
Sub buscaTot()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
        
    
        
        
    strSQL = "select  ti.*,valor from  tab_impostos as ti right join "
    strSQL = strSQL & " sped_redz_tot as r "
    strSQL = strSQL & " on ti.cod_imp=r.cod_imp"
    strSQL = strSQL & " where id_red=" & F_GravaDec(Me.tag)
    strSQL = strSQL & " order by ti.cod_imp "
    abreObjetoAdo rPesq, strSQL
    
    If rPesq.RecordCount = 0 Then
    
        strSQL = "select  ti.*, 0 as valor from tab_impostos as ti order by cod_imp "
        abreObjetoAdo rPesq, strSQL
    End If
    With grid
        .Cols = 3
        .rows = 1
        .Row = 0
        .col = (0): .Text = "Código"
        .col = (1): .Text = "Imposto"
        .col = (2): .Text = "Total"
        
        
    End With
    
    With rPesq
        Do While Not .EOF
            If !ecf_imp = "II" Then
                strSQL = "II"
            ElseIf !ecf_imp = "FF" Then
                strSQL = "FF"
            
            End If
            
            strSQL = Format(!perc_imp, "##00") & "00"
            
            
            
            grid.AddItem !cod_imp & Chr(9) & strSQL & Chr(9) & FormatVal(!valor)
            
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select * from sped_redz where id=" & Me.tag
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            txtCaixa.Text = !caixa
            txtCOO.Text = !num_coo_fin
            txtCRO.Text = !cro
            txtCRZ.Text = !crz
            txtData.Text = !dt_doc
            txtGT.Text = !gt_fin
            txtVlBrt.Text = !vl_brt
            buscaTot
        End If
        
    End With
End Sub
Private Sub Form_Load()
    bActive = False
    txtData.Text = Date
    txtCaixa.Text = strCod_Cai
    buscaTot
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If grid.Row = 0 Then Exit Sub
        grid.col = 2
        txtGrid.Left = grid.CellLeft + grid.Left
        txtGrid.Top = grid.CellTop + grid.Top
        txtGrid.Width = grid.CellWidth
        txtGrid.Height = grid.CellHeight
        txtGrid.Visible = True
        txtGrid.SetFocus
    End If
    
End Sub

Private Sub txtData_GotFocus()
    ganhaFoco txtData
End Sub
Private Sub txtCaixa_GotFocus()
    ganhaFoco txtCaixa
End Sub
Private Sub txtCRO_GotFocus()
    ganhaFoco txtCRO
End Sub

Private Sub txtCRZ_GotFocus()
    ganhaFoco txtCRZ
End Sub
Private Sub txtCOO_GotFocus()
    ganhaFoco txtCOO
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub
Private Sub txtCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub
Private Sub txtCRO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub
Private Sub txtCRZ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub
Private Sub txtCOO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtGrid_GotFocus()
    ganhaFoco txtGrid
End Sub

Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         PerdeFoco txtGrid
    grid.col = 2
    txtGrid.Visible = False
    grid.Text = txtGrid.Text
    If grid.Row < grid.rows - 1 Then
        grid.Row = grid.Row + 1
    End If
    grid.col = 2
    
    grid.SetFocus
    End If
End Sub

Private Sub txtGrid_LostFocus()
    
   txtGrid.Visible = False
    
End Sub

Private Sub txtGT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtGT_LostFocus()
    PerdeFoco txtGT
End Sub

Private Sub txtVlBrt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtGT_GotFocus()
    ganhaFoco txtGT
End Sub
Private Sub txtVlbrt_GotFocus()
    ganhaFoco txtVlBrt
End Sub


Private Sub txtVlBrt_LostFocus()
    PerdeFoco txtVlBrt
End Sub
