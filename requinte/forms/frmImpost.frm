VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmImpost 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impostos"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnEdit 
      Caption         =   "&Editar"
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdnbtnNovo 
      Caption         =   "&Novo"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame frNovo 
      Caption         =   "&Adicionar Novo Imposto:"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cmbxECF 
         Height          =   315
         ItemData        =   "frmImpost.frx":0000
         Left            =   3720
         List            =   "frmImpost.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnOk 
         Caption         =   "Ok"
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtbxInd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtbxDesc 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtbxPerc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ECF:"
         Height          =   195
         Left            =   3720
         TabIndex        =   8
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Índ. ECF:"
         Height          =   195
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perc:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5318
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "frmImpost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnbtnNovo_Click()
    frNovo.Visible = True
    cmndbtnOk.tag = ""
    txtbxPerc.SetFocus
End Sub

Private Sub cmndbtnEdit_Click()
    With msflxgrd
        If .rows = 1 Then Exit Sub
        .col = 0
        cmndbtnOk.tag = .text
        .col = 1
        txtbxDesc.text = .text
        .col = 2
        txtbxPerc.text = .text
        .col = 3
        txtbxInd.text = .text
        frNovo.Visible = True
        txtbxPerc.SetFocus
    End With
End Sub

Private Sub cmndbtnOk_Click()
    If cmndbtnOk.tag = "" Then
        grava
        busca
    Else
        altera
        End If
End Sub
Sub grava()
    Dim strSQL As String
    Dim sCod As String
    sCod = MaxKeySQL("tab_impostos", "cod_imp", "", "I0#")
    strSQL = "Insert into tab_impostos (cod_imp,desc_imp,perc_imp)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxDesc.text & "',"
    strSQL = strSQL & F_GravaDec(txtbxPerc.text) & ")"
    
    Conn.Execute strSQL
    strSQL = "Insert into tab_ecfimp(cod_imp,ecf_imp,impressora_imp)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxInd.text & "',"
    strSQL = strSQL & "'" & Left(cmbxECF.text, 1) & "')  "
    Conn.Execute strSQL
    MsgBox "Imposto Gravado!", vbInformation
    txtbxDesc.text = ""
    frNovo.Visible = False
    
    
End Sub
Sub altera()
    Dim strSQL As String
    strSQL = "update tab_impostos set desc_imp='" & txtbxDesc.text & "',"
    strSQL = strSQL & "perc_imp=" & F_GravaDec(txtbxPerc.text) & " "
    
    strSQL = strSQL & " where cod_imp='" & cmndbtnOk.tag & "' "
    Conn.Execute strSQL
    strSQL = "update tab_ecfimp set ecf_imp='" & txtbxInd.text & "', "
    strSQL = strSQL & "impressora_imp='" & Left(cmbxECF.text, 1) & "' "
    strSQL = strSQL & " where cod_imp='" & cmndbtnOk.tag & "' "
    strSQL = strSQL & " and impressora_imp='" & Left(cmbxECF.text, 1) & "' "
    Conn.Execute strSQL
    MsgBox "Imposto Alterado", vbInformation
    cmndbtnOk.tag = ""
    
    
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    
End Sub

Private Sub Form_Load()
    cmbxECF.ListIndex = 0
    With msflxgrd
        .FixedCols = 0
        .FixedRows = 1
        .rows = 1
        .Cols = 5
        .col = 0: .text = "Cód"
        .col = 1: .text = "Descrição"
        .col = 2: .text = "Perc."
        .col = 3: .text = "Índice ECF"
        .col = 4: .text = "Impressora"
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .ColWidth(2) = 800
        .ColWidth(3) = 1200
        .ColAlignment(3) = 1
    End With
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ti.desc_imp,ti.perc_imp,te.* from tab_impostos as ti,"
    strSQL = strSQL & " tab_ecfimp as te where te.cod_imp=ti.cod_imp order by impressora_imp,ti.cod_imp"
    abreObjetoAdo rPesq, strSQL
    msflxgrd.rows = 1
    With rPesq
    
        Do While Not .EOF
            strSQL = !cod_imp & Chr(9)
            strSQL = strSQL & !desc_imp & Chr(9)
            strSQL = strSQL & !perc_imp & Chr(9)
            strSQL = strSQL & !ecf_imp & Chr(9)
            strSQL = strSQL & !impressora_imp
            msflxgrd.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    If KeyCode = vbKeyDelete Then
        msflxgrd.col = 0
        If MsgBox("Confirma Exclusão do Item?", vbYesNo + vbCritical) = vbYes Then
            strSQL = "delete from tab_impostos where cod_imp='" & msflxgrd.text & "' "
            Conn.Execute strSQL
            busca
        End If
    End If
    
End Sub
