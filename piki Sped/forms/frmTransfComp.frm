VERSION 5.00
Begin VB.Form frmTransfComp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transferência:"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnTransf 
      Caption         =   "&Transferir"
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox cmbxPara 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1920
      Width           =   3855
   End
   Begin VB.ComboBox cmbxDe 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transferir:"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.OptionButton optnCat 
         Caption         =   "&Categorias"
         Height          =   195
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optnMarcas 
         Caption         =   "&Marcas"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Para:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&De:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "frmTransfComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sTipo As String

Private Sub cmbxDe_Click()
    If sTipo = "cat" Then
        cmbxDe.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_Cat='" & cmbxDe.Text & "'")
    Else
        cmbxDe.tag = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxDe.Text & "'")
    End If
End Sub

Private Sub cmbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxPara_Click()
    If sTipo = "mar" Then
        cmbxPara.tag = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxPara.Text & "'")
    Else
        cmbxPara.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_Cat='" & cmbxPara.Text & "'")
    End If
End Sub

Private Sub cmbxPara_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnTransf_Click()
    Dim strSQL As String
    If cmbxDe.tag = "" Then
        MsgBox "Informe o componente que será dissolvido!", vbExclamation
        cmbxDe.SetFocus
        Exit Sub
    End If
    If cmbxPara.tag = "" Then
        MsgBox "Informe o componente que será acrescido!", vbExclamation
        cmbxPara.SetFocus
        Exit Sub
    End If
        
    If sTipo = "cat" Then
        strSQL = "Confirma transferência de todos os produtos "
        strSQL = strSQL & " da Categoria " & cmbxDe.Text
        strSQL = strSQL & " para a Categoria " & cmbxPara.Text
        If MsgBox(strSQL, vbYesNo + vbCritical) = vbYes Then
            strSQL = "update " & TAB_PRODUTOS & " set cod_cat='" & cmbxPara.tag & "' "
            strSQL = strSQL & " where cod_cat='" & cmbxDe.tag & "' "
            conn.Execute strSQL
            MsgBox Banco.RecordsAffected & " Registros foram substituídos!", vbInformation
        
        End If
            
    Else
        strSQL = "Confirma transferência de todos os produtos "
        strSQL = strSQL & " da marca " & cmbxDe.Text
        strSQL = strSQL & " para a Categoria " & cmbxPara.Text
        If MsgBox(strSQL, vbYesNo + vbCritical) = vbYes Then
            strSQL = "update " & TAB_PRODUTOS & " set cod_mar='" & cmbxPara.tag & "' "
            strSQL = strSQL & " where cod_mar='" & cmbxDe.tag & "' "
            conn.Execute strSQL
            MsgBox Banco.RecordsAffected & " Registros foram substituídos!", vbInformation
        
        End If
        
        
    End If
End Sub

Private Sub Form_Activate()
    If sTipo = "cat" Then
        P_PreencheCateg cmbxDe
        P_PreencheCateg cmbxPara
        
        optnCat.Value = True
    Else
        P_PreencheMarca cmbxDe
        P_PreencheMarca cmbxPara
        optnMarcas.Value = True
    End If
End Sub

Private Sub Form_Load()
    sTipo = "cat"
    
End Sub

Private Sub optnCat_Click()
    sTipo = "cat"
    P_PreencheCateg cmbxDe
        P_PreencheCateg cmbxPara
End Sub

Private Sub optnMarcas_Click()
    sTipo = "mar"
    P_PreencheMarca cmbxDe
    P_PreencheMarca cmbxPara
End Sub
