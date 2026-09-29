VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIncPerfil 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Perfis"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmIncPerfil.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.Data dtImp 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid dbGrdImp 
      Bindings        =   "frmIncPerfil.frx":0442
      Height          =   1095
      Left            =   120
      OleObjectBlob   =   "frmIncPerfil.frx":0456
      TabIndex        =   11
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton cmndbtnUnselect 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   4080
      Picture         =   "frmIncPerfil.frx":0E21
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmndbtnSelec 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   4080
      Picture         =   "frmIncPerfil.frx":1263
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrdItens 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1931
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Salvar"
      Height          =   735
      Left            =   2640
      Picture         =   "frmIncPerfil.frx":16A5
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   735
      Left            =   3600
      Picture         =   "frmIncPerfil.frx":1A2F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtbxDesc_Per 
      Height          =   285
      Left            =   120
      MaxLength       =   40
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.TextBox txtbxNome_Per 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Impostos e Taxas &Selecionados"
      Height          =   195
      Left            =   990
      TabIndex        =   9
      Top             =   2760
      Width           =   2250
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Impostos e Taxas Disponíveis"
      Height          =   195
      Left            =   1057
      TabIndex        =   8
      Top             =   1320
      Width           =   2130
   End
   Begin VB.Label Label2 
      Caption         =   "&Descrição do Perfil:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "&Nome do Perfil"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmIncPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub P_Grava()
    Dim strSQL As String
    Dim sCod As String
    Dim iCount As Integer
    sCod = F_BuscaUmCampo("max(cod_per) ", "tab_perfis", "")
    If sCod = "" Then
        sCod = "01"
    Else
        sCod = Format(CInt(sCod) + 1, "0#")
    End If
    strSQL = "insert into tab_perfis(cod_per, nome_per,desc_per) values( "
    strSQL = strSQL & "'" & sCod & "', "
    strSQL = strSQL & "'" & txtbxNome_Per.Text & "', "
    strSQL = strSQL & "'" & txtbxDesc_Per & "')"
    banco.Execute strSQL
    
    With msflxgrdItens
        For iCount = 1 To .Rows - 1
            
            .Row = iCount
            strSQL = "insert into tab_perimp(cod_imp,cod_per,"
            strSQL = strSQL & "tip_imp,val_imp,data_imp) values("
            .Col = 0
            strSQL = strSQL & "'" & .Text & "', "
            
            strSQL = strSQL & "'" & sCod & "', "
            
            .Col = 2
            strSQL = strSQL & "'" & .Text & "', "
            .Col = 3
            strSQL = strSQL & f_GravaDec(.Text) & ","
            .Col = 4
            strSQL = strSQL & F_GravaData(.Text) & ") "
            banco.Execute strSQL
        Next
        MsgBox "perfil cadastrado", vbInformation
        msflxgrdItens.Rows = 1
        txtbxDesc_Per.Text = ""
        txtbxNome_Per.Text = ""
        txtbxNome_Per.SetFocus
    End With
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.Tag = "" Then
        P_Grava
    Else
        P_Altera
    End If
End Sub
Sub P_Altera()
    Dim iCount As Integer
    Dim strSQL As String
    If txtbxDesc_Per.Text = "" Then
        MsgBox "Informe a descrição do perfil!", vbExclamation
        txtbxDesc_Per.SetFocus
        Exit Sub
    ElseIf txtbxNome_Per.Text = "" Then
        MsgBox "Informe o Nome do perfil!", vbExclamation
        txtbxNome_Per.SetFocus
        Exit Sub
    ElseIf msflxgrdItens.Rows = 1 Then
        MsgBox "Não foram selecionados impostos!", vbExclamation
        Exit Sub
    
    ElseIf F_BuscaUmCampo("COD_PER", "TAB_PERFIS", "NOME_PER='" & txtbxNome_Per.Text & "' " _
                    & "And cod_per<>'" & Me.Tag & "'") <> "" Then
        MsgBox "Já existe um perfil com este nome!", vbExclamation
    End If
    strSQL = "Update tab_Perfis set nome_per = '" & txtbxNome_Per.Text & "',"
    strSQL = strSQL & " desc_per='" & txtbxDesc_Per.Text & "' "
    strSQL = strSQL & " where cod_per='" & Me.Tag & "'"
    banco.Execute strSQL
    strSQL = "Delete from tab_perimp where cod_per='" & Me.Tag & "' "
    banco.Execute strSQL

    With msflxgrdItens
        For iCount = 1 To .Rows - 1
            
            .Row = iCount
            strSQL = "insert into tab_perimp(cod_imp,cod_per,"
            strSQL = strSQL & "tip_imp,val_imp,data_imp) values("
            .Col = 0
            strSQL = strSQL & "'" & .Text & "', "
            
            strSQL = strSQL & "'" & Me.Tag & "', "
            
            .Col = 2
            strSQL = strSQL & "'" & .Text & "', "
            .Col = 3
            strSQL = strSQL & f_GravaDec(.Text) & ","
            .Col = 4
            strSQL = strSQL & F_GravaData(.Text) & ") "
            banco.Execute strSQL
        Next
        MsgBox "Perfil Alterado", vbInformation
        Unload Me
    End With

End Sub
Private Sub cmndbtnSelec_Click()
    Dim iCount As Integer
    Dim sTipo As String
    On Error Resume Next
    For iCount = 1 To msflxgrdItens.Rows - 1
        msflxgrdItens.Row = iCount
        msflxgrdItens.Col = 0
        If dtImp.Recordset!cod_Imp = msflxgrdItens.Text Then
            MsgBox "Este intem já está selecionado", vbExclamation
            Exit Sub
        End If
    Next
    
    frmTipo_Imp.Show 1
    If frmTipo_Imp.Tag = "" Then
        sTipo = Chr(9) & "1"
    Else
        sTipo = Chr(9) & "2" & Chr(9) & frmTipo_Imp.txtbxValor_Imp.Text & Chr(9)
        sTipo = sTipo & frmTipo_Imp.txtbxData_imp.Text
    End If
    Unload frmTipo_Imp

    msflxgrdItens.AddItem dtImp.Recordset!cod_Imp & Chr(9) & dtImp.Recordset!nome_Imp & sTipo
End Sub

Private Sub cmndbtnUnselect_Click()
    If msflxgrdItens.Rows = 2 Then
        msflxgrdItens.Rows = 1
    Else
        msflxgrdItens.RemoveItem msflxgrdItens.Row
    End If
End Sub

Private Sub dtImp_Reposition()
    If Trim(txtbxDesc_Per.Text) <> "" And Trim(txtbxNome_Per.Text) <> "" Then
        cmndbtnSelec.Enabled = True
        cmndbtnUnselect.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    Dim strSQL As String
    dtImp.Connect = banco.Connect
    strSQL = "select * from tab_impostos "
    dtImp.RecordSource = strSQL
    dtImp.Refresh
End Sub

Private Sub Form_Load()
    msflxgrdItens.Rows = 1
    msflxgrdItens.Cols = 5
    msflxgrdItens.Col = 0
    msflxgrdItens.Text = "Cód"
    msflxgrdItens.Col = 1
    msflxgrdItens.Text = "Nome do Imposto"
    msflxgrdItens.Col = 2
    msflxgrdItens.Text = "Ti"
    msflxgrdItens.Col = 3
    msflxgrdItens.Text = "Valor"
    msflxgrdItens.Col = 4
    msflxgrdItens.Text = "Data"
    
    
    msflxgrdItens.ColWidth(0) = 600
    msflxgrdItens.ColWidth(1) = 3000
    msflxgrdItens.ColWidth(2) = 300
    msflxgrdItens.ColWidth(3) = 600
    msflxgrdItens.ColWidth(4) = 600
    P_CentraForm Me
End Sub

Private Sub txtbxDesc_Per_GotFocus()
    P_SelecTexto txtbxDesc_Per
End Sub

Private Sub txtbxDesc_Per_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDesc_Per_LostFocus()
    P_AjustaTexto txtbxDesc_Per
    dtImp_Reposition
End Sub

Private Sub txtbxNome_Per_GotFocus()
    P_SelecTexto txtbxNome_Per
End Sub

Private Sub txtbxNome_Per_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Per_LostFocus()
    P_AjustaTexto txtbxNome_Per
    dtImp_Reposition
    
End Sub
