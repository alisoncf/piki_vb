VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmImportCid 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importar Cidades - IBGE"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnImport 
      Caption         =   "Importar"
      Height          =   735
      Left            =   5760
      Picture         =   "frmImportCid.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton cmndbtnFile 
      Caption         =   "..."
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidades Encontradas:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho do Arquivo TXT:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1830
   End
End
Attribute VB_Name = "frmImportCid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnFile_Click()
    Dim sFile As String
    mdiFrm.cmndlg.Filter = "*.txt"
    mdiFrm.cmndlg.ShowOpen
    sFile = mdiFrm.cmndlg.FileName
    Text1.Text = sFile
    abreArquivo sFile
    
    
    
End Sub
Sub abreArquivo(sNomeArq As String)
    Open sNomeArq For Input As #1
    Dim sCod As String
    Dim strSQL As String
    Dim sBuf As String
    msflxgrd.rows = 1
    Do While Not EOF(1)
        Line Input #1, sBuf
        strSQL = Left(sBuf, 7) & Chr(9)
        strSQL = strSQL & Trim(Mid(sBuf, 8, 40)) & Chr(9)
        strSQL = strSQL & Trim(Mid(sBuf, 48, 5)) & Chr(9)
        strSQL = strSQL & Trim(Mid(sBuf, 53, 2))
        msflxgrd.AddItem strSQL
        
        P_Status "Abrindo..."
        
    Loop
    
        
        
        
    Close #1
    P_Status ""
End Sub

Private Sub cmndbtnImport_Click()
    Dim iCount As Integer
    Dim sCod As String
    Dim sNome As String
    Dim sId As String
    Dim sUF As String
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 0
            sCod = .Text
            .Col = 1
            sNome = .Text
            .Col = 2
            sId = .Text
            .Col = 3
            sUF = .Text
            
            P_GravaCidade sCod, sNome, sId, sUF
            P_Status .Row & "/" & .rows
        Next
        
    End With
    P_Status ""
    MsgBox "Concluído", vbInformation
End Sub

Private Sub Form_Load()
    With msflxgrd
        .FixedCols = 0
        .Cols = 4
        .rows = 1
        .Col = 0: .Text = "Código"
        .Col = 1: .Text = "Nome"
        .Col = 2: .Text = "Id"
        .Col = 3: .Text = "UF"
        .ColWidth(1) = 3500
    End With
    
End Sub
