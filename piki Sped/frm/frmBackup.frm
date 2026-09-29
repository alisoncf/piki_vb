VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBackup 
   Caption         =   "Back Up"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   735
      Left            =   6240
      TabIndex        =   4
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdDump 
      Caption         =   "Gerar Dump"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   285
      Left            =   5400
      Picture         =   "frmBackup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "c:\BackUP"
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pasta de Destino"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDump_Click()
    Dim strSQL As String
    Dim i As Integer
    Dim nome_arquivo As String
    
    nome_arquivo = sDataBase & "_BackUp_" & Format(Now, "DD-MM-YYYY_HHMM") & ".sql"
    
    If sPwd <> "" Then
    
    strSQL = "mysqldump -u " & sUid & " -p" & sPwd & " " & sDataBase & " > " & txtPath.Text & "\" & nome_arquivo

    Else
        strSQL = "mysqldump -u " & sUid & " " & sDataBase & " > " & txtPath.Text & "\" & nome_arquivo
    End If
    
    
    P_Status "Gerando Arquivo..."
    strSQL = strSQL
    i = Shell(strSQL)
    
    Sleep (1000)
    
    
    If Dir(txtPath.Text & "\" & nome_arquivo) <> "" Then
        Shell "explorer " & txtPath.Text
    End If
    P_Status ""
    
    
    
End Sub

Private Sub cmdOpen_Click()
    Dim s As String
    mdiFrm.cmndlg.ShowSave
    
    
    
    
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
