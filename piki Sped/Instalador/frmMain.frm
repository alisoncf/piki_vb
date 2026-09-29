VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Instalador Piki"
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar (X)"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Contiuar >>"
      Default         =   -1  'True
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Este aplicativo instala um novo banco de dados do Piki. Clique em continuar para proceder a instalação ou em cancelar para sair."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As ADODB.Connection
Dim usuario As String
Dim senha As String
Dim nomebanco As String



Private Sub Command1_Click()
    instalar
End Sub

Private Sub Command2_Click()
    End
End Sub
Function AbreBanco()
        
    On Error GoTo erro
    DBEngine.DefaultType = dbUseODBC
   
    strConnect = "Driver=" & sDriver & ";"
    
    strConnect = strConnect & "Server=" & sServer & ";Database=" & sDataBase & ";Uid=" & sUid & ";port=3306;pwd=" & sPwd & ";option=" & sOption
    
    'strConnect = "market"
   conn.Open strConnect
   conn_client.CursorLocation = adUseClient
   conn_client.Open strConnect
   conn_client.CursorLocation = adUseClient
        'Set Banco = OpenDatabase(Trim(sDataBase), "", "", strConnect)
'    End If
   
    F_AbreBanco = True
    Exit Function
erro:
    


MsgBox Err.Description & Chr(13) & strConnect, vbCritical, conSigla_Sistema
    F_AbreBanco = False


End Function
Sub instalar()
    Dim driver As String
    frmODBC.Show 1
    
    
    
    
    
    
    driver = frmODBC.lst.Text
    Unload frmODBC
    
    usuario = InputBox("Informe o nome do usuário do mySQL que tem a permissão necessária", "user", "root")
    
    
    
    
End Sub
Private Sub Command3_Click()
    Dim c As New Collection
    Set c = modulo.ListODBCDrivers
    
    
    
End Sub

