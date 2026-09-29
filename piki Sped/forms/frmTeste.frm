VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTeste 
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   975
      Left            =   1800
      TabIndex        =   6
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   735
      Left            =   4800
      TabIndex        =   5
      Top             =   2520
      Width           =   615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgrid 
      Height          =   975
      Left            =   2520
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1720
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Text            =   "Teste"
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Visualizar Impressao"
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   1560
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1560
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
End
Attribute VB_Name = "frmTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i As Single
'    Printer.PaperSize = 256
    Printer.width = Printer.ScaleX(140, vbMillimeters, vbTwips)
    Printer.height = Printer.ScaleX(210, vbMillimeters, vbTwips)
    For i = 0 To Printer.height Step 100
        Printer.Print i
        
    Next
    Printer.EndDoc
    
    
End Sub

Private Sub Command2_Click()
    
    If Not F_OpenLpt Then
        Exit Sub
    End If
    Dim iCount As Integer
    Dim str As String
    Dim Margem As Integer
    For iCount = 1 To 80
        str = str & Left(iCount, 1)
    Next
    P_PrintLpt str, 0, False
    P_PrintLpt str, 0, False
    
    
    If Not F_CloseLpt Then
        Exit Sub
    End If
End Sub

Private Sub Command3_Click()
    Dim obj As Object
    Load frmVisualizadorImpressao
    Set obj = frmVisualizadorImpressao.pcImpressao
    obj.Print Text1.Text
     
    
    frmVisualizadorImpressao.addNewPage
    
    obj.Print "Segunda página"
    frmVisualizadorImpressao.addNewPage
    
    frmVisualizadorImpressao.Show 1
    Unload frmVisualizadorImpressao
    
End Sub

