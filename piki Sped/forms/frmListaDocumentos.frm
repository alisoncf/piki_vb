VERSION 5.00
Begin VB.Form frmListaDocumentos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informa a Lista de Documentos"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Limpar"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Confirmar"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtFormato 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "000000#"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtLista 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Formato:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "frmListaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Lista As String
Private Sub cmdOk_Click()
    Dim c As New Collection
    Set c = F_SeparaTexto(txtLista.Text, ",.-; ", False, txtFormato.Text, "'")
    Lista = CollectionToNotIn(c)
    Me.Hide
    
    
    
End Sub

Private Sub Command1_Click()
    txtLista.Text = ""
End Sub

Private Sub txtFormato_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdOk_Click
    End If
End Sub
