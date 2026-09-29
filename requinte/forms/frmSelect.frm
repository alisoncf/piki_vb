VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSelect 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Selecione"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   6885
      TabIndex        =   5
      Top             =   5325
      Width           =   6885
      Begin VB.CommandButton cmndbtnSelecionar 
         Caption         =   "&Selecionar"
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "E&ditar"
         Height          =   495
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sai&r"
         Height          =   495
         Left            =   5400
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   6855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public sSelect As String
Public sOrd As String
Public cCriterio As New Collection
Dim v(100) As Integer
Public sExpr As String




Private Sub cmndbtnSair_Click()
    Me.tag = ""
    Me.Hide

End Sub


Private Sub cmndbtnSelecionar_Click()
    With msflxgrd
        .col = 0
        Me.tag = .Text
        Me.Hide
    End With
End Sub

Sub busca()

    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iCount As Integer
    Dim jCount As Integer
    On Error GoTo trata_Erro
    
    msflxgrd.Visible = False
    For iCount = 1 To cCriterio.Count
        
        abreObjetoAdo rPesq, sSelect & cCriterio(iCount) & sOrd
        If rPesq.RecordCount > 0 Then
            Exit For
        End If
        If rPesq.RecordCount > 1000 Then
            If MsgBox("A Pesquisa retornará mais de 1000 registros. Pode demorar muito tempo. Deseja continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                rPesq.Close
                Exit Sub
            End If
        End If
    Next
    
    
    If cCriterio.Count = 0 Then
        abreObjetoAdo rPesq, sSelect & sOrd
        If rPesq.RecordCount > 0 Then
            
        End If
        If rPesq.RecordCount > 1000 Then
            If MsgBox("A Pesquisa retornará mais de 1000 registros. Pode demorar muito tempo. Deseja continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                rPesq.Close
                msflxgrd.Visible = True
                Exit Sub
            End If
        End If
    End If
    
    With rPesq
        
        If .RecordCount = 1 Then
           
        ElseIf .RecordCount = 0 Then
            
            Exit Sub
        End If
        
        msflxgrd.rows = 1
        msflxgrd.Cols = .Fields.Count
        
        
        For iCount = 0 To .Fields.Count - 1
            msflxgrd.row = 0
            msflxgrd.col = iCount
            msflxgrd.Text = .Fields(iCount).Name
        Next
        Do While Not .EOF
            strSQL = ""
            For iCount = 0 To .Fields.Count - 1
                strSQL = strSQL & .Fields(iCount).Value & Chr(9)
            Next
            msflxgrd.AddItem strSQL
            .MoveNext
        Loop
        If .RecordCount = 1 Then
            msflxgrd.row = 1
            msflxgrd.col = 1
            Me.tag = msflxgrd.Text
            On Error Resume Next
            msflxgrd.Visible = True
            Me.Hide
            Exit Sub
        End If
        ajustaTamGrid
        ZebraGrid msflxgrd, 1
        If .RecordCount > 1 Then
            Me.caption = "Foram Localizados " & .RecordCount & " Elementos. Selecione algum."
            msflxgrd.TopRow = 1
            msflxgrd.col = 0
            msflxgrd.row = 1
            
'            msflxgrd.SetFocus
        End If
    End With
    
    msflxgrd.Visible = True
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub Form_Activate()
   Form_Resize
    busca
    
End Sub
Sub ajustaTamGrid()
    Dim iCount As Integer
    Dim iCol As Integer
    Dim tMax As String
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            For iCol = 1 To .Cols - 1
                .col = iCol
                If TextWidth(.Text) > .ColWidth(iCol) Then
                     .ColWidth(iCol) = TextWidth(.Text) + 100
                End If
                
            Next
        Next
    End With
End Sub



Private Sub Form_Load()
'    Set cCriterio = Nothing
End Sub

Private Sub Form_Resize()
    msflxgrd.width = Me.width - 200
    msflxgrd.height = pcBot.top - 300
    cmndbtnSair.Left = pcBot.width - cmndbtnSair.width - 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cCriterio = Nothing
End Sub

Private Sub msflxgrd_DblClick()
    cmndbtnSelecionar_Click
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        msflxgrd_DblClick
    End If
End Sub

Private Sub txtbxExpr_Change()
'    P_AjustaTexto txtbxExpr
    
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        msflxgrd.SetFocus
    End If
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub


