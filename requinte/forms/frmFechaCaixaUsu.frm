VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFechaCaixaUsu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Movimentação de Caixa por Usuário:"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtbxData 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8070
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   390
   End
End
Attribute VB_Name = "frmFechaCaixaUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With msflxgrd
        .rows = 1
        .FixedCols = 0
        .Cols = 7
        .col = 0: .text = "Funcionário"
        .col = 1: .text = "Caixa"
        .col = 2: .text = "Turno"
        .col = 3: .text = "À Vista"
        .col = 4: .text = "A Prazo"
        .col = 5: .text = "Finan"
        .col = 6: .text = "Total"
        .ColWidth(0) = 2500
        .ColWidth(1) = 600
        .ColWidth(2) = 400
    End With
    txtbxData.text = Date
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCod As String
    Dim cAv As Currency
    Dim cAp As Currency
    Dim cFinan As Currency
    Dim cTotal As Currency
    Dim iCount As Integer
    Dim rVendas As New ADODB.Recordset
    Dim cTotalCai As Currency
    msflxgrd.rows = 1
    strSQL = "select tf.cod_fun,tf.guerra_fun,tvf.cod_cai,tvf.cod_tur ,sum(tvf.avista_VEN) AS  avista, "
    strSQL = strSQL & " sum(tvf.aprazo_VEN) AS  aprazo,sum(tvf.finan_VEN) AS  finan"
    strSQL = strSQL & " from tab_vendas as tv,tab_venfec as tvf, "
    strSQL = strSQL & "tab_funcionarios as tf "
    strSQL = strSQL & " where tvf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tv.cod_usu=tf.cod_fun "
    strSQL = strSQL & " and tv.dataefet_ven=" & F_GravaData(txtbxData.text)
    strSQL = strSQL & " group by tv.cod_fun,tvf.cod_Cai,tvf.cod_tur  "
    strSQL = strSQL & " order by tf.guerra_fun,cod_cai,cod_tur "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then Exit Sub
        sCod = !cod_fun
        
        Do While Not .EOF
            If sCod <> !cod_fun Or .EOF Then
                
                strSQL = "Total..................:" & Chr(9) & Chr(9) & Chr(9) & FormatVal(cAv) & Chr(9)
                strSQL = strSQL & FormatVal(cAp) & Chr(9) & FormatVal(cFinan) & Chr(9)
                strSQL = strSQL & FormatVal(cTotal)
                msflxgrd.AddItem strSQL
                msflxgrd.row = msflxgrd.rows - 1
                For iCount = 0 To msflxgrd.Cols - 1
                    
                    msflxgrd.col = iCount
                    msflxgrd.CellFontItalic = True
                    msflxgrd.CellFontBold = True
                Next
                
                strSQL = "select * from tab_caixa where cod_usu='" & sCod & "' "
                strSQL = strSQL & " and data_Cai=" & F_GravaData(txtbxData.text)
                strSQL = strSQL & " and num_ven='' "
                strSQL = strSQL & " and ESP_CAI='DIN' "
                strSQL = strSQL & " and excluido_cai=false "
                strSQL = strSQL & "  group by cod_cai "
                strSQL = strSQL & "  order by cod_cai,tipo_cai "
                
                
                Set rVendas = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
                cTotal = 0
                With rVendas
                    Do While Not .EOF
                        strSQL = !historico_cai & Chr(9) & !tipo_cai & Chr(9) & Chr(9)
                        strSQL = strSQL & FormatVal(!valor_Cai)
                        msflxgrd.AddItem strSQL
                        cTotalCai = cTotal + !valor_Cai
                        .MoveNext
                    Loop
                End With
                
            End If
            strSQL = !guerra_fun & Chr(9) & !cod_cai & Chr(9)
            strSQL = strSQL & !cod_tur & Chr(9)
            strSQL = strSQL & FormatVal(!avista) & Chr(9)
            strSQL = strSQL & FormatVal(!aprazo) & Chr(9)
            strSQL = strSQL & FormatVal(!finan) & Chr(9)
            strSQL = strSQL & FormatVal(!avista + !aprazo + !finan)
            
            msflxgrd.AddItem strSQL
            cAv = cAv + !avista
            cAp = cAo + !aprazo
            cFinan = cFinan + !finan
            cTotal = cTotal + !avista + !aprazo + !finan
            .MoveNext
        Loop
        strSQL = "Total..................:" & Chr(9) & Chr(9) & Chr(9) & FormatVal(cAv) & Chr(9)
        strSQL = strSQL & FormatVal(cAp) & Chr(9) & FormatVal(cFinan) & Chr(9)
        strSQL = strSQL & FormatVal(cTotal)
        msflxgrd.AddItem strSQL
        msflxgrd.row = msflxgrd.rows - 1
        For iCount = 0 To msflxgrd.Cols - 1
            
            msflxgrd.col = iCount
            msflxgrd.CellFontItalic = True
            msflxgrd.CellFontBold = True
        Next
        ZebraGrid msflxgrd, 2
    End With
End Sub
