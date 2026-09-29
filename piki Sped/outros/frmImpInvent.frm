VERSION 5.00
Begin VB.Form frmImpInvent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impressão de Inventário:"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Termo de Abertura/Encerramento:"
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   7575
      Begin VB.OptionButton optnEncerra 
         Caption         =   "&Encerramento"
         Height          =   195
         Left            =   4560
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optnAbertura 
         Caption         =   "&Abertura"
         Height          =   195
         Left            =   4560
         TabIndex        =   31
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtbxFolhaAbre 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   27
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtbxNumAbre 
         Height          =   285
         Left            =   2400
         TabIndex        =   26
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtbxDataAbre 
         Height          =   285
         Left            =   3360
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxFinTermo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Text            =   "000"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtbxIniTermo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "001"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmnbtnImprTermoAbert 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   6240
         Picture         =   "frmImpInvent.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Emite Inventário"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Folha:"
         Height          =   195
         Left            =   1680
         TabIndex        =   30
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº"
         Height          =   195
         Left            =   2400
         TabIndex        =   29
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   3360
         TabIndex        =   28
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pág. Final:"
         Height          =   195
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pág. Ini:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton Command1 
         Caption         =   "&Abrir Invent"
         Height          =   615
         Left            =   4920
         Picture         =   "frmImpInvent.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxPag 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   13
         Text            =   "001"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtbxData 
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxInvent 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmndbtnImprInv 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   6240
         Picture         =   "frmImpInvent.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Emite Inventário"
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbxInvent 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folha Ini:"
         Height          =   195
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº"
         Height          =   195
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Inventário de:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Ativo Imobilizado:"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7575
      Begin VB.CommandButton Command2 
         Caption         =   "&Abrir Imob."
         Height          =   615
         Left            =   4920
         Picture         =   "frmImpInvent.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxdataImob 
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxNumImob 
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnImprImob 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   6240
         Picture         =   "frmImpInvent.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Emite Inventário"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxPagImob 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº"
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Folha:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Parâmetros"
      Height          =   1335
      Left            =   120
      TabIndex        =   34
      Top             =   3480
      Width           =   7575
   End
End
Attribute VB_Name = "frmImpInvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function imprimeInv(obj As Object) As Integer
    Dim sEnd As String
    Dim iCount As Integer
    Dim iPag As Integer
    Dim impr As impressao
    Dim rPesq As Recordset
    Dim cTotalTransp As Currency
    Dim strSQL As String
    Dim yFim As Single
    Dim xFim As Single
    strSQL = "Select TI.*,(TI.ef_pro) as estoque,((TI.ef_pro)*TI.CUSTO_PRO) AS TOTAL, TM.*,"
    strSQL = strSQL & " TI.ef_pro * TI.custo_pro as tf,"
    strSQL = strSQL & " TM.* from tab_Invent as TI, TAB_produtos AS TM"
    strSQL = strSQL & " Where "
    strSQL = strSQL & " TI.COD_pro = TM.COD_pro  and "
    strSQL = strSQL & " DATAHORA_INV ='" & Format(cmbxInvent.Text, "YYYY-MM-DD-HH:MM:SS") & "' "
    strSQL = strSQL & " having total > 0 "
    strSQL = strSQL & " ORDER BY TM.NOME_pro"
    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)

    
    On Error GoTo trata_erro
    If rPesq.RecordCount = 0 Then
        MsgBox "Não há o que imprimir!", vbExclamation
        cmbxInvent.SetFocus
        rPesq.Close
    End If

    iPag = Int(txtbxPag.Text)
    obj.ScaleMode = vbMillimeters
    
    obj.FontName = "courier new"
    obj.fontsize = 10
    
    P_CabInv obj, txtbxInvent.Text, txtbxData.Text, txtbxPag.Text
    P_CabInvProd obj, "0,00"
    
    
    
    
    With rPesq
        Do While Not .EOF
            If iCount = 60 Then
                iCount = 0
                iPag = iPag + 1
                P_RodaPeInvProd obj, FormatVal(cTotalTransp)
                yFim = obj.CurrentY
                P_CabInv obj, txtbxInvent.Text, txtbxData.Text, Format(iPag, "00#")
                P_CabInvProd obj, FormatVal(cTotalTransp)
            End If
            
            
            P_ItemInv obj, "", !nome_pro, !unidade_pro, FormatVal(!ef_pro), FormatVal(!custo_pro), FormatVal(!tf)
            
            
            
            iCount = iCount + 1
            cTotalTransp = cTotalTransp + !tf
            .MoveNext
        Loop
        P_RodaPeInvFinal obj, FormatVal(cTotalTransp)
        .Close
    End With
    
    
Exit Function
trata_erro:
    MsgBox "Não foi possível concluir a impressão!" & Chr(13) & Err.Description, vbCritical
    
End Function
Function imprimeImob(obj As Object) As Integer
    Dim sEnd As String
    Dim iCount As Integer
    Dim iPag As Integer
    Dim impr As impressao
    Dim rPesq As Recordset
    Dim cTotalTransp As Currency
    Dim strSQL As String
    Dim yFim As Single
    Dim xFim As Single
    strSQL = "select ti.*,ti.qtd_imo*ti.unit_imo as total from tab_imob as ti order by Desc_imo"
    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)

    
    On Error GoTo trata_erro
    If rPesq.RecordCount = 0 Then
        MsgBox "Não há o que imprimir!", vbExclamation
        cmbxInvent.SetFocus
        
    End If

    
    obj.ScaleMode = vbMillimeters
    
    obj.FontName = "courier new"
    obj.fontsize = 10
    
    P_CabInv obj, txtbxNumImob.Text, txtbxdataImob.Text, txtbxPagImob.Text
    P_CabImo obj
    P_CabInvProd obj, "0,00"
    
    
    
    
    With rPesq
        Do While Not .EOF
            If iCount = 45 Then
                iCount = 0
                iPag = iPag + 1
                P_RodaPeInvProd obj, FormatVal(cTotalTransp)
                yFim = obj.CurrentY
                
                P_CabInv obj, "03", "31/12/2005", Format(iPag, "00#")
                P_CabInvProd obj, FormatVal(cTotalTransp)
            End If
            
            
            P_ItemInv obj, "", !desc_IMO, !unidade_imo, FormatVal(!qtd_imo), FormatVal(!unit_imo), FormatVal(!total)
            
            
            
            iCount = iCount + 1
            cTotalTransp = cTotalTransp + !total
            .MoveNext
        Loop
        P_RodaPeInvFinal obj, FormatVal(cTotalTransp)
        .Close
    End With
    
    
Exit Function
trata_erro:
    MsgBox "Não foi possível concluir a impressão!" & Chr(13) & Err.Description, vbCritical
    
End Function

Sub imprimeTermoAbert(obj As Object)
    Dim sTexto As String
    sTexto = "          ESTAS FOLHAS  CONTÍNUAS,  NUMERADAS POR PROCESSAMENTO "
    sTexto = sTexto & "ELETRÔNICO " & Chr(13) & Chr(13) & "DE DADOS, DE Nºs " & txtbxIniTermo.Text & " a " & txtbxFinTermo
    sTexto = sTexto & ", " & IIf(optnAbertura.Value, "SERVIRÃO ", "SERVIRAM ") & " PARA REGISTRO DE INVENTÁRIO Nº " & txtbxInvent.Text
    sTexto = sTexto & " DA  " & Chr(13) & Chr(13) & "SOCIEDADE MERCANTIL ACIMA QUALIFICADA E " & IIf(optnAbertura.Value, "CONSTARÃO", "CONSTARAM") & " O ESTOQUE FINAL "
    sTexto = sTexto & " APURADO " & Chr(13) & Chr(13) & "EM " & txtbxData.Text & ". "

    P_CabInv obj, txtbxNumAbre.Text, txtbxdataImob.Text, txtbxFolhaAbre.Text
    If optnAbertura.Value Then
        P_CentraImpressao obj, "TERMO DE ABERTURA"
    Else
        P_CentraImpressao obj, "TERMO DE ENCERRAMENTO"
    End If
    obj.Print
    P_ImprX obj, "EMPRESA..................:" & strNome_Emp, 10, False
    obj.Print
    P_ImprX obj, "ENDEREÇO.................:" & strEndereco_Emp & " - " & strBairro_Emp, 10, False
    obj.Print
    P_ImprX obj, "CIDADE-UF................:" & strCidade_Emp & "-" & strUF_Emp, 10, False
    obj.Print
    P_ImprX obj, "CNPJ.....................:" & strCNPJ_Emp, 10, False
    obj.Print
    P_ImprX obj, "INSCRIÇÃO ESTADUAL.......:" & strIE_Emp, 10, False
    obj.Print
    P_ImprX obj, "REG. NA JUNTA COMERCIAL..:" & strRegJunta_Emp, 10, False
    obj.Print
    P_ImprX obj, "ATIVIDADE COMERCIAL......:" & strAtividade_Emp, 10, False
    obj.Print
    P_ImprX obj, "CÓDIGO ATIVIDADE.........:" & strCodAtividade_Emp, 10, False
    obj.Print
    obj.Print
    obj.Print
    P_ImprLinha obj
    obj.Print
    P_ImprX obj, sTexto, 0, False
    
    obj.Print
    obj.Print
    P_ImprX obj, strCidade_Emp & ", " & txtbxdataImob.Text, 20, False
    
    
    obj.Print
    obj.Print
    P_ImprX obj, "  ____________________________", 0, False
    P_ImprX obj, "   " & strNomeResp_Emp, 0, False
    P_ImprX obj, "   SÓCIO GERENTE", 0, False
    
    P_ImprLinha obj
    P_RodapeVis obj
End Sub

Private Sub cmnbtnImprTermoAbert_Click()
    imprimeTermoAbert frmVisualiza.Pic
    frmVisualiza.Show 1
    
    If frmVisualiza.Tag <> "" Then
        imprimeTermoAbert Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub

Private Sub cmndbtnImprImob_Click()
    'imprimeImob frmVisualiza.Pic
    
    'frmVisualiza.Show 1
    'If frmVisualiza.Tag <> "" Then
        imprimeImob Printer
    
        Printer.EndDoc
    'End If
    'Unload frmVisualiza
    
End Sub

Private Sub cmndbtnImprInv_Click()
    imprimeInv frmVisualiza.Pic
    
    frmVisualiza.Show 1
    
    If frmVisualiza.Tag <> "" Then
        imprimeInv Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub

Private Sub Command1_Click()
    frmInventario.Show 0
End Sub

Private Sub Command2_Click()
    frmEdiImo.Show 0
End Sub

Private Sub Form_Load()
    P_BuscaInvent
End Sub
Sub P_BuscaInvent()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select distinct datahora_inv from tab_invent order by datahora_inv desc"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        cmbxInvent.Clear
        Do While Not .EOF
            cmbxInvent.AddItem !DataHora_Inv
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub txtbxData_Change()
    txtbxdataImob.Text = txtbxData.Text
    txtbxDataAbre.Text = txtbxData.Text
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub

Private Sub txtbxdataImob_LostFocus()
    P_AjustaData txtbxdataImob
End Sub

Private Sub txtbxInvent_Change()
    txtbxNumImob.Text = txtbxInvent.Text
    txtbxNumAbre.Text = txtbxInvent.Text
End Sub

Private Sub txtbxInvent_GotFocus()
    P_SelecTexto txtbxInvent
End Sub

Private Sub txtbxInvent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxInvent_LostFocus()
    P_AjustaTexto txtbxInvent
End Sub

Private Sub txtbxPag_GotFocus()
    P_SelecTexto txtbxPag
End Sub

Private Sub txtbxPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxPag_LostFocus()
    P_AjustaTexto txtbxPag
    txtbxPag.Text = FormatVal(txtbxPag.Text, "00#")
End Sub

Private Sub txtbxPagImob_GotFocus()
    P_SelecTexto txtbxPagImob
End Sub

Private Sub txtbxPagImob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPagImob_LostFocus()
    txtbxPagImob.Text = FormatVal(txtbxPagImob.Text, "00#")
    
End Sub
