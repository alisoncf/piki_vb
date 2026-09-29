VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCliAniv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Clientes Aniversariantes"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   10800
      ScaleHeight     =   6465
      ScaleWidth      =   2505
      TabIndex        =   11
      Top             =   960
      Width           =   2535
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Aniversário:"
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   2295
         Begin VB.TextBox txtbxDataFin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   480
            TabIndex        =   26
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   480
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Até:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Movimento"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   2295
         Begin VB.CheckBox chkCompra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Compra"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkAcerto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Acerto"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   2295
         Begin VB.TextBox txtbxE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   600
            TabIndex        =   19
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtbxEntre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   600
            TabIndex        =   17
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Último Movimento:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   0
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&E:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   150
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Entre:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Débitos:"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   2295
         Begin VB.CheckBox chkAVencer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Sem débitos A Vencer"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkVencidos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Sem débitos Vencidos"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   1  'Checked
            Width           =   1935
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   6495
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11456
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   13335
      TabIndex        =   0
      Top             =   0
      Width           =   13365
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ordem:"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   4560
         TabIndex        =   8
         Top             =   0
         Width           =   1695
         Begin VB.OptionButton optnAniv 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Aniversário"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optnNome 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Nome"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmCliAniv.frx":0000
         Left            =   120
         List            =   "frmCliAniv.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   6480
         Picture         =   "frmCliAniv.frx":003F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   8160
         Picture         =   "frmCliAniv.frx":03C9
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   11520
         Picture         =   "frmCliAniv.frx":07D6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico"
         Height          =   615
         Left            =   9840
         Picture         =   "frmCliAniv.frx":08CB
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCliAniv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sTipo As String
    Dim sCodigo As String
    Dim iCount As Integer
    Dim sVencidos As String
    Dim sAvencer As String
    msflxgrd.rows = 1
    
    sVencidos = F_NotIn("distinct cod_ter", "tab_itdoc as ti, tab_doc as td", "ti.seq_doc=td.seq_doc and tipo_ter='C' and ti.reman_itd>0 and ti.venc_itd<=" & F_GravaData(Date))
    sAvencer = F_NotIn("distinct cod_ter", "tab_itdoc as ti, tab_doc as td", "ti.seq_doc=td.seq_doc and tipo_ter='C' and ti.reman_itd>0 and ti.venc_itd>=" & F_GravaData(Date))
    
    
    
    If cmbxTipoTer.ListIndex = 1 Then
        strSQL = "select tc.cod_rep as cod, tc.nome_rep as nome,(tc.nasc_rep) as nasc from tab_representantes as tc "
        strSQL = strSQL & " where right(nasc_rep,5)>='" & Format(txtbxDataIni.Text, "mm-dd") & "'"
        strSQL = strSQL & " and right(nasc_rep,5)<='" & Format(txtbxDataFin.Text, "mm-dd") & "'"
        If optnAniv.Value = False Then
            strSQL = strSQL & " order by nome_rep,nasc "
        Else
            strSQL = strSQL & " order by nasc,nome_rep "
        End If
        sTipo = "R"
    ElseIf cmbxTipoTer.ListIndex = 2 Then
        strSQL = "select tc.cod_fun as cod, tc.nome_fun as nome,(tc.nasc_fun) as nasc from tab_funcionarios as tc "
        strSQL = strSQL & " where right(nasc_fun,5)>='" & Format(txtbxDataIni.Text, "mm-dd") & "'"
        strSQL = strSQL & " and right(nasc_fun,5)<='" & Format(txtbxDataFin.Text, "mm-dd") & "'"
        If optnAniv.Value Then
            strSQL = strSQL & " order by nasc,nome_fun"
        Else
        
            strSQL = strSQL & " order by nome_fun,nasc "
        End If
        sTipo = "U"
    Else
        strSQL = "select tc.cod_cli as cod, max(tv.dataefet_ven) as ultven,count(tv.num_ven) as vendas, tc.nome_cli as nome,tc.datanasc_cli as nasc "
        strSQL = strSQL & " from tab_clientes as tc,tab_vendas as tv  "
        strSQL = strSQL & " where tc.cod_cli=tv.cod_cli  "
        strSQL = strSQL & " and tipocli_ven='C' and tc.cod_cli<>0 "
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and right(datanasc_cli,5)>='" & Format(txtbxDataIni.Text, "mm-dd") & "'"
        End If
        If txtbxDataFin.Text <> "" Then
            strSQL = strSQL & " and right(datanasc_cli,5)<='" & Format(txtbxDataFin.Text, "mm-dd") & "'"
        End If
        
        If txtbxEntre.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxEntre.Text)
        End If
        If txtbxE.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxE.Text)
        End If
        
        If chkAVencer.Value = 1 Then
            strSQL = strSQL & " and tc.cod_cli not in" & sAvencer
        End If
        If chkVencidos.Value = 1 Then
            strSQL = strSQL & " and tc.cod_cli not in" & sVencidos
        End If
        
        
        
        
        
        strSQL = strSQL & " group by tv.cod_cli "
        If optnAniv.Value Then
            strSQL = strSQL & " order by nasc,nome_cli "
        Else
            strSQL = strSQL & " order by nome_cli,nasc  "
        End If
        sTipo = "C"
        
        
    End If
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rPesq
        Do While Not .EOF
            strSQL = !cod & Chr(9) & !nome & Chr(9) & Format(!nasc, "DD/MM") & Chr(9) & Chr(9) & !ultven
            sCodigo = sCodigo & "'" & !cod & "'"
            
            msflxgrd.AddItem strSQL
            .MoveNext
            If Not .EOF Then
                sCodigo = sCodigo & ","
            End If
        Loop
        .Close
    End With
    
    
    strSQL = "select tt.* from tab_telefones as tt, tab_clientes as tc where tipo_ter='" & sTipo & "' "
    strSQL = strSQL & " and tt.cod_cli=tc.cod_cli "
    If sCodigo <> "" Then
        strSQL = strSQL & " and tt.cod_cli in(" & sCodigo & ") "
    End If
    strSQL = strSQL & " order by nome_cli "
    
    
    P_Status "Buscando telefones..."
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rPesq
        Do While Not .EOF
            For iCount = 1 To msflxgrd.rows - 1
                msflxgrd.Col = 0
                msflxgrd.Row = iCount
                If msflxgrd.Text = !cod_cli Then
                    msflxgrd.Col = 3
                    msflxgrd.Text = msflxgrd.Text & " " & !num_tel
                End If
            Next
            
            P_Status "Verificando Telefone " & .AbsolutePosition + 1 & "/" & .RecordCount
            
            .MoveNext
        Loop
    End With
    P_Status ""
    
End Sub
Sub buscaCli()
    Dim strSQL As String
    Dim rPesq As Recordset
    
    
End Sub




Private Sub cmndbtnImprimir_Click()
    P_MsflxToHtml "Listagem de Aniversariantes de " & txtbxDataIni.Text & " Até " & txtbxDataFin.Text, _
        msflxgrd, cmbxTipoTer.Text
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    
End Sub

Private Sub Form_Load()
    With msflxgrd
        .FixedCols = 0
        .rows = 1
        .Cols = 5
        .Col = 0: .Text = "Código"
        .Col = 1: .Text = "Nome"
        .Col = 2: .Text = "Data Nasc."
        .Col = 3: .Text = "Telefones"
        .Col = 4: .Text = "Última Compra"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 3500
        .ColWidth(2) = 1000
        .ColWidth(3) = 4500
        .ColWidth(4) = 2500
        
        
        
        
        
        
    End With
    cmbxTipoTer.ListIndex = 0
End Sub

Private Sub optnAniv_Click()
    busca
End Sub

Private Sub optnNome_Click()
    busca
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxE_GotFocus()
    P_SelecTexto txtbxE
End Sub

Private Sub txtbxE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxE_LostFocus()
    P_AjustaData txtbxE
End Sub

Private Sub txtbxEntre_GotFocus()
    P_SelecTexto txtbxEntre
End Sub

Private Sub txtbxEntre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    txtbxDataIni.Text = Format(txtbxDataIni.Text, "dd/mm")
End Sub
Private Sub txtbxdatafin_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxdatafin_LostFocus()
    txtbxDataFin.Text = Format(txtbxDataFin.Text, "dd/mm")
End Sub

Private Sub txtbxEntre_LostFocus()
    P_AjustaData txtbxEntre
End Sub
