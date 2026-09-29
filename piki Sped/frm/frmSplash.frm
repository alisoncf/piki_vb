VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   5445
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2640
      X2              =   7485
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblAviso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscando Informações..."
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   3840
      TabIndex        =   0
      Top             =   5160
      Width           =   1770
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iVezes As Integer

Private Sub Form_Activate()
    busca
End Sub

Sub busca()


     Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    P_CentraForm Me
    On Error Resume Next
'    If strNome_Usu = "ALISON" Then
'        frmEscolheBanco.Show 1
'    End If
    
    'strCodCid_Emp = "5218300"  'posse
    
    P_Status "Configurando Parâmetros do Sistema..."
    strSQL = "select * from tab_parametros "
    'strImprFiscal = "D" 'daruma
    rPesq.Open strSQL, conn_client
    
    If rPesq.RecordCount = 0 Then
        
        
    Else
        With rPesq
            strNumero_Emp = !numeroemp_par
            strComplemento_Emp = !complementoemp_par
            strLogradouro_Emp = !logradouroemp_par
            
            'strNumero_Emp = "S/N"
            
            strNome_Emp = !Nome_Emp
            strEndereco_Emp = !endereco_Emp
            strBairro_Emp = !Bairro_Emp
            strCidade_Emp = !Cidade_Emp
            strIP_Servidor = !ipserver_par
            strUF_Emp = !uf_emp
            strCEP_Emp = !cep_Emp
            strFone_Emp = !Fone_Emp
            strFax_Emp = !Fax_Emp
            strCNPJ_Emp = F_SoNumero(!CNPJ_Emp)
            strIE_Emp = !Ie_Emp
            strIP_TerminalTEF = !ipterminaltef_par
            strId_Estab = !idestabtef_par
            strFantasia_Emp = !fantasia_emp
            strNomeResp_Emp = !NomeResp_Emp
            strNomeMoeda_Par = !NomeMoeda_Par
            strsimbMoeda_Par = !simbMoeda_Par
            strPluraMoeda_Par = !PLuralMoeda_PAr
            crcTxFederativa_Par = !txfederativa_par
            crcTxAdm = !txadm_par
            crcTxMargemSeg_Par = !margemseg_par
            strPlaVendasEfetuadas = !plavendas_par
            strPlaRecebDupl = !plarecdupl_par
            strPlaPagDupl = !plapagdupl_par
            strPlaNaoClassC = !planaoclassc_par
            strPlaNaoClassD = !planaiclassd_parali
            
            strArquivoServidor = !caminhoexec_par
            If !usatef_par = "S" Then
                bUsaTEF = False
            Else
                bUsaTEF = False
            End If
            strArquivoServidor = !caminhoexec_par
            .Close
        End With
    End If
    P_BuscaDefUsu
'    P_AjustaVenda
    
    'Bematech_FI_AbrePortaSerial
   
   
    
    P_Status ""
    Unload Me
    'mdiFrm.StatusBar1.Panels(2).Text = "type:(" & F_BuscaUmCampo("tipo_usu", _
        "tab_usuarios", "cod_usu='" & strCod_Usu & "'") & ")"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Unload Me
End Sub

Private Sub Timer1_Timer()
    iVezes = iVezes + 1
    'lbl.caption = lbl.caption & "                     "
    'lbl.Refresh
'    Picture1.Left = Picture1.Left - 250
    If iVezes >= 11 Then
        Timer1.Enabled = False
        'Form_KeyDown 32, 0
        lblAviso.caption = "Tecle Algo para Continuar"
    End If
End Sub
