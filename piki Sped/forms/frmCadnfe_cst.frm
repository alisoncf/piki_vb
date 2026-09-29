VERSION 5.00
Begin VB.Form frmCadnfe_cst
 AutoRedraw = -1
Caption = "frmCadnfe_cst"
ClientHeight = 7635
ClientLeft = 60
ClientTop = 345
ClientWidth = 11460
KeyPreview = -1
LinkTopic = "Form1"
ScaleHeight = 7635
ScaleWidth = 11460
Begin VB.PictureBox pcTool 
  Align = 2  
  Appearance = 0
  BackColor = &H80000005
  BorderStyle = 0   
  ForeColor = &H80000008
  Height = 855
  Left = 0
  ScaleHeight = 855
  ScaleWidth = 6930
  TabIndex = 2
  Top = 4095
  Width = 6930
  Begin VB.CommandButton cmdSair
  Caption = "Sair"
  Height = 615
  Left = 5520
  TabIndex = 4
  Top = 120
  Width = 1335
  End
  Begin VB.CommandButton cmdSalvar
  Caption = "Gravar"
  Height = 615
  Left = 4080
  TabIndex = 3
  Top = 120
  Width = 1335
  End
  End

Begin VB.Label lblcodigo
AutoSize = -1
BackStyle = 0
Caption = "codigo"
Height = 195
Left = 240
TabIndex = 0
Top = 40
Width = 630
End
Begin VB.TextBox txtcodigo
Appearance = 0
Height = 285
Left = 240
TabIndex = 1
Top =280
Width = 2775
End

Begin VB.Label lblnome
AutoSize = -1
BackStyle = 0
Caption = "nome"
Height = 195
Left = 240
TabIndex = 0
Top = 640
Width = 630
End
Begin VB.TextBox txtnome
Appearance = 0
Height = 285
Left = 240
TabIndex = 1
Top =880
Width = 2775
End

End
Attribute VB_Name = "frmCadnfe_cst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtcodigo_GotFocus()
 ganhaFoco txtCampo1
End Sub
Private Sub txtcodigo_Keydown(keycode as integer, shift as integer)
 if keycode=vbkeytab then sendkeys chr(9)
End Sub
Private Sub txtcodigo_LostFocus()
 ajustatexto txtcodigo
End Sub
Private Sub txtnome_GotFocus()
 ganhaFoco txtCampo1
End Sub
Private Sub txtnome_Keydown(keycode as integer, shift as integer)
 if keycode=vbkeytab then sendkeys chr(9)
End Sub
Private Sub txtnome_LostFocus()
 ajustatexto txtnome
End Sub

