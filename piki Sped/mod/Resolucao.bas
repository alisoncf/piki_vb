Attribute VB_Name = "Resolucao"
    Public Xtwips As Integer, Ytwips As Integer
    Public Xpixels As Integer, Ypixels As Integer

    Type FRMSIZE
        Height As Long
        Width As Long
    End Type
    Public RePosForm As Boolean
    Public DoResize As Boolean
    
    Public MyForm As FRMSIZE
    Public DesignX As Integer
    Public DesignY As Integer
   Sub Resize_For_Resolution(ByVal SFX As Single, _
      ByVal SFY As Single, MyForm As Form)
      
      Dim I As Integer
      Dim SFFont As Single

      SFFont = (SFX + SFY) / 2  ' escala média
      ' Tamanho dos controles para a nova resolução
      On Error Resume Next  '
      With MyForm
        For I = 0 To .count - 1
         If TypeOf .Controls(I) Is ComboBox Then 'Combobox não altera a propriedade Height
           .Controls(I).Left = .Controls(I).Left * SFX
           .Controls(I).Top = .Controls(I).Top * SFY
           .Controls(I).Width = .Controls(I).Width * SFX
         Else
           .Controls(I).Move .Controls(I).Left * SFX, _
            .Controls(I).Top * SFY, _
            .Controls(I).Width * SFX, _
            .Controls(I).Height * SFY
         End If
           ' Redimensiona e reposiciona antes de alterar o tamanho da fonte
           .Controls(I).FontSize = .Controls(I).FontSize * SFFont
        Next I
        If RePosForm Then
          ' Redimensiona o formulario
          .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
        End If
      End With
   End Sub

Sub ajustaForm(frm As Form)
    Dim ScaleFactorX As Single, ScaleFactorY As Single  ' fatores da escala
      ' Tamanho do formulario em Pixels na resolução de trabalho
      DesignX = 800
      DesignY = 600
      RePosForm = True
      DoResize = False
      ' define os valores da tela
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips
      Xpixels = Screen.Width / Xtwips

      ' Determina os valores da escala
      ScaleFactorX = (Xpixels / DesignX)
      ScaleFactorY = (Ypixels / DesignY)
      ScaleMode = 1  ' twips
      
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, frm
      'Label1.caption = "Resolução Atual " & Str$(Xpixels) + _
       "  por " + Str$(Ypixels)
      MyForm.Height = frm.Height ' Lembra o tamanho atual
      MyForm.Width = frm.Width
End Sub

Sub Resizeform(frm As Form)
    Dim ScaleFactorX As Single, ScaleFactorY As Single

      If Not DoResize Then  ' Para evitar um loop sem fim
         DoResize = True
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = frm.Width / MyForm.Width   ' Quanto mudar ?
      ScaleFactorY = frm.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = frm.Height ' Lembra o tamanho atual
      MyForm.Width = frm.Width
End Sub
