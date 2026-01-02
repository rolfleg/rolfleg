' ---------------------------------------------------------------------------------------

Option Explicit

Private Sub Document_Open()
    AjouterMenuContextuelAvecSousMenu
End Sub

Public Sub AjouterMenuContextuelAvecSousMenu()
    Dim cb As CommandBar
    Dim popup As CommandBarPopup
    Dim bouton As CommandBarButton

    ' Supprimer l'ancien sous-menu si déjà présent
    On Error Resume Next
    Application.CommandBars("Text").Controls("? Mes Macros").Delete
    On Error GoTo 0

    ' Récupérer le menu contextuel (clic droit dans une zone de texte)
    Set cb = Application.CommandBars("Text")

    ' Ajouter un sous-menu au clic droit
    Set popup = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    With popup
        .Caption = "? Mes Macros"
        .Tag = "MonMenuMacros"
        .BeginGroup = True
    End With

    ' Ajouter le bouton 1
    Set bouton = popup.Controls.Add(Type:=msoControlButton)
    With bouton
        .Caption = "Mettre en gras"
        .OnAction = "MettreParagraphesEnNoirGras"
        .FaceId = 59
    End With
Rem
    ' Ajouter le bouton 2
    Set bouton = popup.Controls.Add(Type:=msoControlButton)
    With bouton
        .Caption = "Interligne 1,5"
  '     .OnAction = "call.module3.InterligneUnEtDemiWord"
   '     .OnAction = "MacroInterligneUnEtDemi"
        .FaceId = 138
   End With
    ' Ajouter le bouton 3
    Set bouton = popup.Controls.Add(Type:=msoControlButton)
    With bouton
        .Caption = "Page A4. "
        .OnAction = "AjusterPageA4"
        .FaceId = 294
    End With
     ' Ajouter le bouton 4
    Set bouton = popup.Controls.Add(Type:=msoControlButton)
    With bouton
        .Caption = "MarqueParagraphParEspaceSafe "
        .OnAction = "call.module12.MarqueParagraphParEspaceSafe"
        .FaceId = 295
    End With
Rem
 ' Ajouter le bouton 5
    Set bouton = popup.Controls.Add(Type:=msoControlButton)
    With bouton
        .Caption = "Retraits Ligne. "
        .OnAction = "FormaterRetraitsEtAlignementGauche"
        .FaceId = 297
    End With

    MsgBox "Sous-menu contextuel ajouté avec succès.", vbInformation
End Sub