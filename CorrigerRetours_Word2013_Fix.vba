Option Explicit

' ==============================================================================
' Macro : CorrigerRetours_Word2013
' Auteur : Jules (Assistant IA)
' Description : Corrige les retours à la ligne intempestifs dans Word 2013.
' ==============================================================================

Sub CorrigerRetours_Word2013()
    Dim doc As Document
    Set doc = ActiveDocument

    ' Optimisation de la performance
    Application.ScreenUpdating = False

    ' Définition des classes de caractères pour le français
    Dim min As String
    min = "a-zàâäæçéèêëîïôöœùûüÿ"
    Dim maj As String
    maj = "A-ZÀÂÄÆÇÉÈÊËÎÏÔÖŒÙÛÜŸ"

    ' --------------------------------------------------------------------------
    ' Règle 6 : minuscule + CHR(13) + minuscule -> espace
    ' --------------------------------------------------------------------------
    AppliquerRemplacement doc, "([" & min & "])^13([" & min & "])", "\1 \2"

    ' --------------------------------------------------------------------------
    ' Règle 7 : minuscule + CHR(13) + Majuscule -> ". "
    ' --------------------------------------------------------------------------
    AppliquerRemplacement doc, "([" & min & "])^13([" & maj & "])", "\1. \2"

    ' --------------------------------------------------------------------------
    ' Règle 4 : minuscule + "." + CHR(13) + (minuscule ou «) -> espace
    ' --------------------------------------------------------------------------
    ' Utilisation de [.] au lieu de \. pour éviter les erreurs d'interprétation
    AppliquerRemplacement doc, "([" & min & "])[.]^13([" & min & "«])", "\1 \2"

    ' Rétablir la mise à jour de l'écran
    Application.ScreenUpdating = True

    MsgBox "Traitement terminé avec succès.", vbInformation, "Correction des retours"
End Sub

' ------------------------------------------------------------------------------
' Procédure utilitaire pour appliquer un remplacement par wildcards
' ------------------------------------------------------------------------------
Private Sub AppliquerRemplacement(doc As Document, strFind As String, strReplace As String)
    Dim rng As Range
    ' On ré-initialise la plage à chaque appel pour couvrir tout le document
    Set rng = doc.Content

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = strFind
        .Replacement.Text = strReplace
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
        .MatchCase = True

        ' Exécution du remplacement global
        .Execute Replace:=wdReplaceAll
    End With
End Sub
