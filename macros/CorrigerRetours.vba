Option Explicit

' ==============================================================================
' Macro : CorrigerRetours_Word2013
' Description : Corrige les retours paragraphe intempestifs dans un document Word.
'               Règle 6 : Joint les paragraphes si une minuscule est suivie d'un retour et d'une minuscule.
'               Règle 7 : Ajoute un point si une minuscule est suivie d'un retour et d'une majuscule.
'               Règle 4 : Supprime un point et joint les paragraphes si suivis d'une minuscule ou d'une guillemet.
' ==============================================================================

Sub CorrigerRetours_Word2013()
    Dim low As String, cap As String
    ' Définition des plages de caractères incluant les accents français
    low = "a-zàâäæçéèêëîïôöœùûüÿ"
    cap = "A-ZÀ-ÂÄÆÇÉÈÊËÎÏÔÖŒÙÛÜŸ"

    ' --- Nettoyage préliminaire ---
    ' Supprimer les espaces et tabulations autour des retours paragraphe pour un matching fiable.
    ' On utilise (^13) pour capturer le retour et le réinsérer avec \1, ce qui préserve sa nature.
    AppliquerRegle "[ ^t]@(^13)", "\1"
    AppliquerRegle "(^13)[ ^t]@", "\1"

    ' Traiter les sauts de ligne manuels en les convertissant en retours paragraphe
    AppliquerRegleSimple "^l", "^p"

    ' --- Application des règles de ponctuation (Mode Wildcards) ---

    ' Règle 6 : minuscule + return + minuscule -> espace
    AppliquerRegle "([" & low & "])^13([" & low & "])", "\1 \2"

    ' Règle 7 : minuscule + return + Majuscule -> ". "
    AppliquerRegle "([" & low & "])^13([" & cap & "])", "\1. \2"

    ' Règle 4 : minuscule + "." + return + (minuscule ou «) -> minuscule + " " + (minuscule ou «)
    AppliquerRegle "([" & low & "])[.]^13([" & low & "«])", "\1 \2"

    ' Nettoyage final : supprimer les doubles espaces éventuellement créés
    AppliquerRegleSimple "  ", " "

    MsgBox "Traitement terminé. Les retours paragraphe ont été corrigés.", vbInformation, "Correction"
End Sub

Private Sub AppliquerRegle(ByVal sTrouver As String, ByVal sRemplacer As String)
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = sTrouver
        .Replacement.Text = sRemplacer
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Sub AppliquerRegleSimple(ByVal sTrouver As String, ByVal sRemplacer As String)
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = sTrouver
        .Replacement.Text = sRemplacer
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub
