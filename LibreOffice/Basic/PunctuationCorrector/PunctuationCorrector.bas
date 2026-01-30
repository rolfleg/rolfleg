REM  *****  BASIC  *****

' Module pour la correction de la ponctuation selon des règles spécifiques.
' Auteur: Jules (Assistant IA)
' Date: 2024-07-29

Global g_oNomsPropres As Object ' Dictionnaire des noms propres

' Procédure principale qui orchestre la correction.
Sub CorrectPunctuation()
    On Error GoTo ErrorHandler

    Dim oDoc As Object
    oDoc = ThisComponent

    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox "Cette macro ne peut être exécutée que sur un document texte.", 16, "Erreur"
        Exit Sub
    End If

    LockControllers(oDoc)

    ' Charger le dictionnaire de noms propres
    If Not LoadNomsPropres() Then
        MsgBox "Impossible de charger le dictionnaire des noms propres. Vérifiez que le fichier NomsPropres.txt existe au même endroit que la macro.", 16, "Erreur"
        GoTo ErrorHandlerExit ' Sortir proprement
    End If

    ' Appliquer les règles de correction
    AppliquerRegles123(oDoc)
    AppliquerRegle4(oDoc)

    MsgBox "Correction de la ponctuation terminée.", 64, "Succès"

ErrorHandlerExit:
    UnlockControllers(oDoc)
    Exit Sub

ErrorHandler:
    Dim errorMsg As String
    errorMsg = "Une erreur inattendue est survenue." & Chr(13) & _
               "Module: PunctuationCorrector" & Chr(13) & _
               "Ligne: " & Erl & Chr(13) & _
               "Erreur " & Err & ": " & Err.Description
    MsgBox errorMsg, 16, "Erreur d'exécution"
    Resume ErrorHandlerExit
End Sub

' Verrouille les contrôleurs du document pour améliorer les performances.
Sub LockControllers(oDoc As Object)
    If Not IsNull(oDoc) AND oDoc.hasControllers() Then
        oDoc.lockControllers
    End If
End Sub

' Déverrouille les contrôleurs du document et rafraîchit la vue.
Sub UnlockControllers(oDoc As Object)
    If Not IsNull(oDoc) AND oDoc.hasControllers() Then
        oDoc.unlockControllers
    End If
End Sub

' Charge le fichier NomsPropres.txt dans le dictionnaire global g_oNomsPropres.
Function LoadNomsPropres() As Boolean
    On Error GoTo LoadError

    Dim oPathSubst As Object
    Dim sUserScriptsUrl As String
    Dim sDictionaryUrl As String
    Dim oFileAccess As Object
    Dim oInputStream As Object
    Dim oTextInputStream As Object
    Dim sLine As String

    LoadNomsPropres = False

    ' Initialiser le dictionnaire global
    g_oNomsPropres = CreateUnoService("com.sun.star.container.EnumerableMap")

    ' 1. Obtenir le chemin du répertoire des scripts de l'utilisateur.
    ' C'est la méthode la plus robuste pour trouver les fichiers associés à une macro.
    oPathSubst = createUnoService("com.sun.star.comp.framework.PathSubstitution")
    sUserScriptsUrl = oPathSubst.getSubstituteVariableValue("$(user)") & "/Scripts/Basic/PunctuationCorrector/"

    ' 2. Construire le chemin complet du fichier dictionnaire en format URL.
    sDictionaryUrl = sUserScriptsUrl & "NomsPropres.txt"

    ' 3. Lire le fichier dictionnaire en utilisant les services UNO.
    oFileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not oFileAccess.exists(sDictionaryUrl) Then
        Exit Function ' Le fichier n'existe pas, la fonction renverra False.
    End If

    oInputStream = oFileAccess.openFileRead(sDictionaryUrl)
    oTextInputStream = createUnoService("com.sun.star.io.TextInputStream")
    oTextInputStream.setInputStream(oInputStream)
    oTextInputStream.setEncoding("UTF-8") ' Assurer la compatibilité avec les accents.

    ' 4. Remplir le dictionnaire avec le contenu du fichier.
    While Not oTextInputStream.isEOF()
        sLine = Trim(oTextInputStream.readLine())
        If sLine <> "" Then
            ' Stocker les noms en minuscules pour une comparaison insensible à la casse.
            g_oNomsPropres.put(LCase(sLine), 1)
        End If
    Wend

    oTextInputStream.closeInput()
    LoadNomsPropres = True
    Exit Function

LoadError:
    ' En cas d'erreur (ex: permissions de lecture), la fonction renverra False.
    LoadNomsPropres = False
End Function

' Applique les règles 1, 2 et 3 en utilisant la recherche et le remplacement par expressions régulières.
Sub AppliquerRegles123(oDoc As Object)
    Dim oSearchDesc As Object
    Dim oFound As Object
    Dim i As Long
    Dim oRange As Object
    Dim sFullMatch As String, sCleanedMatch As String
    Dim aWords() As String
    Dim sWord1 As String, sWord2 As String
    Dim sReplacement As String
    Dim bIsMaj As Boolean, bIsMinus As Boolean, bEstNomPropre As Boolean

    ' 1. Créer un descripteur de recherche.
    oSearchDesc = oDoc.createSearchDescriptor()

    ' 2. Définir l'expression régulière.
    '    - \b\w*[[:lower:]]  : un mot se terminant par une minuscule. \b = limite de mot.
    '    - \.?\s+             : un point optionnel (\.?) suivi d'un ou plusieurs espaces (\s+).
    '    - \b\w+\b            : un mot complet.
    oSearchDesc.setSearchString("\b\w*[[:lower:]]\b\.?\s+\b\w+\b")
    oSearchDesc.setPropertyValue("SearchRegularExpression", True)

    ' 3. Trouver toutes les occurrences.
    oFound = oDoc.findAll(oSearchDesc)

    If IsNull(oFound) Or oFound.getCount() = 0 Then
        Exit Sub ' Aucune correspondance trouvée.
    End If

    ' 4. Parcourir les résultats en ordre inverse pour éviter les problèmes d'indexation.
    For i = oFound.getCount() - 1 To 0 Step -1
        oRange = oFound.getByIndex(i)
        sFullMatch = oRange.getString()

        ' 5. Extraire les deux mots de la chaîne trouvée.
        sCleanedMatch = Replace(sFullMatch, ".", " ") ' Remplacer le point pour faciliter le split.
        aWords = Split(Trim(sCleanedMatch), " ")

        ' Filtrer les chaînes vides potentielles résultant de multiples espaces.
        Dim tempArr() As String
        Dim word As Variant
        Dim count As Integer: count = -1
        For Each word In aWords
            If Trim(word) <> "" Then
                count = count + 1
                ReDim Preserve tempArr(count)
                tempArr(count) = Trim(word)
            End If
        Next word

        If count < 1 Then Continue For ' Si on n'a pas au moins deux mots, on ignore.
        aWords = tempArr

        sWord1 = aWords(0)
        sWord2 = aWords(1)

        ' 6. Analyser le deuxième mot et appliquer les règles.
        bEstNomPropre = IsNomPropre(sWord2)

        Dim sFirstChar As String: sFirstChar = Left(sWord2, 1)
        bIsMaj = (sFirstChar >= "A" And sFirstChar <= "Z") Or (sFirstChar >= "À" And sFirstChar <= "Ö") ' Inclure les accents majuscules
        bIsMinus = (sFirstChar >= "a" And sFirstChar <= "z") Or (sFirstChar >= "à" And sFirstChar <= "ö") ' Inclure les accents minuscules

        ' Règle 1: Mot se terminant par minuscule + Nom Propre
        ' -> Pas de point, un seul espace.
        If bEstNomPropre Then
            sReplacement = sWord1 & " " & sWord2

        ' Règle 2: Mot se terminant par minuscule + Mot commençant par minuscule
        ' -> Pas de point, un seul espace.
        ElseIf bIsMinus Then
            sReplacement = sWord1 & " " & sWord2

        ' Règle 3: Mot se terminant par minuscule + Mot commençant par majuscule (qui n'est pas un nom propre)
        ' -> Mettre un point et garder l'espace.
        ElseIf bIsMaj And Not bEstNomPropre Then
            sReplacement = sWord1 & ". " & sWord2

        ' Cas par défaut: ne rien changer si aucune règle ne s'applique.
        Else
            sReplacement = sFullMatch
        End If

        ' 7. Appliquer le remplacement si nécessaire.
        If sReplacement <> sFullMatch Then
            oRange.setString(sReplacement)
        End If
    Next i
End Sub

' Applique la règle 4 en fusionnant les paragraphes sous conditions.
Sub AppliquerRegle4(oDoc As Object)
    Dim oParas As Object
    Dim i As Long
    Dim oPara As Object, oNextPara As Object
    Dim sParaText As String, sNextParaText As String
    Dim oCursor As Object
    Dim sNextWord As String
    Dim sFirstCharNext As String

    oParas = oDoc.getText().getParagraphs()

    ' Parcourir les paragraphes en ordre inverse, de l'avant-dernier au premier.
    For i = oParas.getCount() - 2 To 0 Step -1
        oPara = oParas.getByIndex(i)
        oNextPara = oParas.getByIndex(i + 1)

        sParaText = Trim(oPara.getString())
        sNextParaText = Trim(oNextPara.getString())

        ' Condition 1: Le paragraphe actuel se termine par une minuscule suivie d'un point.
        If Right(sParaText, 1) = "." Then
            Dim sCharBeforePoint As String
            If Len(sParaText) > 1 Then
                sCharBeforePoint = Mid(sParaText, Len(sParaText) - 1, 1)

                ' Vérifier si le caractère est une minuscule (y compris les accents).
                If (sCharBeforePoint >= "a" And sCharBeforePoint <= "z") Or (sCharBeforePoint >= "à" And sCharBeforePoint <= "ö") Then

                    ' Condition 2: Le paragraphe suivant n'est pas vide.
                    If sNextParaText <> "" Then
                        sFirstCharNext = Left(sNextParaText, 1)

                        ' Condition 3: Le paragraphe suivant commence par une minuscule ou un guillemet ouvrant.
                        If (sFirstCharNext >= "a" And sFirstCharNext <= "z") Or (sFirstCharNext >= "à" And sFirstCharNext <= "ö") Or sFirstCharNext = "«" Then

                            ' Condition 4: Le premier mot du paragraphe suivant n'est pas un nom propre.
                            sNextWord = Split(sNextParaText, " ")(0)
                            ' Nettoyer le mot s'il commence par un guillemet.
                            If Left(sNextWord, 1) = "«" Then sNextWord = Mid(sNextWord, 2)

                            If Not IsNomPropre(sNextWord) Then
                                ' Toutes les conditions sont remplies, on fusionne.

                                ' 1. Créer un curseur qui englobe la fin du premier paragraphe
                                '    et la totalité du second.
                                oCursor = oPara.getEnd()
                                oCursor.goLeft(1, False) ' Se placer avant le point.
                                oCursor.gotoRange(oNextPara.getEnd(), True) ' Sélectionner jusqu'à la fin du paragraphe suivant.

                                ' 2. Construire la chaîne de remplacement.
                                Dim sMergedText As String
                                sMergedText = " " & sNextParaText ' Espace + texte du para suivant.

                                ' 3. Remplacer la sélection (point + saut de para + texte du para suivant)
                                '    par la nouvelle chaîne.
                                oCursor.setString(sMergedText)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Vérifie si un mot est un nom propre (insensible à la casse).
Function IsNomPropre(sWord As String) As Boolean
    IsNomPropre = False ' Placeholder
    If IsNull(g_oNomsPropres) Or IsEmpty(g_oNomsPropres) Then Exit Function
    If g_oNomsPropres.hasByName(LCase(sWord)) Then
        IsNomPropre = True
    End If
End Function
