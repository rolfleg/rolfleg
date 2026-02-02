' Macro de correction de ponctuation pour LibreOffice Writer
' Cette macro gère 4 règles spécifiques basées sur la casse et un dictionnaire de noms propres.
' Auteur: Jules
' Date: 2024-05-23

Sub CorrigerPonctuationAvancee()
    Dim oDoc As Object
    oDoc = ThisComponent

    ' Vérification du type de document
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox "Cette macro doit être exécutée dans LibreOffice Writer.", 16, "Erreur"
        Exit Sub
    End If

    Dim oProperNouns As Object
    oProperNouns = LoadProperNouns()

    ' Verrouillage des contrôleurs pour améliorer les performances
    oDoc.lockControllers()
    On Error GoTo Cleanup

    ' Règle 1, 2, 3 : Transitions entre mots dans le même paragraphe
    ProcessRules1To3(oDoc, oProperNouns)

    ' Règle 4 : Transitions entre paragraphes
    ProcessRule4(oDoc, oProperNouns)

Cleanup:
    oDoc.unlockControllers()
    If Err <> 0 Then
        MsgBox "Une erreur est survenue : " & Err.Description, 48, "Erreur"
    Else
        MsgBox "Correction de la ponctuation terminée.", 64, "Succès"
    End If
End Sub

' Charge le dictionnaire des noms propres depuis le profil utilisateur
Function LoadProperNouns() As Object
    Dim oMap As Object
    oMap = CreateUnoService("com.sun.star.container.EnumerableMap")
    oMap.create("string", "boolean")

    Dim oPathSub As Object
    oPathSub = CreateUnoService("com.sun.star.util.PathSubstitution")
    Dim sUserPath As String
    sUserPath = oPathSub.getSubstituteVariableValue("$(user)")

    ' Chemin spécifié : Scripts/python/noms_propres.txt dans le dossier utilisateur
    ' $(user) correspond au dossier de profil LibreOffice de l'utilisateur
    Dim sFileUrl As String
    sFileUrl = sUserPath & "/Scripts/python/noms_propres.txt"

    Dim oSFA As Object
    oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")

    If oSFA.exists(sFileUrl) Then
        Dim oInStream As Object
        oInStream = oSFA.openFileRead(sFileUrl)
        Dim oTextIn As Object
        oTextIn = CreateUnoService("com.sun.star.io.TextInputStream")
        oTextIn.setInputStream(oInStream)
        oTextIn.setEncoding("UTF-8")

        Do While Not oTextIn.isEOF()
            Dim sLine As String
            sLine = Trim(oTextIn.readLine())
            If sLine <> "" Then
                If Not oMap.containsKey(sLine) Then
                    oMap.put(sLine, True)
                End If
            End If
        Loop
        oInStream.closeInput()
    End If

    LoadProperNouns = oMap
End Function

' Gère les règles 1, 2 et 3
Sub ProcessRules1To3(oDoc As Object, oProperNouns As Object)
    Dim oSD As Object
    oSD = oDoc.createSearchDescriptor()
    oSD.SearchRegularExpression = True
    ' Recherche : une minuscule, optionnellement un point, des espaces, puis un mot
    ' Note : \p{Ll} correspond à une lettre minuscule Unicode.
    ' On utilise les propriétés en minuscules (text, string) comme demandé.
    oSD.SearchString = "\p{Ll}\.?\s+[\p{L}«]+"

    Dim oFound As Object
    oFound = oDoc.findAll(oSD)

    Dim i As Long
    For i = oFound.Count - 1 To 0 Step -1
        Dim oRange As Object
        oRange = oFound.getByIndex(i)
        Dim sMatch As String
        sMatch = oRange.string

        Dim sFirstChar As String
        sFirstChar = Left(sMatch, 1)

        ' Extraction du deuxième mot
        Dim j As Long
        Dim sRem As String
        sRem = Mid(sMatch, 2)
        j = 1
        ' Saute le point et les espaces (y compris espace insécable Chr(160))
        Do While j <= Len(sRem)
            Dim sC As String
            sC = Mid(sRem, j, 1)
            If sC <> "." And sC <> " " And sC <> Chr(160) Then Exit Do
            j = j + 1
        Loop

        Dim sWord2 As String
        sWord2 = Mid(sRem, j)
        If sWord2 = "" Then GoTo NextMatch

        Dim sFirstCharW2 As String
        sFirstCharW2 = Left(sWord2, 1)

        Dim sNewSep As String
        sNewSep = ""

        ' Règle 1 : Suivi d'un nom propre -> Espace seul
        If oProperNouns.containsKey(sWord2) Then
            sNewSep = " "
        ' Règle 2 : Suivi d'une minuscule -> Espace seul
        ElseIf IsLowercaseLetter(sFirstCharW2) Then
            sNewSep = " "
        ' Règle 3 : Suivi d'une majuscule (non nom propre) -> Point + Espace
        ElseIf IsUppercaseLetter(sFirstCharW2) Then
            sNewSep = ". "
        End If

        If sNewSep <> "" Then
            Dim sNewText As String
            sNewText = sFirstChar & sNewSep & sWord2
            If oRange.string <> sNewText Then
                oRange.string = sNewText
            End If
        End If
NextMatch:
    Next i
End Sub

' Gère la règle 4 (Fusion de paragraphes)
Sub ProcessRule4(oDoc As Object, oProperNouns As Object)
    Dim oSD As Object
    oSD = oDoc.createSearchDescriptor()
    oSD.SearchRegularExpression = True
    ' Recherche une minuscule suivie d'un point en fin de paragraphe ($)
    oSD.SearchString = "\p{Ll}\.$"

    Dim oFound As Object
    oFound = oDoc.findAll(oSD)

    Dim i As Long
    For i = oFound.Count - 1 To 0 Step -1
        Dim oRange As Object
        oRange = oFound.getByIndex(i)

        ' Utilisation de la propriété minuscule 'text'
        Dim oCursor As Object
        oCursor = oDoc.text.createCursorByRange(oRange)

        If oCursor.gotoNextParagraph(False) Then
            ' Analyse du premier mot du paragraphe suivant
            Dim oWordCursor As Object
            oWordCursor = oDoc.text.createCursorByRange(oCursor)
            oWordCursor.gotoEndOfWord(True)
            Dim sNextWord As String
            sNextWord = oWordCursor.string

            If sNextWord <> "" Then
                Dim sFC As String
                sFC = Left(sNextWord, 1)

                ' Condition Règle 4 : Commence par minuscule ou « ET n'est pas un nom propre
                If (sFC = "«" Or IsLowercaseLetter(sFC)) And (Not oProperNouns.containsKey(sNextWord)) Then
                    ' 1. Supprimer le point (on garde la minuscule)
                    oRange.string = Left(oRange.string, 1)

                    ' 2. Fusionner les paragraphes avec un espace
                    Dim oMergeCursor As Object
                    oMergeCursor = oDoc.text.createCursorByRange(oRange)
                    oMergeCursor.gotoEndOfParagraph(False)
                    oMergeCursor.goRight(1, True) ' Sélectionne la marque de paragraphe
                    oMergeCursor.string = " "
                End If
            End If
        End If
    Next i
End Sub

' Fonction utilitaire : vérifie si un caractère est une lettre minuscule Unicode
Function IsLowercaseLetter(sChar As String) As Boolean
    If sChar = "" Then Exit Function
    ' LCase et UCase sont différents pour les lettres, identiques pour les symboles
    IsLowercaseLetter = (LCase(sChar) = sChar And UCase(sChar) <> sChar)
End Function

' Fonction utilitaire : vérifie si un caractère est une lettre majuscule Unicode
Function IsUppercaseLetter(sChar As String) As Boolean
    If sChar = "" Then Exit Function
    IsUppercaseLetter = (UCase(sChar) = sChar And LCase(sChar) <> sChar)
End Function
