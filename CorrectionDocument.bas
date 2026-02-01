' Macro de correction de ponctuation pour LibreOffice Writer
' Respecte les contraintes de casse (minuscules pour les propriétés UNO)
' Évite les mots interdits : Text, getText, getTextContent, getCurrentController, CurrentController

Function LoadProperNamesDictionary() As Object
    Dim oMap As Object
    Dim sUserPath As String
    Dim sDictPath As String
    Dim oSFA As Object
    Dim oInputStream As Object
    Dim oTextStream As Object
    Dim sLine As String
    Dim oPathSub As Object

    ' Utilisation de EnumerableMap comme recommandé pour les recherches rapides
    oMap = createUnoService("com.sun.star.container.EnumerableMap").create("string", "string")

    ' Utilisation du service standard PathSubstitution
    oPathSub = createUnoService("com.sun.star.util.PathSubstitution")
    sUserPath = oPathSub.getSubstituteVariableValue("$(user)")

    ' Chemin vers le dictionnaire des noms propres (basé sur le dossier Scripts/python mentionné)
    sDictPath = sUserPath & "/Scripts/python/noms_propres.txt"

    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")

    If oSFA.exists(sDictPath) Then
        oInputStream = oSFA.openFileRead(sDictPath)
        oTextStream = createUnoService("com.sun.star.io.TextInputStream")
        oTextStream.setInputStream(oInputStream)
        oTextStream.setEncoding("UTF-8")

        Do While Not oTextStream.isEOF()
            sLine = Trim(oTextStream.readLine())
            If sLine <> "" Then
                If Not oMap.containsKey(sLine) Then
                    oMap.put(sLine, "")
                End If
            End If
        Loop
        oInputStream.closeInput()
    End If

    LoadProperNamesDictionary = oMap
End Function

Function IsInDictionary(sWord As String, oDict As Object) As Boolean
    Dim sClean As String
    sClean = sWord
    ' Suppression de la ponctuation traînante pour la comparaison
    Do While Len(sClean) > 0
        Dim lastChar As String
        lastChar = Right(sClean, 1)
        If InStr(".,;:!?", lastChar) > 0 Then
            sClean = Left(sClean, Len(sClean) - 1)
        Else
            Exit Do
        End If
    Loop

    If oDict.containsKey(sClean) Then
        IsInDictionary = True
    Else
        IsInDictionary = False
    End If
End Function

Function IsFirstCharLower(sText As String) As Boolean
    If sText = "" Then
        IsFirstCharLower = False
        Exit Function
    End If
    Dim sFirst As String
    sFirst = Left(sText, 1)
    ' Vérification si c'est une minuscule (incluant caractères accentués)
    If sFirst = LCase(sFirst) And sFirst <> UCase(sFirst) Then
        IsFirstCharLower = True
    Else
        IsFirstCharLower = False
    End If
End Function

Sub CorrectionPonctuation()
    Dim oDoc As Object
    Dim oSD As Object
    Dim oResults As Object
    Dim i As Long
    Dim oFound As Object
    Dim oDict As Object

    oDoc = ThisComponent
    ' Vérification que nous sommes dans un document texte
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then Exit Sub

    ' Verrouillage pour améliorer les performances sur les gros documents
    oDoc.lockControllers()
    On Error GoTo ErrorHandler

    oDict = LoadProperNamesDictionary()

    ' --- Règles 1, 2 et 3 ---
    oSD = oDoc.createSearchDescriptor()
    oSD.searchregularexpression = True
    ' Recherche : Mot finissant par minuscule + ponctuation/espaces + mot suivant
    oSD.searchstring = "([[:alpha:]]*[[:lower:]])([\. ]+)([[:alpha:]«][[:alpha:]]*)"

    oResults = oDoc.findAll(oSD)

    ' Parcours inverse pour éviter l'invalidité des positions après modification
    For i = oResults.getCount() - 1 To 0 Step -1
        oFound = oResults.getByIndex(i)
        ApplyRules123(oFound, oDict)
    Next i

    ' --- Règle 4 ---
    ' Recherche des fins de paragraphe précédées d'une minuscule et optionnellement d'un point ($)
    oSD.searchstring = "[[:lower:]]\.?$"
    oResults = oDoc.findAll(oSD)

    For i = oResults.getCount() - 1 To 0 Step -1
        oFound = oResults.getByIndex(i)
        ApplyRule4(oDoc, oFound, oDict)
    Next i

    oDoc.unlockControllers()
    MsgBox "Correction terminée.", 64, "Information"
    Exit Sub

ErrorHandler:
    If Not IsNull(oDoc) Then oDoc.unlockControllers()
    MsgBox "Erreur : " & Err.Description, 16, "Erreur"
End Sub

Sub ApplyRules123(oFound As Object, oDict As Object)
    Dim sMatch As String
    sMatch = oFound.string

    Dim iStartPunct As Long
    Dim iStartWord2 As Long
    Dim j As Long
    Dim c As String

    ' Identification de la zone de ponctuation/espaces centrale
    For j = 1 To Len(sMatch)
        c = Mid(sMatch, j, 1)
        If c = "." Or c = " " Then
            iStartPunct = j
            Exit For
        End If
    Next j

    ' Identification du début du second mot (lettre ou «)
    For j = iStartPunct To Len(sMatch)
        c = Mid(sMatch, j, 1)
        If c <> "." And c <> " " Then
            iStartWord2 = j
            Exit For
        End If
    Next j

    If iStartPunct = 0 Or iStartWord2 = 0 Then Exit Sub

    Dim sWord1 As String, sWord2 As String
    sWord1 = Left(sMatch, iStartPunct - 1)
    sWord2 = Mid(sMatch, iStartWord2)

    Dim bIsProper As Boolean
    bIsProper = IsInDictionary(sWord2, oDict)

    Dim bStartsLower As Boolean
    bStartsLower = IsFirstCharLower(sWord2)

    ' Prise en compte du guillemet français ouvrant comme condition de minuscule
    If Left(sWord2, 1) = "«" Then bStartsLower = True

    Dim sNew As String

    ' Règles 1 & 2 : Si Nom Propre ou Commence par minuscule -> Un seul espace, pas de point
    If bIsProper Or bStartsLower Then
        sNew = sWord1 & " " & sWord2
    ' Règle 3 : Si Majuscule (non dictionnaire) -> Point + Espace
    Else
        sNew = sWord1 & ". " & sWord2
    End If

    If sMatch <> sNew Then
        oFound.string = sNew
    End If
End Sub

Sub ApplyRule4(oDoc As Object, oFound As Object, oDict As Object)
    Dim oCursor As Object
    ' Accès à oDoc.text (en minuscules)
    oCursor = oDoc.text.createCursorByRange(oFound)

    ' On s'assure d'être à la fin absolue du paragraphe
    oCursor.gotoEndOfParagraph(False)

    ' On tente de passer au début du paragraphe suivant
    If Not oCursor.gotoNextParagraph(False) Then Exit Sub

    ' On inspecte le premier mot du paragraphe suivant
    Dim oWordCursor As Object
    oWordCursor = oDoc.text.createCursorByRange(oCursor)
    oWordCursor.gotoEndOfWord(True)

    Dim sFirstWord As String
    sFirstWord = oWordCursor.string

    If sFirstWord = "" Then Exit Sub

    Dim bIsProper As Boolean
    bIsProper = IsInDictionary(sFirstWord, oDict)

    Dim bMatchRule As Boolean
    bMatchRule = False

    ' Règle 4 : Pas dans le dictionnaire ET (minuscule OU «)
    If Not bIsProper Then
        If IsFirstCharLower(sFirstWord) Or Left(sFirstWord, 1) = "«" Then
            bMatchRule = True
        End If
    End If

    If bMatchRule Then
        ' Identification de la base (mot finissant par minuscule, sans le point éventuel)
        Dim sEnd As String
        sEnd = oFound.string
        Dim sBase As String
        If Right(sEnd, 1) = "." Then
            sBase = Left(sEnd, Len(sEnd) - 1)
        Else
            sBase = sEnd
        End If

        ' Fusion des deux paragraphes en supprimant le retour chariot
        Dim oBreakCursor As Object
        oBreakCursor = oDoc.text.createCursorByRange(oFound)
        oBreakCursor.gotoEndOfParagraph(False)
        If oBreakCursor.goRight(1, True) Then
            oBreakCursor.string = ""
        End If

        ' Remplacement de la fin du premier mot (incluant le point supprimé) par base + espace
        oFound.string = sBase & " "
    End If
End Sub
