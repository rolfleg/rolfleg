' Macro LibreOffice Writer pour la correction de la ponctuation et de la casse
' Développé pour Windows 11, LibreOffice dernière version.

Sub CorrectionPonctuation
    Dim oDoc As Object
    oDoc = ThisComponent

    ' Vérification que le document est un document texte
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox "Cette macro ne fonctionne que dans LibreOffice Writer."
        Exit Sub
    End If

    Dim oDic As Object
    oDic = ChargerDictionnaireNomsPropres()

    ' Verrouiller l'affichage pour améliorer les performances
    oDoc.lockControllers()
    On Error GoTo GestionErreur

    ' Application des règles 1, 2 et 3 (transitions entre mots)
    AppliquerRegles123(oDoc, oDic)

    ' Application de la règle 4 (fusion de paragraphes)
    AppliquerRegle4(oDoc, oDic)

    oDoc.unlockControllers()
    MsgBox "Traitement terminé avec succès."
    Exit Sub

GestionErreur:
    oDoc.unlockControllers()
    MsgBox "Une erreur est survenue : " & Err.Description
End Sub

' Charge le dictionnaire des noms propres depuis $(user)/Scripts/python/noms_propres.txt
Function ChargerDictionnaireNomsPropres() As Object
    Dim oMap As Object
    ' Utilisation de EnumerableMap comme suggéré par les bonnes pratiques du projet
    On Error Resume Next
    oMap = createUnoService("com.sun.star.container.EnumerableMap")
    ' Si EnumerableMap n'est pas disponible, on utilise une Collection Basic
    If IsNull(oMap) Then
        oMap = New Collection
    End If
    On Error GoTo 0

    Dim oPathSub As Object
    oPathSub = createUnoService("com.sun.star.util.PathSubstitution")
    Dim sUserPath As String
    sUserPath = oPathSub.getSubstituteVariableValue("$(user)")
    Dim sFileUrl As String
    sFileUrl = sUserPath & "/Scripts/python/noms_propres.txt"

    Dim oSFA As Object
    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")

    If oSFA.exists(sFileUrl) Then
        Dim oInputStream As Object
        oInputStream = oSFA.openFileRead(sFileUrl)
        Dim oTextStream As Object
        oTextStream = createUnoService("com.sun.star.io.TextInputStream")
        oTextStream.setInputStream(oInputStream)
        oTextStream.setEncoding("UTF-8")

        Do While Not oTextStream.isEOF()
            Dim sLine As String
            sLine = Trim(oTextStream.readLine())
            If sLine <> "" Then
                On Error Resume Next
                If TypeOf oMap Is Collection Then
                    oMap.Add(True, sLine)
                Else
                    oMap.put(sLine, True)
                End If
                On Error GoTo 0
            End If
        Loop
        oInputStream.closeInput()
    End If

    ChargerDictionnaireNomsPropres = oMap
End Function

' Vérifie si un mot est présent dans le dictionnaire
Function EstDansDictionnaire(oDic As Object, sMot As String) As Boolean
    If IsNull(oDic) Then
        EstDansDictionnaire = False
        Exit Function
    End If

    ' Nettoyer le mot des éventuels guillemets ou ponctuation résiduelle
    Dim sNettoye As String
    sNettoye = sMot
    If Left(sNettoye, 1) = "«" Then sNettoye = Mid(sNettoye, 2)
    Do While Len(sNettoye) > 0
        Dim c As String
        c = Right(sNettoye, 1)
        If InStr(".,;:!?»", c) > 0 Then
            sNettoye = Left(sNettoye, Len(sNettoye) - 1)
        Else
            Exit Do
        End If
    Loop

    On Error Resume Next
    If TypeOf oDic Is Collection Then
        Dim v As Variant
        v = oDic.Item(sNettoye)
        EstDansDictionnaire = Not IsEmpty(v)
    Else
        EstDansDictionnaire = oDic.containsKey(sNettoye)
    End If
    On Error GoTo 0
End Function

' Règles 1, 2, 3 : Gestion des transitions de mots au sein d'un paragraphe
Sub AppliquerRegles123(oDoc As Object, oDic As Object)
    Dim oSearch As Object
    oSearch = oDoc.createSearchDescriptor()
    oSearch.searchregularexpression = True
    ' Recherche : mot finissant par minuscule + ponctuation optionnelle + espaces + mot suivant
    ' \b\w*[[:lower:]] -> Mot se terminant par une minuscule
    ' \s*\.?\s+ -> Espaces, point optionnel, puis un ou plusieurs espaces
    ' \w+ -> Mot suivant
    oSearch.searchstring = "\b\w*[[:lower:]]\s*\.?\s+\w+"

    Dim oFound As Object
    oFound = oDoc.findAll(oSearch)
    If oFound.getCount() = 0 Then Exit Sub

    Dim i As Long
    ' Itération inverse pour éviter les décalages d'index lors des modifications
    For i = oFound.getCount() - 1 To 0 Step -1
        Dim oRange As Object
        oRange = oFound.getByIndex(i)
        Dim sMatch As String
        sMatch = oRange.string

        ' Analyse du match pour extraire Word1 et Word2
        Dim iFirstSep As Long
        iFirstSep = TrouverPremierSeparateur(sMatch)
        If iFirstSep > 0 Then
            Dim sWord1 As String
            sWord1 = Left(sMatch, iFirstSep - 1)

            ' Trouver le début du Word2
            Dim iWord2Start As Long
            iWord2Start = iFirstSep
            Do While iWord2Start <= Len(sMatch)
                Dim c As String
                c = Mid(sMatch, iWord2Start, 1)
                If c <> " " And c <> "." Then Exit Do
                iWord2Start = iWord2Start + 1
            Loop

            Dim sWord2 As String
            sWord2 = Mid(sMatch, iWord2Start)

            ' Vérification des conditions
            Dim bWord2DansDic As Boolean
            bWord2DansDic = EstDansDictionnaire(oDic, sWord2)

            Dim sPremCharWord2 As String
            sPremCharWord2 = Left(sWord2, 1)
            Dim bWord2CommenceMinuscule As Boolean
            bWord2CommenceMinuscule = (sPremCharWord2 = LCase(sPremCharWord2)) And (sPremCharWord2 <> UCase(sPremCharWord2))
            Dim bWord2CommenceMajuscule As Boolean
            bWord2CommenceMajuscule = (sPremCharWord2 = UCase(sPremCharWord2)) And (sPremCharWord2 <> LCase(sPremCharWord2))

            Dim sNouveau As String
            sNouveau = ""

            ' Règle 1 : Suivi d'un nom propre -> Pas de point, un seul espace
            If bWord2DansDic Then
                sNouveau = sWord1 & " " & sWord2
            ' Règle 2 : Suivi d'une minuscule -> Pas de point, un seul espace
            ElseIf bWord2CommenceMinuscule Then
                sNouveau = sWord1 & " " & sWord2
            ' Règle 3 : Suivi d'une majuscule (hors dictionnaire) -> Point et espace
            ElseIf bWord2CommenceMajuscule And Not bWord2DansDic Then
                sNouveau = sWord1 & ". " & sWord2
            End If

            If sNouveau <> "" And sNouveau <> sMatch Then
                oRange.string = sNouveau
            End If
        End If
    Next i
End Sub

' Règle 4 : Fusion de paragraphes
Sub AppliquerRegle4(oDoc As Object, oDic As Object)
    Dim oSearch As Object
    oSearch = oDoc.createSearchDescriptor()
    oSearch.searchregularexpression = True
    ' Recherche une fin de paragraphe : mot finissant par minuscule + point + fin de para ($)
    oSearch.searchstring = "\b\w*[[:lower:]]\.$"

    Dim oFound As Object
    oFound = oDoc.findAll(oSearch)
    If oFound.getCount() = 0 Then Exit Sub

    Dim i As Long
    For i = oFound.getCount() - 1 To 0 Step -1
        Dim oRange As Object
        oRange = oFound.getByIndex(i)

        ' Création d'un curseur pour manipuler la structure
        Dim oCursor As Object
        oCursor = oDoc.text.createCursorByRange(oRange)

        ' Sélectionner le saut de paragraphe qui suit
        If oCursor.goRight(1, True) Then
            ' Vérifier le contenu du paragraphe suivant
            Dim oNextParaCursor As Object
            oNextParaCursor = oDoc.text.createCursorByRange(oCursor.getEnd())
            oNextParaCursor.gotoEndOfParagraph(True)

            Dim sNextParaText As String
            sNextParaText = LTrim(oNextParaCursor.string)

            If sNextParaText <> "" Then
                Dim sFirstChar As String
                sFirstChar = Left(sNextParaText, 1)

                ' Condition : commence par minuscule ou «
                Dim bConditionCase As Boolean
                bConditionCase = (sFirstChar = LCase(sFirstChar) And sFirstChar <> UCase(sFirstChar)) Or (sFirstChar = "«")

                If bConditionCase Then
                    ' Extraire le premier mot pour vérification dictionnaire
                    Dim sFirstWord As String
                    sFirstWord = ExtrairePremierMot(sNextParaText)

                    If Not EstDansDictionnaire(oDic, sFirstWord) Then
                        ' Action : Supprimer point et saut de paragraphe, ajouter espace
                        Dim sBaseText As String
                        sBaseText = oRange.string
                        If Right(sBaseText, 1) = "." Then
                            sBaseText = Left(sBaseText, Len(sBaseText) - 1)
                        End If
                        ' Remplacement du texte sélectionné (incluant le saut de paragraphe)
                        oCursor.string = sBaseText & " "
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Helper : Trouve l'index du premier espace ou point servant de séparateur
Function TrouverPremierSeparateur(sText As String) As Long
    Dim i As Long
    For i = 1 To Len(sText)
        Dim c As String
        c = Mid(sText, i, 1)
        If c = " " Or c = "." Then
            TrouverPremierSeparateur = i
            Exit Function
        End If
    Next i
    TrouverPremierSeparateur = 0
End Function

' Helper : Extrait le premier mot alphabétique
Function ExtrairePremierMot(sText As String) As String
    Dim i As Long
    Dim iStart As Long: iStart = 0
    Dim sResult As String: sResult = ""

    ' Sauter les caractères non alphabétiques au début (comme «)
    For i = 1 To Len(sText)
        Dim c As String
        c = Mid(sText, i, 1)
        If (LCase(c) <> UCase(c)) Then
            iStart = i
            Exit For
        End If
    Next i

    If iStart > 0 Then
        For i = iStart To Len(sText)
            Dim c2 As String
            c2 = Mid(sText, i, 1)
            If (LCase(c2) <> UCase(c2)) Then
                sResult = sResult & c2
            Else
                Exit For
            End If
        Next i
    End If
    ExtrairePremierMot = sResult
End Function
