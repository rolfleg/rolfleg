REM  *****  BASIC  *****

Sub TraitementPunctuation
    Dim oDoc As Object
    oDoc = ThisComponent

    ' Vérifier que nous sommes dans un document texte
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox "Cette macro doit être exécutée dans LibreOffice Writer.", 16, "Erreur"
        Exit Sub
    End If

    Dim dictionnaire As Object
    dictionnaire = ChargerDictionnaire()

    ' Verrouiller l'affichage pour plus de rapidité
    oDoc.lockControllers()
    On Error GoTo ErrorHandler

    ' Appliquer les règles
    AppliquerRegle4(oDoc, dictionnaire)
    AppliquerRegles123(oDoc, dictionnaire)

    oDoc.unlockControllers()
    MsgBox "Le traitement est terminé.", 64, "Succès"
    Exit Sub

ErrorHandler:
    If Not IsNull(oDoc) Then oDoc.unlockControllers()
    MsgBox "Une erreur est survenue : " & Err.Description, 16, "Erreur"
End Sub

Function ChargerDictionnaire() As Object
    Dim oDict As Object
    ' Utilisation d'une collection pour stocker les noms propres
    oDict = New Collection

    Dim oPathSub As Object
    Dim sUserPath As String
    Dim sFullURL As String
    Dim oSFA As Object
    Dim oInStream As Object
    Dim oTextStream As Object
    Dim sLine As String

    On Error Resume Next
    oPathSub = createUnoService("com.sun.star.util.PathSubstitution")
    sUserPath = oPathSub.getSubstituteVariableValue("$(user)")
    sFullURL = sUserPath & "/Scripts/python/DictionnaireNomsPropres.txt"

    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not oSFA.exists(sFullURL) Then
        ' Fallback sur le chemin Windows spécifique mentionné par l'utilisateur
        sFullURL = convertToURL("C:\Users\jeanp\AppData\LibreOffice\4\user\Scripts\python\DictionnaireNomsPropres.txt")
    End If

    If oSFA.exists(sFullURL) Then
        oInStream = oSFA.openFileRead(sFullURL)
        oTextStream = createUnoService("com.sun.star.io.TextInputStream")
        oTextStream.setInputStream(oInStream)
        oTextStream.setEncoding("UTF-8")

        Do While Not oTextStream.isEOF()
            sLine = Trim(oTextStream.readLine())
            ' Ignorer les commentaires et lignes vides
            If sLine <> "" And Left(sLine, 1) <> "#" Then
                ' On utilise le mot comme clé pour la collection
                oDict.Add(sLine, sLine)
            End If
        Loop
        oInStream.closeInput()
    End If
    On Error GoTo 0
    ChargerDictionnaire = oDict
End Function

Function EstDansDictionnaire(sMot As String, dictionnaire As Object) As Boolean
    Dim bResult As Boolean
    bResult = False
    On Error Resume Next
    Dim sVal As String
    sVal = dictionnaire.Item(sMot)
    If Err.Number = 0 Then
        bResult = True
    End If
    On Error GoTo 0
    EstDansDictionnaire = bResult
End Function

Function IsLowerCase(s As String) As Boolean
    If s = "" Then IsLowerCase = False : Exit Function
    IsLowerCase = (LCase(s) = s And UCase(s) <> s)
End Function

Function IsUpperCase(s As String) As Boolean
    If s = "" Then IsUpperCase = False : Exit Function
    IsUpperCase = (UCase(s) = s And LCase(s) <> s)
End Function

Sub AppliquerRegles123(oDoc As Object, dictionnaire As Object)
    ' Règle 1: minuscule + Nom Propre -> Espace uniquement
    ' Règle 2: minuscule + minuscule -> Espace uniquement (supprimer le point si besoin)
    ' Règle 3: minuscule + Majuscule (non NP) -> Point + garder l'espace

    Dim oSearch As Object
    oSearch = oDoc.createSearchDescriptor()
    oSearch.searchregularexpression = True
    ' Chercher une minuscule suivie éventuellement d'un point et d'espaces, puis un mot
    oSearch.searchstring = "\p{Ll}\.?[:space:]+\w+"

    Dim oFound As Object
    oFound = oDoc.findAll(oSearch)
    If oFound.Count = 0 Then Exit Sub

    Dim i As Integer
    Dim oRange As Object
    Dim sMatch As String
    Dim sPrefix As String
    Dim sMiddle As String
    Dim sWord As String
    Dim nLastSpace As Integer
    Dim sNew As String

    ' Parcours inverse
    For i = oFound.Count - 1 To 0 Step -1
        oRange = oFound.getByIndex(i)
        sMatch = oRange.string

        ' Trouver le début du dernier mot (dernier espace)
        nLastSpace = InStrRev(sMatch, " ")
        If nLastSpace > 1 Then
            sPrefix = Left(sMatch, 1)
            sWord = Mid(sMatch, nLastSpace + 1)
            ' sMiddle contient ce qu'il y a entre la minuscule et le dernier mot
            sMiddle = Mid(sMatch, 2, nLastSpace - 1)

            Dim bProper As Boolean
            bProper = EstDansDictionnaire(sWord, dictionnaire)

            Dim bLower As Boolean
            bLower = IsLowerCase(Left(sWord, 1))

            Dim bUpper As Boolean
            bUpper = IsUpperCase(Left(sWord, 1))

            sNew = sMatch

            ' Règle 1 & 2: Pas de point, un seul espace
            If bProper Or bLower Then
                sNew = sPrefix & " " & sWord
            ' Règle 3: Majuscule (non NP) -> Ajouter un point si manquant, garder l'espace existant
            ElseIf bUpper And Not bProper Then
                If InStr(sMiddle, ".") = 0 Then
                    sNew = sPrefix & "." & sMiddle & sWord
                End If
            End If

            If sNew <> sMatch Then
                oRange.string = sNew
            End If
        End If
    Next i
End Sub

Sub AppliquerRegle4(oDoc As Object, dictionnaire As Object)
    ' Règle 4: Si minuscule + "." + retour paragraphe, et mot suivant (non NP) commence par minuscule ou «
    ' alors fusionner avec un espace.

    Dim oSearch As Object
    oSearch = oDoc.createSearchDescriptor()
    oSearch.searchregularexpression = True
    ' On cherche une minuscule suivie d'un point optionnel à la fin du paragraphe
    oSearch.searchstring = "\p{Ll}\.?$"

    Dim oFound As Object
    oFound = oDoc.findAll(oSearch)
    If oFound.Count = 0 Then Exit Sub

    Dim i As Integer
    Dim oRange As Object
    Dim oCursor As Object
    Dim oNextCursor As Object
    Dim sFirstWord As String
    Dim c1 As String

    ' Parcours inverse pour ne pas invalider les positions lors de la fusion
    For i = oFound.Count - 1 To 0 Step -1
        oRange = oFound.getByIndex(i)
        ' Utilisation de .text en minuscule comme demandé
        oCursor = oDoc.text.createCursorByRange(oRange)
        oCursor.gotoEndOfParagraph(False)

        oNextCursor = oDoc.text.createCursorByRange(oCursor)
        If oNextCursor.gotoNextParagraph(False) Then
            ' Aller au début du paragraphe suivant et sélectionner le premier mot
            oNextCursor.gotoStartOfParagraph(False)
            oNextCursor.gotoEndOfWord(True)
            ' Utilisation de .string en minuscule
            sFirstWord = Trim(oNextCursor.string)

            If sFirstWord <> "" Then
                c1 = Left(sFirstWord, 1)
                ' Conditions: (commence par minuscule ou «) ET (n'est pas dans le dictionnaire)
                If (IsLowerCase(c1) Or c1 = "«") And Not EstDansDictionnaire(sFirstWord, dictionnaire) Then
                    ' Fusionner
                    oCursor = oDoc.text.createCursorByRange(oRange)
                    oCursor.collapseToStart()
                    oCursor.goRight(1, False) ' Se placer après la minuscule
                    oCursor.gotoEndOfParagraph(True) ' Sélectionner le point s'il existe
                    oCursor.goRight(1, True) ' Sélectionner le retour paragraphe
                    oCursor.string = " "
                End If
            End If
        End If
    Next i
End Sub
