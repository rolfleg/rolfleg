' Macro LibreOffice Writer pour la gestion de la ponctuation et des noms propres
' Développé pour Windows 11 et la dernière version de LibreOffice

' --- Fonctions de chargement du dictionnaire ---

Function GetDictionaryPath() As String
    Dim oPathSub As Object
    oPathSub = CreateUnoService("com.sun.star.util.PathSubstitution")
    ' Chemin vers le dictionnaire dans le dossier utilisateur
    ' C:\Users\jeanp\AppData\LibreOffice\4\user\Scripts\python\DictionnaireNomsPropres.txt
    GetDictionaryPath = oPathSub.getSubstituteVariableValue("$(user)") & "/Scripts/python/DictionnaireNomsPropres.txt"
End Function

Function LoadProperNouns(sFilePath As String) As Object
    Dim oMap As Object
    oMap = CreateUnoService("com.sun.star.container.EnumerableMap")
    ' Initialisation de la map : clés de type string, valeurs de type boolean
    oMap.create("string", "boolean")

    Dim oSFA As Object
    oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")

    If Not oSFA.exists(sFilePath) Then
        LoadProperNouns = oMap
        Exit Function
    End If

    Dim oInStream As Object
    oInStream = oSFA.openFileRead(sFilePath)
    Dim oTextIn As Object
    oTextIn = CreateUnoService("com.sun.star.io.TextInputStream")
    oTextIn.setInputStream(oInStream)
    oTextIn.setEncoding("UTF-8")

    While Not oTextIn.isEOF()
        Dim sLine As String
        sLine = Trim(oTextIn.readLine())
        ' Ignorer les lignes vides et les commentaires
        If sLine <> "" And Left(sLine, 1) <> "#" Then
            oMap.put(sLine, True)
        End If
    Wend
    oInStream.closeInput()
    LoadProperNouns = oMap
End Function

' --- Fonctions utilitaires ---

Function IsLowercase(sChar As String) As Boolean
    If sChar = "" Then
        IsLowercase = False
        Exit Function
    End If
    ' LCase et UCase gèrent les caractères accentués en Basic
    IsLowercase = (LCase(sChar) = sChar And UCase(sChar) <> sChar)
End Function

Function IsInDictionary(oMap As Object, sWord As String) As Boolean
    If IsNull(oMap) Then
        IsInDictionary = False
        Exit Function
    End If
    IsInDictionary = oMap.containsKey(sWord)
End Function

Sub ApplyRules123(oDoc As Object, oProperNouns As Object)
    Dim oSearch As Object
    oSearch = oDoc.createSearchDescriptor()
    oSearch.searchregularexpression = True
    ' On cherche la fin du premier mot : lettre minuscule + point optionnel + espaces
    oSearch.searchstring = "\p{Ll}\.?[ ]+"

    Dim oFound As Object
    oFound = oDoc.findAll(oSearch)
    If IsNull(oFound) Then Exit Sub
    If oFound.Count = 0 Then Exit Sub

    ' Service pour extraire les groupes de capture du texte trouvé
    Dim oTS As Object
    oTS = CreateUnoService("com.sun.star.util.TextSearch")
    Dim oOptions As New com.sun.star.util.SearchOptions
    oOptions.algorithmType = com.sun.star.util.SearchAlgorithms.REGEXP
    oOptions.searchString = "(\p{Ll})(\.?)(\s+)"
    oTS.setOptions(oOptions)

    Dim i As Integer
    ' Parcours inverse pour éviter de décaler les positions dans le document
    For i = oFound.Count - 1 To 0 Step -1
        Dim oRange As Object
        oRange = oFound.getByIndex(i)
        Dim sMatch As String
        sMatch = oRange.string

        ' On utilise un curseur pour obtenir le mot suivant
        Dim oCursor As Object
        oCursor = oRange.text.createCursorByRange(oRange)
        oCursor.collapseToEnd()
        oCursor.gotoEndOfWord(True)
        Dim sNextWord As String
        sNextWord = oCursor.string

        If sNextWord <> "" Then
            Dim oRes As Object
            oRes = oTS.searchForward(sMatch, 0, Len(sMatch))

            If oRes.subregexpressions >= 4 Then
                Dim sG1 As String, sG2 As String, sG3 As String
                sG1 = Mid(sMatch, oRes.startoffset(1) + 1, oRes.endoffset(1) - oRes.startoffset(1))
                sG2 = Mid(sMatch, oRes.startoffset(2) + 1, oRes.endoffset(2) - oRes.startoffset(2))
                sG3 = Mid(sMatch, oRes.startoffset(3) + 1, oRes.endoffset(3) - oRes.startoffset(3))

                Dim sNew As String
                Dim bChanged As Boolean : bChanged = False

                If IsLowercase(Left(sNextWord, 1)) Then
                    ' Règle 2 : suivie d'une minuscule -> un seul espace, pas de point
                    sNew = sG1 & " "
                    bChanged = True
                Else
                    ' Majuscule
                    If IsInDictionary(oProperNouns, sNextWord) Then
                        ' Règle 1 : nom propre -> un seul espace, pas de point
                        sNew = sG1 & " "
                        bChanged = True
                    Else
                        ' Règle 3 : hors dictionnaire -> ajouter point, garder espace
                        sNew = sG1 & "." & sG3
                        If sG2 <> "." Or sNew <> sMatch Then
                            bChanged = True
                        End If
                    End If
                End If

                If bChanged And sNew <> sMatch Then
                    oRange.string = sNew
                End If
            End If
        End If
    Next i
End Sub

Sub ApplyRule4(oDoc As Object, oProperNouns As Object)
    Dim oSearch As Object
    oSearch = oDoc.createSearchDescriptor()
    oSearch.searchregularexpression = True
    ' Règle 4 : Minuscule, point optionnel, à la fin d'un paragraphe
    oSearch.searchstring = "\p{Ll}\.?\s*$"

    Dim oFound As Object
    oFound = oDoc.findAll(oSearch)
    If IsNull(oFound) Then Exit Sub
    If oFound.Count = 0 Then Exit Sub

    Dim i As Integer
    For i = oFound.Count - 1 To 0 Step -1
        Dim oRange As Object
        oRange = oFound.getByIndex(i)

        ' Utilisation d'un curseur pour explorer le paragraphe suivant
        Dim oCursor As Object
        oCursor = oRange.text.createCursorByRange(oRange)

        If oCursor.gotoNextParagraph(False) Then
            ' Vérifier le premier mot du paragraphe suivant
            oCursor.gotoEndOfWord(True)
            Dim sNextWord As String
            sNextWord = oCursor.string

            ' Conditions Règle 4 : mot suivant hors dico ET (minuscule OU «)
            If (Not IsInDictionary(oProperNouns, sNextWord)) And _
               (IsLowercase(Left(sNextWord, 1)) Or Left(sNextWord, 1) = "«") Then

                Dim oMergeCursor As Object
                oMergeCursor = oRange.text.createCursorByRange(oRange)
                ' Se placer après la lettre minuscule
                oMergeCursor.collapseToStart()
                oMergeCursor.goRight(1, False)
                ' Sélectionner jusqu'au début du paragraphe suivant
                oMergeCursor.gotoNextParagraph(True)

                ' Fusionner avec un espace
                oMergeCursor.string = " "
            End If
        End If
    Next i
End Sub

Sub PunctuationMacro
    Dim oDoc As Object
    oDoc = ThisComponent

    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox "Cette macro ne fonctionne que dans LibreOffice Writer.", 16, "Erreur"
        Exit Sub
    End If

    oDoc.lockControllers()
    On Error GoTo ErrorHandler

    Dim sDictPath As String
    sDictPath = GetDictionaryPath()

    Dim oProperNouns As Object
    oProperNouns = LoadProperNouns(sDictPath)

    ' Règles 1, 2 et 3
    ApplyRules123(oDoc, oProperNouns)

    ' Règle 4 (fusion de paragraphes)
    ApplyRule4(oDoc, oProperNouns)

    oDoc.unlockControllers()
    MsgBox "Traitement de la ponctuation terminé.", 64, "Succès"
    Exit Sub

ErrorHandler:
    If Not IsNull(oDoc) Then oDoc.unlockControllers()
    MsgBox "Une erreur est survenue : " & Err.Description & " (Ligne " & Erl & ")", 16, "Erreur"
End Sub
