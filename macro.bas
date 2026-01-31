REM  *****  BASIC  *****

' Macro de correction de ponctuation pour LibreOffice Writer
' Développé par Jules, Expert en Basic LibreOffice
' Destinataire : J.c.mourocq@gmail.com

Option Explicit

''' <summary>
''' Récupère le chemin du dictionnaire des noms propres.
''' </summary>
Function GetDicoPath() As String
    Dim oPathSubst As Object
    Dim sUserPath As String
    oPathSubst = CreateUnoService("com.sun.star.comp.framework.PathSubstitution")
    sUserPath = oPathSubst.getSubstituteVariableValue("$(user)")
    ' Utilisation du chemin spécifié par l'utilisateur ou substitution standard
    GetDicoPath = sUserPath & "/Scripts/python/DicoNomsPropres.txt"
End Function

''' <summary>
''' Charge le dictionnaire des noms propres dans un objet EnumerableMap.
''' </summary>
Function LoadProperNouns() As Object
    Dim sUrl As String
    Dim oFileAccess As Object
    Dim oInputStream As Object
    Dim oTextStream As Object
    Dim oMap As Object
    Dim sLine As String

    sUrl = GetDicoPath()

    ' Tentative de création de l'EnumerableMap
    On Error Resume Next
    ' En LibreOffice Basic, la création d'un EnumerableMap nécessite parfois une factory.
    ' On tente d'abord par le service, avec une sécurité IsNull plus tard.
    oMap = CreateUnoService("com.sun.star.container.EnumerableMap")

    oFileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not oFileAccess.exists(sUrl) Then
        LoadProperNouns = Nothing
        Exit Function
    End If

    oInputStream = oFileAccess.openFileRead(sUrl)
    oTextStream = CreateUnoService("com.sun.star.io.TextInputStream")
    oTextStream.setInputStream(oInputStream)
    oTextStream.setEncoding("UTF-8")

    Do While Not oTextStream.isEOF()
        sLine = Trim(oTextStream.readLine())
        If sLine <> "" Then
            ' EnumerableMap utilise put(key, value)
            If Not IsNull(oMap) Then
                oMap.put(sLine, True)
            End If
        End If
    Loop
    oInputStream.closeInput()
    On Error GoTo 0

    LoadProperNouns = oMap
End Function

''' <summary>
''' Procédure principale de correction de la ponctuation.
''' </summary>
Sub CorrectPunctuation()
    Dim oDoc As Object
    Dim oDico As Object
    Dim oSearch As Object
    Dim oResults As Object
    Dim i As Long
    Dim oFound As Object
    Dim sMatch As String
    Dim sWord2 As String
    Dim sFirstCharWord2 As String
    Dim sNewSeparator As String
    Dim sTargetString As String
    Dim iWord2Pos As Integer
    Dim k As Integer
    Dim c As String

    oDoc = ThisComponent
    ' Vérification que nous sommes dans un document texte
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then Exit Sub

    ' Chargement du dictionnaire
    oDico = LoadProperNouns()

    ' Optimisation : on bloque les contrôleurs
    oDoc.lockControllers()
    On Error GoTo Finalize

    ' --- PARTIE 1 : Ponctuation intra-paragraphe (Règles 1, 2, 3) ---
    oSearch = oDoc.createSearchDescriptor()
    oSearch.SearchRegularExpression = True
    ' Motif : une minuscule, suivie d'un point facultatif et d'espaces, puis d'un mot
    oSearch.SearchString = "([a-zà-ÿ])(\.?\s+)([A-ZÀ-ßa-zà-ÿ«]+)"

    oResults = oDoc.findAll(oSearch)

    ' Traitement en sens inverse pour ne pas invalider les positions
    For i = oResults.getCount() - 1 To 0 Step -1
        oFound = oResults.getByIndex(i)
        sMatch = oFound.string

        ' Localiser le début du deuxième mot (après l'espace/ponctuation)
        iWord2Pos = 0
        For k = 2 To Len(sMatch)
            c = Mid(sMatch, k, 1)
            If c <> " " And c <> "." And Asc(c) > 32 Then
                iWord2Pos = k
                Exit For
            End If
        Next k

        If iWord2Pos > 0 Then
            sWord2 = Mid(sMatch, iWord2Pos)
            sFirstCharWord2 = Left(sWord2, 1)
            sNewSeparator = ""

            ' Règle 2 : minuscule + mot commençant par une minuscule -> espace seul
            If IsLowercaseChar(sFirstCharWord2) Then
                sNewSeparator = " "

            ' Règles 1 et 3 : minuscule + mot commençant par une MAJUSCULE
            ElseIf IsUppercaseChar(sFirstCharWord2) Then
                Dim bInDico As Boolean
                bInDico = False
                ' Pas de court-circuit en Basic, donc on imbrique les If
                If Not IsNull(oDico) Then
                    If oDico.containsKey(sWord2) Then
                        bInDico = True
                    End If
                End If

                If bInDico Then
                    ' Règle 1 : Présent dans le dictionnaire -> espace seul
                    sNewSeparator = " "
                Else
                    ' Règle 3 : Absent du dictionnaire -> point + espace
                    sNewSeparator = ". "
                End If
            End If

            ' Application du changement si nécessaire
            If sNewSeparator <> "" Then
                sTargetString = Left(sMatch, 1) & sNewSeparator & sWord2
                If sMatch <> sTargetString Then
                    oFound.string = sTargetString
                End If
            End If
        End If
    Next i

    ' --- PARTIE 2 : Ponctuation inter-paragraphe (Règle 4) ---
    ' Utilisation de curseurs pour plus de stabilité lors des fusions
    Dim oCursor1 As Object
    Dim oCursor2 As Object
    Dim bMerged As Boolean
    Dim sPara1 As String
    Dim sPara2 As String
    Dim sFirstWord2 As String

    oCursor1 = oDoc.text.createTextCursor()
    oCursor1.gotoStart(False)

    Do
        oCursor2 = oDoc.text.createTextCursorByRange(oCursor1)
        ' Tente de passer au paragraphe suivant
        If Not oCursor2.gotoNextParagraph(False) Then Exit Do

        bMerged = False
        sPara1 = oCursor1.string
        sPara2 = oCursor2.string

        If Len(sPara1) >= 2 And Len(sPara2) > 0 Then
            ' Vérification fin Para 1 : minuscule + point
            If Right(sPara1, 1) = "." And IsLowercaseChar(Mid(sPara1, Len(sPara1) - 1, 1)) Then
                ' Vérification début Para 2 : minuscule ou «
                If IsLowercaseChar(Left(sPara2, 1)) Or Left(sPara2, 1) = "«" Then
                    sFirstWord2 = ExtractFirstWord(sPara2)

                    Dim bShouldMerge As Boolean
                    bShouldMerge = False
                    If IsNull(oDico) Then
                        bShouldMerge = True
                    Else
                        If Not oDico.containsKey(sFirstWord2) Then
                            bShouldMerge = True
                        End If
                    End If

                    If bShouldMerge Then
                        ' Suppression du point à la fin du premier paragraphe
                        Dim oTempCursor As Object
                        oTempCursor = oDoc.text.createTextCursorByRange(oCursor1.getEnd())
                        oTempCursor.goLeft(1, True)
                        If oTempCursor.string = "." Then
                            oTempCursor.string = ""
                        End If

                        ' Fusion en remplaçant le saut de paragraphe par un espace
                        oTempCursor = oDoc.text.createTextCursorByRange(oCursor1.getEnd())
                        oTempCursor.gotoRange(oCursor2.getStart(), True)
                        oTempCursor.string = " "
                        bMerged = True
                    End If
                End If
            End If
        End If

        ' Si on a fusionné, oCursor1 couvre maintenant le texte fusionné.
        ' On ne passe pas au suivant pour permettre de fusionner avec le paragraphe d'après.
        If Not bMerged Then
            If Not oCursor1.gotoNextParagraph(False) Then Exit Do
        End If
    Loop

    MsgBox "Correction de la ponctuation terminée.", 64, "Jules Macro"

Finalize:
    oDoc.unlockControllers()
End Sub

''' <summary>
''' Extrait le premier mot d'une chaîne, en ignorant les guillemets ouvrants.
''' </summary>
Function ExtractFirstWord(s As String) As String
    Dim i As Integer
    Dim res As String
    Dim c As String
    res = s
    If Left(res, 1) = "«" Then res = Mid(res, 2)

    For i = 1 To Len(res)
        c = Mid(res, i, 1)
        If Not IsLowercaseChar(c) And Not IsUppercaseChar(c) Then
            res = Left(res, i - 1)
            Exit For
        End If
    Next i
    ExtractFirstWord = res
End Function

''' <summary>
''' Vérifie si un caractère est une minuscule (incluant les accents français).
''' </summary>
Function IsLowercaseChar(sChar As String) As Boolean
    Dim iCode As Integer
    If Len(sChar) = 0 Then
        IsLowercaseChar = False
        Exit Function
    End If
    iCode = Asc(sChar)
    IsLowercaseChar = (iCode >= 97 And iCode <= 122) Or (iCode >= 224 And iCode <= 255)
End Function

''' <summary>
''' Vérifie si un caractère est une majuscule (incluant les accents français).
''' </summary>
Function IsUppercaseChar(sChar As String) As Boolean
    Dim iCode As Integer
    If Len(sChar) = 0 Then
        IsUppercaseChar = False
        Exit Function
    End If
    iCode = Asc(sChar)
    IsUppercaseChar = (iCode >= 65 And iCode <= 90) Or (iCode >= 192 And iCode <= 221)
End Function
