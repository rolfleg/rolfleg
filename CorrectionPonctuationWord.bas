Attribute VB_Name = "CorrectionPonctuation"

' Macro de correction de ponctuation pour Microsoft Word 2013+
' Développé par Jules, Expert en VBA
' Destinataire : J.c.mourocq@gmail.com

Option Explicit

''' <summary>
''' Charge le dictionnaire des noms propres depuis le profil utilisateur.
''' </summary>
Function ChargerDictionnaire() As Object
    Dim dico As Object
    Dim fso As Object
    Dim userProfile As String
    Dim paths(1) As String
    Dim filePath As String
    Dim i As Integer
    Dim txtStream As Object
    Dim line As String

    Set dico = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    userProfile = Environ("USERPROFILE")

    ' Chemins possibles
    paths(0) = userProfile & "\DictionnaireNomsPropres.txt"
    paths(1) = userProfile & "\AppData\Roaming\LibreOffice\4\user\Scripts\python\DictionnaireNomsPropres.txt"

    filePath = ""
    For i = 0 To UBound(paths)
        If fso.FileExists(paths(i)) Then
            filePath = paths(i)
            Exit For
        End If
    Next i

    ' Si aucun n'existe, on crée un par défaut dans le profil
    If filePath = "" Then
        filePath = paths(0)
        On Error Resume Next
        Set txtStream = fso.CreateTextFile(filePath, True)
        If Err.Number = 0 Then
            txtStream.WriteLine "Jean-Claude"
            txtStream.WriteLine "Microsoft"
            txtStream.WriteLine "Paris"
            txtStream.Close
        End If
        On Error GoTo 0
    End If

    ' Chargement du contenu
    If fso.FileExists(filePath) Then
        On Error Resume Next
        Set txtStream = fso.OpenTextFile(filePath, 1, False) ' ForReading, False = ASCII
        If Err.Number = 0 Then
            Do While Not txtStream.AtEndOfStream
                line = Trim(txtStream.ReadLine)
                If line <> "" Then
                    If Not dico.Exists(line) Then dico.Add line, True
                End If
            Loop
            txtStream.Close
        End If
        On Error GoTo 0
    End If

    Set ChargerDictionnaire = dico
End Function

''' <summary>
''' Vérifie si un caractère est une minuscule (incluant les accents).
''' </summary>
Function IsMinuscule(s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    IsMinuscule = (s = LCase(s) And s <> UCase(s))
End Function

''' <summary>
''' Vérifie si un caractère est une majuscule (incluant les accents).
''' </summary>
Function IsMajuscule(s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    IsMajuscule = (s = UCase(s) And s <> LCase(s))
End Function

''' <summary>
''' Applique les règles 1, 2 et 3 (intra-paragraphe).
''' </summary>
Sub AppliquerRegles123(doc As Object, dico As Object)
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    Dim firstChar As String, separator As String, secondWord As String
    Dim firstCharSecondWord As String, newText As String, cleanedWord As String
    Dim rng As Object

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = False
    ' Motif : [Lettre minuscule] [Points/Espaces facultatifs] [Mot commençant par lettre ou «]
    regEx.Pattern = "([a-zà-âæçéèêëîïôœùûüÿ])([\. ]+)([a-zA-Zà-âæçéèêëîïôœùûüÿÀ-ÂÆÇÉÈÊËÎÏÔŒÙÛÜŸŸ«][a-zA-Zà-âæçéèêëîïôœùûüÿÀ-ÂÆÇÉÈÊËÎÏÔŒÙÛÜŸŸ-]*)"

    On Error Resume Next
    Set matches = regEx.Execute(doc.Range.Text)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    For i = matches.Count - 1 To 0 Step -1
        Set match = matches.Item(i)
        If match.SubMatches.Count >= 3 Then
            firstChar = match.SubMatches(0)
            separator = match.SubMatches(1)
            secondWord = match.SubMatches(2)

            If secondWord <> "" Then
                firstCharSecondWord = Left(secondWord, 1)
                newText = ""

                ' Règle 2 : le mot suivant commence par une minuscule
                If IsMinuscule(firstCharSecondWord) Then
                    newText = firstChar & " " & secondWord

                ' Règle 1 et 3 : le mot suivant commence par une majuscule ou «
                ElseIf IsMajuscule(firstCharSecondWord) Or firstCharSecondWord = "«" Then
                    cleanedWord = secondWord
                    If Left(cleanedWord, 1) = "«" Then cleanedWord = Mid(cleanedWord, 2)
                    ' Retirer ponctuation collée à la fin si nécessaire
                    Do While Len(cleanedWord) > 0 And Not (IsMinuscule(Right(cleanedWord, 1)) Or IsMajuscule(Right(cleanedWord, 1)))
                        cleanedWord = Left(cleanedWord, Len(cleanedWord) - 1)
                    Loop

                    If dico.Exists(cleanedWord) Then
                        ' Règle 1 : Nom propre -> espace seul
                        newText = firstChar & " " & secondWord
                    Else
                        ' Règle 3 : Pas un nom propre -> point + espace
                        newText = firstChar & ". " & secondWord
                    End If
                End If

                If newText <> "" And match.Value <> newText Then
                    Set rng = doc.Range(match.FirstIndex, match.FirstIndex + match.Length)
                    ' Vérification de sécurité pour ne pas modifier si le texte a bougé (peu probable ici)
                    If rng.Text = match.Value Then
                        rng.Text = newText
                    End If
                End If
            End If
        End If
    Next i
End Sub

''' <summary>
''' Applique les règles 5 et 6 (sauts de ligne et autres CHR).
''' </summary>
Sub AppliquerRegles56(doc As Object, dico As Object)
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    Dim firstChar As String, separator As String, secondWord As String
    Dim firstCharSecondWord As String, newText As String, cleanedWord As String
    Dim rng As Object

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = False
    ' Motif : [Minuscule] [Un ou plusieurs sauts \r\n\v] [Mot]
    regEx.Pattern = "([a-zà-âæçéèêëîïôœùûüÿ])([\r\n\v]+)([a-zA-Zà-âæçéèêëîïôœùûüÿÀ-ÂÆÇÉÈÊËÎÏÔŒÙÛÜŸŸ«][a-zA-Zà-âæçéèêëîïôœùûüÿÀ-ÂÆÇÉÈÊËÎÏÔŒÙÛÜŸŸ-]*)"

    On Error Resume Next
    Set matches = regEx.Execute(doc.Range.Text)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    For i = matches.Count - 1 To 0 Step -1
        Set match = matches.Item(i)
        If match.SubMatches.Count >= 3 Then
            firstChar = match.SubMatches(0)
            separator = match.SubMatches(1)
            secondWord = match.SubMatches(2)

            If secondWord <> "" Then
                firstCharSecondWord = Left(secondWord, 1)
                newText = ""

                cleanedWord = secondWord
                If Left(cleanedWord, 1) = "«" Then cleanedWord = Mid(cleanedWord, 2)
                Do While Len(cleanedWord) > 0 And Not (IsMinuscule(Right(cleanedWord, 1)) Or IsMajuscule(Right(cleanedWord, 1)))
                    cleanedWord = Left(cleanedWord, Len(cleanedWord) - 1)
                Loop

                ' Règle 5 & 6 : Cas minuscule ou nom propre
                If IsMinuscule(firstCharSecondWord) Or dico.Exists(cleanedWord) Then
                    newText = firstChar & " " & secondWord
                ' Règle 5 & 6 : Cas majuscule (non nom propre) ou «
                ElseIf IsMajuscule(firstCharSecondWord) Or firstCharSecondWord = "«" Then
                    newText = firstChar & ". " & secondWord
                End If

                If newText <> "" And match.Value <> newText Then
                    Set rng = doc.Range(match.FirstIndex, match.FirstIndex + match.Length)
                    If rng.Text = match.Value Then
                        rng.Text = newText
                    End If
                End If
            End If
        End If
    Next i
End Sub

''' <summary>
''' Procédure principale pour lancer toutes les corrections.
''' </summary>
Sub CorrectionPonctuationTotale()
    Dim doc As Object
    Dim dico As Object

    Set doc = ActiveDocument
    Set dico = ChargerDictionnaire()

    If dico Is Nothing Then
        MsgBox "Erreur lors du chargement du dictionnaire.", vbCritical
        Exit Sub
    End If

    ' Optimisation des performances
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Correction Ponctuation"

    ' Exécution des règles dans l'ordre logique
    AppliquerRegles123 doc, dico
    AppliquerRegle4 doc, dico
    AppliquerRegles56 doc, dico

    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True

    MsgBox "La correction de la ponctuation est terminée.", vbInformation, "Jules Macro"
End Sub

''' <summary>
''' Applique la règle 4 (fusion de paragraphes).
''' </summary>
Sub AppliquerRegle4(doc As Object, dico As Object)
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    Dim firstChar As String, secondWord As String
    Dim cleanedWord As String
    Dim rng As Object

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = False
    ' Motif : [Minuscule] [Point] [Retour Paragraphe \r] [Minuscule ou «]
    regEx.Pattern = "([a-zà-âæçéèêëîïôœùûüÿ])\.\r([a-zà-âæçéèêëîïôœùûüÿ«][a-zà-âæçéèêëîïôœùûüÿ-]*)"

    On Error Resume Next
    Set matches = regEx.Execute(doc.Range.Text)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    For i = matches.Count - 1 To 0 Step -1
        Set match = matches.Item(i)
        If match.SubMatches.Count >= 2 Then
            firstChar = match.SubMatches(0)
            secondWord = match.SubMatches(1)

            cleanedWord = secondWord
            If Left(cleanedWord, 1) = "«" Then cleanedWord = Mid(cleanedWord, 2)

            ' Retirer ponctuation collée à la fin du mot extrait
            Do While Len(cleanedWord) > 0 And Not (IsMinuscule(Right(cleanedWord, 1)) Or IsMajuscule(Right(cleanedWord, 1)))
                cleanedWord = Left(cleanedWord, Len(cleanedWord) - 1)
            Loop

            If Not dico.Exists(cleanedWord) Then
                Set rng = doc.Range(match.FirstIndex, match.FirstIndex + match.Length)
                If rng.Text = match.Value Then
                    rng.Text = firstChar & " " & secondWord
                End If
            End If
        End If
    Next i
End Sub
