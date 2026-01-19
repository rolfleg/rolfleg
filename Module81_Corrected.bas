REM  *****  BASIC  *****

Option Explicit

' ================================================================
'                 PROCÉDURE PRINCIPALE
' ================================================================
sub main
    ' Supprime tous les points pour repartir d'une base propre.
    ' Assurez-vous que Module81.SupprimerTousLesPoints existe et fonctionne comme attendu.
    ' On le met en commentaire pour le moment pour se concentrer sur la logique d'ajout/suppression.
    ' Module81.SupprimerTousLesPoints

    ' Lance le traitement global basé sur les règles.
    Etape2_TraitementGlobal_VFinale()
end sub

' ================================================================
'    NOUVELLE ARCHITECTURE : TRAITEMENT GLOBAL DU DOCUMENT
' ================================================================
Public Sub Etape2_TraitementGlobal_VFinale()
    Dim oDoc As Object
    oDoc = ThisComponent
    If oDoc Is Nothing Then
        MsgBox "Erreur : aucun document actif.", 16, "Erreur"
        Exit Sub
    End If

    Dim startTime As Double
    startTime = Timer

    MsgBox "Début du traitement global du document...", 64, "Information"

    ' ---- CHARGEMENT DU DICTIONNAIRE ----
    Dim oDicoNomsPropres As Object
    oDicoNomsPropres = CharpZEAWYtiB6bJ16NuLbGCc6CZ6jJdKfb63()
    If oDicoNomsPropres.Count = 0 Then
        MsgBox "Avertissement : Le dictionnaire des noms propres est vide ou n'a pas pu être chargé.", 48, "Avertissement"
    End If

    ' ---- APPLICATION SÉQUENTIELLE DES RÈGLES ----
    ' L'ordre est crucial pour éviter les conflits entre les règles.

    ' 1. On fusionne les paragraphes qui n'auraient pas dû être coupés.
    Regle4_FusionnerParagraphes oDoc, oDicoNomsPropres

    ' 2. On ajoute les points là où il en manque.
    Regle3_AjouterPoints oDoc, oDicoNomsPropres

    ' 3. On supprime les points en trop.
    Regles1et2_SupprimerPointsIncorrects oDoc, oDicoNomsPropres

    MsgBox "Traitement terminé en " & Format(Timer - startTime, "0.00") & " secondes.", 64, "Terminé"

End Sub

' ================================================================
' RÈGLE 4 : Fusionner paragraphes (minuscule. Chr(13) minuscule)
' ================================================================
Sub Regle4_FusionnerParagraphes(oDoc As Object, oDico As Object)
    Dim oSearch As Object
    oSearch = oDoc.createSearchDescriptor()

    ' Expression régulière pour trouver: [minuscule][point][retour paragraphe][minuscule]
    oSearch.SearchString = "([a-zà-ÿ])\.(\n)([a-zà-ÿ])"
    oSearch.setPropertyValue("SearchRegularExpression", True)

    Dim oFound As Object
    oFound = oDoc.findAll(oSearch)

    If IsNull(oFound) Or oFound.getCount() = 0 Then Exit Sub

    Dim i As Long
    ' Itérer de la fin vers le début pour ne pas invalider les index
    For i = oFound.getCount() - 1 To 0 Step -1
        Dim oRange As Object
        oRange = oFound.getByIndex(i)

        Dim oCursor As Object
        oCursor = oDoc.Text.createTextCursorByRange(oRange)
        oCursor.gotoEnd(False)
        oCursor.gotoStartOfWord(False)
        oCursor.gotoEndOfWord(True)

        Dim sNextWord As String
        sNextWord = oCursor.getString()

        If Not EstDansDictionnaire(sNextWord, oDico) Then
            Dim oSubSearch As Object
            oSubSearch = oRange.createSearchDescriptor()
            oSubSearch.SearchString = "\.\n"
            oSubSearch.setPropertyValue("SearchRegularExpression", True)

            Dim oPointBreak As Object
            oPointBreak = oRange.findOne(oSubSearch)

            If Not IsNull(oPointBreak) Then
                oPointBreak.setString(" ")
            End If
        End If
    Next i
End Sub

' ================================================================
' RÈGLE 3 : Ajouter les points (minuscule Majuscule)
' ================================================================
Sub Regle3_AjouterPoints(oDoc As Object, oDico As Object)
    Dim oSearch As Object
    oSearch = oDoc.createSearchDescriptor()

    ' Expression régulière pour trouver: [minuscule][espace][Majuscule]
    oSearch.SearchString = "([a-zà-ÿ])( )([A-ZÀ-Ö])"
    oSearch.setPropertyValue("SearchRegularExpression", True)

    Dim oFound As Object
    oFound = oDoc.findAll(oSearch)

    If IsNull(oFound) Or oFound.getCount() = 0 Then Exit Sub

    Dim i As Long
    For i = oFound.getCount() - 1 To 0 Step -1
        Dim oRange As Object
        oRange = oFound.getByIndex(i)

        Dim oCursor As Object
        oCursor = oDoc.Text.createTextCursorByRange(oRange)
        oCursor.gotoEnd(False)
        oCursor.gotoStartOfWord(False)
        oCursor.gotoEndOfWord(True)

        Dim sNextWord As String
        sNextWord = oCursor.getString()

        If Not EstDansDictionnaire(sNextWord, oDico) Then
            ' Le TextRange trouvé contient "e M". Remplacer l'espace par ". "
            Dim oSubSearch As Object
            oSubSearch = oRange.createSearchDescriptor()
            oSubSearch.SearchString = " "

            Dim oSpace As Object
            oSpace = oRange.findOne(oSubSearch)

            If Not IsNull(oSpace) Then
                oSpace.setString(". ")
            End If
        End If
    Next i
End Sub

' ================================================================
' RÈGLES 1 & 2 : Supprimer les points incorrects
' ================================================================
Sub Regles1et2_SupprimerPointsIncorrects(oDoc As Object, oDico As Object)
    ' ---- Règle 2 : Supprimer point avant minuscule ----
    Dim oSearch As Object
    oSearch = oDoc.createSearchDescriptor()
    ' Expression régulière : [minuscule][point][espace][minuscule]
    oSearch.SearchString = "([a-zà-ÿ])(\. )([a-zà-ÿ])"
    oSearch.setPropertyValue("SearchRegularExpression", True)

    ' Remplacer directement toutes les occurrences, car il n'y a pas de condition
    oDoc.replaceAll(oSearch, "$1 $3") ' $1 est le 1er groupe (lettre), $3 le 3e.

    ' ---- Règle 1 : Supprimer point avant nom propre ----
    oSearch = oDoc.createSearchDescriptor()
    ' Expression régulière : [minuscule][point][espace][Majuscule]
    oSearch.SearchString = "([a-zà-ÿ])(\. )([A-ZÀ-Ö])"
    oSearch.setPropertyValue("SearchRegularExpression", True)

    Dim oFound As Object
    oFound = oDoc.findAll(oSearch)

    If IsNull(oFound) Or oFound.getCount() = 0 Then Exit Sub

    Dim i As Long
    For i = oFound.getCount() - 1 To 0 Step -1
        Dim oRange As Object
        oRange = oFound.getByIndex(i)

        Dim oCursor As Object
        oCursor = oDoc.Text.createTextCursorByRange(oRange)
        oCursor.gotoEnd(False)
        oCursor.gotoStartOfWord(False)
        oCursor.gotoEndOfWord(True)

        Dim sNextWord As String
        sNextWord = oCursor.getString()

        ' Si le mot EST un nom propre, on supprime le point.
        If EstDansDictionnaire(sNextWord, oDico) Then
            oRange.setString(Replace(oRange.getString(), ". ", " "))
        End If
    Next i
End Sub


' ================================================================
'     SECTION DES FONCTIONS UTILITAIRES CONSERVÉES
' ================================================================

' ================================================================
'     Dictionnaire noms propres
' ================================================================
Function EstDansDictionnaire(mot As String, aDico As Object) As Boolean
    If Not IsObject(aDico) Then
        EstDansDictionnaire = False
        Exit Function
    End If

    Dim m As String
    m = NettoyerMot(mot)
    If Len(m) = 0 Then
        EstDansDictionnaire = False
        Exit Function
    End If

    EstDansDictionnaire = aDico.Exists(UCase(m))
End Function

Function NettoyerMot(mot As String) As String
    Dim i As Long
    Dim r As String
    r = ""

    For i = 1 To Len(mot)
        Dim c As String
        c = Mid(mot, i, 1)
        If c Like "[A-Za-zÀ-ÖØ-öø-ÿ]" Then
            r = r & c
        Else
            Exit For
        End If
    Next i

    NettoyerMot = r
End Function

' ================================================================
'     Vérification des Minuscules / Majuscules
' ================================================================
Function EstMinuscule(c As String) As Boolean
    If Len(c) <> 1 Then
        EstMinuscule = False
        Exit Function
    End If

    If (c >= "a" And c <= "z") Or (c >= "à" And c <= "ÿ") Then
        EstMinuscule = True
    Else
        EstMinuscule = False
    End If
End Function

Function EstMajuscule(c As String) As Boolean
    If Len(c) <> 1 Then
        EstMajuscule = False
        Exit Function
    End If

    If (c >= "A" And c <= "Z") Or (c >= "À" And c <= "Ö") Then
        EstMajuscule = True
    Else
        EstMajuscule = False
    End If
End Function

' ================================================================
'     Chargement du dictionnaire des Noms Propres (Version Corrigée pour LibreOffice)
' ================================================================
Function CharpZEAWYtiB6bJ16NuLbGCc6CZ6jJdKfb63() As Object
    Dim d As Object
    d = CreateObject("Scripting.Dictionary") ' Compatible avec LO sur Windows
    d.CompareMode = 1 ' Mode insensible à la casse

    Dim sChemin As String
    sChemin = "C:\Users\jeanp\AppData\Roaming\LibreOffice\4\user\wordbook\noms propres.dic"
    Dim sUrl As String
    sUrl = ConvertToURL(sChemin)

    Dim oSFA As Object
    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")

    If Not oSFA.exists(sUrl) Then
        MsgBox "Le fichier dictionnaire '" & sChemin & "' n'a pas été trouvé.", 48, "Avertissement"
        CharpZEAWYtiB6bJ16NuLbGCc6CZ6jJdKfb63 = d
        Exit Function
    End If

    Dim oInputStream As Object
    Dim oTextStream As Object
    Dim ligne As String

    On Error GoTo ErrorHandler

    oInputStream = oSFA.openFileRead(sUrl)
    oTextStream = createUnoService("com.sun.star.io.TextInputStream")
    oTextStream.setInputStream(oInputStream)
    oTextStream.setEncoding("UTF-8") ' Encodage standard

    Do While Not oTextStream.isEOF()
        ligne = oTextStream.readLine()
        ligne = Trim(ligne)
        If Len(ligne) > 0 Then
            ' Ignorer les en-têtes ou métadonnées du dictionnaire
            If Left(ligne, 1) <> "[" And InStr(ligne, ":") = 0 Then
                If Not d.Exists(UCase(ligne)) Then
                    d.Add UCase(ligne), 1
                End If
            End If
        End If
    Loop

    ' Fermer les flux
    oTextStream.closeInput()
    oInputStream.closeInput()

    CharpZEAWYtiB6bJ16NuLbGCc6CZ6jJdKfb63 = d
    Exit Function

ErrorHandler:
    MsgBox "Erreur lors de la lecture du fichier dictionnaire : " & Err.Description, 16, "Erreur"
    On Error Resume Next ' Empêche une erreur dans le gestionnaire d'erreur
    If Not IsNull(oTextStream) Then oTextStream.closeInput()
    If Not IsNull(oInputStream) Then oInputStream.closeInput()
    CharpZEAWYtiB6bJ16NuLbGCc6CZ6jJdKfb63 = d
End Function
