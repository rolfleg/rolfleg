Option Explicit

' Macro pour LibreOffice Writer
' Fusionne les paragraphes si le précédent se termine par une minuscule
' et que le suivant commence par un nom propre présent dans le dictionnaire.

Sub FusionParagrapheNomPropre()
    Dim oDoc As Object
    Dim oText As Object
    Dim oCursor As Object
    Dim oNextCursor As Object
    Dim oCollection As Object
    Dim sPrevText As String
    Dim sCurrText As String
    Dim sFirstWord As String
    Dim sLastChar As String
    Dim bMerged As Boolean

    oDoc = ThisComponent
    ' Vérification du type de document (meilleure pratique)
    If Not oDoc.supportsservice("com.sun.star.text.TextDocument") Then
        MsgBox "Cette macro ne fonctionne que dans un document LibreOffice Writer.", 16, "Erreur"
        Exit Sub
    End If

    ' Utilisation d'une Collection native pour le dictionnaire
    oCollection = New Collection
    If Not ChargerDicoNomsPropres(oCollection) Then
        MsgBox "Impossible de charger le dictionnaire des noms propres (Noms Propres.dic)." & Chr(10) & _
               "Vérifiez qu'il existe dans votre dossier wordbook du profil utilisateur.", 48, "Avertissement"
        ' On arrête le traitement si le dictionnaire est indispensable
        Exit Sub
    End If

    oText = oDoc.text
    ' On utilise un TextCursor pour l'itération, plus stable que l'énumération lors de modifs de structure
    oCursor = oText.createcursor()
    oCursor.gotostart(False)

    ' Optimisation de l'affichage
    oDoc.lockcontrollers()
    On Error GoTo ErrorHandler

    Do
        ' Sélectionner le paragraphe actuel pour analyse
        oCursor.gotostartofparagraph(False)
        oCursor.gotoendofparagraph(True)
        sPrevText = Trim(oCursor.string)
        bMerged = False

        If Len(sPrevText) > 0 Then
            sLastChar = Right(sPrevText, 1)
            ' Vérifie si le paragraphe se termine par une minuscule
            If EstMinuscule(sLastChar) Then
                ' Créer un curseur auxiliaire pour examiner le paragraphe suivant
                oNextCursor = oText.createcursorbyrange(oCursor)
                oNextCursor.collapsetoend()

                ' Essayer de passer au paragraphe suivant
                If oNextCursor.gotonextparagraph(False) Then
                    oNextCursor.gotoendofparagraph(True)
                    sCurrText = LTrim(oNextCursor.string)
                    sFirstWord = PremierMot(sCurrText)

                    ' Si le premier mot est dans le dictionnaire, on fusionne
                    If sFirstWord <> "" Then
                        If ElementDansCollection(oCollection, sFirstWord) Then
                            ' Fusion : on remplace le saut de paragraphe par un espace
                            oCursor.collapsetoend()
                            oCursor.goright(1, True) ' Sélectionne le saut de paragraphe
                            oCursor.string = " "
                            ' bMerged reste True pour re-tester le nouveau paragraphe fusionné (fusion en chaîne)
                            bMerged = True
                        End If
                    End If
                End If
            End If
        End If

        ' Si pas de fusion effectuée, on passe au paragraphe suivant
        If Not bMerged Then
            If Not oCursor.gotonextparagraph(False) Then Exit Do
        End If
    Loop

    oDoc.unlockcontrollers()
    MsgBox "Traitement terminé !", 64, "Information"
    Exit Sub

ErrorHandler:
    If Not IsNull(oDoc) Then oDoc.unlockcontrollers()
    MsgBox "Une erreur est survenue : " & Err.Description, 16, "Erreur"
End Sub

'---------------------------------------------------------
' Charge le dictionnaire depuis le profil utilisateur
'---------------------------------------------------------
Function ChargerDicoNomsPropres(oCollection As Object) As Boolean
    Dim sPath As String
    Dim sUrl As String
    Dim oSFA As Object
    Dim oInStream As Object
    Dim oTextStream As Object
    Dim sLine As String
    Dim bDictStarted As Boolean
    Dim oPathSubst As Object

    On Error GoTo ErrH

    ' Utilisation de PathSubstitution pour localiser le dossier utilisateur (portable)
    oPathSubst = createUnoService("com.sun.star.util.PathSubstitution")
    sPath = oPathSubst.getsubstitutevariablevalue("$(user)") & "/wordbook/Noms Propres.dic"
    sUrl = ConvertToURL(sPath)

    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not oSFA.exists(sUrl) Then
        ChargerDicoNomsPropres = False
        Exit Function
    End If

    ' Ouverture du fichier avec gestion de l'encodage UTF-8
    oInStream = oSFA.openfileread(sUrl)
    oTextStream = createUnoService("com.sun.star.io.TextInputStream")
    oTextStream.setinputstream(oInStream)
    oTextStream.setencoding("UTF-8")

    bDictStarted = False
    Do While Not oTextStream.iseof()
        sLine = Trim(oTextStream.readline())
        ' Format LO .dic : les mots commencent après "---"
        If sLine = "---" Then
            bDictStarted = True
        ElseIf bDictStarted Then
            If sLine <> "" Then
                ' Ajout à la collection (clé = mot pour recherche rapide)
                On Error Resume Next
                oCollection.Add(sLine, sLine)
                On Error GoTo ErrH
            End If
        End If
    Loop

    oInStream.closeinput()
    ChargerDicoNomsPropres = True
    Exit Function

ErrH:
    If Not IsNull(oInStream) Then oInStream.closeinput()
    ChargerDicoNomsPropres = False
End Function

'---------------------------------------------------------
' Extrait le premier mot d'une chaîne et retire la ponctuation
'---------------------------------------------------------
Function PremierMot(s As String) As String
    Dim i As Integer
    Dim sTemp As String
    sTemp = LTrim(s)
    ' On coupe au premier espace
    i = InStr(sTemp, " ")
    If i > 0 Then
        sTemp = Left(sTemp, i - 1)
    End If
    ' On retire la ponctuation de fin attachée au mot (ex: "Jean-Pierre,")
    Do While Len(sTemp) > 0
        If InStr(",;:!?.", Right(sTemp, 1)) > 0 Then
            sTemp = Left(sTemp, Len(sTemp) - 1)
        Else
            Exit Do
        End If
    Loop
    PremierMot = sTemp
End Function

'---------------------------------------------------------
' Vérifie si un caractère est une minuscule (gère les accents)
'---------------------------------------------------------
Function EstMinuscule(c As String) As Boolean
    If Len(c) <> 1 Then
        EstMinuscule = False
    Else
        ' Comparaison insensible à la casse vs sensible en LO Basic
        EstMinuscule = (c = LCase(c) And c <> UCase(c))
    End If
End Function

'---------------------------------------------------------
' Vérifie l'existence d'une clé dans la collection (évite l'erreur si absente)
'---------------------------------------------------------
Function ElementDansCollection(col As Object, sKey As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = col.Item(sKey)
    ElementDansCollection = (Err = 0)
    On Error GoTo 0
End Function
