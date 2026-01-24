REM  *****  BASIC  *****

Option Explicit

Sub MainCorrectionPonctuation()
    Dim oDoc As Object
    Dim oText As Object
    Dim aNomsPropres As Object

    On Error GoTo GestionErreurs

    oDoc = ThisComponent
    If IsNull(oDoc) Or Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox "Cette macro doit être exécutée dans un document LibreOffice Writer.", 16, "Erreur"
        Exit Sub
    End If

    ' Afficher un message de confirmation avant de commencer
    If MsgBox("Cette macro va modifier le document actif pour corriger la ponctuation." & Chr(13) & "Voulez-vous continuer ?", 36, "Confirmation") <> 6 Then
        Exit Sub ' 6 = vbYes
    End If

    oText = oDoc.getText()

    ' Charger le dictionnaire des noms propres
    aNomsPropres = ChargerDictionnaireNomsPropres()
    If IsNull(aNomsPropres) Then
        Exit Sub ' Le message d'erreur est déjà géré dans la fonction
    End If

    ' Désactiver la mise à jour de l'affichage pour améliorer les performances
    oDoc.lockControllers()

    ' Appliquer les règles de correction
    AppliquerRegle1Et3(oDoc, aNomsPropres)
    AppliquerRegle2(oDoc)
    AppliquerRegle4(oDoc, aNomsPropres)

    ' Réactiver la mise à jour de l'affichage
    oDoc.unlockControllers()

    MsgBox "Correction de la ponctuation terminée.", 64, "Succès"

    Exit Sub

GestionErreurs:
    ' S'assurer que les controllers sont réactivés en cas d'erreur
    If Not IsNull(oDoc) And oDoc.hasControllers() And oDoc.isLocked() Then
        oDoc.unlockControllers()
    End If
    MsgBox "Une erreur inattendue est survenue : " & Err.Description, 16, "Erreur d'exécution"
End Sub

Function ChargerDictionnaireNomsPropres() As Object
    Dim sCheminDic As String
    Dim sURL As String
    Dim oSimpleFileAccess As Object
    Dim oInputStream As Object
    Dim oNomsPropres As Object
    Dim oPathSubst As Object
    Dim sUserProfileURL As String

    On Error GoTo GestionErreurs

    ' Utilisation d'un service de map pour une recherche rapide des noms propres.
    oNomsPropres = createUnoService("com.sun.star.container.EnumerableMap")

    ' Déterminer dynamiquement le chemin du dictionnaire pour la portabilité
    oPathSubst = createUnoService("com.sun.star.comp.framework.PathSubstitution")
    sUserProfileURL = oPathSubst.getSubstituteVariableValue("$(user)")
    sURL = sUserProfileURL & "/wordbook/noms propres.dic"

    ' Convertir l'URL en chemin système lisible pour les messages d'erreur
    sCheminDic = ConvertFromURL(sURL)

    oSimpleFileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")

    If oSimpleFileAccess.exists(sURL) Then
        oInputStream = oSimpleFileAccess.openFileRead(sURL)

        Dim oTextInputStream As Object
        oTextInputStream = createUnoService("com.sun.star.io.TextInputStream")
        oTextInputStream.setInputStream(oInputStream)
        oTextInputStream.setEncoding("UTF-8")

        Do While Not oTextInputStream.isEOF()
            Dim sLine As String
            sLine = oTextInputStream.readLine()
            If Trim(sLine) <> "" Then
                If Not oNomsPropres.hasByName(sLine) Then
                    oNomsPropres.putByName(sLine, True)
                End If
            End If
        Loop

        oInputStream.closeInput()
        oTextInputStream.closeInput()

        ChargerDictionnaireNomsPropres = oNomsPropres
    Else
        MsgBox "Le fichier dictionnaire ""noms propres.dic"" est introuvable." & Chr(13) & "Chemin attendu : " & sCheminDic, 16, "Erreur Fichier"
        ChargerDictionnaireNomsPropres = Null
    End If

    Exit Function

GestionErreurs:
    MsgBox "Erreur lors du chargement du dictionnaire : " & Err.Description, 16, "Erreur"
    ChargerDictionnaireNomsPropres = Null
End Function

Sub AppliquerRegle1Et3(oDoc As Object, aNomsPropres As Object)
    ' Règle 1: mot_minuscule [.] [espace] Nom_Propre -> mot_minuscule [espace] Nom_Propre
    ' Règle 3: mot_minuscule [espace] Mot_Majuscule_NON_Nom_Propre -> mot_minuscule. [espace] Mot_Majuscule
    Dim oSearch As Object
    Dim oTrouve As Object
    Dim i As Long

    oSearch = oDoc.createSearchDescriptor()
    ' Recherche un mot se terminant par une minuscule, suivi par un espace (ou un point et un espace), puis un mot avec une majuscule.
    oSearch.SearchString = "\b\w*[[:lower:]](\. | )[[:upper:]][^ ]*\b"
    oSearch.SearchRegularExpression = True

    oTrouve = oDoc.findAll(oSearch)

    If Not IsNull(oTrouve) Then
        For i = oTrouve.getCount() - 1 To 0 Step -1
            Dim oPlage As Object
            Dim sTextePlage As String
            Dim aMots As Variant
            Dim sMot2 As String

            oPlage = oTrouve.getByIndex(i)
            sTextePlage = oPlage.getString()

            ' Isoler le deuxième mot (celui avec la majuscule)
            If InStr(sTextePlage, ". ") > 0 Then
                aMots = Split(sTextePlage, ". ")
            Else
                aMots = Split(sTextePlage, " ")
            End If
            sMot2 = aMots(UBound(aMots))

            ' Vérifier si le mot est dans le dictionnaire
            If aNomsPropres.hasByName(sMot2) Then
                ' Règle 1: C'est un nom propre, on s'assure qu'il n'y a pas de point avant.
                oPlage.setString(Replace(sTextePlage, ". ", " "))
            Else
                ' Règle 3: Ce n'est pas un nom propre, on s'assure qu'il y a un point.
                If InStr(sTextePlage, ". ") = 0 Then
                    oPlage.setString(Replace(sTextePlage, " ", ". "))
                End If
            End If
        Next i
    End If
End Sub

Sub AppliquerRegle2(oDoc As Object)
    ' Règle 2: mot_terminant_minuscule. mot_commencant_minuscule -> mot_terminant_minuscule mot_commencant_minuscule
    Dim oSearch As Object
    Dim oTrouve As Object
    Dim i As Long

    oSearch = oDoc.createSearchDescriptor()

    ' Recherche un mot se terminant par une minuscule, suivi d'un point, d'un ou plusieurs espaces,
    ' puis d'un mot commençant par une minuscule ou un guillemet ouvrant.
    oSearch.SearchString = "\b\w*[[:lower:]]\. +\b(?:«|[[:lower:]])\w*"
    oSearch.SearchRegularExpression = True

    oTrouve = oDoc.findAll(oSearch)

    If Not IsNull(oTrouve) Then
        For i = oTrouve.getCount() - 1 To 0 Step -1
            Dim oPlage As Object
            Dim sTextePlage As String
            Dim sTexteNettoye As String
            Dim posPoint As Integer

            oPlage = oTrouve.getByIndex(i)
            sTextePlage = oPlage.getString()

            ' Trouve la position du premier point pour le supprimer
            posPoint = InStr(sTextePlage, ".")
            If posPoint > 0 Then
                ' Reconstruit la chaîne: partie avant le point + espace + partie après le point (nettoyée des espaces en trop)
                sTexteNettoye = Left(sTextePlage, posPoint - 1) & " " & Trim(Mid(sTextePlage, posPoint + 1))
                oPlage.setString(sTexteNettoye)
            End If
        Next i
    End If
End Sub

Sub AppliquerRegle4(oDoc As Object, aNomsPropres As Object)
    ' Règle 4: mot_minuscule.[FIN_PARAGRAPHE]mot_minuscule_NON_Nom_Propre -> mot_minuscule [espace] mot_minuscule
    Dim oSearch As Object
    Dim oTrouve As Object
    Dim i As Long

    ' Recherche un mot se terminant par une minuscule et un point à la fin d'un paragraphe.
    oSearch = oDoc.createSearchDescriptor()
    oSearch.SearchString = "\b\w*[[:lower:]]\.$"
    oSearch.SearchRegularExpression = True

    oTrouve = oDoc.findAll(oSearch)

    If Not IsNull(oTrouve) Then
        For i = oTrouve.getCount() - 1 To 0 Step -1
            Dim oPlage As Object
            Dim oCurseur As Object
            Dim sPremierMotParagrapheSuivant As String

            oPlage = oTrouve.getByIndex(i)

            ' Créer un curseur à la fin de la plage trouvée pour inspecter le paragraphe suivant
            oCurseur = oPlage.getText().createTextCursorByRange(oPlage.getEnd())

            ' Si le curseur peut se déplacer au paragraphe suivant...
            If oCurseur.gotoNextParagraph(False) Then
                ' Aller au début de ce paragraphe suivant
                oCurseur.gotoStartOfParagraph(False)
                ' L'étendre jusqu'au premier mot
                oCurseur.gotoEndOfWord(True)

                sPremierMotParagrapheSuivant = oCurseur.getString()

                ' Vérifier si le mot commence par une minuscule ou "«" et n'est pas un nom propre
                If (LCase(sPremierMotParagrapheSuivant) = sPremierMotParagrapheSuivant Or Left(sPremierMotParagrapheSuivant, 1) = "«") And _
                   Not aNomsPropres.hasByName(sPremierMotParagrapheSuivant) Then

                    ' 1. Supprimer le point à la fin du premier paragraphe
                    oPlage.setString(Left(oPlage.getString(), Len(oPlage.getString()) - 1))

                    ' 2. Fusionner les paragraphes
                    ' Créer un curseur à la fin du premier paragraphe (où se trouvait le point)
                    Dim oCurseurFusion As Object
                    oCurseurFusion = oPlage.getText().createTextCursorByRange(oPlage.getEnd())
                    ' Sélectionner le saut de paragraphe (1 caractère)
                    oCurseurFusion.goRight(1, True)
                    ' Le remplacer par un espace
                    oCurseurFusion.setString(" ")
                End If
            End If
        Next i
    End If
End Sub
