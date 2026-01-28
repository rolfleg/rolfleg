REM  *****  BASIC  *****

Option Explicit

' =================================================================================================
' === FONCTION : getDictionnaireNomsPropres
' === DESCRIPTION : Charge un dictionnaire de noms propres à partir d'un fichier texte dans
' ===             le profil utilisateur de LibreOffice. Si le fichier n'existe pas, il est créé.
' === RETOUR : Un objet com.sun.star.container.EnumerableMap contenant les noms propres.
' =================================================================================================
Function getDictionnaireNomsPropres()
    ' --- Déclarations des variables ---
    Dim oPathSubst As Object
    Dim sUserProfileURL As String
    Dim sDicoURL As String
    Dim sDicoFileName As String

    sDicoFileName = "DictionnaireNomsPropres.txt"

    On Error GoTo ErreurSysteme

    ' --- 1. Obtenir le chemin du profil utilisateur de manière dynamique ---
    ' Cela rend la macro portable sur différentes installations de LibreOffice.
    oPathSubst = createUnoService("com.sun.star.comp.framework.PathSubstitution")
    sUserProfileURL = oPathSubst.getSubstituteVariableValue("$(user)")

    ' --- 2. Construire le chemin complet vers le fichier dictionnaire ---
    ' On s'assure que l'URL se termine par un "/" avant d'ajouter le nom du fichier.
    If Right(sUserProfileURL, 1) <> "/" Then
        sUserProfileURL = sUserProfileURL & "/"
    End If
    sDicoURL = sUserProfileURL & sDicoFileName

    ' --- 3. Vérifier si le fichier dictionnaire existe. Le créer si nécessaire. ---
    Dim oSFA As Object
    Dim oOutputStream As Object
    Dim oTextOutputStream As Object
    Dim sContenuInitial As String

    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")

    If Not oSFA.exists(sDicoURL) Then
        ' Le fichier n'existe pas, on le crée avec un contenu par défaut.
        ' On utilise Chr(10) (Line Feed) pour la compatibilité entre systèmes.
        sContenuInitial = "Jean-Claude" & Chr(10) & "Microsoft" & Chr(10) & "Paris"

        ' On utilise un flux (stream) pour écrire le fichier avec l'encodage UTF-8
        oOutputStream = createUnoService("com.sun.star.io.Pipe")
        oTextOutputStream = createUnoService("com.sun.star.io.TextOutputStream")
        oTextOutputStream.setOutputStream(oOutputStream)
        oTextOutputStream.setEncoding("UTF-8")
        oTextOutputStream.writeString(sContenuInitial)
        ' Il faut fermer le flux de texte pour que le contenu soit "flushé" vers le pipe
        oTextOutputStream.closeOutput()

        ' Le SimpleFileAccess prend en charge l'interface com.sun.star.io.XInputStream
        ' Le pipe fournit cette interface via sa méthode getInputStream()
        oSFA.writeFile(sDicoURL, oOutputStream.getInputStream())

        ' --- 4. Vérification post-création ---
        ' On s'assure que le fichier existe bien avant de continuer.
        If Not oSFA.exists(sDicoURL) Then
            MsgBox "Échec de la création du fichier dictionnaire à l'emplacement : " & Chr(13) & ConvertFromURL(sDicoURL), 16, "Erreur critique"
            getDictionnaireNomsPropres = CreateUnoService("com.sun.star.container.EnumerableMap") ' Retourne un dico vide
            Exit Function
        End If

        MsgBox("Le fichier DictionnaireNomsPropres.txt a été créé dans votre profil." & Chr(13) & _
               "Chemin : " & ConvertFromURL(sDicoURL) & Chr(13) & Chr(13) & _
               "Veuillez le personnaliser avec vos propres noms.", 64, "Information")
    End If

    ' --- 5. Lire le fichier et charger les mots dans un dictionnaire (EnumerableMap) ---
    Dim oDictionnaire As Object
    Dim oInputStream As Object
    Dim oTextInputStream As Object
    Dim sLigne As String

    oDictionnaire = CreateUnoService("com.sun.star.container.EnumerableMap")
    oInputStream = oSFA.openFileRead(sDicoURL)
    oTextInputStream = createUnoService("com.sun.star.io.TextInputStream")
    oTextInputStream.setInputStream(oInputStream)
    oTextInputStream.setEncoding("UTF-8") ' Important pour les accents

    ' On boucle sur chaque ligne du fichier
    Do While Not oTextInputStream.isEOF()
        sLigne = oTextInputStream.readLine()
        sLigne = Trim(sLigne) ' On nettoie les espaces superflus
        If sLigne <> "" Then
            ' On stocke le mot comme clé pour une recherche rapide. La valeur (True) n'a pas d'importance.
            oDictionnaire.put(sLigne, True)
        End If
    Loop

    oTextInputStream.closeInput() ' On ferme le flux

    ' --- On retourne le dictionnaire rempli ---
    getDictionnaireNomsPropres = oDictionnaire

    ' --- Gestion des erreurs ---
    Exit Function ' Sortie normale en fin de fonction

ErreurSysteme:
    MsgBox "Une erreur système est survenue dans getDictionnaireNomsPropres." & Chr(13) & _
           "Description : " & Err.Description, 16, "Erreur fatale"
    ' On retourne un objet vide pour que la macro principale puisse gérer l'échec
    getDictionnaireNomsPropres = CreateUnoService("com.sun.star.container.EnumerableMap")

End Function

' =================================================================================================
' === PROCÉDURE : CorrectionPonctuationAvancee
' === DESCRIPTION : Applique plusieurs règles de correction de ponctuation avancées sur le
' ===             document texte actuellement ouvert.
' =================================================================================================
Sub CorrectionPonctuationAvancee()
    ' --- Déclarations des variables ---
    Dim oDoc As Object
    Dim oDictionnaire As Object

    ' --- Initialisation ---
    oDoc = ThisComponent
    If IsNull(oDoc) Or Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox("Cette macro doit être exécutée sur un document texte.", 16, "Erreur")
        Exit Sub
    End If

    ' --- Chargement du dictionnaire ---
    oDictionnaire = getDictionnaireNomsPropres()

    ' On vérifie si le chargement du dictionnaire a réussi.
    ' getDictionnaireNomsPropres retourne un objet vide en cas d'erreur.
    If IsNull(oDictionnaire) Or oDictionnaire.getCount() = 0 Then
        MsgBox("Le dictionnaire des noms propres n'a pas pu être chargé ou est vide." & Chr(13) & _
               "Vérifiez le fichier DictionnaireNomsPropres.txt dans votre profil utilisateur." & Chr(13) & _
               "La macro va s'arrêter.", 16, "Erreur Dictionnaire")
        Exit Sub
    End If

    ' --- Application des règles de correction ---
    oDoc.lockControllers() ' Accélère les modifications en désactivant le rafraîchissement de l'écran

    AppliquerRegle2(oDoc)
    AppliquerRegles1et3(oDoc, oDictionnaire)
    AppliquerRegle4(oDoc, oDictionnaire)

    oDoc.unlockControllers() ' Réactive le rafraîchissement

    ' --- Finalisation ---
    MsgBox("La correction de la ponctuation est terminée.", 64, "Opération réussie")

End Sub


' =================================================================================================
' === RÈGLE 2 : Mot en minuscule suivi d'un point puis d'un mot en minuscule
' =================================================================================================
Sub AppliquerRegle2(oDoc As Object)
    ' DESCRIPTION : Si un mot se terminant par une minuscule est suivi d'un mot commençant
    '               par une minuscule, supprime le point et ne met qu'un seul espace.
    On Error GoTo ErreurRegle2

    Dim oSearch As Object
    Dim oResult As Object
    Dim i As Long
    Dim oRange As Object
    Dim sTrouve As String
    Dim sRemplace As String
    Dim iPos As Integer

    oSearch = oDoc.createSearchDescriptor()
    oSearch.SearchRegularExpression = True
    ' Regex:
    ' \b\p{L}*[[:lower:]] -> Un mot (lettres Unicode) se terminant par une minuscule.
    ' \.                  -> Un point littéral.
    ' \s+                 -> Un ou plusieurs espaces.
    ' \b[[:lower:]]\p{L}* -> Un mot (lettres Unicode) commençant par une minuscule.
    oSearch.setSearchString("\b\p{L}*[[:lower:]]\.\s+\b[[:lower:]]\p{L}*")

    oResult = oDoc.findAll(oSearch)

    If Not IsNull(oResult) Then
        For i = oResult.getCount() - 1 To 0 Step -1
            oRange = oResult.getByIndex(i)
            sTrouve = oRange.getString()

            iPos = InStr(sTrouve, ".")
            If iPos > 0 Then
                sRemplace = Left(sTrouve, iPos - 1) & " " & Trim(Mid(sTrouve, iPos + 1))
                oRange.setString(sRemplace)
            End If
        Next i
    End If

    Exit Sub
ErreurRegle2:
    MsgBox "Une erreur est survenue lors de l'application de la Règle 2." & Chr(13) & _
           "Description : " & Err.Description, 16, "Erreur Règle 2"
End Sub

' =================================================================================================
' === RÈGLES 1 & 3 : Mot en minuscule suivi d'un mot en MAJUSCULE
' =================================================================================================
Sub AppliquerRegles1et3(oDoc As Object, oDictionnaire As Object)
    ' DESCRIPTION : Gère deux cas :
    ' Règle 1: Si le 2e mot est un nom propre, supprime le point.
    ' Règle 3: Si le 2e mot n'est pas un nom propre, ajoute un point.
    On Error GoTo ErreurRegles1et3

    Dim oSearch As Object
    Dim oResult As Object
    Dim i As Long
    Dim oRange As Object
    Dim sTrouve As String
    Dim sPremierMorceau As String, sSecondMorceau As String, sPremierMorceauSansPoint As String
    Dim bPointPresent As Boolean
    Dim iPosEspace As Integer

    oSearch = oDoc.createSearchDescriptor()
    oSearch.SearchRegularExpression = True
    ' Regex:
    ' \b\p{L}*[[:lower:]] -> Mot se terminant par une minuscule.
    ' \.?                 -> Suivi d'un point optionnel.
    ' \s+                 -> Suivi d'un ou plusieurs espaces.
    ' \b[[:upper:]]\p{L}* -> Suivi d'un mot commençant par une majuscule.
    oSearch.setSearchString("\b\p{L}*[[:lower:]]\.?\s+\b[[:upper:]]\p{L}*")

    oResult = oDoc.findAll(oSearch)

    If Not IsNull(oResult) Then
        For i = oResult.getCount() - 1 To 0 Step -1
            oRange = oResult.getByIndex(i)
            sTrouve = oRange.getString()

            ' On sépare les deux parties de la chaîne trouvée
            iPosEspace = InStr(sTrouve, " ")
            sPremierMorceau = RTrim(Left(sTrouve, iPosEspace - 1))
            sSecondMorceau = Trim(Mid(sTrouve, iPosEspace))

            bPointPresent = (Right(sPremierMorceau, 1) = ".")

            If bPointPresent Then
                sPremierMorceauSansPoint = Left(sPremierMorceau, Len(sPremierMorceau) - 1)
            Else
                sPremierMorceauSansPoint = sPremierMorceau
            End If

            ' --- Application de la logique ---
            If oDictionnaire.hasByName(sSecondMorceau) Then
                ' RÈGLE 1 : Le mot est un nom propre. Il NE FAUT PAS de point.
                If bPointPresent Then
                    oRange.setString(sPremierMorceauSansPoint & " " & sSecondMorceau)
                End If
            Else
                ' RÈGLE 3 : Le mot n'est pas un nom propre. Il FAUT un point.
                If Not bPointPresent Then
                    oRange.setString(sPremierMorceauSansPoint & ". " & sSecondMorceau)
                End If
            End If
        Next i
    End If

    Exit Sub
ErreurRegles1et3:
    MsgBox "Une erreur est survenue lors de l'application des Règles 1 & 3." & Chr(13) & _
           "Description : " & Err.Description, 16, "Erreur Règles 1 & 3"
End Sub

' =================================================================================================
' === RÈGLE 4 : Fusion de paragraphes
' =================================================================================================
Sub AppliquerRegle4(oDoc As Object, oDictionnaire As Object)
    ' DESCRIPTION : Fusionne un paragraphe avec le suivant sous certaines conditions.
    On Error GoTo ErreurRegle4

    Dim oParas As Object
    Dim i As Long
    Dim oParaActuel As Object, oParaSuivant As Object
    Dim sTexteParaActuel As String, sTexteParaSuivant As String
    Dim oCursor As Object
    Dim sDernierMot As String
    Dim sPremierCarSuivant As String
    Dim aMots As Variant

    oParas = oDoc.getText().createEnumeration()
    Dim oParaArray() As Object
    Dim nCount As Long
    nCount = 0

    ' On stocke les paragraphes dans un tableau pour un accès inversé facile.
    Do While oParas.hasMoreElements()
        Redim Preserve oParaArray(nCount)
        oParaArray(nCount) = oParas.nextElement()
        nCount = nCount + 1
    Loop

    ' On parcourt les paragraphes de l'avant-dernier au premier.
    For i = UBound(oParaArray) - 1 To 0 Step -1
        oParaActuel = oParaArray(i)
        oParaSuivant = oParaArray(i+1)

        sTexteParaActuel = Trim(oParaActuel.getString())
        sTexteParaSuivant = Trim(oParaSuivant.getString())

        ' Condition 1: Le paragraphe actuel n'est pas vide et se termine par un point.
        If sTexteParaActuel <> "" And Right(sTexteParaActuel, 1) = "." Then

            ' Condition 2: Le paragraphe suivant n'est pas vide.
            If sTexteParaSuivant <> "" Then
                sPremierCarSuivant = Left(sTexteParaSuivant, 1)

                ' Condition 3: Le paragraphe suivant commence par une minuscule ou un guillemet.
                If (sPremierCarSuivant >= "a" And sPremierCarSuivant <= "z") Or sPremierCarSuivant = "«" Then

                    ' On isole le dernier mot du paragraphe actuel
                    aMots = Split(sTexteParaActuel, " ")
                    sDernierMot = aMots(UBound(aMots))
                    ' On enlève le point final
                    sDernierMot = Left(sDernierMot, Len(sDernierMot) - 1)

                    ' Condition 4: Le dernier mot n'est pas un nom propre.
                    If Not oDictionnaire.hasByName(sDernierMot) Then

                        ' --- Toutes les conditions sont remplies : on fusionne. ---
                        ' On utilise un curseur pour modifier le document sans perdre le formatage.

                        ' 1. On supprime le point à la fin du premier paragraphe.
                        oCursor = oParaActuel.getText().createCursorByRange(oParaActuel)
                        oCursor.gotoEnd(False)      ' Aller à la fin du texte du paragraphe
                        oCursor.goLeft(1, True)     ' Sélectionner le dernier caractère (le point)
                        oCursor.setString("")       ' Le supprimer

                        ' 2. On remplace la marque de paragraphe par un espace pour fusionner.
                        oCursor.gotoEnd(False)      ' Revenir à la fin
                        oCursor.goRight(1, True)    ' Sélectionner la marque de paragraphe
                        oCursor.setString(" ")      ' Remplacer par un espace

                    End If
                End If
            End If
        End If
    Next i

    Exit Sub
ErreurRegle4:
    MsgBox "Une erreur est survenue lors de l'application de la Règle 4." & Chr(13) & _
           "Description : " & Err.Description, 16, "Erreur Règle 4"
End Sub
