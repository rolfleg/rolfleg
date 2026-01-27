REM  *****  BASIC  *****

' #####################################################################
' #                                                                   #
' #    MACRO DE CORRECTION DE PONCTUATION AVANCÉE POUR WRITER         #
' #                                                                   #
' #    Auteur : Jules (Assistant IA)                                  #
' #    Date   : 24/07/2024                                            #
' #    Version: 1.0                                                   #
' #                                                                   #
' #####################################################################

' =====================================================================
' = DÉBUT DE LA PROCÉDURE PRINCIPALE                                  =
' =====================================================================
Sub CorrectionPonctuationAvancee()
    ' --- Déclaration des variables ---
    Dim oDoc As Object
    Dim oDictionnaire As Object
    Dim oSearch As Object
    Dim oFound As Object
    Dim oRange As Object
    Dim i As Long
    Dim sTexte As String
    Dim aMots() As String
    Dim sMot2 As String

    ' --- Initialisation ---
    On Error GoTo ErreurGenerale
    oDoc = ThisComponent

    ' --- Étape 1 : Charger le dictionnaire des noms propres ---
    ' On appelle la fonction de support qui va chercher, créer si besoin,
    ' et charger le dictionnaire dans un objet optimisé pour la recherche.
    oDictionnaire = getDictionnaireNomsPropres()
    If IsNull(oDictionnaire) Or IsEmpty(oDictionnaire) Then
        ' Si la fonction n'a pas pu charger le dictionnaire, on arrête la macro.
        ' Un message d'erreur aura déjà été affiché par la fonction elle-même.
        Exit Sub
    End If

    ' --- Étape 2 : Préparer le descripteur de recherche ---
    ' C'est l'outil principal de LibreOffice pour trouver du texte.
    ' On le configure pour utiliser les expressions régulières (RegEx).
    oSearch = oDoc.createSearchDescriptor()
    oSearch.SearchRegularExpression = True
    oSearch.SearchCaseSensitive = True ' Important pour différencier majuscules et minuscules

    ' ---------------------------------------------------------------------
    ' --- APPLICATION DES RÈGLES DE CORRECTION                          ---
    ' ---------------------------------------------------------------------
    ' IMPORTANT : On parcourt TOUJOURS les résultats de la fin vers le début
    ' (For i = oFound.Count - 1 To 0 Step -1). Si on le faisait du début
    ' à la fin, chaque modification changerait la position des correspondances
    ' suivantes, ce qui corromprait le document.
    ' ---------------------------------------------------------------------

    ' --- Règle 4 : Corriger [minuscule].(Retour Paragraphe)[minuscule ou «] ---
    ' Explication du RegEx : \b(\w*[[:lower:]])\.\n([[:lower:]]|«)
    ' \b                  : Limite de mot (pour ne pas commencer au milieu d'un mot)
    ' (\w*[[:lower:]])    : Un mot se terminant par une minuscule
    ' \.                  : Un point littéral
    ' \n                  : Un saut de ligne/paragraphe
    ' ([[:lower:]]|«)     : Une minuscule OU un guillemet ouvrant
    oSearch.SearchString = "\b(\w*[[:lower:]])\.\n([[:lower:]]|«)"
    oFound = oDoc.findAll(oSearch)
    For i = oFound.Count - 1 To 0 Step -1
        oRange = oFound.getByIndex(i)
        sTexte = oRange.string

        ' On doit extraire le premier mot de la ligne suivante pour le vérifier
        Dim iPos As Integer
        Dim sPartie2 As String
        Dim sMotSuivant As String
        iPos = InStr(sTexte, Chr(10)) ' Chr(10) est le saut de ligne
        sPartie2 = Trim(Mid(sTexte, iPos + 1))

        If InStr(sPartie2, " ") > 0 Then
            sMotSuivant = Left(sPartie2, InStr(sPartie2, " ") - 1)
        Else
            sMotSuivant = sPartie2
        End If

        If Left(sMotSuivant, 1) = "«" Then sMotSuivant = Trim(Mid(sMotSuivant, 2))

        ' Si le mot suivant n'est PAS dans le dictionnaire, on corrige.
        If NOT oDictionnaire.hasByName(sMotSuivant) Then
            oRange.string = Replace(sTexte, "." & Chr(10), " ")
        End If
    Next i

    ' --- Règle 2 : Corriger [minuscule]. [minuscule] ---
    ' Explication du RegEx : \b(\w*[[:lower:]])\. ([[:lower:]])
    ' \b(\w*[[:lower:]])    : Un mot se terminant par une minuscule
    ' \.                  : Un point littéral suivi d'un espace
    ' ([[:lower:]])        : Une minuscule
    oSearch.SearchString = "\b(\w*[[:lower:]])\. ([[:lower:]])"
    oFound = oDoc.findAll(oSearch)
    For i = oFound.Count - 1 To 0 Step -1
        oRange = oFound.getByIndex(i)
        ' Remplacement simple : on enlève le point.
        oRange.string = Replace(oRange.string, ". ", " ")
    Next i

    ' --- Règle 1 : Corriger [minuscule]. [MAJUSCULE] si MAJUSCULE est un nom propre ---
    ' Explication du RegEx : \b(\w*[[:lower:]])\. ([[:upper:]]\w*)\b
    ' \b(\w*[[:lower:]])    : Un mot se terminant par une minuscule
    ' \.                  : Un point littéral suivi d'un espace
    ' ([[:upper:]]\w*)\b   : Un mot commençant par une majuscule
    oSearch.SearchString = "\b(\w*[[:lower:]])\. ([[:upper:]]\w*)\b"
    oFound = oDoc.findAll(oSearch)
    For i = oFound.Count - 1 To 0 Step -1
        oRange = oFound.getByIndex(i)
        sTexte = oRange.string
        ' On extrait le deuxième mot (le nom propre potentiel)
        aMots = Split(sTexte, ". ")
        sMot2 = aMots(1)

        ' Si le mot EST dans le dictionnaire, on enlève le point.
        If oDictionnaire.hasByName(sMot2) Then
            oRange.string = Replace(sTexte, ". ", " ")
        End If
    Next i

    ' --- Règle 3 : Corriger [minuscule] [MAJUSCULE] si MAJUSCULE n'est pas un nom propre ---
    ' Explication du RegEx : \b(\w*[[:lower:]]) ([[:upper:]]\w*)\b
    ' \b(\w*[[:lower:]])    : Un mot se terminant par une minuscule
    ' (espace)            : Un espace simple
    ' ([[:upper:]]\w*)\b   : Un mot commençant par une majuscule
    oSearch.SearchString = "\b(\w*[[:lower:]]) ([[:upper:]]\w*)\b"
    oFound = oDoc.findAll(oSearch)
    For i = oFound.Count - 1 To 0 Step -1
        oRange = oFound.getByIndex(i)
        sTexte = oRange.string
        ' On extrait le deuxième mot
        aMots = Split(sTexte, " ")
        sMot2 = aMots(1)

        ' Si le mot n'est PAS dans le dictionnaire, on ajoute un point.
        If NOT oDictionnaire.hasByName(sMot2) Then
            oRange.string = Replace(sTexte, " ", ". ")
        End If
    Next i

    MsgBox "Correction de la ponctuation terminée avec succès.", 64, "Opération Réussie"
    Exit Sub

ErreurGenerale:
    MsgBox "Une erreur inattendue est survenue." & Chr(13) & Chr(13) & "Description : " & Err.Description, 16, "Erreur Critique"
End Sub

' =====================================================================
' = DÉBUT DE LA FONCTION DE SUPPORT POUR LE DICTIONNAIRE              =
' =====================================================================
Function getDictionnaireNomsPropres() As Object
    ' --- Déclaration des variables ---
    Dim oPathSubst As Object
    Dim sProfilUserURL As String
    Dim sCheminDocURL As String
    Dim sfa As Object
    Dim oDictionnaire As Object
    Dim oInputStream As Object
    Dim oOutputStream As Object
    Dim sLigne As String

    On Error GoTo ErreurDictionnaire

    ' --- Étape 1 : Obtenir le chemin du profil utilisateur de manière dynamique ---
    ' Cette méthode est la plus fiable pour trouver le dossier utilisateur
    ' peu importe le système d'exploitation (Windows, Linux, Mac).
    oPathSubst = CreateUnoService("com.sun.star.comp.framework.PathSubstitution")
    sProfilUserURL = oPathSubst.getSubstituteVariableValue("$(user)")

    ' --- Étape 2 : Construire le chemin complet vers le fichier dictionnaire ---
    ' Note : Le chemin du prompt original pointe vers 'python', nous le respectons.
    ' Si ce dossier n'existe pas, la macro retournera une erreur gérée.
    sCheminDocURL = sProfilUserURL & "/Scripts/python/DictionnaireNomsPropres.txt"

    ' --- Étape 3 : Accéder au système de fichiers ---
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")

    ' --- Étape 4 : Vérifier si le fichier existe. Le créer sinon. ---
    If NOT sfa.exists(sCheminDocURL) Then
        ' Le fichier n'existe pas, on le crée avec des exemples.
        oOutputStream = sfa.openFileWrite(sCheminDocURL)
        Dim oTextOutputStream As Object
        oTextOutputStream = CreateUnoService("com.sun.star.io.TextOutputStream")
        oTextOutputStream.setOutputStream(oOutputStream)

        ' Écriture des exemples
        oTextOutputStream.writeString("Paris" & Chr(13))
        oTextOutputStream.writeString("Jean" & Chr(13))
        oTextOutputStream.writeString("Microsoft" & Chr(13))

        oTextOutputStream.closeOutput()
        oOutputStream.closeOutput()
    End If

    ' --- Étape 5 : Lire le fichier et le charger en mémoire ---
    ' On utilise un EnumerableMap car il est très rapide pour vérifier
    ' si un élément existe (méthode hasByName).
    oDictionnaire = CreateUnoService("com.sun.star.container.EnumerableMap")
    oInputStream = sfa.openFileRead(sCheminDocURL)

    Dim oTextInputStream As Object
    oTextInputStream = CreateUnoService("com.sun.star.io.TextInputStream")
    oTextInputStream.setInputStream(oInputStream)
    oTextInputStream.setEncoding("UTF-8") ' Assurer la compatibilité des caractères

    While NOT oTextInputStream.isEOF()
        sLigne = Trim(oTextInputStream.readLine())
        If sLigne <> "" Then
            ' On ajoute le mot au dictionnaire. La valeur (1) n'a pas d'importance.
            oDictionnaire.putByName(sLigne, 1)
        End If
    Wend

    oInputStream.closeInput()
    getDictionnaireNomsPropres = oDictionnaire
    Exit Function

ErreurDictionnaire:
    MsgBox "Erreur lors du chargement du dictionnaire des noms propres." & Chr(13) & Chr(13) & _
           "Vérifiez que le chemin existe et que vous avez les permissions :" & Chr(13) & _
           ConvertToURL(sProfilUserURL & "/Scripts/python/") & Chr(13) & Chr(13) & _
           "Description de l'erreur : " & Err.Description, 16, "Erreur Fichier Dictionnaire"
    getDictionnaireNomsPropres = Null
End Function
