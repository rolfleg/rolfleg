Rem Module 76
' Traduction optimisée avec conservation des styles et insertion rapide
' Améliorations significatives de la vitesse en traitant le texte par blocs

'====================================================
' TRADUCTION OPTIMISÉE AVEC CONSERVATION DES STYLES
' Version Word 2007+ - Conserve formatage complet
'====================================================

Sub TraduireDocument_OpenAI_Word_Rapide()
    ' Affiche un message indiquant le module actif (pour le débogage)
    MsgBox "Nous sommes dans le module 76, Insertion rapide et optimisée."

    Dim apiKey As String
    ' Récupère la clé API OpenAI de la variable d'environnement JCM5
    apiKey = Environ$("JCM5")
    If apiKey = "" Then
        MsgBox "Clé API absente (variable JCM5)", vbCritical
        Exit Sub
    End If

    ' Taille maximale d'un segment de texte à envoyer à l'API (en caractères)
    Dim maxSegment As Long
    maxSegment = 15000

    ' Gestionnaire d'erreurs
    On Error GoTo GestionErreur

    ' S'assure qu'un document est bien actif
    Dim doc As Document
    Set doc = ActiveDocument

    If doc.Path = "" Then
        MsgBox "Veuillez d'abord sauvegarder le document.", vbExclamation
        Exit Sub
    End If

    If doc.Paragraphs.Count = 0 Then
        MsgBox "Document vide.", vbExclamation
        Exit Sub
    End If

    ' Crée un dossier "Traduction" dans le répertoire du document s'il n'existe pas
    Dim dossierTraduction As String
    dossierTraduction = doc.Path & "\Traduction"
    If Dir(dossierTraduction, vbDirectory) = "" Then
        MkDir dossierTraduction
    End If

    ' Calcule le nombre total de caractères et estime le nombre de segments nécessaires
    Dim totalCaracteres As Long
    totalCaracteres = Len(doc.Range.text)
    Dim nbSegments As Long
    nbSegments = (totalCaracteres \ maxSegment) + 1

    ' Prépare le nom de base pour les fichiers de sortie
    Dim nomBase As String
    nomBase = Left(doc.Name, InStrRev(doc.Name, ".") - 1)
    If nomBase = "" Then nomBase = doc.Name

    ' Propose à l'utilisateur de sauvegarder ou non les segments individuels
    Dim sauvegarderSegments As Boolean
    Dim reponse As Integer

    reponse = MsgBox("Voulez-vous sauvegarder les segments individuels ?" & vbCrLf & vbCrLf & _
                     "NON = Plus rapide (recommandé)" & vbCrLf & _
                     "OUI = Sauvegarde de " & nbSegments & " fichiers (plus lent)", _
                     vbYesNoCancel + vbQuestion, "Option de sauvegarde")

    If reponse = vbCancel Then Exit Sub
    sauvegarderSegments = (reponse = vbYes)

    ' Construit et affiche un message de confirmation avec les détails de la traduction
    Dim message As String
    message = "Document : " & totalCaracteres & " caractères" & vbCrLf
    message = message & "Paragraphes : " & doc.Paragraphs.Count & vbCrLf
    message = message & "Segments : " & nbSegments & vbCrLf & vbCrLf
    message = message & "NOUVEAU : Conservation des styles !" & vbCrLf
    message = message & "(gras, italique, couleurs, polices...)" & vbCrLf & vbCrLf

    If sauvegarderSegments Then
        message = message & "Mode : Sauvegarde complète" & vbCrLf
    Else
        message = message & "Mode : RAPIDE" & vbCrLf
    End If

    message = message & "Temps estimé : " & EstimerTemps(nbSegments, sauvegarderSegments) & vbCrLf & vbCrLf
    message = message & "Continuer ?"

    If MsgBox(message, vbYesNo + vbInformation, "Confirmation") = vbNo Then
        Exit Sub
    End If

    ' Désactive la mise à jour de l'écran pour accélérer le traitement
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True

    Dim startTime As Double
    startTime = Timer

    ' Crée un nouveau document pour y placer la traduction
    Dim docFusionne As Document
    Set docFusionne = Documents.Add
    docFusionne.ShowGrammaticalErrors = False
    docFusionne.ShowSpellingErrors = False

    ' Collection pour stocker les styles de chaque paragraphe
    Dim stylesParas As New Collection

    Dim i As Long
    Dim segmentTextes As String
    Dim segmentDebut As Long
    Dim traduction As String
    Dim para As Paragraph
    Dim numeroSegment As Long
    Dim pourcentage As Long
    ' *** CORRECTION : variable dédiée pour le suivi des styles ***
    Dim styleIndex As Long
    styleIndex = 1

    ' *** AMÉLIORATION : Séparateur de paragraphe unique ***
    Const PARA_SEPARATOR As String = "[|||]"

    segmentTextes = ""
    numeroSegment = 0
    segmentDebut = 1

    ' Initialise l'objet pour les requêtes HTTP
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    ' Boucle sur tous les paragraphes du document source
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)

        ' Capture le style du paragraphe et l'ajoute à la collection
        stylesParas.Add CaptureStyleParagraphe(para)

        pourcentage = CLng((i / doc.Paragraphs.Count) * 100)

        ' *** AMÉLIORATION : Utilise le séparateur pour joindre les paragraphes ***
        Dim textePara As String
        textePara = para.Range.text
        ' Supprime le dernier caractère (marque de paragraphe) pour éviter les doublons
        If Len(textePara) > 0 Then
            textePara = Left(textePara, Len(textePara) - 1)
        End If

        ' Si le segment actuel dépasse la taille max, on le traite
        If Len(segmentTextes) + Len(textePara) + Len(PARA_SEPARATOR) > maxSegment And Len(segmentTextes) > 0 Then
            numeroSegment = numeroSegment + 1

            ' Met à jour la barre de statut avec la progression
            Dim tempsEcoule As Long
            tempsEcoule = CLng(Timer - startTime)
            Dim tempsRestant As Long
            If numeroSegment > 1 Then
                tempsRestant = CLng((tempsEcoule / (numeroSegment - 1)) * (nbSegments - numeroSegment + 1))
            End If

            Application.StatusBar = "Segment " & numeroSegment & "/" & nbSegments & _
                                  " (" & pourcentage & "%) | Ecoulé: " & FormatTemps(tempsEcoule) & _
                                  " | Restant: " & FormatTemps(tempsRestant)

            ' Traduit le segment
            traduction = TraduireSegment_Optimise(segmentTextes, apiKey, http)

            ' Gestion d'erreur de traduction
            If Left(traduction, 8) = "(Erreur" Then
                Application.StatusBar = False
                Application.ScreenUpdating = True
                MsgBox traduction, vbExclamation
                docFusionne.Close SaveChanges:=wdDoNotSaveChanges
                Exit Sub
            End If

            ' *** AMÉLIORATION : Nouvelle fonction pour insérer et styler ***
            AjouterTexteEtStylesRapide docFusionne, traduction, stylesParas, styleIndex

            ' Sauvegarde le segment si l'option est activée
            If sauvegarderSegments Then
                SauvegarderSegmentDocx traduction, numeroSegment, dossierTraduction
            End If

            ' Réinitialise le segment avec le paragraphe actuel
            segmentTextes = textePara
            segmentDebut = i
        Else
            ' Ajoute le paragraphe au segment en cours
            If Len(segmentTextes) > 0 Then
                segmentTextes = segmentTextes & PARA_SEPARATOR & textePara
            Else
                segmentTextes = textePara
            End If
        End If
    Next i

    ' Traite le dernier segment restant
    If Len(segmentTextes) > 0 Then
        numeroSegment = numeroSegment + 1
        Application.StatusBar = "Traduction du dernier segment..."
        traduction = TraduireSegment_Optimise(segmentTextes, apiKey, http)

        If Left(traduction, 8) <> "(Erreur" Then
            AjouterTexteEtStylesRapide docFusionne, traduction, stylesParas, styleIndex
            If sauvegarderSegments Then
                SauvegarderSegmentDocx traduction, numeroSegment, dossierTraduction
            End If
        End If
    End If

    ' Supprime le dernier paragraphe s'il est vide (souvent le cas)
    If docFusionne.Paragraphs.Count > 0 Then
        If Len(Trim(docFusionne.Paragraphs(docFusionne.Paragraphs.Count).Range.text)) <= 1 Then
            docFusionne.Paragraphs(docFusionne.Paragraphs.Count).Range.Delete
        End If
    End If

    ' Sauvegarde le document final
    Application.StatusBar = "Sauvegarde finale..."
    SauvegarderDocumentFusionne docFusionne, nomBase, dossierTraduction

    ' Réactive la mise à jour de l'écran et nettoie la barre de statut
    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' Affiche le message de fin avec les statistiques
    Dim tempsTotal As Long
    tempsTotal = CLng(Timer - startTime)

    Dim messageFin As String
    messageFin = "Traduction terminée !" & vbCrLf & vbCrLf
    messageFin = messageFin & "• Paragraphes : " & doc.Paragraphs.Count & vbCrLf
    messageFin = messageFin & "• Segments : " & numeroSegment & vbCrLf
    messageFin = messageFin & "• Styles conservés : OUI" & vbCrLf
    messageFin = messageFin & "• Temps total : " & FormatTemps(tempsTotal) & vbCrLf
    If numeroSegment > 0 Then
        messageFin = messageFin & "• Temps/segment : " & FormatTemps(CLng(tempsTotal / numeroSegment)) & vbCrLf & vbCrLf
    End If
    messageFin = messageFin & "Fichier : Fr_" & nomBase & ".docx" & vbCrLf
    messageFin = messageFin & "Dossier : " & dossierTraduction

    MsgBox messageFin, vbInformation, "Traduction terminée"
    Exit Sub

GestionErreur:
    ' En cas d'erreur, restaure l'affichage et affiche le message d'erreur
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'====================================================
' CAPTURER LE STYLE D'UN PARAGRAPHE
'====================================================
Function CaptureStyleParagraphe(para As Paragraph) As String
    On Error Resume Next
    Dim style As String

    ' Format : StyleNom|Police|Taille|Gras|Italique|Couleur|Alignement
    style = para.style.NameLocal & "|"
    style = style & para.Range.Font.Name & "|"
    style = style & CStr(para.Range.Font.Size) & "|"
    style = style & CStr(para.Range.Font.Bold) & "|"
    style = style & CStr(para.Range.Font.Italic) & "|"
    style = style & CStr(para.Range.Font.Color) & "|"
    style = style & CStr(para.Alignment)

    CaptureStyleParagraphe = style
End Function

'====================================================
' APPLIQUER LE STYLE À UN PARAGRAPHE
'====================================================
Sub AppliquerStyleParagraphe(para As Paragraph, styleInfo As String)
    On Error Resume Next

    Dim parties() As String
    parties = Split(styleInfo, "|")

    If UBound(parties) < 6 Then Exit Sub

    ' Appliquer le style nommé
    If parties(0) <> "" Then para.style = parties(0)

    ' Appliquer la police
    If parties(1) <> "" Then para.Range.Font.Name = parties(1)

    ' Appliquer la taille
    If IsNumeric(parties(2)) Then para.Range.Font.Size = CSng(parties(2))

    ' Appliquer gras/italique
    para.Range.Font.Bold = (parties(3) = "-1")
    para.Range.Font.Italic = (parties(4) = "-1")

    ' Appliquer couleur
    If IsNumeric(parties(5)) Then para.Range.Font.Color = CLng(parties(5))

    ' Appliquer alignement
    If IsNumeric(parties(6)) Then para.Alignment = CInt(parties(6))
End Sub

'====================================================
' AJOUTER TEXTE ET STYLES (NOUVELLE MÉTHODE RAPIDE)
'====================================================
Sub AjouterTexteEtStylesRapide(docDest As Document, texteTraduit As String, _
                               stylesParas As Collection, ByRef styleIndex As Long)
    ' *** AMÉLIORATION MAJEURE DE LA PERFORMANCE ***
    ' Insère le bloc de texte en une seule fois, puis applique les styles.

    ' 1. Préparer le texte en remplaçant le séparateur par des marques de paragraphe Word
    Const PARA_SEPARATOR As String = "[|||]"
    Dim textePourWord As String
    ' *** CORRECTION : Gère également les sauts de ligne manuels (Chr(11)) ***
    textePourWord = Replace(texteTraduit, PARA_SEPARATOR, vbCr)
    textePourWord = Replace(textePourWord, Chr(11), vbLf) ' Remplacer par un saut de ligne standard

    ' 2. Insérer le bloc de texte à la fin du document
    Dim rng As Range
    Set rng = docDest.Range
    rng.Collapse Direction:=wdCollapseEnd
    rng.InsertAfter textePourWord
    ' S'assurer que le dernier paragraphe est créé s'il n'y en a pas
    If Not rng.text Like "*" & vbCr Then
        rng.InsertParagraphAfter
    End If

    ' 3. Boucler sur les nouveaux paragraphes pour appliquer les styles
    Dim paragraphesCrees As Long
    paragraphesCrees = UBound(Split(textePourWord, vbCr)) + 1

    Dim i As Long
    Dim paraIndexDoc As Long

    ' Calculer l'index du premier paragraphe que nous venons d'ajouter
    Dim premierParaIndex As Long
    premierParaIndex = docDest.Paragraphs.Count - paragraphesCrees + 1
    If premierParaIndex <= 0 Then premierParaIndex = 1 ' Sécurité

    ' Boucle pour appliquer les styles
    For i = 0 To paragraphesCrees - 1
        paraIndexDoc = premierParaIndex + i
        If paraIndexDoc <= docDest.Paragraphs.Count And styleIndex <= stylesParas.Count Then
            Dim paraCree As Paragraph
            Set paraCree = docDest.Paragraphs(paraIndexDoc)
            AppliquerStyleParagraphe paraCree, stylesParas(styleIndex)
            styleIndex = styleIndex + 1
        Else
            ' Sortir si on dépasse le nombre de styles ou de paragraphes
            Exit For
        End If
    Next i
End Sub

'====================================================
' TRADUCTION OPTIMISÉE
'====================================================
Function TraduireSegment_Optimise(segment As String, apiKey As String, http As Object) As String
    On Error GoTo ErrAPI

    If Len(Trim(segment)) = 0 Then
        TraduireSegment_Optimise = ""
        Exit Function
    End If

    ' *** AMÉLIORATION : Mise à jour du prompt système pour préserver le séparateur ET les sauts de ligne manuels ***
    Dim systemPrompt As String
    systemPrompt = "You are an expert translator. Translate the following text to French. " & _
                 "It is CRITICAL that you PRESERVE the `[|||]` paragraph separators EXACTLY as they appear. " & _
                 "It is ALSO CRITICAL to preserve the `\u000B` characters, which represent manual line breaks. " & _
                 "Do not add, remove, or change them in any way."

    Dim json As String
    json = "{""model"":""gpt-4o-mini""," & _
           """messages"":[" & _
           "{""role"":""system"",""content"":""" & EscapeJSON(systemPrompt) & """}," & _
           "{""role"":""user"",""content"":""" & EscapeJSON(segment) & """}" & _
           "],""temperature"":0.3}"

    http.Open "POST", "https://api.openai.com/v1/chat/completions", False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Authorization", "Bearer " & apiKey

    http.send json

    If http.Status <> 200 Then
        TraduireSegment_Optimise = "(Erreur HTTP " & http.Status & ")"
        Exit Function
    End If

    Dim resp As String
    resp = http.responseText

    TraduireSegment_Optimise = ExtraireContent(resp)
    Exit Function

ErrAPI:
    TraduireSegment_Optimise = "(Erreur API : " & Err.Description & ")"
End Function

'====================================================
' SAUVEGARDE ET MANIPULATION DE FICHIERS (INCHANGÉ)
'====================================================

Sub SauvegarderSegmentDocx(texte As String, numero As Long, dossier As String)
    On Error Resume Next
    Dim docSegment As Document
    Set docSegment = Documents.Add
    docSegment.ShowGrammaticalErrors = False
    docSegment.ShowSpellingErrors = False

    Dim rng As Range
    Set rng = docSegment.Range
    rng.text = texte

    Dim nomFichier As String
    nomFichier = dossier & "\Segment_" & PadLeft(CStr(numero), 3, "0") & ".docx"
    docSegment.SaveAs nomFichier, FileFormat:=wdFormatXMLDocument
    docSegment.Close SaveChanges:=wdDoNotSaveChanges
End Sub

Sub SauvegarderDocumentFusionne(docFusionne As Document, nomBase As String, dossier As String)
    On Error Resume Next
    Dim nomFichier As String
    nomFichier = dossier & "\Fr_" & nomBase & ".docx"
    docFusionne.SaveAs nomFichier, FileFormat:=wdFormatXMLDocument
End Sub

'====================================================
' PARSING ET FORMATAGE JSON (INCHANGÉ)
'====================================================

Function ExtraireContent(resp As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(resp, """content"":""")
    If p1 = 0 Then p1 = InStr(resp, """content"": """)
    If p1 = 0 Then
        ExtraireContent = "(Erreur: content introuvable)"
        Exit Function
    End If

    If InStr(p1, resp, """content"": """) = p1 Then
        p1 = p1 + 12
    Else
        p1 = p1 + 11
    End If

    p2 = InStr(p1, resp, """,")
    If p2 = 0 Then p2 = InStr(p1, resp, """" & vbLf)
    If p2 = 0 Or p2 <= p1 Then
        ExtraireContent = "(Erreur: fin introuvable)"
        Exit Function
    End If

    Dim txt As String
    txt = Mid$(resp, p1, p2 - p1)
    ' Nettoyage standard des caractères échappés JSON
    ' *** CORRECTION : Gère le \n comme un séparateur de paragraphe vbCr, et \u000B pour les sauts de ligne manuels ***
    txt = Replace(txt, "\n", vbCr)
    txt = Replace(txt, "\r", "")
    txt = Replace(txt, "\t", vbTab)
    txt = Replace(txt, "\""", """")
    txt = Replace(txt, "\/", "/")
    txt = Replace(txt, "\\", "\")
    txt = Replace(txt, "\u000B", Chr(11))

    ExtraireContent = txt
End Function


Function EscapeJSON(s As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String

    result = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        Select Case ch
            Case "\"
                result = result & "\\"
            Case """"
                result = result & "\"""
            Case vbCr
                result = result & "\n"
            Case vbLf
                ' Ignorer
            Case vbTab
                result = result & "\t"
            ' *** CORRECTION : Gère explicitement le saut de ligne manuel ***
            Case Chr(11)
                result = result & "\u000B"
            Case Else
                If Asc(ch) < 32 Then
                    ' Conserve les autres caractères de contrôle si nécessaire
                    result = result & "\u" & Right("0000" & LCase(Hex(Asc(ch)))), 4)
                Else
                    result = result & ch
                End If
        End Select
    Next i
    EscapeJSON = result
End Function

'====================================================
' UTILITAIRES (INCHANGÉ)
'====================================================
Function FormatTemps(secondes As Long) As String
    Dim h As Long, m As Long, s As Long
    h = secondes \ 3600
    m = (secondes Mod 3600) \ 60
    s = secondes Mod 60

    If h > 0 Then
        FormatTemps = h & "h " & m & "m " & s & "s"
    ElseIf m > 0 Then
        FormatTemps = m & "m " & s & "s"
    Else
        FormatTemps = s & "s"
    End If
End Function

Function EstimerTemps(nbSegments As Long, avecSauvegarde As Boolean) As String
    Dim tempsParSegment As Long
    ' Réduction du temps estimé car la nouvelle méthode est plus rapide
    If avecSauvegarde Then
        tempsParSegment = 20
    Else
        tempsParSegment = 10
    End If

    Dim tempsTotal As Long
    tempsTotal = nbSegments * tempsParSegment
    EstimerTemps = FormatTemps(tempsTotal)
End Function

Function PadLeft(texte As String, longueur As Integer, caractere As String) As String
    PadLeft = String(longueur - Len(texte), caractere) & texte
End Function
