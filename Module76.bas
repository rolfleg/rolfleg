Attribute VB_Name = "Module76"
Rem Module 76
' TRADUCTION SÉCURISÉE AVEC CONSERVATION DES IMAGES ET FORMES
' Crée une copie du document et ne traduit que le texte,
' préservant ainsi tous les éléments non textuels (images, dessins).

Option Explicit

'====================================================
' TRADUCTION AVEC CONSERVATION DES ÉLÉMENTS GRAPHIQUES
'====================================================

Public Sub TraduireDocument_OpenAI_Word_Module76()
    ' Affiche un message pour identifier la version de la macro en cours d'exécution.
    MsgBox "Module 76 : Traduction sécurisée qui conserve les images et les formes."

    ' Récupération de la clé API OpenAI depuis les variables d'environnement.
    Dim apiKey As String
    apiKey = Environ$("JCM5")
    If apiKey = "" Then
        MsgBox "La clé API OpenAI (variable d'environnement JCM5) est introuvable.", vbCritical, "Erreur Critique"
        Exit Sub
    End If

    ' Paramètre de la taille maximale pour les segments de texte à envoyer à l'API.
    Dim maxSegment As Long
    maxSegment = 15000 ' Corresponds à environ 4000 tokens, une limite de sécurité.

    ' Gestionnaire d'erreurs pour intercepter les problèmes imprévus.
    On Error GoTo GestionErreur

    ' Vérifie si un document est bien actif.
    If ActiveDocument Is Nothing Then
        MsgBox "Aucun document n'est ouvert.", vbExclamation, "Opération impossible"
        Exit Sub
    End If

    Dim docSource As Document
    Set docSource = ActiveDocument

    ' S'assure que le document a été enregistré au moins une fois pour avoir un chemin d'accès.
    If docSource.Path = "" Then
        MsgBox "Veuillez d'abord sauvegarder le document avant de lancer la traduction.", vbExclamation, "Action requise"
        Exit Sub
    End If

    ' S'assure que le document n'est pas vide.
    If docSource.Paragraphs.Count = 0 Then
        MsgBox "Le document est vide et ne peut pas être traduit.", vbExclamation, "Document vide"
        Exit Sub
    End If

    ' Avertissement sur la nature de la macro et demande de confirmation.
    Dim confirmation As VbMsgBoxResult
    confirmation = MsgBox("Cette macro va créer une copie de votre document (avec le préfixe 'Fr_') pour y effectuer la traduction." & vbCrLf & vbCrLf & _
                          "Elle préservera les images, dessins et autres formes en ignorant les paragraphes qui en contiennent." & vbCrLf & vbCrLf & _
                          "Le document original ne sera pas modifié." & vbCrLf & vbCrLf & _
                          "Voulez-vous continuer ?", vbYesNo + vbQuestion, "Confirmation de la traduction")

    If confirmation = vbNo Then
        MsgBox "Traduction annulée par l'utilisateur.", vbInformation
        Exit Sub
    End If

    ' --- Création de la copie du document ---
    Application.ScreenUpdating = False ' Désactive le rafraîchissement de l'écran pour la performance.

    Dim nomBase As String
    Dim posPoint As Long
    posPoint = InStrRev(docSource.Name, ".")

    ' Gère le cas où le nom de fichier n'a pas d'extension pour éviter une erreur.
    If posPoint > 0 Then
        nomBase = Left(docSource.Name, posPoint - 1)
    Else
        nomBase = docSource.Name
    End If

    Dim nouveauNom As String
    nouveauNom = docSource.Path & "\Fr_" & nomBase & ".docx"

    Dim docTraduit As Document
    ' Copie le document source et travaille sur la copie.
    Set docTraduit = docSource.SaveAs2(FileName:=nouveauNom, AddToRecentFiles:=False)

    ' --- Nouvelle Logique de Traduction "au fil de l'eau" ---

    Dim startTime As Double
    startTime = Timer

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    Dim i As Long
    Dim parasIgnorés As Long
    Dim parasTraduits As Long
    Dim segmentTexte As String
    Dim indicesParasSegment As New Collection

    parasIgnorés = 0
    parasTraduits = 0

    ' Boucle unique qui parcourt tout le document une seule fois.
    For i = 1 To docTraduit.Paragraphs.Count

        Dim para As Paragraph
        Set para = docTraduit.Paragraphs(i)

        Dim estDernierPara As Boolean
        estDernierPara = (i = docTraduit.Paragraphs.Count)

        Dim contientImage As Boolean
        contientImage = (para.Range.InlineShapes.Count > 0 Or para.Range.ShapeRange.Count > 0)

        Dim estTexteVide As Boolean
        estTexteVide = (Len(Trim(para.Range.Text)) <= 1)

        ' Condition pour traiter le segment accumulé :
        ' 1. On rencontre une image
        ' 2. On a atteint la taille max
        ' 3. C'est le dernier paragraphe
        Dim declencherTraduction As Boolean
        declencherTraduction = contientImage Or Len(segmentTexte) > maxSegment Or estDernierPara

        If Not contientImage And Not estTexteVide Then
            ' Si c'est du texte, on l'ajoute au segment.
            indicesParasSegment.Add i
            segmentTexte = segmentTexte & para.Range.Text & "[|||]"
        Else
            parasIgnorés = parasIgnorés + 1
        End If

        If declencherTraduction And Len(segmentTexte) > 0 Then

            ' Met à jour la barre de statut
            Dim pourcentage As Long
            pourcentage = CLng((i / docTraduit.Paragraphs.Count) * 100)
            Application.StatusBar = "Traduction en cours... " & pourcentage & "%"

            ' Enlève le dernier séparateur
            segmentTexte = Left(segmentTexte, Len(segmentTexte) - 5)

            ' Traduit le segment
            Dim traduction As String
            traduction = TraduireSegment_Optimise(segmentTexte, apiKey, http)

            If Left(traduction, 8) = "(Erreur" Then
                MsgBox "La traduction a échoué. " & traduction, vbCritical, "Erreur de traduction"
                docTraduit.Close SaveChanges:=False
                GoTo FinMacro
            End If

            ' Remplace le texte dans le document
            RemplacerTexteParIndices docTraduit, indicesParasSegment, traduction
            parasTraduits = parasTraduits + indicesParasSegment.Count

            ' Réinitialise pour le prochain segment
            segmentTexte = ""
            Set indicesParasSegment = New Collection
        End If
    Next i

    ' --- Finalisation ---
    Dim tempsTotal As Long
    tempsTotal = CLng(Timer - startTime)

    docTraduit.Save
    docTraduit.Close

    Dim messageFin As String
    messageFin = "Traduction terminée avec succès !" & vbCrLf & vbCrLf & _
                 "• Paragraphes traduits : " & parasTraduits & vbCrLf & _
                 "• Paragraphes ignorés (avec images ou vides) : " & parasIgnorés & vbCrLf & _
                 "• Temps total : " & FormatTemps(tempsTotal) & vbCrLf & vbCrLf & _
                 "Le document traduit a été sauvegardé sous :" & vbCrLf & nouveauNom

    MsgBox messageFin, vbInformation, "Traduction terminée"

FinMacro:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

GestionErreur:
    ' Réactive les mises à jour de l'écran en cas d'erreur.
    Application.ScreenUpdating = True
    ' Affiche un message d'erreur détaillé.
    MsgBox "Une erreur inattendue est survenue." & vbCrLf & vbCrLf & _
           "Erreur n° : " & Err.Number & vbCrLf & _
           "Description : " & Err.Description, vbCritical, "Erreur d'Exécution"
End Sub

'====================================================
' SECTION DES FONCTIONS UTILITAIRES
' Fonctions réutilisées du Module 75 pour la communication
' avec l'API OpenAI et la manipulation de chaînes.
'====================================================

'----------------------------------------------------
' TRADUCTION D'UN SEGMENT DE TEXTE VIA API OPENAI
'----------------------------------------------------
Private Function TraduireSegment_Optimise(ByVal segment As String, ByVal apiKey As String, ByVal http As Object) As String
    On Error GoTo ErrAPI

    If Len(Trim(segment)) = 0 Then
        TraduireSegment_Optimise = ""
        Exit Function
    End If

    Dim systemPrompt As String
    systemPrompt = "You are a professional document translator. Translate the text to French. The user's text may contain paragraphs separated by '[|||]'. You MUST preserve this '[|||]' separator exactly as it is in the output. Also preserve all special line breaks like \u000B."

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

'----------------------------------------------------
' EXTRACTION DU CONTENU DE LA RÉPONSE JSON
'----------------------------------------------------
Private Function ExtraireContent(ByVal resp As String) As String
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
    txt = Replace(txt, "\n", Chr(13))
    txt = Replace(txt, "\r", "")
    txt = Replace(txt, "\t", vbTab)
    txt = Replace(txt, "\""", """")
    txt = Replace(txt, "\/", "/")
    txt = Replace(txt, "\\", "\")
    txt = Replace(txt, "\u000B", Chr(11)) ' Gestion des sauts de ligne manuels

    ExtraireContent = txt
End Function

'----------------------------------------------------
' ÉCHAPPEMENT DES CARACTÈRES SPÉCIAUX POUR JSON
'----------------------------------------------------
Private Function EscapeJSON(ByVal s As String) As String
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
            Case Chr(13)
                result = result & "\n"
            Case Chr(10)
                ' Ignorer le Chr(10) (Line Feed) qui suit souvent un Chr(13)
            Case vbTab
                result = result & "\t"
            Case Chr(11)
                ' Préserve les sauts de ligne manuels
                result = result & "\u000B"
            Case Else
                If Asc(ch) < 32 Then
                    result = result & "\u" & Right("0000" & Hex(Asc(ch)), 4)
                Else
                    result = result & ch
                End If
        End Select
    Next i
    EscapeJSON = result
End Function

'----------------------------------------------------
' FORMATAGE DU TEMPS EN HEURES, MINUTES, SECONDES
'----------------------------------------------------
Private Function FormatTemps(ByVal secondes As Long) As String
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

'----------------------------------------------------
' REMPLACEMENT DU TEXTE PAR INDICES DE PARAGRAPHES
'----------------------------------------------------
Private Sub RemplacerTexteParIndices(ByVal doc As Document, ByVal indices As Collection, ByVal traduction As String)
    On Error GoTo GestionErreurRemplacement

    Dim traductions() As String
    traductions = Split(traduction, "[|||]")

    ' Le parcours en sens inverse est la clé pour éviter l'erreur 438.
    ' En modifiant les paragraphes de la fin vers le début, on ne décale pas
    ' les indices des paragraphes qui n'ont pas encore été traités.
    Dim i As Long
    For i = indices.Count To 1 Step -1

        If i - 1 <= UBound(traductions) Then
            Dim paraIndex As Long
            paraIndex = indices(i)

            ' Vérifie si l'indice est toujours valide dans le document.
            If paraIndex <= doc.Paragraphs.Count Then
                Dim para As Paragraph
                Set para = doc.Paragraphs(paraIndex)

                ' Supprime l'ancien texte tout en conservant le style.
                If para.Range.Characters.Count > 1 Then
                    para.Range.Characters(1, para.Range.Characters.Count - 1).Delete
                End If

                ' Nettoie le texte traduit et l'insère.
                para.Range.InsertBefore CleanTranslatedText(traductions(i - 1))
            End If
        End If
    Next i

    Exit Sub

GestionErreurRemplacement:
    ' Gère une erreur éventuelle lors du remplacement.
    MsgBox "Une erreur est survenue lors de la mise à jour du document." & vbCrLf & _
           "Erreur: " & Err.Description, vbCritical
End Sub

'----------------------------------------------------
' NETTOYAGE DU TEXTE TRADUIT
'----------------------------------------------------
Private Function CleanTranslatedText(ByVal text As String) As String
    ' Supprime les espaces de début et de fin.
    Dim cleanedText As String
    cleanedText = Trim(text)

    ' Supprime la marque de paragraphe finale si l'API en a ajouté une.
    ' C'est la cause probable de la création de paragraphes vides
    ' qui corrompait la structure du document.
    If Len(cleanedText) > 0 Then
        If Right(cleanedText, 1) = Chr(13) Then
            cleanedText = Left(cleanedText, Len(cleanedText) - 1)
        End If
    End If

    ' Retourne le texte nettoyé après un dernier trim.
    CleanTranslatedText = Trim(cleanedText)
End Function
