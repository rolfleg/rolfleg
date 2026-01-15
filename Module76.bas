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
    nomBase = Left(docSource.Name, InStrRev(docSource.Name, ".") - 1)
    If nomBase = "" Then nomBase = docSource.Name

    Dim nouveauNom As String
    nouveauNom = docSource.Path & "\Fr_" & nomBase & ".docx"

    Dim docTraduit As Document
    ' Copie le document source et travaille sur la copie.
    Set docTraduit = docSource.SaveAs2(FileName:=nouveauNom, AddToRecentFiles:=False)

    ' --- Identification des paragraphes à traduire ---
    Application.StatusBar = "Analyse du document pour identifier les paragraphes textuels..."

    Dim parasATraduire As New Collection
    Dim para As Paragraph
    Dim parasIgnorés As Long
    Dim i As Long
    parasIgnorés = 0

    ' Boucle sur tous les paragraphes du document à traduire.
    For i = 1 To docTraduit.Paragraphs.Count
        Set para = docTraduit.Paragraphs(i)

        ' On vérifie si le paragraphe contient des formes flottantes (Shapes)
        ' ou des formes ancrées dans le texte (InlineShapes).
        If para.Range.InlineShapes.Count > 0 Or para.Range.ShapeRange.Count > 0 Then
            parasIgnorés = parasIgnorés + 1
        ' On s'assure aussi que le paragraphe n'est pas vide ou composé uniquement d'espaces.
        ' La longueur > 1 est pour ignorer les paragraphes qui n'ont que la marque de fin (Chr(13)).
        ElseIf Len(Trim(para.Range.Text)) > 1 Then
            parasATraduire.Add para
        Else
            ' Les paragraphes vides sont également ignorés.
            parasIgnorés = parasIgnorés + 1
        End If

        ' Met à jour la barre de statut pour donner un retour visuel à l'utilisateur.
        If i Mod 10 = 0 Then
            Application.StatusBar = "Analyse du document... Paragraphe " & i & " sur " & docTraduit.Paragraphs.Count
        End If
    Next i

    Application.StatusBar = False

    ' --- Fin de l'identification ---

    ' --- Traduction des paragraphes identifiés ---
    If parasATraduire.Count = 0 Then
        MsgBox "Aucun paragraphe textuel à traduire n'a été trouvé. Le document est peut-être composé uniquement d'images ou de paragraphes vides.", vbInformation, "Traduction terminée"
        docTraduit.Close SaveChanges:=False ' Ferme sans sauvegarder si rien n'a été fait
        Exit Sub
    End If

    Dim startTime As Double
    startTime = Timer

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    Dim segmentTexte As String
    Dim parasDuSegment As New Collection
    Dim traduction As String
    Dim parasTraduits As Long
    parasTraduits = 0

    ' Boucle sur les paragraphes à traduire pour créer des segments
    For i = 1 To parasATraduire.Count
        Set para = parasATraduire(i)

        ' Ajoute le paragraphe et son texte au segment en cours
        parasDuSegment.Add para
        ' Utilise un séparateur unique pour que l'IA le préserve
        segmentTexte = segmentTexte & para.Range.Text & "[|||]"

        ' Vérifie si le segment atteint la taille maximale ou si c'est le dernier paragraphe
        If Len(segmentTexte) > maxSegment Or i = parasATraduire.Count Then

            ' Met à jour la barre de statut
            Dim pourcentage As Long
            pourcentage = CLng((i / parasATraduire.Count) * 100)
            Application.StatusBar = "Traduction en cours... " & pourcentage & "%"

            ' Enlève le dernier séparateur avant d'envoyer
            segmentTexte = Left(segmentTexte, Len(segmentTexte) - 5)

            ' Appel à l'API pour la traduction
            traduction = TraduireSegment_Optimise(segmentTexte, apiKey, http)

            If Left(traduction, 8) = "(Erreur" Then
                MsgBox "La traduction a échoué. " & traduction, vbCritical, "Erreur de traduction"
                docTraduit.Close SaveChanges:=False
                GoTo FinMacro
            End If

            ' Remplace le texte original par le texte traduit
            RemplacerTexteParagraphes parasDuSegment, traduction
            parasTraduits = parasTraduits + parasDuSegment.Count

            ' Réinitialise les variables pour le prochain segment
            segmentTexte = ""
            Set parasDuSegment = New Collection
        End If
    Next i

    ' --- Finalisation ---
    Dim tempsTotal As Long
    tempsTotal = CLng(Timer - startTime)

    docTraduit.Save
    docTraduit.Close

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Dim messageFin As String
    messageFin = "Traduction terminée avec succès !" & vbCrLf & vbCrLf & _
                 "• Paragraphes traduits : " & parasTraduits & vbCrLf & _
                 "• Paragraphes ignorés : " & parasIgnorés & vbCrLf & _
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
' REMPLACEMENT DU TEXTE DANS LES PARAGRAPHES CIBLES
'----------------------------------------------------
Private Sub RemplacerTexteParagraphes(ByVal paras As Collection, ByVal traduction As String)
    ' Sépare la traduction en utilisant le même délimiteur qu'à l'envoi.
    Dim traductions() As String
    traductions = Split(traduction, "[|||]")

    Dim i As Long
    ' Boucle en sens inverse pour éviter l'erreur 438.
    ' En modifiant le document de la fin vers le début, les indices des paragraphes
    ' non encore traités ne sont pas affectés par les modifications.
    For i = paras.Count To 1 Step -1
        If i - 1 <= UBound(traductions) Then
            Dim para As Paragraph
            Set para = paras(i)

            ' S'assure que le paragraphe est toujours valide avant de le modifier.
            If Not para Is Nothing Then
                If para.Range.Characters.Count > 1 Then
                    ' Supprime le contenu existant tout en préservant la marque de paragraphe et son style.
                    para.Range.Characters(1, para.Range.Characters.Count - 1).Delete
                End If

                ' Insère le texte traduit au début du paragraphe vide.
                para.Range.InsertBefore Trim(traductions(i - 1))
            End If
        End If
    Next i
End Sub
