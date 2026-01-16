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

    ' --- Architecture en 3 Phases : Lecture, Traduction, Écriture ---

    Dim startTime As Double
    startTime = Timer

    ' --- PHASE 1: LECTURE SEULE ---
    ' On parcourt le document une seule fois pour collecter les informations
    ' sur les paragraphes à traduire, sans faire aucune modification.
    Application.StatusBar = "Phase 1/3 : Analyse du document..."

    Dim parasATraduireInfo As New Collection
    Dim parasIgnorés As Long: parasIgnorés = 0
    Dim i As Long
    Dim paraIndex As Long ' Déclaration unique pour éviter l'erreur.

    For i = 1 To docTraduit.Paragraphs.Count
        Dim para As Paragraph
        Set para = docTraduit.Paragraphs(i)

        If para.Range.InlineShapes.Count > 0 Or para.Range.ShapeRange.Count > 0 Or Len(Trim(para.Range.Text)) <= 1 Then
            parasIgnorés = parasIgnorés + 1
        Else
            ' Stocke un Array contenant l'index et le texte original.
            parasATraduireInfo.Add Array(i, para.Range.Text)
        End If
    Next i

    If parasATraduireInfo.Count = 0 Then
        MsgBox "Aucun paragraphe textuel à traduire n'a été trouvé.", vbInformation, "Traduction terminée"
        docTraduit.Close SaveChanges:=False
        GoTo FinMacro
    End If

    ' --- PHASE 2: TRADUCTION (HORS DOCUMENT) ---
    ' On travaille uniquement sur les données en mémoire pour la traduction.
    Application.StatusBar = "Phase 2/3 : Traduction du texte..."

    Dim translationsDict As Object
    Set translationsDict = CreateObject("Scripting.Dictionary")

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    Dim segmentTexte As String
    Dim indicesParasSegment As New Collection

    For i = 1 To parasATraduireInfo.Count
        Dim paraInfo As Variant: paraInfo = parasATraduireInfo(i)
        paraIndex = paraInfo(0) ' Assignation (pas de déclaration).
        Dim paraTexte As String: paraTexte = paraInfo(1)

        indicesParasSegment.Add paraIndex
        segmentTexte = segmentTexte & paraTexte & "[|||]"

        ' Déclenche la traduction si le segment est plein ou si c'est la fin.
        If Len(segmentTexte) > maxSegment Or i = parasATraduireInfo.Count Then
            Dim pourcentage As Long
            pourcentage = CLng((i / parasATraduireInfo.Count) * 100)
            Application.StatusBar = "Phase 2/3 : Traduction en cours... " & pourcentage & "%"

            segmentTexte = Left(segmentTexte, Len(segmentTexte) - 5)

            Dim traduction As String
            traduction = TraduireSegment_Optimise(segmentTexte, apiKey, http)

            If Left(traduction, 8) = "(Erreur" Then
                MsgBox "La traduction a échoué. " & traduction, vbCritical, "Erreur de traduction"
                docTraduit.Close SaveChanges:=False
                GoTo FinMacro
            End If

            Dim traductionsSegment() As String
            traductionsSegment = Split(traduction, "[|||]")

            ' Stocke les traductions dans le dictionnaire.
            Dim j As Long
            For j = 0 To UBound(traductionsSegment)
                If j < indicesParasSegment.Count Then
                    translationsDict.Add key:=indicesParasSegment(j + 1), item:=CleanTranslatedText(traductionsSegment(j))
                End If
            Next j

            segmentTexte = ""
            Set indicesParasSegment = New Collection
        End If
    Next i

    ' --- PHASE 3: ÉCRITURE SEULE (EN ORDRE INVERSE) ---
    ' On modifie le document en une seule passe, de la fin vers le début,
    ' pour garantir une stabilité maximale.
    Application.StatusBar = "Phase 3/3 : Mise à jour du document..."

    Dim keys As Variant: keys = translationsDict.keys

    ' Tri simple des clés en ordre décroissant.
    Dim k As Long, temp As Variant
    For i = LBound(keys) To UBound(keys) - 1
        For k = i + 1 To UBound(keys)
            If keys(i) < keys(k) Then
                temp = keys(i): keys(i) = keys(k): keys(k) = temp
            End If
        Next k
    Next i

    For i = LBound(keys) To UBound(keys)
        paraIndex = keys(i)

        ' Vérifie si le paragraphe existe toujours.
        If paraIndex <= docTraduit.Paragraphs.Count Then
            Set para = docTraduit.Paragraphs(paraIndex)

            ' Définit une plage qui couvre tout le contenu du paragraphe SAUF la marque de fin.
            ' C'est la méthode la plus sûre pour remplacer du texte.
            Dim contentRange As Range
            Set contentRange = docTraduit.Range(Start:=para.Range.Start, End:=para.Range.End - 1)

            ' Remplace directement le texte de cette plage.
            contentRange.Text = translationsDict(paraIndex)
        End If
    Next i

    ' --- Finalisation ---
    Dim tempsTotal As Long
    tempsTotal = CLng(Timer - startTime)
    Dim parasTraduits As Long: parasTraduits = translationsDict.Count

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
' NETTOYAGE DU TEXTE TRADUIT (VERSION ROBUSTE)
'----------------------------------------------------
Private Function CleanTranslatedText(ByVal text As String) As String
    ' Remplace TOUS les sauts de paragraphe (Chr(13)) par un espace.
    ' C'est la correction la plus critique : l'API peut insérer des sauts
    ' de paragraphe au milieu du texte, ce qui scinde un paragraphe en deux,
    ' corrompt la structure du document et provoque l'erreur 438.
    ' Cette ligne garantit que la structure des paragraphes reste stable.
    Dim cleanedText As String
    cleanedText = Replace(text, Chr(13), " ")

    ' Supprime les espaces superflus au début et à la fin.
    CleanTranslatedText = Trim(cleanedText)
End Function
