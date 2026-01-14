Sub TraduireDocument_OpenAI_Word_Rapide()
    Dim apiKey As String
    apiKey = Environ$("JCM5")
    If apiKey = "" Then
        MsgBox "Clé API absente (variable JCM5)", vbCritical
        Exit Sub
    End If

    ' *** OPTIMISATION 1 : SEGMENTS PLUS PETITS = RÉPONSES PLUS RAPIDES ***
    Dim maxSegment As Long
    maxSegment = 15000  ' Réduit de 30000 à 15000 (réponses 2x plus rapides)

    On Error GoTo GestionErreur

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

    ' Détecter les images et les dessins
    Dim nbImages As Long
    nbImages = doc.InlineShapes.Count + doc.Shapes.Count

    If nbImages > 0 Then
        Dim msg As String
        msg = "Ce document contient " & nbImages & " image(s) ou dessin(s)." & vbCrLf & vbCrLf
        msg = msg & "La macro ne traduira que le texte." & vbCrLf & vbCrLf
        msg = msg & "Voulez-vous continuer ?"

        If MsgBox(msg, vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    ' Créer le dossier Traduction
    Dim dossierTraduction As String
    dossierTraduction = doc.Path & "\Traduction"
    If Dir(dossierTraduction, vbDirectory) = "" Then
        MkDir dossierTraduction
    End If

    Dim totalCaracteres As Long
    totalCaracteres = Len(doc.Range.text)
    Dim nbSegments As Long
    nbSegments = (totalCaracteres \ maxSegment) + 1

    Dim nomBase As String
    nomBase = Left(doc.Name, InStrRev(doc.Name, ".") - 1)
    If nomBase = "" Then nomBase = doc.Name

    ' *** OPTIMISATION 2 : OPTION POUR DÉSACTIVER LA SAUVEGARDE DES SEGMENTS ***
    Dim sauvegarderSegments As Boolean
    Dim reponse As Integer

    reponse = MsgBox("Voulez-vous sauvegarder les segments individuels ?" & vbCrLf & vbCrLf & _
                     "NON = Plus rapide (recommandé)" & vbCrLf & _
                     "OUI = Sauvegarde de " & nbSegments & " fichiers (plus lent)", _
                     vbYesNoCancel + vbQuestion, "Option de sauvegarde")

    If reponse = vbCancel Then Exit Sub
    sauvegarderSegments = (reponse = vbYes)

    Dim message As String
    message = "Document : " & totalCaracteres & " caractères" & vbCrLf
    message = message & "Paragraphes : " & doc.Paragraphs.Count & vbCrLf
    message = message & "Segments : " & nbSegments & vbCrLf
    message = message & "Taille segment : " & maxSegment & " car." & vbCrLf & vbCrLf

    If sauvegarderSegments Then
        message = message & "Mode : Sauvegarde complète (plus lent)" & vbCrLf
    Else
        message = message & "Mode : RAPIDE (pas de sauvegarde segments)" & vbCrLf
    End If

    message = message & vbCrLf & "Temps estimé : " & EstimerTemps(nbSegments, sauvegarderSegments) & vbCrLf & vbCrLf
    message = message & "Continuer ?"

    If MsgBox(message, vbYesNo + vbInformation, "Confirmation") = vbNo Then
        Exit Sub
    End If

    ' *** OPTIMISATION 3 : DÉSACTIVER MISE À JOUR ÉCRAN ET CALCULS ***
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True

    Dim startTime As Double
    startTime = Timer

    ' *** CRÉER LE DOCUMENT FUSIONNÉ ***
    Dim docFusionne As Document
    Set docFusionne = Documents.Add

    ' *** OPTIMISATION 4 : DÉSACTIVER LES VÉRIFICATIONS ORTHOGRAPHIQUES ***
    docFusionne.ShowGrammaticalErrors = False
    docFusionne.ShowSpellingErrors = False

    Dim i As Long
    Dim segmentTextes As String
    Dim traduction As String
    Dim para As Paragraph
    Dim numeroSegment As Long
    Dim pourcentage As Long
    Dim textePara As String

    segmentTextes = ""
    numeroSegment = 0

    ' *** OPTIMISATION 5 : RÉUTILISER LE MÊME OBJET HTTP ***
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        textePara = para.Range.text

        pourcentage = CLng((i / doc.Paragraphs.Count) * 100)

        If Len(segmentTextes) + Len(textePara) > maxSegment And Len(segmentTextes) > 0 Then
            numeroSegment = numeroSegment + 1

            ' Calculer temps écoulé et temps restant
            Dim tempsEcoule As Long
            tempsEcoule = CLng(Timer - startTime)
            Dim tempsRestant As Long
            If numeroSegment > 0 Then
                tempsRestant = CLng((tempsEcoule / numeroSegment) * (nbSegments - numeroSegment))
            End If

            Application.StatusBar = "Segment " & numeroSegment & "/" & nbSegments & _
                                  " (" & pourcentage & "%) | Ecoulé: " & FormatTemps(tempsEcoule) & _
                                  " | Restant: " & FormatTemps(tempsRestant)

            ' *** APPEL API OPTIMISÉ ***
            traduction = TraduireSegment_Optimise(segmentTextes, apiKey, http)

            If Left(traduction, 8) = "(Erreur" Then
                Application.StatusBar = False
                Application.ScreenUpdating = True
                MsgBox traduction, vbExclamation
                docFusionne.Close SaveChanges:=wdDoNotSaveChanges
                Exit Sub
            End If

            ' *** OPTIMISATION 6 : INSERTION DIRECTE SANS MANIPULATION EXCESSIVE ***
            AjouterTexteRapide docFusionne, traduction

            ' Sauvegarde conditionnelle
            If sauvegarderSegments Then
                SauvegarderSegmentDocx traduction, numeroSegment, dossierTraduction
            End If

            segmentTextes = textePara
        Else
            segmentTextes = segmentTextes & textePara
        End If
    Next i

    ' Dernier segment
    If Len(segmentTextes) > 0 Then
        numeroSegment = numeroSegment + 1
        Application.StatusBar = "Traduction du dernier segment..."
        traduction = TraduireSegment_Optimise(segmentTextes, apiKey, http)

        If Left(traduction, 8) <> "(Erreur" Then
            AjouterTexteRapide docFusionne, traduction
            If sauvegarderSegments Then
                SauvegarderSegmentDocx traduction, numeroSegment, dossierTraduction
            End If
        End If
    End If

    ' Nettoyer paragraphe final
    If docFusionne.Paragraphs.Count > 0 Then
        If Len(Trim(docFusionne.Paragraphs(docFusionne.Paragraphs.Count).Range.text)) <= 1 Then
            docFusionne.Paragraphs(docFusionne.Paragraphs.Count).Range.Delete
        End If
    End If

    ' Sauvegarder le document fusionné
    Application.StatusBar = "Sauvegarde finale..."
    SauvegarderDocumentFusionne docFusionne, nomBase, dossierTraduction

    Application.StatusBar = False
    Application.ScreenUpdating = True

    Dim tempsTotal As Long
    tempsTotal = CLng(Timer - startTime)

    Dim messageFin As String
    messageFin = "Traduction terminée !" & vbCrLf & vbCrLf
    messageFin = messageFin & "• Paragraphes : " & doc.Paragraphs.Count & vbCrLf
    messageFin = messageFin & "• Segments : " & numeroSegment & vbCrLf
    messageFin = messageFin & "• Temps total : " & FormatTemps(tempsTotal) & vbCrLf
    messageFin = messageFin & "• Temps/segment : " & FormatTemps(CLng(tempsTotal / numeroSegment)) & vbCrLf & vbCrLf
    messageFin = messageFin & "Fichier : Fr_" & nomBase & ".docx" & vbCrLf
    messageFin = messageFin & "Dossier : " & dossierTraduction

    MsgBox messageFin, vbInformation, "Traduction terminée"
    Exit Sub

GestionErreur:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'====================================================
' TRADUCTION OPTIMISÉE AVEC HTTP RÉUTILISÉ
'====================================================
Function TraduireSegment_Optimise(segment As String, apiKey As String, http As Object) As String
    On Error GoTo ErrAPI

    If Len(Trim(segment)) = 0 Then
        TraduireSegment_Optimise = ""
        Exit Function
    End If

    ' *** OPTIMISATION 7 : PROMPT PLUS COURT ***
    Dim systemPrompt As String
    systemPrompt = "Translate to French. PRESERVE ALL paragraph marks and line breaks EXACTLY."

    Dim json As String
    json = "{""model"":""gpt-4o""," & _
           """messages"":[" & _
           "{""role"":""system"",""content"":""" & EscapeJSON(systemPrompt) & """}," & _
           "{""role"":""user"",""content"":""" & EscapeJSON(segment) & """}" & _
           "],""temperature"":0.3}"

    ' *** RÉUTILISER L'OBJET HTTP (PAS DE CRÉATION/DESTRUCTION) ***
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
' AJOUT RAPIDE SANS MANIPULATIONS EXCESSIVES
'====================================================
Sub AjouterTexteRapide(docDest As Document, texte As String)
    Dim paragraphes() As String
    Dim i As Long
    Dim p As String
    Dim rng As Range

    paragraphes = Split(texte, Chr(13))
    Set rng = docDest.Range
    rng.Collapse Direction:=wdCollapseEnd

    For i = LBound(paragraphes) To UBound(paragraphes)
        p = Replace(paragraphes(i), Chr(10), "")

        If Len(p) > 0 Then
            If InStr(p, Chr(11)) > 0 Then
                InsererAvecSautsLigne rng, p
            Else
                rng.InsertAfter p
                rng.Collapse Direction:=wdCollapseEnd
            End If
        End If

        If i < UBound(paragraphes) Or Len(p) > 0 Then
            rng.InsertParagraphAfter
            rng.Collapse Direction:=wdCollapseEnd
        End If
    Next i
End Sub

Sub InsererAvecSautsLigne(rng As Range, texte As String)
    Dim parties() As String
    Dim i As Long
    parties = Split(texte, Chr(11))

    For i = LBound(parties) To UBound(parties)
        If Len(parties(i)) > 0 Then
            rng.InsertAfter parties(i)
            rng.Collapse Direction:=wdCollapseEnd
        End If
        If i < UBound(parties) Then
            rng.InsertBreak Type:=wdLineBreak
            rng.Collapse Direction:=wdCollapseEnd
        End If
    Next i
End Sub

Sub SauvegarderSegmentDocx(texte As String, numero As Long, dossier As String)
    On Error Resume Next
    Dim docSegment As Document
    Set docSegment = Documents.Add
    docSegment.ShowGrammaticalErrors = False
    docSegment.ShowSpellingErrors = False

    AjouterTexteRapide docSegment, texte

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

Function ExtraireContent(resp As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(resp, """content"":""")
    If p1 = 0 Then p1 = InStr(resp, """content"": """)
    If p1 = 0 Then
        ExtraireContent = "(Erreur: content introuvable)"
        Exit Function
    End If

    ' Utiliser If au lieu de IIf pour compatibilité
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
            Case Chr(13)
                result = result & "\n"
            Case Chr(10)
                ' Ignorer
            Case vbTab
                result = result & "\t"
            Case Chr(11)
                result = result & "\u000B"
            Case Else
                If Asc(ch) < 32 Then
                    result = result & "\u" & Right("0000" & hex(Asc(ch)), 4)
                Else
                    result = result & ch
                End If
        End Select
    Next i
    EscapeJSON = result
End Function

'====================================================
' UTILITAIRES
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
    If avecSauvegarde Then
        tempsParSegment = 60  ' 1 min par segment avec sauvegarde (optimisé)
    Else
        tempsParSegment = 30  ' 30 sec par segment sans sauvegarde
    End If

    Dim tempsTotal As Long
    tempsTotal = nbSegments * tempsParSegment
    EstimerTemps = FormatTemps(tempsTotal)
End Function

Function PadLeft(texte As String, longueur As Integer, caractere As String) As String
    Dim result As String
    result = texte
    Do While Len(result) < longueur
        result = caractere & result
    Loop
    PadLeft = result
End Function
