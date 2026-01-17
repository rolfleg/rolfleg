Cette procédure, Comporte des erreurs.
Rem Nous sommes dans le module 73.
Rem  En provenance du Module 76
' Copie du module 75 avec ajout de thermomètre.
Rem Module 75 avec Thermomètre
' Conservation des styles + Thermomètre de progression visuel

'====================================================
' VARIABLES GLOBALES POUR LE THERMOMÈTRE
'====================================================
Public AnnulerTraduction As Boolean

'====================================================
' TRADUCTION AVEC THERMOMÈTRE VISUEL
'====================================================
Sub TraduireDocument_OpenAI_Word_Rapide()
    Dim apiKey As String
    apiKey = Environ$("JCM5")
    If apiKey = "" Then
        MsgBox "Clé API absente (variable JCM5)", vbCritical
        Exit Sub
    End If

    Dim maxSegment As Long
    maxSegment = 15000

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

    Dim dossierTraduction As String
    dossierTraduction = doc.Path & "\Traduction"
    If Dir(dossierTraduction, vbDirectory) = "" Then
        MkDir dossierTraduction
    End If

    Dim totalCaracteres As Long
    totalCaracteres = Len(doc.Range.Text)
    Dim nbSegments As Long
    nbSegments = (totalCaracteres \ maxSegment) + 1

    Dim nomBase As String
    nomBase = Left(doc.Name, InStrRev(doc.Name, ".") - 1)
    If nomBase = "" Then nomBase = doc.Name

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
    message = message & "Segments : " & nbSegments & vbCrLf & vbCrLf
    message = message & "✓ Conservation des styles" & vbCrLf
    message = message & "✓ Thermomètre de progression" & vbCrLf & vbCrLf

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

    ' *** AFFICHER LE THERMOMÈTRE ***
    AnnulerTraduction = False
    Load frmProgressionTraduction
    frmProgressionTraduction.Show vbModeless

    Application.ScreenUpdating = False

    Dim startTime As Double
    startTime = Timer

    Dim docFusionne As Document
    Set docFusionne = Documents.Add
    docFusionne.ShowGrammaticalErrors = False
    docFusionne.ShowSpellingErrors = False

    Dim stylesParas As New Collection

    Dim i As Long
    Dim segmentTextes As String
    Dim segmentDebut As Long
    Dim traduction As String
    Dim para As Paragraph
    Dim numeroSegment As Long
    Dim textePara As String
    Dim caracteresTraites As Long

    segmentTextes = ""
    numeroSegment = 0
    segmentDebut = 1
    caracteresTraites = 0

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    For i = 1 To doc.Paragraphs.Count
        ' *** VÉRIFIER ANNULATION ***
        If AnnulerTraduction Then
            frmProgressionTraduction.Hide
            Unload frmProgressionTraduction
            Application.ScreenUpdating = True
            docFusionne.Close SaveChanges:=wdDoNotSaveChanges
            MsgBox "Traduction annulée par l'utilisateur.", vbInformation
            Exit Sub
        End If

        Set para = doc.Paragraphs(i)
        textePara = para.Range.Text

        ' *** CORRECTION 1 : Tester AVANT de capturer le style ***
        If Len(segmentTextes) + Len(textePara) > maxSegment And Len(segmentTextes) > 0 Then
            ' Le segment actuel est complet, on le traduit SANS le paragraphe actuel
            numeroSegment = numeroSegment + 1

            Dim tempsEcoule As Long
            tempsEcoule = CLng(Timer - startTime)
            Dim tempsRestant As Long
            If numeroSegment > 0 Then
                tempsRestant = CLng((tempsEcoule / numeroSegment) * (nbSegments - numeroSegment))
            End If

            ' *** MISE À JOUR DU THERMOMÈTRE ***
            Dim pourcentageGlobal As Long
            pourcentageGlobal = CLng((caracteresTraites / totalCaracteres) * 100)

            MettreAJourThermometre numeroSegment, nbSegments, pourcentageGlobal, _
                                  tempsEcoule, tempsRestant, "Traduction en cours..."

            traduction = TraduireSegment_Optimise(segmentTextes, apiKey, http)

            If Left(traduction, 8) = "(Erreur" Then
                frmProgressionTraduction.Hide
                Unload frmProgressionTraduction
                Application.ScreenUpdating = True
                MsgBox traduction, vbExclamation
                docFusionne.Close SaveChanges:=wdDoNotSaveChanges
                Exit Sub
            End If

            ' i-1 car le paragraphe i n'est PAS dans ce segment
            AjouterTexteAvecStyles docFusionne, traduction, stylesParas

            If sauvegarderSegments Then
                MettreAJourThermometre numeroSegment, nbSegments, pourcentageGlobal, _
                                      tempsEcoule, tempsRestant, "Sauvegarde segment..."
                SauvegarderSegmentDocx traduction, numeroSegment, dossierTraduction
            End If

            caracteresTraites = caracteresTraites + Len(segmentTextes)

            ' Le paragraphe actuel devient le DÉBUT du nouveau segment
            segmentTextes = textePara
            segmentDebut = i

            ' Vider la collection de styles et capturer le style du nouveau segment
            Set stylesParas = New Collection
            Dim styleInfo As String
            styleInfo = CaptureStyleParagraphe(para)
            stylesParas.Add styleInfo
        Else
            ' On ajoute le paragraphe au segment actuel
            segmentTextes = segmentTextes & textePara

            ' Capturer le style
            Dim styleInfo2 As String
            styleInfo2 = CaptureStyleParagraphe(para)
            stylesParas.Add styleInfo2
        End If
    Next i

    ' Dernier segment
    If Len(segmentTextes) > 0 Then
        numeroSegment = numeroSegment + 1
        MettreAJourThermometre numeroSegment, nbSegments, 95, _
                              CLng(Timer - startTime), 0, "Dernier segment..."

        traduction = TraduireSegment_Optimise(segmentTextes, apiKey, http)

        If Left(traduction, 8) <> "(Erreur" Then
            AjouterTexteAvecStyles docFusionne, traduction, stylesParas
            If sauvegarderSegments Then
                SauvegarderSegmentDocx traduction, numeroSegment, dossierTraduction
            End If
        End If
    End If

    If docFusionne.Paragraphs.Count > 0 Then
        If Len(Trim(docFusionne.Paragraphs(docFusionne.Paragraphs.Count).Range.Text)) <= 1 Then
            docFusionne.Paragraphs(docFusionne.Paragraphs.Count).Range.Delete
        End If
    End If

    MettreAJourThermometre nbSegments, nbSegments, 100, _
                          CLng(Timer - startTime), 0, "Sauvegarde finale..."

    SauvegarderDocumentFusionne docFusionne, nomBase, dossierTraduction

    ' *** FERMER LE THERMOMÈTRE ***
    frmProgressionTraduction.Hide
    Unload frmProgressionTraduction

    Application.ScreenUpdating = True

    Dim tempsTotal As Long
    tempsTotal = CLng(Timer - startTime)

    Dim messageFin As String
    messageFin = "✓ Traduction terminée !" & vbCrLf & vbCrLf
    messageFin = messageFin & "• Paragraphes : " & doc.Paragraphs.Count & vbCrLf
    messageFin = messageFin & "• Segments : " & numeroSegment & vbCrLf
    messageFin = messageFin & "• Styles conservés : OUI" & vbCrLf
    messageFin = messageFin & "• Temps total : " & FormatTemps(tempsTotal) & vbCrLf
    messageFin = messageFin & "• Temps/segment : " & FormatTemps(CLng(tempsTotal / numeroSegment)) & vbCrLf & vbCrLf
    messageFin = messageFin & "Fichier : Fr_" & nomBase & ".docx" & vbCrLf
    messageFin = messageFin & "Dossier : " & dossierTraduction

    MsgBox messageFin, vbInformation, "Traduction terminée"
    Exit Sub

GestionErreur:
    On Error Resume Next
    frmProgressionTraduction.Hide
    Unload frmProgressionTraduction
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'====================================================
' MISE À JOUR DU THERMOMÈTRE
'====================================================
Sub MettreAJourThermometre(segmentActuel As Long, segmentsTotal As Long, _
                          pourcentage As Long, tempsEcoule As Long, _
                          tempsRestant As Long, statut As String)
    On Error Resume Next

    With frmProgressionTraduction
        .lblSegment.Caption = "Segment " & segmentActuel & " / " & segmentsTotal
        .lblPourcentage.Caption = pourcentage & " %"
        .lblTempsEcoule.Caption = "Écoulé : " & FormatTemps(tempsEcoule)
        .lblTempsRestant.Caption = "Restant : " & FormatTemps(tempsRestant)
        .lblStatut.Caption = statut

        ' Mise à jour de la barre de progression
        If pourcentage >= 0 And pourcentage <= 100 Then
            .barProgression.Width = CLng((.Width - 40) * (pourcentage / 100))
        End If

        .Repaint
        DoEvents
    End With
End Sub

'====================================================
' CAPTURER LE STYLE D'UN PARAGRAPHE
'====================================================
Function CaptureStyleParagraphe(para As Paragraph) As String
    On Error Resume Next
    Dim style As String

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

    If parties(0) <> "" Then para.style = parties(0)
    If parties(1) <> "" Then para.Range.Font.Name = parties(1)
    If parties(2) <> "" And IsNumeric(parties(2)) Then para.Range.Font.Size = CSng(parties(2))

    If parties(3) = "-1" Then
        para.Range.Font.Bold = True
    ElseIf parties(3) = "0" Then
        para.Range.Font.Bold = False
    End If

    If parties(4) = "-1" Then
        para.Range.Font.Italic = True
    ElseIf parties(4) = "0" Then
        para.Range.Font.Italic = False
    End If

    If parties(5) <> "" And IsNumeric(parties(5)) Then para.Range.Font.Color = CLng(parties(5))
    If parties(6) <> "" And IsNumeric(parties(6)) Then para.Alignment = CInt(parties(6))
End Sub

'====================================================
' AJOUTER TEXTE AVEC STYLES CONSERVÉS
'====================================================
Sub AjouterTexteAvecStyles(docDest As Document, texte As String, stylesParas As Collection)
    Dim paragraphes() As String
    Dim i As Long
    Dim p As String
    Dim rng As Range
    Dim indexStyle As Long

    paragraphes = Split(texte, Chr(13))
    Set rng = docDest.Range
    rng.Collapse Direction:=wdCollapseEnd

    indexStyle = 1

    For i = LBound(paragraphes) To UBound(paragraphes)
        p = Replace(paragraphes(i), Chr(10), "")

        If Len(p) > 0 Or i < UBound(paragraphes) Then
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

                If indexStyle <= stylesParas.Count Then
                    Dim paraCree As Paragraph
                    Set paraCree = docDest.Paragraphs(docDest.Paragraphs.Count)
                    AppliquerStyleParagraphe paraCree, stylesParas(indexStyle)
                End If

                indexStyle = indexStyle + 1
                rng.Collapse Direction:=wdCollapseEnd
            End If
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

'====================================================
' TRADUCTION OPTIMISÉE
'====================================================
Function TraduireSegment_Optimise(segment As String, apiKey As String, http As Object) As String
    On Error GoTo ErrAPI

    If Len(Trim(segment)) = 0 Then
        TraduireSegment_Optimise = ""
        Exit Function
    End If

    Dim systemPrompt As String
    systemPrompt = "Translate to French. PRESERVE ALL paragraph marks and line breaks EXACTLY."

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

Sub SauvegarderSegmentDocx(texte As String, numero As Long, dossier As String)
    On Error Resume Next
    Dim docSegment As Document
    Set docSegment = Documents.Add
    docSegment.ShowGrammaticalErrors = False
    docSegment.ShowSpellingErrors = False

    Dim rng As Range
    Set rng = docSegment.Range
    rng.Text = texte

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
    ' *** CORRECTION 2 : Ne plus convertir \u000B en retour, car il ne devrait pas être envoyé ***
    ' txt = Replace(txt, "\u000B", Chr(11))

    ExtraireContent = txt
End Function

'====================================================
' *** CORRECTION 2 : Ne plus échapper Chr(11) ***
'====================================================
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
            ' *** CORRECTION 2 : Laisser Chr(11) tel quel, ne pas l'échapper ***
            ' Case Chr(11)
            '     result = result & "\u000B"
            Case Else
                If Asc(ch) < 32 And Asc(ch) <> 11 Then
                    result = result & "\u" & Right("0000" & Hex(Asc(ch)), 4)
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
        tempsParSegment = 60
    Else
        tempsParSegment = 30
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