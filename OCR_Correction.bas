Option Explicit

' =====================================================
' MACRO PRINCIPALE
' =====================================================
Sub ExtraireTexteDesImagesVersNouveauDoc()
    ' Vérification du type de document (Memory)
    If Not ThisComponent.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox "Cette macro ne fonctionne que dans un document texte LibreOffice Writer.", 48, "Erreur"
        Exit Sub
    End If

    Dim reponse As Integer
    reponse = MsgBox("Voulez-vous continuer ?", vbQuestion + vbYesNo, "Confirmation")
    If reponse = 7 Then Exit Sub ' 7 = vbNo

    Dim oDocOrigine     As Object
    Dim oDocCible       As Object
    Dim oShape          As Object
    Dim oGraphic        As Object
    Dim oGraphicObjects As Object

    Dim i               As Long
    Dim imgCount        As Long
    Dim totalImages     As Long

    Dim api_key         As String
    Dim sBase64         As String
    Dim sTexteExtrait   As String
    Dim sTexteNet       As String

    Dim oTextCible      As Object
    Dim oCursorCible    As Object

    Dim sUrlOrigine     As String
    Dim sDossier        As String
    Dim sNomFichier     As String
    Dim sUrlCible       As String
    Dim sUrlTemp        As String
    Dim nWidth          As Long
    Dim nHeight         As Long

    ' =========================
    ' CLÉ API
    ' =========================
    api_key = Environ("jcm9")

    If api_key = "" Then
        MsgBox "Clé API introuvable dans les variables d'environnement (jcm9).", 48, "Erreur clé API"
        Exit Sub
    End If

    oDocOrigine = ThisComponent

    If Not oDocOrigine.haslocation() Then
        MsgBox "Enregistrez d'abord le document.", 48, "Erreur"
        Exit Sub
    End If

    sUrlOrigine = oDocOrigine.getlocation()

    ' Verrouiller les contrôleurs pour la performance (Memory)
    oDocOrigine.lockControllers()

    On Error GoTo FinErreur

    ' ══════════════════════════════════════════════════════════
    ' Collecte de toutes les images
    ' ══════════════════════════════════════════════════════════
    oGraphicObjects = oDocOrigine.graphicobjects
    totalImages = oGraphicObjects.count

    If totalImages = 0 Then
        MsgBox "Aucune image extractible trouvée dans le document.", 48, "Information"
        GoTo Fin
    End If

    ' Nouveau document cible
    oDocCible = StarDesktop.loadcomponentfromurl( _
        "private:factory/swriter", "_blank", 0, Array())

    oTextCible   = oDocCible.text
    oCursorCible = oTextCible.createtextcursor()
    oTextCible.insertstring(oCursorCible, "--- RÉSULTATS OCR ---" & Chr(13) & Chr(13), False)

    imgCount = 0

    ' ══════════════════════════════════════════════════════════
    ' Traitement de chaque image
    ' ══════════════════════════════════════════════════════════
    For i = 0 To totalImages - 1
        oShape = oGraphicObjects.getbyindex(i)
        imgCount = imgCount + 1

        oTextCible.insertstring(oCursorCible, "--- Image " & imgCount & " ---" & Chr(13), False)

        ' Dimensions
        nWidth  = 0
        nHeight = 0
        On Error Resume Next
        nWidth  = oShape.width
        nHeight = oShape.height
        On Error GoTo FinErreur

        If nWidth  = 0 Then nWidth  = 5000
        If nHeight = 0 Then nHeight = 5000

        ' Récupération de l'objet Graphic
        ' Retrait du 'f' parasite
        oGraphic = Null
        On Error Resume Next
        oGraphic = oShape.graphic
        On Error GoTo FinErreur

        If IsNull(oGraphic) Or IsEmpty(oGraphic) Then
            oTextCible.insertstring(oCursorCible, _
                "[Erreur : objet Graphic inaccessible]" & Chr(13) & Chr(13), False)
        Else
            ' Export PNG temporaire
            sUrlTemp = ExporterGraphicTemp(oGraphic, imgCount)

            If sUrlTemp <> "" Then
                ' Insertion visuelle dans le doc cible
                InsererImageDepuisURL oDocCible, oTextCible, oCursorCible, sUrlTemp, nWidth, nHeight

                ' Conversion Base64 (Correction : Utilisation de SimpleFileAccess)
                sBase64 = conversionFichierEnBase64(sUrlTemp)

                ' Nettoyage fichier temporaire
                Dim oSFA As Object
                oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
                If oSFA.exists(sUrlTemp) Then oSFA.kill(sUrlTemp)

                If sBase64 <> "" Then
                    sTexteExtrait = AppelerOpenAIVision(sBase64, api_key)
                Else
                    sTexteExtrait = "[Erreur de conversion Base64]"
                End If
            Else
                oTextCible.insertstring(oCursorCible, "[Erreur export image]" & Chr(13), False)
                sTexteExtrait = ""
            End If

            sTexteNet = Trim(sTexteExtrait)
            If EstTexteSignificatif(sTexteNet) Then
                oTextCible.insertstring(oCursorCible, sTexteNet & Chr(13) & Chr(13), False)
            Else
                oTextCible.insertstring(oCursorCible, Chr(13), False)
            End If
        End If

    Next i

    ' Sauvegarde du document résultat
    sDossier    = ExtraireDossier(sUrlOrigine)
    sNomFichier = ExtraireNomFichier(sUrlOrigine)
    sUrlCible   = sDossier & "OCR_" & sNomFichier

    oDocCible.storeasurl(sUrlCible, Array())

    MsgBox imgCount & " image(s) traitée(s)." & Chr(13) & ConvertFromURL(sUrlCible), 64, "OCR terminé"

Fin:
    oDocOrigine.unlockControllers()
    Exit Sub

FinErreur:
    MsgBox "Une erreur est survenue : " & Err.Description, 16, "Erreur"
    Resume Fin
End Sub


' =====================================================
' EXPORTER UN GRAPHIC EN FICHIER PNG TEMPORAIRE
' =====================================================
Function ExporterGraphicTemp(oGraphic As Object, id As Long) As String
    On Error GoTo GestionErreur

    Dim oProvider     As Object
    Dim oPathSettings As Object
    Dim sTempDir      As String
    Dim sUrlTemp      As String
    Dim args(1) As New com.sun.star.beans.PropertyValue

    oPathSettings = CreateUnoService("com.sun.star.util.PathSettings")
    sTempDir = oPathSettings.temp
    ' S'assurer que le dossier temporaire se termine par /
    If Right(sTempDir, 1) <> "/" Then sTempDir = sTempDir & "/"

    sUrlTemp = sTempDir & "ocr_temp_" & id & ".png"

    oProvider = CreateUnoService("com.sun.star.graphic.GraphicProvider")
    args(0).Name  = "URL"
    args(0).Value = sUrlTemp
    args(1).Name  = "MimeType"
    args(1).Value = "image/png"

    oProvider.storegraphic(oGraphic, args())

    Dim oSFA As Object
    oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If oSFA.exists(sUrlTemp) Then
        ExporterGraphicTemp = sUrlTemp
    Else
        ExporterGraphicTemp = ""
    End If
    Exit Function

GestionErreur:
    ExporterGraphicTemp = ""
End Function


' =====================================================
' INSÉRER UNE IMAGE DE MANIÈRE INCORPORÉE
' =====================================================
Sub InsererImageDepuisURL(oDoc As Object, oText As Object, oCursor As Object, _
                           sUrl As String, nWidth As Long, nHeight As Long)
    On Error GoTo GestionErreur

    Dim oImgShape As Object
    Dim oProvider As Object
    Dim args(0)   As New com.sun.star.beans.PropertyValue

    oImgShape = oDoc.createinstance("com.sun.star.text.TextGraphicObject")
    oImgShape.anchortype = com.sun.star.text.TextContentAnchorType.AS_CHARACTER
    oImgShape.width  = nWidth
    oImgShape.height = nHeight

    oProvider = CreateUnoService("com.sun.star.graphic.GraphicProvider")
    args(0).Name  = "URL"
    args(0).Value = sUrl
    oImgShape.graphic = oProvider.querygraphic(args())

    oText.inserttextcontent(oCursor, oImgShape, False)
    oText.insertstring(oCursor, Chr(13), False)
    Exit Sub

GestionErreur:
    oText.insertstring(oCursor, "[Erreur insertion visuelle]" & Chr(13), False)
End Sub


' =====================================================
' CONVERSION BASE64 VIA CERTUTIL (Optimisé LibreOffice)
' =====================================================
Function conversionFichierEnBase64(sUrlImage As String) As String
    On Error GoTo GestionErreur

    Dim oShell        As Object
    Dim sImageFile    As String
    Dim sTextFile     As String
    Dim sUrlTextFile  As String
    Dim sBase64       As String
    Dim sCmd          As String
    Dim oSFA          As Object
    Dim oInStream     As Object
    Dim oTextStream   As Object

    sImageFile = ConvertFromURL(sUrlImage)
    sTextFile  = Left(sImageFile, InStrRev(sImageFile, ".")) & "txt"
    sUrlTextFile = ConvertToURL(sTextFile)

    oShell = CreateObject("WScript.Shell")
    sCmd = "cmd /c certutil -encode """ & sImageFile & """ """ & sTextFile & """"
    oShell.run(sCmd, 0, True)

    oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If oSFA.exists(sUrlTextFile) Then
        oInStream = oSFA.openfileread(sUrlTextFile)
        oTextStream = CreateUnoService("com.sun.star.io.TextInputStream")
        oTextStream.setinputstream(oInStream)
        oTextStream.setencoding("UTF-8")

        sBase64 = ""
        Do While Not oTextStream.iseof()
            sBase64 = sBase64 & oTextStream.readline()
        Loop
        oInStream.closeinput()
        oSFA.kill(sUrlTextFile)

        sBase64 = Replace(sBase64, "-----BEGIN CERTIFICATE-----", "")
        sBase64 = Replace(sBase64, "-----END CERTIFICATE-----", "")
        sBase64 = Replace(sBase64, Chr(13), "")
        sBase64 = Replace(sBase64, Chr(10), "")
        sBase64 = Replace(sBase64, " ", "")

        conversionFichierEnBase64 = Trim(sBase64)
    Else
        conversionFichierEnBase64 = ""
    End If
    Exit Function

GestionErreur:
    conversionFichierEnBase64 = ""
End Function


' =====================================================
' DÉTECTION D'UN TEXTE OCR SIGNIFICATIF
' =====================================================
Function EstTexteSignificatif(sTexte As String) As Boolean
    Dim sLower As String
    sLower = LCase(Trim(sTexte))

    If sLower = ""                                        Then EstTexteSignificatif = False : Exit Function
    If Left(sLower, 7)  = "[aucun "                       Then EstTexteSignificatif = False : Exit Function
    If Left(sLower, 8)  = "[erreur "                      Then EstTexteSignificatif = False : Exit Function
    If InStr(sLower, "no text")                   > 0     Then EstTexteSignificatif = False : Exit Function
    If InStr(sLower, "no visible text")           > 0     Then EstTexteSignificatif = False : Exit Function
    If InStr(sLower, "does not contain any text") > 0     Then EstTexteSignificatif = False : Exit Function
    If InStr(sLower, "there is no text")          > 0     Then EstTexteSignificatif = False : Exit Function
    If InStr(sLower, "image contains no text")    > 0     Then EstTexteSignificatif = False : Exit Function
    If InStr(sLower, "no readable text")          > 0     Then EstTexteSignificatif = False : Exit Function
    If InStr(sLower, "aucun texte")               > 0     Then EstTexteSignificatif = False : Exit Function
    If InStr(sLower, "pas de texte")              > 0     Then EstTexteSignificatif = False : Exit Function
    If Len(Trim(sTexte)) < 3                              Then EstTexteSignificatif = False : Exit Function

    EstTexteSignificatif = True
End Function


' =====================================================
' APPEL API OPENAI VISION
' =====================================================
Function AppelerOpenAIVision(sBase64 As String, apiKey As String) As String
    On Error GoTo GestionErreur

    Dim oHttp    As Object
    Dim url      As String
    Dim payload  As String
    Dim response As String

    url = "https://api.openai.com/v1/chat/completions"

    oHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    oHttp.open("POST", url, False)
    oHttp.settimeouts(5000, 10000, 60000, 60000)
    oHttp.setrequestheader("Content-Type",  "application/json")
    oHttp.setrequestheader("Authorization", "Bearer " & apiKey)

    ' Correction : gpt-4o-mini (Memory)
    payload = "{""model"":""gpt-4o-mini""," & _
              """messages"":[{" & _
              """role"":""user""," & _
              """content"":[" & _
              "{""type"":""text"",""text"":""Please perform OCR on this image and return all the text you can read in it. Return only the extracted text, nothing else. If there is no text, reply exactly: no text""}," & _
              "{""type"":""image_url"",""image_url"":{""url"":""data:image/png;base64," & sBase64 & """}}" & _
              "]}]," & _
              """max_tokens"":2000}"

    oHttp.send(payload)
    response = oHttp.responsetext

    If InStr(response, """error""") > 0 Then
        AppelerOpenAIVision = "[Erreur API] " & response
        Exit Function
    End If

    AppelerOpenAIVision = ExtraireContent(response)
    Exit Function

GestionErreur:
    AppelerOpenAIVision = "[Erreur réseau ou API OpenAI : " & Err.Description & "]"
End Function


' =====================================================
' EXTRACTION DU TEXTE DANS LE JSON
' =====================================================
Function ExtraireContent(sJson As String) As String
    Dim p1     As Long
    Dim pStart As Long
    Dim p2     As Long

    p1 = InStr(sJson, """content""")
    If p1 = 0 Then
        ExtraireContent = "[Aucun texte détecté]"
        Exit Function
    End If

    pStart = InStr(p1 + 9, sJson, """")
    If pStart = 0 Then
        ExtraireContent = "[Aucun texte détecté]"
        Exit Function
    End If

    pStart = pStart + 1
    p2 = pStart

    Do While p2 <= Len(sJson)
        If Mid(sJson, p2, 1) = """" Then
            If Mid(sJson, p2 - 1, 1) <> "\" Then Exit Do
        End If
        p2 = p2 + 1
    Loop

    ExtraireContent = Mid(sJson, pStart, p2 - pStart)
    ExtraireContent = Replace(ExtraireContent, "\n",  Chr(13))
    ExtraireContent = Replace(ExtraireContent, "\r",  "")
    ExtraireContent = Replace(ExtraireContent, "\""", """")
    ExtraireContent = Replace(ExtraireContent, "\\",  "\")

    ' Nettoyage : suppression des sauts de paragraphe finaux (Memory)
    Do While Right(ExtraireContent, 1) = Chr(13)
        ExtraireContent = Left(ExtraireContent, Len(ExtraireContent) - 1)
    Loop
End Function


' =====================================================
' OUTILS UTILITAIRES
' =====================================================
Function ExtraireDossier(sUrl As String) As String
    ExtraireDossier = Left(sUrl, InStrRev(sUrl, "/"))
End Function

Function ExtraireNomFichier(sUrl As String) As String
    ExtraireNomFichier = Mid(sUrl, InStrRev(sUrl, "/") + 1)
End Function
