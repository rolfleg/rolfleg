Sub Supprimer_paragraphes_en_double3()
    SupprimerParagraphesDoublonsDoublonsComplet3
End Sub

Sub SupprimerParagraphesDoublonsDoublonsComplet3()
    Dim reponse As Integer
    reponse = ConfirmDialog()

    If reponse <> 6 Then Exit Sub

    Dim oDoc As Object, oNewDoc As Object
    Dim oText As Object, oNewText As Object
    Dim oEnum As Object, oPara As Object
    Dim paras() As Object
    Dim texts() As String
    Dim n As Long, i As Long, j As Long
    Dim countSuppr As Long
    Dim oSI As Object
    Dim sText As String, sTextClean As String
    Dim sDocURL As String, sFileName As String, sExt As String
    Dim sBackupURL As String, sDelDocURL As String
    Dim k As Long
    Dim lastDotPos As Long
    Dim dummy() As New com.sun.star.beans.PropertyValue
    Dim oCursor As Object

    oDoc = ThisComponent

    If Not oDoc.hasLocation() Then
        ShowMessage "Veuillez enregistrer votre document une première fois avant de lancer la macro.", 48, "Sauvegarde requise"
        Exit Sub
    End If

    sDocURL = oDoc.getURL()

    lastDotPos = InStrRev(sDocURL, ".")
    If lastDotPos > 0 Then
        sExt = Mid(sDocURL, lastDotPos)
        sFileName = Left(sDocURL, lastDotPos - 1)
    Else
        sExt = ""
        sFileName = sDocURL
    End If

    k = 1
    Do
        sBackupURL = sFileName & "_Backup_" & k & sExt
        k = k + 1
    Loop While FileExists(sBackupURL)

    On Error Resume Next
    oDoc.storeToURL(sBackupURL, dummy())
    If Err <> 0 Then
        ShowMessage "Erreur lors de la création de la sauvegarde.", 16, "Erreur"
        Exit Sub
    End If

    lastDotPos = InStrRev(sDocURL, "/")
    If lastDotPos > 0 Then
        sDelDocURL = Left(sDocURL, lastDotPos) & "Del_" & Mid(sDocURL, lastDotPos + 1)
    Else
        sDelDocURL = "Del_" & sDocURL
    End If

    oNewDoc = Null

    oText = oDoc.Text
    oEnum = oText.createEnumeration()

    n = 0
    Do While oEnum.hasMoreElements()
        oPara = oEnum.nextElement()
        On Error Resume Next
        sText = oPara.getString()
        If Err = 0 Then
            sTextClean = Trim(sText)

            If sTextClean <> "" And Not EstTexteAIgnorer(sTextClean) Then
                ReDim Preserve texts(n)
                ReDim Preserve paras(n)
                texts(n) = sTextClean
                paras(n) = oPara
                n = n + 1
            End If
        End If
    Loop

    If n = 0 Then
        ShowMessage "Aucun paragraphe à traiter trouvé.", 64, "LibreOffice Writer"
        Exit Sub
    End If

    oSI = Null
    On Error Resume Next
    If Not IsNull(oDoc.CurrentController) Then
        If Not IsNull(oDoc.CurrentController.Frame) Then
            oSI = oDoc.CurrentController.Frame.createStatusIndicator()
        End If
    End If

    If Not IsNull(oSI) Then
        oSI.start("Suppression et extraction des doublons...", n)
    End If

    countSuppr = 0

    For i = n - 1 To 1 Step -1
        If Not IsNull(oSI) Then oSI.setValue(n - i)
        For j = 0 To i - 1
            If texts(i) <> "" And texts(i) = texts(j) Then
                If IsNull(oNewDoc) Or IsEmpty(oNewDoc) Then
                    oNewDoc = StarDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, dummy())
                    oNewText = oNewDoc.Text
                End If

                If countSuppr > 0 Then
                    oNewText.insertControlCharacter(oNewText.getEnd(), com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
                End If
                oNewText.insertString(oNewText.getEnd(), texts(i), False)

                ' Suppression unitaire sécurisée du paragraphe doublon via un curseur ancré précisément sur le paragraphe
                On Error Resume Next
                oCursor = oText.createTextCursorByRange(paras(i))
                oCursor.gotoStartOfParagraph(False)
                oCursor.gotoEndOfParagraph(True)
                If Not oCursor.goRight(1, True) Then
                    oCursor.gotoEndOfParagraph(False)
                    oCursor.gotoStartOfParagraph(True)
                    oCursor.goLeft(1, True)
                End If
                oCursor.setString("")
                If Err = 0 Then
                    countSuppr = countSuppr + 1
                    texts(i) = ""
                End If
                Exit For
            End If
        Next j
    Next i

    If Not IsNull(oSI) Then oSI.end()

    ' Nettoyage final des lignes vides générées par les doublons supprimés
    SupprimerParagraphesVidesSilencieux(oDoc)

    If countSuppr > 0 Then
        If FileExists(sDelDocURL) Then
            On Error Resume Next
            Kill sDelDocURL
        End If

        oNewDoc.storeAsURL(sDelDocURL, dummy())
        oNewDoc.close(True)
        ShowMessage "Traitement terminé avec succès !" & Chr(13) & _
                    "- Sauvegarde créée." & Chr(13) & _
                    "- " & countSuppr & " paragraphe(s) en double extrait(s) dans le fichier Del_ et supprimé(s).", 64, "LibreOffice Writer"
    Else
        ShowMessage "Traitement terminé avec succès !" & Chr(13) & _
                    "- Sauvegarde créée." & Chr(13) & _
                    "- Aucun paragraphe en double n'a été trouvé. Aucun fichier Del_ n'a été créé.", 64, "LibreOffice Writer"
    End If
End Sub

Sub SupprimerParagraphesVides()
    Dim oDoc As Object
    Dim countEmpty As Long
    oDoc = ThisComponent
    countEmpty = SupprimerParagraphesVidesSilencieux(oDoc)
    ShowMessage "Traitement terminé !" & Chr(13) & countEmpty & " paragraphe(s) vide(s) ont été supprimé(s).", 64, "LibreOffice Writer"
End Sub

Function SupprimerParagraphesVidesSilencieux(oDoc As Object) As Long
    Dim oText As Object
    Dim oEnum As Object
    Dim oPara As Object
    Dim sText As String
    Dim countEmpty As Long
    Dim paras() As Object
    Dim n As Long, i As Long
    Dim oCursor As Object

    oText = oDoc.Text
    oEnum = oText.createEnumeration()

    n = 0
    Do While oEnum.hasMoreElements()
        oPara = oEnum.nextElement()
        On Error Resume Next
        sText = oPara.getString()
        If Err = 0 And Trim(sText) = "" Then
            ReDim Preserve paras(n)
            paras(n) = oPara
            n = n + 1
        End If
    Loop

    If n = 0 Then
        SupprimerParagraphesVidesSilencieux = 0
        Exit Function
    End If

    countEmpty = 0
    For i = n - 1 To 0 Step -1
        On Error Resume Next
        oCursor = oText.createTextCursorByRange(paras(i))
        oCursor.gotoStartOfParagraph(False)
        oCursor.gotoEndOfParagraph(True)
        If Not oCursor.goRight(1, True) Then
            oCursor.gotoEndOfParagraph(False)
            oCursor.gotoStartOfParagraph(True)
            oCursor.goLeft(1, True)
        End If
        oCursor.setString("")
        If Err = 0 Then
            countEmpty = countEmpty + 1
        End If
    Next i

    SupprimerParagraphesVidesSilencieux = countEmpty
End Function

Function EstTexteAIgnorer(sChaine As String) As Boolean
    Dim sLower As String
    sLower = LCase(Trim(sChaine))

    If InStr(sLower, "page suivante") > 0 _
    Or InStr(sLower, "confidentiel") > 0 _
    Or InStr(sLower, "majic eyes only") > 0 _
    Or InStr(sLower, "majic. eyes. only") > 0 _
    Or InStr(sLower, "lignes illisibles") > 0 _
    Or InStr(sLower, "top secret") > 0 _
    Or InStr(sLower, "censuré") > 0 _
    Or InStr(sLower, "ultra") > 0 _
    Or InStr(sLower, "interplanetary. phenomenon. unit") > 0 Then
        EstTexteAIgnorer = True
    Else
        EstTexteAIgnorer = False
    End If
End Function

Function FileExists(sURL As String) As Boolean
    Dim oSFA As Object
    On Error Resume Next
    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")
    FileExists = oSFA.exists(sURL)
End Function

Function ConfirmDialog() As Integer
    Dim res As Integer
    On Error Resume Next
    res = MsgBox("Suppression des paragraphes en double" & Chr(10) & _
                 "Création d'un fichier avec préfixe Del_ contenant" & Chr(10) & _
                 "les paragraphes en double (si détectés)", 4 + 32, "Confirmation d'arrêt")
    If res = 0 Then
        ConfirmDialog = 6
    Else
        ConfirmDialog = res
    End If
End Function

Sub ShowMessage(sMsg As String, nType As Integer, sTitle As String)
    On Error Resume Next
    MsgBox sMsg, nType, sTitle
End Sub

Function InStrRev(sTarget As String, sSearch As String) As Long
    Dim pos As Long, lastPos As Long
    pos = InStr(sTarget, sSearch)
    Do While pos > 0
        lastPos = pos
        pos = InStr(pos + 1, sTarget, sSearch)
    Loop
    InStrRev = lastPos
End Function
