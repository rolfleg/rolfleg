Option Explicit

Sub FusionParagrapheNomPropre()
    Dim oDoc As Object
    Dim oText As Object
    Dim oCursor As Object
    Dim colDico As New Collection
    Dim sLastChar As String
    Dim sFirstWord As String
    Dim sPrevText As String
    Dim sCurrText As String
    Dim oEndParaA As Object

    oDoc = ThisComponent
    ' Vérification que nous sommes bien dans un document Writer
    If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
        MsgBox "Cette macro doit être exécutée dans LibreOffice Writer.", 16, "Erreur"
        Exit Sub
    End If

    ' Chargement du dictionnaire des noms propres
    If Not ChargerDicoNomsPropres(colDico) Then
        MsgBox "Impossible de charger le dictionnaire des noms propres. Vérifiez le chemin du fichier.", 48, "Attention"
        Exit Sub
    End If

    oText = oDoc.Text
    oCursor = oText.createTextCursor()
    oCursor.gotoStart(False)

    ' Parcours des paragraphes
    Do
        ' On sélectionne le contenu du paragraphe actuel (A) pour analyse
        oCursor.gotoEndOfParagraph(True)
        sPrevText = Trim(oCursor.getString())

        ' Mémorise la fin du paragraphe actuel
        oEndParaA = oCursor.getEnd()

        ' On déplace le curseur au début du paragraphe suivant (B)
        oCursor.collapseToEnd()
        If Not oCursor.gotoNextParagraph(False) Then Exit Do ' Plus de paragraphe suivant

        ' On sélectionne le contenu du paragraphe suivant (B)
        oCursor.gotoEndOfParagraph(True)
        sCurrText = Trim(oCursor.getString())

        If Len(sPrevText) > 0 Then
            sLastChar = Right(sPrevText, 1)

            ' Règle : dernier caractère du paragraphe précédent est une minuscule
            If EstMinuscule(sLastChar) Then
                sFirstWord = PremierMot(sCurrText)

                ' Règle : le premier mot du paragraphe suivant est dans le dictionnaire
                If sFirstWord <> "" Then
                    If ExisteDansCollection(colDico, sFirstWord) Then
                        ' --- FUSION ---
                        ' On revient à la fin du paragraphe A
                        oCursor.gotoRange(oEndParaA, False)
                        ' On sélectionne le saut de paragraphe qui suit
                        oCursor.goRight(1, True)
                        ' On remplace le saut par un espace
                        oCursor.setString(" ")

                        ' Après la fusion, on se repositionne au début du paragraphe fusionné
                        ' pour vérifier s'il doit lui-même être fusionné avec son nouveau suivant.
                        oCursor.gotoStartOfParagraph(False)
                        GoTo continue_loop
                    End If
                End If
            End If
        End If

        ' Si pas de fusion, on prépare la boucle suivante en se plaçant au début du paragraphe B
        oCursor.collapseToStart()

continue_loop:
    Loop

    MsgBox "Traitement terminé !", 64, "Information"
End Sub

'---------------------------------------------------------
' Charge le dictionnaire depuis le fichier texte spécifié
'---------------------------------------------------------
Function ChargerDicoNomsPropres(col As Collection) As Boolean
    Dim sPath As String
    Dim n As Integer
    Dim l As String
    Dim bOk As Boolean

    ' Chemin spécifié par l'utilisateur
    sPath = "C:\Users\jeanp\AppData\Roaming\LibreOffice\4\user\wordbook\Noms Propres.dic"

    On Error GoTo ErrH

    ' Vérifier si le fichier existe
    If Dir(sPath) = "" Then
        ChargerDicoNomsPropres = False
        Exit Function
    End If

    n = FreeFile
    Open sPath For Input As #n

    bOk = False
    Do Until EOF(n)
        Line Input #n, l
        l = Trim(l)
        ' On commence à charger après la ligne "---"
        If l = "---" Then
            bOk = True
        ElseIf bOk Then
            If l <> "" Then
                ' On ajoute à la collection en utilisant le mot comme clé
                ' On utilise On Error Resume Next pour ignorer les doublons
                On Error Resume Next
                col.Add True, l
                On Error GoTo ErrH
            End If
        End If
    Loop

    Close #n
    ChargerDicoNomsPropres = True
    Exit Function

ErrH:
    If n > 0 Then Close #n
    MsgBox "Erreur lors du chargement du dictionnaire : " & Err.Description, 16, "Erreur"
    ChargerDicoNomsPropres = False
End Function

'---------------------------------------------------------
' Extrait le premier mot d'une chaîne
'---------------------------------------------------------
Function PremierMot(s As String) As String
    Dim sTrimmed As String
    Dim arr() As String
    sTrimmed = LTrim(s)
    arr = Split(sTrimmed, " ")
    If UBound(arr) >= 0 Then
        PremierMot = arr(0)
    Else
        PremierMot = ""
    End If
End Function

'---------------------------------------------------------
' Vérifie si un caractère est une lettre minuscule
'---------------------------------------------------------
Function EstMinuscule(c As String) As Boolean
    If Len(c) <> 1 Then
        EstMinuscule = False
    Else
        ' Une minuscule est égale à sa version LCase et différente de sa version UCase
        EstMinuscule = (c = LCase(c) And c <> UCase(c))
    End If
End Function

'---------------------------------------------------------
' Vérifie si une clé existe dans une Collection
'---------------------------------------------------------
Function ExisteDansCollection(col As Collection, sKey As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = col.Item(sKey)
    ExisteDansCollection = (Err = 0)
    On Error GoTo 0
End Function
