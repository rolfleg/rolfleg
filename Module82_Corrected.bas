REM  *****  BASIC  *****

Option Explicit

Public Sub Etape2_AjouterPoints_Stable_VFinale()
    Dim oDoc As Object
    oDoc = ThisComponent
    If oDoc Is Nothing Then
        MsgBox "Erreur : aucun document actif.", 16, "Erreur"
        Exit Sub
    End If

    Dim oEnum As Object
    Dim oTextElements As Object
    oTextElements = oDoc.Text
    oEnum = oTextElements.createEnumeration()

    Dim startTime As Double
    startTime = Timer
    Dim paraCount As Long
    paraCount = 0

    ' ---- CHARGER DICTIONNAIRE NOMSPROPRES ----
    Dim aDico As Object
    aDico = Cha()

    On Error GoTo ErrorHandler

    ' ---------------------------
    '   ANALYSE PAR PARAGRAPHES
    ' ---------------------------
    Do While oEnum.hasMoreElements()
        Dim oPara As Object
        oPara = oEnum.nextElement()

        Dim sTxt As String
        On Error Resume Next
        sTxt = oPara.getString()
        On Error GoTo 0

        If Len(Trim(sTxt)) > 0 Then
            Dim sMod As String
            sMod = AjouterPointsSimple(sTxt, aDico)

            If sMod <> sTxt Then
                oPara.setString(sMod)
            End If
            paraCount = paraCount + 1
        End If
    Loop

Finaliser:
    Dim sTempsEcoule As String
    sTempsEcoule = Format(Timer - startTime, "0.00")
    MsgBox "Terminé : " & paraCount & " paragraphes traités en " & sTempsEcoule & " sec.", 64, "Rapport d'exécution"
    Exit Sub

ErrorHandler:
    MsgBox "Erreur durant le traitement : " & Err.Description, 16, "Erreur Critique"
    Resume Finaliser
End Sub

' ================================================================ ' Analyse texte + insertion automatique des points ' ================================================================ Function AjouterPointsSimple(sTexte As String, aDico As Object) As String Dim resultat As String resultat = "" Dim motCourant As String motCourant = ""

Dim i As Long, L As Long
L = Len(sTexte)

For i = 1 To L
    Dim c As String
    c = Mid(sTexte, i, 1)

    ' Si séparateur
    If c = " " Or c = Chr(13) Or c = Chr(10) Then
        If Len(motCourant) > 0 Then
            Dim prochainMot As String
            prochainMot = ExtraireProchainMotSimple(sTexte, i)

            If Len(prochainMot) > 0 Then
                Dim d As String
                d = Right(motCourant, 1)
                Dim p As String
                p = Left(prochainMot, 1)

                ' Règle 1 : minuscule → minuscule = PAS DE POINT
                If EstMinuscule(d) And EstMinuscule(p) Then
                    resultat = resultat & motCourant & " "

                ' Règle 2 : minuscule → majuscule
                ElseIf EstMinuscule(d) And EstMajuscule(p) Then
                    If EstDansDictionnaire(prochainMot, aDico) Then
                        resultat = resultat & motCourant & " "
                    Else
                        resultat = resultat & motCourant & ". "
                    End If
                Else
                    resultat = resultat & motCourant & " "
                End If
            Else
                resultat = resultat & motCourant & " "
            End If
            motCourant = ""
        Else
            resultat = resultat & c
        End If
    Else
        motCourant = motCourant & c
    End If
Next i

' Ajouter dernier mot
If Len(motCourant) > 0 Then
    resultat = resultat & motCourant
End If

' Nettoyage espaces doublés
Do While InStr(resultat, "  ") > 0
    resultat = Replace(resultat, "  ", " ")
Loop

AjouterPointsSimple = Trim(resultat)
End Function

' ================================================================ ' Extraction prochain mot ' ================================================================ Function ExtraireProchainMotSimple(sTexte As String, posActuelle As Long) As String Dim L As Long L = Len(sTexte) Dim i As Long i = posActuelle + 1 Dim mot As String mot = ""

' Passer séparateurs
Do While i <= L And (Mid(sTexte, i, 1) = " " Or Mid(sTexte, i, 1) = Chr(13) Or Mid(sTexte, i, 1) = Chr(10))
    i = i + 1
Loop

' Lire mot
Do While i <= L
    Dim c As String
    c = Mid(sTexte, i, 1)
    If c = " " Or c = Chr(13) Or c = Chr(10) Then Exit Do
    mot = mot & c
    i = i + 1
Loop

ExtraireProchainMotSimple = mot
End Function

' ================================================================ ' Dictionnaire noms propres ' ================================================================ Function EstDansDictionnaire(mot As String, aDico As Object) As Boolean If Not IsObject(aDico) Then EstDansDictionnaire = False Exit Function End If

Dim m As String
m = NettoyerMot(mot)
If Len(m) = 0 Then
    EstDansDictionnaire = False
    Exit Function
End If

EstDansDictionnaire = aDico.Exists(UCase(m))
End Function

Function NettoyerMot(mot As String) As String Dim i As Long Dim r As String r = ""

For i = 1 To Len(mot)
    Dim c As String
    c = Mid(mot, i, 1)
    If c Like "[A-Za-zÀ-ÖØ-öø-ÿ]" Then
        r = r & c
    Else
        Exit For
    End If
Next i

NettoyerMot = r
End Function

' ================================================================ ' Minuscules / Majuscules ' ================================================================ Function EstMinuscule(c As String) As Boolean If Len(c) = 0 Then EstMinuscule = False Exit Function End If

If (c >= "a" And c <= "z") Then
    EstMinuscule = True
    Exit Function
End If

If c Like "[àâäéèêëîïôöùûüç]" Then
    EstMinuscule = True
    Exit Function
End If

EstMinuscule = False
End Function

Function EstMajuscule(c As String) As Boolean If Len(c) = 0 Then EstMajuscule = False Exit Function End If

If (c >= "A" And c <= "Z") Then
    EstMajuscule = True
    Exit Function
End If

If c Like "[ÀÂÄÉÈÊËÎÏÔÖÙÛÜÇ]" Then
    EstMajuscule = True
    Exit Function
End If

EstMajuscule = False
End Function

' ================================================================
' Charger dictionnaire NomsPropres (Version Corrigée pour LibreOffice)
' ================================================================
Function Cha() As Object
    Dim d As Object
    d = CreateObject("Scripting.Dictionary") ' Compatible avec LO sur Windows
    d.CompareMode = 1 ' Mode insensible à la casse

    Dim sChemin As String
    sChemin = "C:\Users\jeanp\AppData\Roaming\LibreOffice\4\user\wordbook\noms propres.dic"
    Dim sUrl As String
    sUrl = ConvertToURL(sChemin)

    Dim oSFA As Object
    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")

    If Not oSFA.exists(sUrl) Then
        MsgBox "Le fichier dictionnaire '" & sChemin & "' n'a pas été trouvé.", 48, "Avertissement"
        Cha = d
        Exit Function
    End If

    Dim oInputStream As Object
    Dim oTextStream As Object
    Dim ligne As String

    On Error GoTo ErrorHandler_Cha

    oInputStream = oSFA.openFileRead(sUrl)
    oTextStream = createUnoService("com.sun.star.io.TextInputStream")
    oTextStream.setInputStream(oInputStream)
    oTextStream.setEncoding("UTF-8") ' Encodage standard

    Do While Not oTextStream.isEOF()
        ligne = oTextStream.readLine()
        ligne = Trim(ligne)
        If Len(ligne) > 0 Then
            ' Ignorer les en-têtes ou métadonnées du dictionnaire
            If Left(ligne, 1) <> "[" And InStr(ligne, ":") = 0 Then
                If Not d.Exists(UCase(ligne)) Then
                    d.Add UCase(ligne), 1
                End If
            End If
        End If
    Loop

    ' Fermer les flux
    oTextStream.closeInput()
    oInputStream.closeInput()

    Cha = d
    Exit Function

ErrorHandler_Cha:
    MsgBox "Erreur lors de la lecture du fichier dictionnaire : " & Err.Description, 16, "Erreur"
    On Error Resume Next ' Empêche une erreur dans le gestionnaire d'erreur
    If Not IsNull(oTextStream) Then oTextStream.closeInput()
    If Not IsNull(oInputStream) Then oInputStream.closeInput()
    Cha = d
End Function
