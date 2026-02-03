Sub PunctuationCorrection
    dim odoc as object
    odoc = thiscomponent

    if not odoc.supportsservice("com.sun.star.text.TextDocument") then
        exit sub
    end if

    dim odict as object
    odict = loadpropernouns()

    odoc.lockcontrollers()
    on error goto error_handler

    dim osearch as object
    osearch = odoc.createsearchdescriptor()
    osearch.searchregularexpression = true

    ' --- Règles 1, 2, 3 (Dans le même paragraphe ou séparés par des espaces) ---
    ' Recherche : lettre minuscule, point optionnel, espaces, et le mot suivant
    osearch.searchstring = "[[:lower:]]\.?\s+[^ \r\n\t]+"

    dim ofoundall as object
    ofoundall = odoc.findall(osearch)

    dim i as long
    dim orange as object
    dim smatch as string
    dim nspacepos as integer
    dim sbefore as string
    dim safter as string
    dim sword2 as string
    dim sfirstchar2 as string
    dim bhasperiod as boolean
    dim slastchar1 as string
    dim bchanged as boolean
    dim snew as string
    dim sprefix as string

    for i = ofoundall.count - 1 to 0 step -1
        orange = ofoundall.getbyindex(i)
        smatch = orange.string

        nspacepos = instr(smatch, " ")
        if nspacepos > 0 then
            sbefore = left(smatch, nspacepos - 1)
            safter = mid(smatch, nspacepos) ' Inclut l'espace ou les espaces

            bhasperiod = (right(sbefore, 1) = ".")
            if bhasperiod then
                slastchar1 = mid(sbefore, len(sbefore) - 1, 1)
            else
                slastchar1 = right(sbefore, 1)
            end if

            dim ssecondpart as string
            ssecondpart = ltrim(safter)
            sword2 = getfirstword(ssecondpart)

            if sword2 <> "" then
                sfirstchar2 = left(sword2, 1)
                dim bindict as boolean
                bindict = odict.containskey(sword2)

                bchanged = false

                ' Règle 1 : Suivi d'un nom propre (dans dict) -> Pas de point, 1 espace
                if bindict then
                    snew = slastchar1 & " " & ssecondpart
                    bchanged = true
                ' Règle 2 : Suivi d'une minuscule -> Pas de point, 1 espace
                elseif islowercase(sfirstchar2) then
                    snew = slastchar1 & " " & ssecondpart
                    bchanged = true
                ' Règle 3 : Suivi d'une majuscule (hors dict) -> Point et 1 espace
                elseif isuppercase(sfirstchar2) and not bindict then
                    snew = slastchar1 & ". " & ssecondpart
                    bchanged = true
                end if

                if bchanged then
                    if bhasperiod then
                        sprefix = left(sbefore, len(sbefore) - 2)
                    else
                        sprefix = left(sbefore, len(sbefore) - 1)
                    end if

                    if (sprefix & snew) <> smatch then
                        orange.string = sprefix & snew
                    end if
                end if
            end if
        end if
    next i

    ' --- Règle 4 (Fusion de paragraphes) ---
    ' Recherche : lettre minuscule, point, et fin de paragraphe
    osearch.searchstring = "[[:lower:]]\.\$"
    ofoundall = odoc.findall(osearch)

    for i = ofoundall.count - 1 to 0 step -1
        orange = ofoundall.getbyindex(i)
        dim ocursor as object
        ocursor = odoc.text.createcursorbyrange(orange)

        if ocursor.gotonextparagraph(false) then
            ocursor.gotostartofparagraph(false)
            ocursor.gotoendofword(true)
            dim snextword as string
            snextword = ocursor.string

            if snextword <> "" then
                dim sfirstchar4 as string
                sfirstchar4 = left(snextword, 1)
                dim swordfordict as string
                if sfirstchar4 = "«" then
                    swordfordict = cleanword(mid(snextword, 2))
                else
                    swordfordict = cleanword(snextword)
                end if

                ' Règle 4 : Si minuscule ou « ET pas dans dictionnaire
                if (islowercase(sfirstchar4) or sfirstchar4 = "«") and not odict.containskey(swordfordict) then
                    ocursor.gotorange(orange, false)
                    ocursor.goleft(0, false)
                    ocursor.goright(1, false) ' Passer la lettre minuscule
                    ocursor.gotoendofparagraph(true) ' Sélectionner le point et la fin de para
                    if ocursor.goright(1, true) then ' Sélectionner le saut de paragraphe
                        ocursor.string = " "
                    end if
                end if
            end if
        end if
    next i

    msgbox "Correction de la ponctuation terminée.", 64, "Information"
    goto end_sub

error_handler:
    msgbox "Erreur " & err & " : " & err.description, 16, "Erreur"

end_sub:
    odoc.unlockcontrollers()
End Sub

Function loadpropernouns() as object
    dim odict as object
    odict = createunoservice("com.sun.star.container.EnumerableMap")
    odict.create("string", "boolean")

    dim opathsub as object
    opathsub = createunoservice("com.sun.star.util.PathSubstitution")
    dim suserpath as string
    suserpath = opathsub.getsubstitutevariablevalue("$(user)")
    dim sfileurl as string
    sfileurl = suserpath & "/Scripts/python/noms_propres.txt"

    dim osfa as object
    osfa = createunoservice("com.sun.star.ucb.SimpleFileAccess")

    if osfa.exists(sfileurl) then
        on error resume next
        dim oinstream as object
        oinstream = osfa.openfileread(sfileurl)
        dim otextin as object
        otextin = createunoservice("com.sun.star.io.TextInputStream")
        otextin.setinputstream(oinstream)
        otextin.setencoding("UTF-8")

        do while not otextin.iseof()
            dim sline as string
            sline = trim(otextin.readline())
            if sline <> "" then
                if not odict.containskey(sline) then
                    odict.put(sline, true)
                end if
            end if
        loop
        oinstream.closeinput()
        on error goto 0
    end if

    loadpropernouns = odict
End Function

Function islowercase(s as string) as boolean
    if s = "" then
        islowercase = false
    else
        islowercase = (s = lcase(s)) and (s <> ucase(s))
    end if
End Function

Function isuppercase(s as string) as boolean
    if s = "" then
        isuppercase = false
    else
        isuppercase = (s = ucase(s)) and (s <> lcase(s))
    end if
End Function

Function getfirstword(s as string) as string
    dim k as integer
    dim res as string
    res = ""
    for k = 1 to len(s)
        dim c as string
        c = mid(s, k, 1)
        if instr(" " & chr(13) & chr(10) & chr(9), c) > 0 then
            exit for
        else
            res = res & c
        end if
    next k
    getfirstword = res
End Function

Function cleanword(s as string) as string
    dim i as integer
    dim res as string
    res = ""
    for i = 1 to len(s)
        dim c as string
        c = mid(s, i, 1)
        if c like "[A-Za-zÀ-ÿ]" then
            res = res & c
        else
            exit for
        end if
    next i
    cleanword = res
End Function
