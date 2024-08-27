
' Fichier clsBF.vb
' ----------------

Friend Class clsBF

    ' D'après le fichier d'origine en VB6 :
    ' BF
    ' ------------------------------------------------------------
    ' Module Base de Faits pour Turbo-EXPERT
    ' (c) Philippe LARVET Avril 96
    ' Nouvelle version du 8 juillet 96
    ' ------------------------------------------------------------
    ' Version reprise en VB6 - fin mai 02
    ' ------------------------------------------------------------

#Region "Déclarations et initialisations"

    Friend m_colFaits As Hashtable
    Friend m_colFaitsI As New Collection() ' Faits initiaux
    Friend m_colFaitsIJustes As New Collection() ' Faits initiaux justes
    Friend m_iNbFaitsInitiauxDefinis%

    Private m_oBR As clsBR
    Private m_oDico As clsDico
    Private m_colCR As Specialized.StringCollection

    Friend m_config As clsVBBBox.TConfig

    ' Sert seulement pour valider le test des chaînes
    Private Const iCodeChaine% = -1

    Friend Sub New(oBR As clsBR, oDico As clsDico, colCR As Specialized.StringCollection)
        m_oBR = oBR
        m_oDico = oDico
        m_colCR = colCR
    End Sub

    Private Sub AjouterMsg(sMessage$)
        m_colCR.Add(sMessage)
    End Sub

#End Region

#Region "Gestion des faits initiaux"

    Friend Function bChargerFaitsInitiauxSession(colFaits As Collection) As Boolean

        ' Charger les faits initiaux de la session

        m_colFaits = New Hashtable()

        Dim structfait As clsVBBBox.TFait

        ' Première façon d'énumérer la collection de type Hashtable
        'Dim myEnumerator As IDictionaryEnumerator = myHashTable.GetEnumerator()
        'While myEnumerator.MoveNext()
        '    myEnumerator.Key, myEnumerator.Value
        'End While

        ' Seconde façon peut être un peu plus élégante
        Dim de As DictionaryEntry
        For Each de In m_oDico.m_colDico

            Dim var As clsDico.TVar = CType(de.Value, clsDico.TVar)

            ' Pour les variables intermédiaires, on ne peut préjuger 
            '  de leur valeur initiale que s'il y a une valeur par défaut
            '  dans le dictionnaire
            If m_oDico.bIntermediaire(var.sVariable) Then GoTo VarSuivante
            If m_oDico.sValDefVar(var.sVariable) = "" Then GoTo VarSuivante

            ' Stockage de la variable (chargement du fait dans la BF)
            Dim fait As clsDico.TPremisse = Nothing
            fait.sVar = var.sVariable
            fait.sVar2 = "" ' AAA
            fait.oper = clsDico.TOper.Egal '"="
            fait.sVal = m_oDico.sValDefVar(var.sVariable)
            If fait.sVal <> "" AndAlso fait.sVal.Chars(0) = sGm Then fait.type = clsDico.TTypeVar.Chaine

            fait.sValDebut = fait.sVal
            fait.sRegleApp = ""
            fait.sReglesApp = ""
            Dim rFiabl! = m_oDico.rFiabDef(var.sVariable)
            fait.rFiab = m_oDico.rFiabDef(var.sVariable)
            fait.rFiabOrig = fait.rFiab
            fait.sRemarque = ""

            Dim sCle$ = var.sVariable
            m_colFaits.Add(sCle, fait) ' Hashtable

VarSuivante:
        Next de

        For Each structfait In colFaits

            If Not m_oDico.bVarExiste(structfait.sVar) Then
                AjouterMsg("Erreur : variable inconnue : " & structfait.sVar)
                Return False
            End If

            If m_oDico.bIntermediaire(structfait.sVar) Then
                AjouterMsg("Erreur : " & structfait.sVar & " : les variables intermédiaires")
                AjouterMsg("ne peuvent être initialisées comme les faits initiaux")
                Return False
            End If

            ' Définition d'un fait à indéfini pour pouvoir changer 
            '  la valeur par défaut de la variable
            If structfait.sVal = "?" Then
                If m_oDico.sValDefVar(structfait.sVar) = "" Then GoTo FaitSuivant
                ' On retire le fait
                m_colFaits.Remove(structfait.sVar)
                GoTo FaitSuivant
            End If

            Dim fait As clsDico.TPremisse
            fait.sVar2 = ""
            fait.oper = clsDico.TOper.Egal
            fait.sVal = structfait.sVal
            Dim rFiabl! = structfait.rFiab
            fait.rFiab = structfait.rFiab
            fait.rFiabOrig = fait.rFiab
            fait.sDateOrig = ""

            Dim StrVal$ = structfait.sVal

            If InStr(StrVal, "/") > 0 Then
                ' Ici c'est une date
                fait.sDateOrig = StrVal
                If Not clsUtil.bInverserDate(StrVal) Then
                    AjouterMsg("Erreur dans la session, variable : " & structfait.sVar & " :")
                    AjouterMsg("Date invalide : " & StrVal)
                    Return False
                End If
                ' On considère une date comme un numérique
                fait.type = clsDico.TTypeVar.Numerique

            Else
                Dim sCar$ = Left(StrVal, 1)
                Dim j% = InStr("0123456789-", sCar)
                If j > 0 Then
                    fait.type = clsDico.TTypeVar.Numerique
                Else
                    ' C'est une chaîne ; on lui rajoute des " si elle n'en a pas
                    If sCar <> sGm Then StrVal = sGm & StrVal & sGm
                    fait.type = clsDico.TTypeVar.Chaine
                End If
            End If
            fait.sVal = StrVal

            fait.sVar = structfait.sVar
            fait.sValDebut = fait.sVal ' Pour le bilan
            fait.sRegleApp = ""
            fait.sReglesApp = ""
            fait.sRemarque = structfait.sRemarque

            ' Gestion de la configuration
            Dim sNomVar$ = structfait.sVar
            Dim sVal$ = fait.sVal
            bGestionConfig(sNomVar, sVal, m_config)

            Dim sCle$ = sNomVar
            If bVarExisteDansBF(sNomVar) Then
                ' Modification du fait par rapport à sa valeur par défaut 
                m_colFaits.Item(sCle) = fait
            Else
                ' Ajout du fait
                m_colFaits.Add(sCle, fait)
            End If

FaitSuivant:
        Next structfait

        m_iNbFaitsInitiauxDefinis = 0

        For Each de In m_colFaits
            Dim prem As clsDico.TPremisse = CType(de.Value, clsDico.TPremisse)
            If m_oDico.bConfig(prem.sVar) Then GoTo PremisseSuivante
            Dim StrFait$ = m_oDico.sComposerHypothese(prem)

            Dim sFiab$ = ""
            If prem.rFiab <> clsVBBBox.rCodeFiabIndefini Then _
                sFiab = " (" & Format(prem.rFiab, clsVBBBox.sFormatFiab) & ")"

            StrFait &= sFiab

            m_colFaitsI.Add(StrFait)
            If prem.sVal <> "" And prem.sVal <> clsUtil.sFaux Then
                If Not m_oDico.bConstante(prem.sVar) Then
                    m_colFaitsIJustes.Add(StrFait)
                End If
            End If
            If prem.sVal <> "" Then m_iNbFaitsInitiauxDefinis += 1
PremisseSuivante:
        Next de

        bChargerFaitsInitiauxSession = True

    End Function

    Friend Function bGestionConfig(sVarConfig$, sValeurDef$, ByRef config As clsVBBBox.TConfig) As Boolean

        ' Gestion de la configuration
        bGestionConfig = True
        Select Case sVarConfig
            Case clsVBBBox.sConf_bLogiqueNonMonotone
                config.bLogiqueNonMonotone = clsUtil.bValeurNulleOuVrai(sValeurDef)
            Case clsVBBBox.sConf_bAutoriserReglesContr
                config.bAutoriserReglesContradictoires = clsUtil.bValeurNulleOuVrai(sValeurDef)
            Case clsVBBBox.sConf_bLogiqueFloue
                config.bLogiqueFloue = clsUtil.bValeurNulleOuVrai(sValeurDef)
            Case clsVBBBox.sConf_bLogiqueFloueInterpretee
                config.bLogiqueFloueInterpretee = clsUtil.bValeurNulleOuVrai(sValeurDef)
                If config.bLogiqueFloueInterpretee Then config.bLogiqueFloue = True
            Case Else
                bGestionConfig = False
        End Select

    End Function

#End Region

#Region "Interrogation de la base de faits"

    Friend Function bVarExisteDansBF(sVar$) As Boolean
        bVarExisteDansBF = m_colFaits.ContainsKey(sVar)
    End Function

    Friend Function fait(sVar$) As clsDico.TPremisse
        fait = CType(m_colFaits.Item(sVar), clsDico.TPremisse)
    End Function

    Friend Function bTrouverVar(R%, P%, ByRef sFait$) As Boolean

        ' Retourner True s'il existe un fait de la BF de même nom de variable 
        '  que celui de Premisse n° P de la règle R (retourner aussi la variable)

        Dim de As DictionaryEntry
        For Each de In m_colFaits
            Dim prem As clsDico.TPremisse = CType(de.Value, clsDico.TPremisse)
            If prem.sVar = m_oBR.m_aRegles(R).aPremisses(P).sVar Then _
        sFait = prem.sVar : Return True
        Next de
        Return False

    End Function

    Friend Function bExisteDansBF(ByRef zon As clsDico.TPremisse, ByRef sFait$) As Boolean

        ' Vérifier si une prémisse ou une conclusion existe déjà 
        '  telle quelle dans la BF : si oui, renvoyer le Fait

        Dim de As DictionaryEntry
        For Each de In m_colFaits
            Dim prem As clsDico.TPremisse = CType(de.Value, clsDico.TPremisse)
            If prem.sVar = zon.sVar And prem.oper = zon.oper Then sFait = prem.sVar : Return True
        Next de
        Return False

    End Function

    Friend Function bVerifieeDansBF(ByRef zon As clsDico.TPremisse) As Boolean

        ' Contrôler si une prémisse ou une conclusion est vérifiée dans la BF

        Dim de As DictionaryEntry
        For Each de In m_colFaits
            Dim prem As clsDico.TPremisse = CType(de.Value, clsDico.TPremisse)
            If prem.sVar = zon.sVar And prem.oper = zon.oper And prem.sVal = zon.sVal Then Return True
        Next de
        Return False

    End Function

#End Region

#Region "Vérification d'une prémisse de la BR dans la BF"

    Friend Function bPremisseVraieDansBF(R%, P%, sFait$, ByRef rMinFiab!, ByRef rFiabFait!) As Boolean

        ' Vérifier si une prémisse de la BR est vérifiée dans la BF 
        '---------------------------------------------------------
        ' On procède à des changements de variables :
        ' opérateur du fait -> opFait
        ' valeur num du fait (fait.Valeur) -> iValFait si Valeur numérique
        ' valeur str du fait (fait.Valeur) -> sValFait si Valeur string
        ' valeur num de la prém. (Premisse(R,P).Valeur) -> iValPremisse si num
        ' valeur str de la prém. (Premisse(R,P).Valeur) -> sValPremisse si str
        ' Une fonction spécifique permet de déterminer iValPremisse
        '  au cas où sa valeur est dans une autre variable :
        '  (exemple : "si max_débits > moyenne")
        '---------------------------------------------------------

        'rMinFiab : si un fait à une fiab < à rMinFiab, on met à jour rMinFiab

        Dim fait As clsDico.TPremisse = CType(m_colFaits.Item(sFait), clsDico.TPremisse)
        Dim opFait As clsDico.TOper = fait.oper
        Dim sValFait$ = fait.sVal
        Dim sValPremisse$ = m_oBR.m_aRegles(R).aPremisses(P).sVal
        Dim bPremisseVerifiable As Boolean
        Dim iValPremisse% = iLireValeurPremisse(R, P, bPremisseVerifiable)
        If Not bPremisseVerifiable Then Return False

        Dim bFaitVerifiable As Boolean
        Dim iValFait% = iValeurFait(sFait, bFaitVerifiable)
        If Not bFaitVerifiable Then Return False
        Dim opPremisse As clsDico.TOper = m_oBR.m_aRegles(R).aPremisses(P).oper
        rFiabFait = fait.rFiab

        ' Ce cas ne peut jamais arriver :
        '  - au départ, toutes les fiab. sont >= 0 dans VBBrainBox.mdb ;
        '  - en mode logique floue non interpretée, 
        '    on ne change pas le cours de l'expertise ;
        '  - en mode logique floue interpretée, toutes les fiab. 
        '    sont remises en >= 0 et les faits sont changés
        'If rFiabFait < 0 And rFiabFait <> clsVBBBox.rCodeFiabIndefini Then
        '    If sValFait = clsUtil.sVrai Then
        '        sValFait = clsUtil.sFaux : rFiabFait *= -1
        '    ElseIf sValFait = clsUtil.sFaux Then
        '        sValFait = clsUtil.sVrai : rFiabFait *= -1
        '    End If
        'End If

        bPremisseVraieDansBF = bExaminerPremisse(
            iValPremisse, iValFait, sValPremisse, sValFait, opPremisse, opFait)
        If Not bPremisseVraieDansBF Then Exit Function

        If rFiabFait <> clsVBBBox.rCodeFiabIndefini And m_config.bLogiqueFloue Then
            If rMinFiab = clsVBBBox.rCodeFiabIndefini Then
                rMinFiab = rFiabFait
            Else
                ' Si la loqique floue n'est pas interprétée, les fiab. peuvent
                '  être négatives, cela n'aurait pas de sens d'inverser toute
                '  la fiabilité du résultat de la règle pour une seule prémisse < 0
                '  on calcule donc en valeur absolue
                If Math.Abs(rFiabFait) < rMinFiab Then rMinFiab = Math.Abs(rFiabFait)
            End If
        End If

    End Function

    Private Function iLireValeurPremisse%(R%, P%, ByRef bPremisseVerifiable As Boolean)

        ' Cette fonction intérroge la BR et la BF

        iLireValeurPremisse = 0

        ' Test s'il n'y a pas de seconde variable
        If m_oBR.m_aRegles(R).aPremisses(P).type <> clsDico.TTypeVar.Reference Then

            ' Valeur est numérique
            If m_oBR.m_aRegles(R).aPremisses(P).type = clsDico.TTypeVar.Numerique Then
                iLireValeurPremisse = CInt(Val(m_oBR.m_aRegles(R).aPremisses(P).sVal))
            Else ' Valeur est une string
                iLireValeurPremisse = iCodeChaine
            End If

        Else

            ' On recherche dans la BF si la référence (seconde variable)
            '  peut être remplacée par sa valeur dans la BF
            Dim sVar2$ = m_oBR.m_aRegles(R).aPremisses(P).sVar2
            If bVarExisteDansBF(sVar2) Then
                iLireValeurPremisse = iValeurFait(sVar2, bPremisseVerifiable)
                If Not bPremisseVerifiable Then Exit Function
            Else
                ' La prémisse n'est plus vérifiable (= plus valide)
                '  dès lors qu'il y a bien une seconde variable, mais
                '  que celle-ci n'a pas de valeur affectée dans la BF
                bPremisseVerifiable = False : Exit Function
            End If
        End If
        bPremisseVerifiable = True

    End Function

    Private Function iValeurFait%(sVar$, ByRef bFaitVerifiable As Boolean)

        iValeurFait = 0

        Dim prem As clsDico.TPremisse = CType(m_colFaits.Item(sVar), clsDico.TPremisse)
        If prem.type = clsDico.TTypeVar.Numerique Then
            If prem.sVal = "" Then Exit Function
            iValeurFait = CInt(Val(prem.sVal))
        Else ' Valeur est une string
            iValeurFait = iCodeChaine
        End If
        bFaitVerifiable = True

    End Function

    Private Function bExaminerPremisse(iValPremisse%, iValFait%, sValPremisse$, sValFait$,
        opPremisse As clsDico.TOper, opFait As clsDico.TOper) As Boolean

        Dim bRes As Boolean = False
        Select Case opPremisse
            Case clsDico.TOper.Egal
                bRes = bEgal(iValPremisse, iValFait, sValPremisse, sValFait)

            Case clsDico.TOper.Sup
                Select Case opFait
                    Case clsDico.TOper.Egal : bRes = bSuper(iValPremisse, iValFait)
                    Case clsDico.TOper.Sup : bRes = bSupEgal(iValPremisse, iValFait)
                    Case clsDico.TOper.SupEgal : bRes = bSupEgal(iValPremisse, iValFait)
                End Select

            Case clsDico.TOper.Inf
                Select Case opFait
                    Case clsDico.TOper.Egal : bRes = bInfer(iValPremisse, iValFait)
                    Case clsDico.TOper.Inf : bRes = bInfEgal(iValPremisse, iValFait)
                    Case clsDico.TOper.InfEgal : bRes = bInfer(iValPremisse, iValFait)
                End Select

            Case clsDico.TOper.SupEgal
                Select Case opFait
                    Case clsDico.TOper.Egal : bRes = bSupEgal(iValPremisse, iValFait)
                    Case clsDico.TOper.Sup
                        If (iValFait >= (iValPremisse - 1)) Then bRes = True
                    Case clsDico.TOper.SupEgal : bRes = bSupEgal(iValPremisse, iValFait)
                End Select

            Case clsDico.TOper.InfEgal
                Select Case opFait
                    Case clsDico.TOper.Egal : bRes = bInfEgal(iValPremisse, iValFait)
                    Case clsDico.TOper.Inf
                        If (iValFait <= (iValPremisse + 1)) Then bRes = True
                    Case clsDico.TOper.InfEgal : bRes = bInfEgal(iValPremisse, iValFait)
                End Select

            Case clsDico.TOper.Different
                Select Case opFait
                    Case clsDico.TOper.Egal
                        bRes = bDiff(iValPremisse, iValFait, sValPremisse, sValFait)
                    Case clsDico.TOper.Sup : bRes = bSupEgal(iValPremisse, iValFait)
                    Case clsDico.TOper.Inf : bRes = bInfEgal(iValPremisse, iValFait)
                    Case clsDico.TOper.SupEgal : bRes = bSuper(iValPremisse, iValFait)
                    Case clsDico.TOper.InfEgal : bRes = bInfer(iValPremisse, iValFait)
                    Case clsDico.TOper.Different
                        bRes = bEgal(iValPremisse, iValFait, sValPremisse, sValFait)
                End Select

        End Select
        bExaminerPremisse = bRes

    End Function

    Private Function bEgal(iValPremisse%, iValFait%, sValPremisse$, sValFait$) As Boolean
        If iValFait = iValPremisse And sValFait = sValPremisse Then Return True
        Return False
    End Function

    Private Function bSuper(iValPremisse%, iValFait%) As Boolean
        If iValFait > iValPremisse Then Return True
        Return False
    End Function

    Private Function bInfer(iValPremisse%, iValFait%) As Boolean
        If iValFait < iValPremisse Then Return True
        Return False
    End Function

    Private Function bSupEgal(iValPremisse%, iValFait%) As Boolean
        If iValFait >= iValPremisse Then Return True
        Return False
    End Function

    Private Function bInfEgal(iValPremisse%, iValFait%) As Boolean
        If iValFait <= iValPremisse Then Return True
        Return False
    End Function

    Private Function bDiff(iValPremisse%, iValFait%, sValPremisse$, sValFait$) As Boolean
        If iValFait <> iValPremisse Or sValPremisse <> sValFait Then Return True
        Return False
    End Function

#End Region

#Region "Ajout et modification d'un fait"

    Friend Sub AjouterFait(R%, C%, rMinFiab!, ByRef rFiab!)

        ' Ajouter un fait dans la BF

        Dim fait As clsDico.TPremisse
        fait = m_oBR.m_aRegles(R).aConclusions(C)
        Dim sVar$ = fait.sVar
        Dim sRegle$ = m_oBR.m_aRegles(R).sRegle
        fait.sRegleApp = sRegle
        If fait.sReglesApp = "" Then
            fait.sReglesApp = sRegle
        Else
            fait.sReglesApp &= ", " & sRegle
        End If
        Dim sRegleApp$ = fait.sRegleApp
        Dim sReglesApp$ = fait.sReglesApp

        Dim rFiabRegle! = m_oBR.m_aRegles(R).rFiab
        If rFiabRegle = clsVBBBox.rCodeFiabIndefini And rMinFiab = clsVBBBox.rCodeFiabIndefini Then
            rFiab = clsVBBBox.rCodeFiabIndefini
        Else
            If rFiabRegle = clsVBBBox.rCodeFiabIndefini Then rFiabRegle = 1
            If rMinFiab = clsVBBBox.rCodeFiabIndefini Then rMinFiab = 1
            ' Formule de Mycin (1975)
            ' Si on déduit un fait par une règle à partir de plusieurs faits
            '  alors la fiabilité de ce nouveau fait est le produit
            '  de la fiabilité de la règle par le min. des fiabilités des faits
            rFiab = rFiabRegle * rMinFiab
        End If
        fait.rFiab = rFiab
        fait.rFiabOrig = clsVBBBox.rCodeFiabIndefini

        Dim sCle$ = sVar
        m_colFaits.Add(sCle, fait)

    End Sub

    Friend Function bMAJFait(sFait$, R%, C%, rMinFiab!, ByRef sMajFiab$, ByRef sErr$) As Boolean

        ' Mettre à jour un fait dans la BF

        Dim fait As clsDico.TPremisse = CType(m_colFaits.Item(sFait), clsDico.TPremisse)

        Dim sMemValDebut$ = fait.sValDebut
        Dim sMemVal$ = fait.sVal
        Dim sMemRegleApp$ = fait.sRegleApp
        Dim sVal$ = m_oBR.m_aRegles(R).aConclusions(C).sVal
        Dim sVar$ = m_oBR.m_aRegles(R).aConclusions(C).sVar
        sErr = ""

        If sVal <> sMemVal And sMemVal <> "" Then
            Dim sAvert$ = "Règle " & m_oBR.m_aRegles(R).sRegle & vbCrLf
            sAvert &= "Variable : " & sVar & " : " & sVal & " <> " & sMemVal
            If m_config.bLogiqueNonMonotone Then
                sErr = "Attention : Logique non monotone :" & vbCrLf
                sErr &= "La variable " & sVar & vbCrLf
                sErr &= "possèdent une valeur par défaut : " & sMemValDebut & vbCrLf
                sErr &= sAvert & vbCrLf
            Else
                sErr = "Erreur : En logique monotone, un fait défini" & vbCrLf
                sErr &= "dans la session ou par défaut ou bien déduit" & vbCrLf
                sErr &= "ne peut pas changer de valeur" & vbCrLf
                sErr &= sAvert & vbCrLf
                sErr &= "Solution : n'initialisez par la variable," & vbCrLf
                sErr &= "ou bien vérifiez les règles," & vbCrLf
                sErr &= "ou alors passez en logique non monotone :" & vbCrLf
                sErr &= "ajoutez Config_bLogiqueNonMonotone dans les variables" & vbCrLf
                Return False
            End If
        End If

        Dim rAncFiab! = fait.rFiab

        fait = MAJPremisse(m_oBR.m_aRegles(R).aConclusions(C), fait)

        Dim sRegleApp$ = m_oBR.m_aRegles(R).sRegle
        If fait.sReglesApp = "" Then
            fait.sReglesApp = sRegleApp
        Else
            fait.sReglesApp &= ", " & sRegleApp
        End If

        ' Si double changement de val du fait depuis la val de début : contrad.
        If sMemVal <> fait.sVal And sMemVal <> sMemValDebut Then
            sErr = "Il y a une contradiction entre la règle " & m_oBR.m_aRegles(R).sRegle & vbCrLf
            sErr &= "et la règle " & sMemRegleApp & vbCrLf
            sErr &= "Variable : " & fait.sVar & " : " & fait.sVal & " <> " & sMemVal
            If m_config.bAutoriserReglesContradictoires Then
                sErr = "Attention : " & sErr & vbCrLf
                sErr &= "Dans ce cas, le simple chaînage avant en régime irrévocable" & vbCrLf
                sErr &= "peut être insuffisant à trouver tous les faits déductibles" & vbCrLf
            Else
                sErr = "Erreur : " & sErr
                Return False
            End If
        End If

        sMajFiab = ""
        Dim rFiab!
        Dim rFiabRegle! = m_oBR.m_aRegles(R).rFiab

        If rFiabRegle = clsVBBBox.rCodeFiabIndefini And rMinFiab = clsVBBBox.rCodeFiabIndefini And
            rAncFiab = clsVBBBox.rCodeFiabIndefini Then

            rFiab = clsVBBBox.rCodeFiabIndefini

        Else

            If rFiabRegle = clsVBBBox.rCodeFiabIndefini Then rFiabRegle = 1
            If rMinFiab = clsVBBBox.rCodeFiabIndefini Then rMinFiab = 1
            Dim rNouvFiab! = rFiabRegle * rMinFiab

            Dim bFiabCompatibles As Boolean
            Dim bInverserFiab As Boolean
            ' Vérification de la valeur des faits : les règles de Mycin
            '  ne marchent que si les faits ne changent pas de valeur ; 
            '  et pour les booléens, il faut intégrer le fait que Faux 
            '  est le contraire de Vrai
            If (sMemVal = clsUtil.sVrai And fait.sVal = clsUtil.sFaux) Or
                (sMemVal = clsUtil.sFaux And fait.sVal = clsUtil.sVrai) Then
                bInverserFiab = True
                bFiabCompatibles = True
            ElseIf sMemVal = fait.sVal Then
                bFiabCompatibles = True
                'Else
                ' ToDo : construire une liste de valeurs, et calculer les fiabilités
                '  pour chacune d'entre elles : faire une classe clsValeur
            End If

            If rAncFiab = clsVBBBox.rCodeFiabIndefini Then bFiabCompatibles = False

            If Not bFiabCompatibles Then
                rFiab = rNouvFiab
            Else
                ' Formules associatives de MYCIN (1975)
                ' Si un fait reçoit plusieurs fiabilités, on les combine
                'http://www.computing.surrey.ac.uk/research/ai/PROFILE/mycin.html
                If bInverserFiab Then rAncFiab *= -1
                If rAncFiab >= 0 And rNouvFiab >= 0 Then
                    rFiab = rAncFiab + rNouvFiab - rAncFiab * rNouvFiab
                ElseIf rAncFiab < 0 And rNouvFiab < 0 Then
                    ' Ce cas ne se produit pas car rFiabRegle >= 0 et 
                    '  rMinFiab >= 0 donc rNouvFiab >= 0
                    rFiab = rAncFiab + rNouvFiab + rAncFiab * rNouvFiab
                Else
                    rFiab = (rAncFiab + rNouvFiab) /
                        (1 - Math.Min(Math.Abs(rAncFiab), Math.Abs(rNouvFiab)))
                End If
            End If

            sMajFiab = "(" & Format(rFiab, clsVBBBox.sFormatFiab) & ")"

            If m_config.bLogiqueFloueInterpretee And rFiab < 0 And
                (fait.sVal = clsUtil.sFaux Or fait.sVal = clsUtil.sVrai) Then
                sErr = "Logique floue interprétée : le fait : " & vbCrLf &
                    fait.sVar & " = " & fait.sVal & " (" &
                    Format(rFiab, clsVBBBox.sFormatFiab) & ")"
                sMajFiab = "(" & Format(rFiab, clsVBBBox.sFormatFiab) & ") -> "
                rFiab *= -1
                If fait.sVal = clsUtil.sFaux Then
                    fait.sVal = clsUtil.sVrai
                Else
                    fait.sVal = clsUtil.sFaux
                End If
                sErr &= " devient : " & vbCrLf &
                    fait.sVar & " = " & fait.sVal & " (" &
                    Format(rFiab, clsVBBBox.sFormatFiab) & ")" & vbCrLf
                sMajFiab &= fait.sVal & " (" & Format(rFiab, clsVBBBox.sFormatFiab) & ")"
            End If

        End If
        fait.rFiab = rFiab
        If rFiab = clsVBBBox.rCodeFiabIndefini Then sMajFiab = ""

        ' sRegleApp n'est plus utilisée dans le bilan, on affiche toutes les
        '  règles appliquées, on conserve quand même sRegleApp dans le cas
        '  de grosses applications où il faudra limiter les infos.
        If m_config.bLogiqueFloue Then
            ' On mémorise la règle la plus utile
            If rFiab >= rAncFiab Then
                fait.sRegleApp = m_oBR.m_aRegles(R).sRegle ' La nouvelle
            Else
                fait.sRegleApp = sMemRegleApp ' L'ancienne
            End If
        Else
            ' On mémorise la dernière règle appliquée
            fait.sRegleApp = m_oBR.m_aRegles(R).sRegle
        End If

        Dim sCle$ = sFait
        m_colFaits.Item(sCle) = fait
        bMAJFait = True

    End Function

    Private Function MAJPremisse(ByRef premNouv As clsDico.TPremisse,
            ByRef premActuelle As clsDico.TPremisse) As clsDico.TPremisse

        MAJPremisse = premActuelle ' Conservation des champs actuels

        MAJPremisse.sVar = premNouv.sVar ' Mise à jour des nouveaux champs
        MAJPremisse.sVal = premNouv.sVal

    End Function

#End Region

End Class