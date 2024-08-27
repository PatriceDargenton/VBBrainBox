
' Fichier clsDico.vb
' ------------------

Friend Class clsDico

    ' D'apr�s le fichier d'origine en VB6 :
    ' DicD
    ' ------------------------------------------------------------
    ' Module Dictionnaire des libell�s des variables des R�gles
    ' (c) Philippe LARVET Avril 96
    ' Nouvelle version du 27 mai 96 sans ptr (Pascal)
    ' ------------------------------------------------------------
    ' Version VB6 mai 02
    ' ------------------------------------------------------------

#Region "D�clarations et initialisations"

    Private m_asOper$() = {"=", ">", "<", ">=", "<=", "<>"}
    Private m_asOperCompatTE$() = {"=", ">", "<", "G", "L", "D"}
    Friend Enum TOper
        Egal      ' =
        Sup       ' >
        Inf       ' <
        SupEgal   ' >=
        InfEgal   ' <=
        Different ' <>
    End Enum

    Friend Structure TVar
        Dim sVariable$, sValeurDef$, sConstante$, sDescription$
        Dim bConst, bIntermediaire, bConfig As Boolean
        Dim rFiab!
        ' En mode fichier, on doit faire correspondre l'ordre des faits initiaux
        '  avec l'ordre de chargement du dico
        Dim iNumVar%
    End Structure

    Private m_iNbVariables% ' cf. TVar.iNumVar
    Friend m_iNbVarInitiales%

    Friend Enum TTypeVar
        Numerique ' Num�rique ou date
        Chaine    ' Cha�ne de caract�res
        Reference ' R�f�rence � une variable ou une constante du dico
    End Enum

    Friend Structure TPremisse ' Type pr�misse pour BR et BF
        Dim sVar$         ' Nom de la variable
        Dim oper As TOper 'op�rateur de comparaison
        Dim sVal$         ' Valeur en String de la variable 1
        Dim sVar2$        ' Seconde variable (r�f�rence)
        Dim sValDebut$    ' Valeur initiale de la variable 1, pour le bilan
        ' Derni�re r�gle appliqu�e ayant entrain� une m.�.j. du fait
        Dim sRegleApp$
        Dim sReglesApp$   ' Liste des r�gles appliqu�es
        Dim rFiab!, rFiabOrig!
        Dim sDateOrig$    ' Date dans le format d'origine pour l'exportation
        Dim type As TTypeVar
        Dim sRemarque$    ' Pour d�tailler la valeur du fait dans le rapport
    End Structure

    Friend m_colDico As Hashtable
    Private m_colCR As Specialized.StringCollection

    Friend Sub New(colCR As Specialized.StringCollection)
        m_colCR = colCR
    End Sub

    Private Sub AjouterMsg(sMessage$)
        m_colCR.Add(sMessage)
    End Sub

#End Region

#Region "Chargement du dictionnaire"

    Private Sub InitDico()
        m_colDico = New Hashtable()
    End Sub

    Friend Sub ChargerDico(colVar As Collection)

        InitDico()
        Dim var As TVar
        For Each var In colVar
            ' Valeur ind�finie
            If var.sValeurDef = "?" Then var.sValeurDef = ""
            ' C'est une cha�ne ; on lui rajoute des " si elle n'en a pas
            var.sValeurDef = sTraiterGuillemets(var.sValeurDef)
            If var.sConstante <> "" Then
                If bVarExiste(var.sConstante) Then _
                var.sValeurDef = sValDefVar(var.sConstante)
            End If
            Dim sCle$ = var.sVariable
            m_colDico.Add(sCle, var)
        Next var

    End Sub

    Friend Function sTraiterGuillemets$(sValeur$)
        ' Ajouter des " � la valeur si c'est une cha�ne repr�sentant 
        '  une valeur non num�rique et si ce n'est pas une date ("/")
        If sValeur <> "" AndAlso sValeur.Chars(0) <> sGm AndAlso
            (Not IsNumeric(sValeur)) AndAlso InStr(sValeur, "/") = 0 Then
            sTraiterGuillemets = sGm & sValeur & sGm
        Else
            sTraiterGuillemets = sValeur
        End If
    End Function

    Friend Function bChargerDico(sCheminDico$, ByRef colVar As Collection) As Boolean

        ' Charger le dictionnaire en mode fichier

        bChargerDico = False

        InitDico()
        Dim sr As New IO.StreamReader(sCheminDico, clsUtil.encodageVB6)
        Dim bVarInterm As Boolean
        m_iNbVariables = 0
        Do

            Dim sLigne$ = sr.ReadLine
            If sLigne Is Nothing Then Exit Do

            Dim car$ = Left(sLigne, 1)
            Select Case car
                Case "*" : GoTo LigneSuivante
        ' Les var. interm. sont s�par�es des autres par une ligne de tirets
                Case "-" : sLigne = "" : bVarInterm = True
                Case "=" : Exit Do
            End Select

            If sLigne <> "" Then

                If InStr(sLigne, " ") > 0 Then
                    AjouterMsg("Erreur : les variables doivent �tre sans espace :")
                    AjouterMsg(sLigne) : GoTo Err
                End If

                If bVarExiste(sLigne) Then
                    AjouterMsg("Erreur : variable d�j� d�finie :")
                    AjouterMsg(sLigne) : GoTo Err
                End If

                Dim var As TVar = Nothing
                m_iNbVariables += 1
                var.iNumVar = m_iNbVariables
                var.sVariable = sLigne
                If Not bVarInterm Then
                    ' En mode fichier, la valeur par d�faut des 
                    '  faits initiaux est unique : "FAUX"
                    var.sValeurDef = clsVBBBox.sValFaitInitialDefautModeFichier
                    m_iNbVarInitiales += 1
                    var.bIntermediaire = False
                Else
                    var.bIntermediaire = True
                    var.sValeurDef = clsVBBBox.sValFaitIntermediaireDefautModeFichier
                End If
                var.bConfig = False
                If bNomVarConfig(var.sVariable) Then
                    var.bConfig = True
                    var.sValeurDef = clsUtil.sVrai
                End If
                Dim sCle$ = sLigne
                m_colDico.Add(sCle, var)

                colVar.Add(sLigne, sLigne)
            End If

LigneSuivante:
        Loop While True

Fin:
        bChargerDico = True

Err:
        sr.Close()

    End Function

    Friend Function sNomVar$(iNumVar%)
        ' Trouver la variable iNumVar (en mode fichier seulement)
        Dim de As DictionaryEntry
        For Each de In m_colDico
            Dim var As clsDico.TVar = CType(de.Value, clsDico.TVar)
            If var.iNumVar = iNumVar Then sNomVar = var.sVariable : Exit Function
        Next de
        sNomVar = ""
    End Function

#End Region

#Region "Interrogation du dictionnaire"

    Friend Function ConvOper(sOper$) As TOper
        ' Interpr�tation du mode fichier
        ConvOper = Nothing
        Select Case sOper
            Case "=" : ConvOper = TOper.Egal
            Case ">" : ConvOper = TOper.Sup
            Case "<" : ConvOper = TOper.Inf
            Case "G" : ConvOper = TOper.SupEgal
            Case "L" : ConvOper = TOper.InfEgal
            Case "D" : ConvOper = TOper.Different
        End Select
    End Function

    Friend Function sConvOper$(Oper As TOper, Optional bCompatTurboExpert As Boolean = False)

        If bCompatTurboExpert Then
            sConvOper = m_asOperCompatTE(Oper)
        Else
            sConvOper = m_asOper(Oper)
        End If

    End Function

    Friend Function bVarExiste(sVar$) As Boolean
        bVarExiste = m_colDico.ContainsKey(sVar)
    End Function

    Friend Function bIntermediaire(sVar$) As Boolean
        bIntermediaire = CType(m_colDico.Item(sVar), TVar).bIntermediaire
    End Function

    Friend Function sValDefVar$(sVar$)
        sValDefVar = CType(m_colDico.Item(sVar), TVar).sValeurDef
    End Function

    Friend Function rFiabDef!(sVar$)
        rFiabDef = CType(m_colDico.Item(sVar), TVar).rFiab
    End Function

    Friend Function bConstante(sVar$) As Boolean
        bConstante = CType(m_colDico.Item(sVar), TVar).bConst
    End Function

    Friend Function bConfig(sVar$) As Boolean
        bConfig = CType(m_colDico.Item(sVar), TVar).bConfig
    End Function

    Friend Function bNomVarConfig(sVar$) As Boolean
        If sVar.Length >= 7 AndAlso sVar.Substring(0, 7) = "Config_" Then Return True
        Return False
    End Function

#End Region

#Region "Analyse d'une pr�misse de r�gle"

    '--------------------------------------------------------------------------
    'EXTRACTION DES VARIABLES et VALEURS D'UNE LIGNE-PREMISSE, CONCL ou FAIT
    '--------------------------------------------------------------------------
    Friend Function DecomposerHypothese(sParam$,
            ByRef bErr As Boolean, ByRef sErr$) As clsDico.TPremisse

        DecomposerHypothese = Nothing
        Dim type As TTypeVar
        Dim sVal$ ' Valeur en string de la variable 1
        Dim oper As clsDico.TOper
        Dim sVar$ = ""
        Dim sVar2$ = ""
        Dim iLenPrm% = Len(sParam) ' Les zones � extraire sont dans sParam
        Dim i% = InStr(sParam, " ")
        If i > 0 Then ' Ici, il y a un op�rateur et une valeur

            sVar = Mid(sParam, 1, i - 1)
            If Not bVarExiste(sVar) Then
                bErr = True
                sErr = "Variable inconnue : " & sVar
                'Exit Function
                Return Nothing
            End If
            oper = ConvOper(Mid(sParam, i + 1, 1))

            sVal = Mid(sParam, i + 3, iLenPrm - (i + 2))
            Dim sCar$ = Left(sVal, 1)
            Dim j% = InStr("0123456789-", sCar)

            If j = 0 Then

                ' Ici, sVal contient une valeur-cha�ne ou une seconde variable
                '  si c'est une valeur-cha�ne, elle commence par " 
                If sCar = sGm Then
                    ' Ici, valeur-cha�ne
                    type = TTypeVar.Chaine
                Else
                    ' Ici, seconde variable (r�f�rence)

                    If Not bVarExiste(sVal) Then
                        bErr = True
                        sErr = "Variable inconnue : " & sVar
                        'Exit Function
                        Return Nothing
                    End If

                    sVar2 = sVal
                    If bConstante(sVar2) Then
                        sVal = sValDefVar(sVar2)
                        type = TTypeVar.Chaine
                        If IsNumeric(sVal) Then type = TTypeVar.Numerique
                        If InStr(sVal, "/") > 0 Then type = TTypeVar.Numerique ' Date
                    End If
                End If

            Else
                ' Ici, valeur num�rique
                ' cas particulier du "%":
                '  (on teste si �a a une importance...)
                'If Right$(sVal, 1) = "%" Then _
                '   sVal = Left$(sVal, Len(sVal) - 1)
                type = TTypeVar.Numerique
            End If

        Else

            ' Ici, pas d'op�rateur => la var. est un bool�en
            '  soit 'pr�l�vement' , soit 'non_pr�l�vement'
            If Left(sParam, 4) = "non_" Then ' En mode fichier seulement
                sVar = Right(sParam, Len(sParam) - 4)
                sVal = clsUtil.sFaux
            Else
                sVar = sParam
                sVal = clsUtil.sVrai
            End If
            oper = TOper.Egal
            type = TTypeVar.Chaine

        End If

        DecomposerHypothese.sDateOrig = ""
        If InStr(sVal, "/") > 0 Then
            DecomposerHypothese.sDateOrig = sVal
            clsUtil.bInverserDate(sVal)
        End If

        DecomposerHypothese.sVar = sVar
        DecomposerHypothese.oper = oper
        DecomposerHypothese.sVal = sVal
        DecomposerHypothese.sVar2 = sVar2
        DecomposerHypothese.rFiab = clsVBBBox.rCodeFiabIndefini
        DecomposerHypothese.type = type

    End Function

#End Region

#Region "Traduction d'une pr�misse de r�gle en fran�ais"

    Friend Function sComposerHypothese$(ByRef hyp As TPremisse)

        ' On re�oit ici une pr�misse, et on traduit dans une string
        '  l'expression de la premisse ; cette fonction est l'inverse de 
        '  de DecomposerHypothese 

        Dim sLigne$ = hyp.sVar
        sLigne &= " " & m_asOper(hyp.oper)

        If hyp.type = TTypeVar.Reference Then
            sLigne &= " " & hyp.sVar2
        Else
            Dim sVal$ = hyp.sVal
            If (sVal Is Nothing) Or sVal = "" Then sVal = "?"
            sLigne &= " " & sVal
        End If

        sComposerHypothese = sLigne

    End Function

#End Region

End Class