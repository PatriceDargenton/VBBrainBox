
' Fichier clsBR.vb
' ----------------

Friend Class clsBR

    ' D'apr�s le fichier d'origine en VB6 :
    ' BR
    ' ------------------------------------------------------------
    ' Module Base de R�gles pour Turbo-EXPERT
    ' (c) Philippe LARVET Avril 96
    ' Nouvelle version du 27 mai
    ' ------------------------------------------------------------
    ' Version VB6 mai 02
    ' ------------------------------------------------------------

#Region "D�clarations et initialisations"

    ' Structures de donn�es de la BR :

    Friend Structure THypothese ' Hypoth�se ou conclusion de r�gle
        Dim sRegle$, sVar$, sVal$, sOp$
        Dim bConclusion As Boolean
        Dim rFiabRegle!
    End Structure

    Friend Structure TRegle
        Dim aPremisses() As clsDico.TPremisse
        Dim aConclusions() As clsDico.TPremisse
        ' Bool�en pour indiquer si la r�gle a d�j� �t� utili�e
        Dim bDeduction As Boolean
        Dim sRegle$ ' Nom de la r�gle
        Dim rFiab!
    End Structure

    Friend m_aRegles() As TRegle ' ToDo : Utiliser une Collection : + simple
    Friend m_iNbRegles As Integer

    ' Fin des Structures de donn�es de la BR

    ' Donn�es pour le chargement de la BR en mode fichier :
    Private m_iNbPremisses%, m_iNbConclusions%
    Private m_bErr As Boolean, m_sErr$
    Private m_oDico As clsDico
    Private m_colCR As Specialized.StringCollection

    Friend Sub New(oDico As clsDico, colCR As Specialized.StringCollection)
        ' Note : en VB .NET, il n'est plus n�cessaire de passer les objets par 
        '  ref�rence. De plus, le est plus rapide (m�me pour les objets), 
        '  ce qui n'est pas le cas en VB6. Explication : en VB .NET
        '  si on utilise ByVal, l'objet est copi� une fois, mais il est copi�
        '  2 fois dans le cas du ByRef, dixit VB.NET Professionnel de Wrox Team
        m_oDico = oDico
        m_colCR = colCR
    End Sub

    Private Sub AjouterMsg(sMessage$)
        m_colCR.Add(sMessage)
    End Sub

    Friend Sub InitDeductions()
        Dim R%
        For R = 1 To m_iNbRegles
            m_aRegles(R).bDeduction = False
        Next R
    End Sub

    Friend Sub Initialiser() ' Init BR avant chargement d'une nouvelle BR
        m_iNbRegles = 0
    End Sub

#End Region

#Region "Gestion du mode base de donn�es"

    Friend Function bBDChargerBR(colRegles As Collection) As Boolean

        Initialiser()
        Dim hyp As THypothese
        Dim prem As clsDico.TPremisse
        Dim sMemRegle$ = ""
        m_bErr = False
        Dim iNbPremisses%, iNbConclusions%
        For Each hyp In colRegles

            If hyp.sRegle <> sMemRegle Then
                sMemRegle = hyp.sRegle
                m_iNbRegles += 1
                ReDim Preserve m_aRegles(m_iNbRegles)
                m_aRegles(m_iNbRegles).sRegle = hyp.sRegle
                iNbConclusions = 0
                iNbPremisses = 0
                ReDim m_aRegles(m_iNbRegles).aConclusions(0)
                ReDim m_aRegles(m_iNbRegles).aPremisses(0)
                m_aRegles(m_iNbRegles).rFiab = hyp.rFiabRegle
            End If

            Select Case hyp.sOp
                Case ">=" : hyp.sOp = "G"
                Case "<=" : hyp.sOp = "L"
                Case "<>" : hyp.sOp = "D"
            End Select

            Select Case hyp.sVal
                Case "VRAI", "FAUX" : hyp.sVal = sGm & hyp.sVal & sGm
            End Select

            Dim sParam$ = hyp.sVar & " " & hyp.sOp & " " & hyp.sVal
            prem = m_oDico.DecomposerHypothese(sParam, m_bErr, m_sErr)
            If m_bErr Then Return False

            If hyp.bConclusion Then
                iNbConclusions += 1
                ReDim Preserve m_aRegles(m_iNbRegles).aConclusions(iNbConclusions)
                m_aRegles(m_iNbRegles).aConclusions(iNbConclusions) = prem
            Else
                iNbPremisses += 1
                ReDim Preserve m_aRegles(m_iNbRegles).aPremisses(iNbPremisses)
                m_aRegles(m_iNbRegles).aPremisses(iNbPremisses) = prem
            End If

        Next
        bBDChargerBR = True

    End Function

#End Region

#Region "Gestion du mode fichier"

    Friend Function bChargerBR(sCheminFichierBR$) As Boolean

        bChargerBR = False

        Initialiser()
        Dim sr As New IO.StreamReader(sCheminFichierBR, clsUtil.encodageVB6)
        Dim car$ = ""
        Dim iCar% = sr.Read()
        If iCar <> -1 Then car = Convert.ToChar(iCar)
        If car <> "R" Or iCar = -1 Then
            AjouterMsg("Ce fichier n'est pas une Base de R�gles !")
            GoTo Fin
        End If

        ' Chargement de la BR 
        Dim bPremisse As Boolean
        Dim sNumRegle$ = "."
        m_bErr = False

        Do
            Dim sEnreg$ = sr.ReadLine
            If sEnreg Is Nothing Then Exit Do

            sEnreg = RTrim(sEnreg)
            Dim iLenEnreg% = Len(sEnreg)

            car = Left(sEnreg, 1)
            Select Case car
                Case "R" : sNumRegle = Mid(sEnreg, 1, iLenEnreg)
                Case "s"
                    If Left(sEnreg, 3) = "si " Then
                        bPremisse = True
                        ExtraireSi(sEnreg, sNumRegle)
                    End If
                Case "e"
                    If Left(sEnreg, 3) = "et " Then _
                ExtraireEt(sEnreg, bPremisse)
                Case "a"
                    If Left(sEnreg, 6) = "alors " Then
                        bPremisse = False
                        extraire_alors(sEnreg)
                    End If
                Case "-" : sNumRegle = "."
                Case "=" : Exit Do
            End Select

        Loop While Not m_bErr

        bChargerBR = Not m_bErr
        If m_bErr Then
            AjouterMsg("Erreur R�gle : " & sNumRegle)
            AjouterMsg(m_sErr)
        End If

Fin:
        sr.Close()

    End Function

    Private Sub ExtraireSi(sEnreg$, sNumRegle$)

        m_iNbRegles += 1
        ReDim Preserve m_aRegles(m_iNbRegles)
        m_aRegles(m_iNbRegles).rFiab = clsVBBBox.rCodeFiabIndefini
        m_iNbPremisses = 0
        m_aRegles(m_iNbRegles).sRegle = sNumRegle
        Dim iLenEnreg% = Len(sEnreg)
        Dim sParam$ = Mid(sEnreg, 4, iLenEnreg - 3)
        StockerPremisse(sParam)

    End Sub

    Private Sub ExtraireEt(sEnreg$, bPremisse As Boolean)

        Dim iLenEnreg% = Len(sEnreg)
        Dim sParam$ = Mid(sEnreg, 4, iLenEnreg - 3)
        If bPremisse Then
            StockerPremisse(sParam)
        Else
            StockerConclusion(sParam)
        End If

    End Sub

    Private Sub extraire_alors(sEnreg$)

        m_iNbConclusions = 0
        Dim iLenEnreg% = Len(sEnreg)
        Dim sParam$ = Mid(sEnreg, 7, iLenEnreg - 6)
        StockerConclusion(sParam)

    End Sub

    Private Sub StockerPremisse(sParam$)
        m_iNbPremisses += 1
        ReDim Preserve m_aRegles(m_iNbRegles).aPremisses(m_iNbPremisses)
        m_aRegles(m_iNbRegles).aPremisses(m_iNbPremisses) =
        m_oDico.DecomposerHypothese(sParam, m_bErr, m_sErr)
    End Sub

    Private Sub StockerConclusion(sParam$)
        m_iNbConclusions += 1
        ReDim Preserve m_aRegles(m_iNbRegles).aConclusions(m_iNbConclusions)
        m_aRegles(m_iNbRegles).aConclusions(m_iNbConclusions) =
        m_oDico.DecomposerHypothese(sParam, m_bErr, m_sErr)
    End Sub

#End Region

#Region "Traduction d'une r�gle appliqu�e en fran�ais"

    Friend Sub ExprimerRegleOk(iNumRegle%, bLogiqueFloue As Boolean,
            colFiab As Specialized.StringCollection,
            colFiabC As Specialized.StringCollection)

        Dim P%, C%, sLigne$
        Dim iNbPremisses% = m_aRegles(iNumRegle).aPremisses.GetUpperBound(0)
        For P = 1 To iNbPremisses
            sLigne = m_oDico.sComposerHypothese(m_aRegles(iNumRegle).aPremisses(P))
            If P = 1 Then sLigne = "si " & sLigne Else sLigne = "et " & sLigne
            If bLogiqueFloue Then
                Dim sFiab$ = colFiab(P - 1)
                Dim rFiab! = CSng(sFiab)
                If rFiab <> clsVBBBox.rCodeFiabIndefini Then sLigne &= " (" & sFiab & ")"
            End If
            AjouterMsg(sLigne)
        Next P

        Dim iNbConclusions% = m_aRegles(iNumRegle).aConclusions.GetUpperBound(0)
        For C = 1 To iNbConclusions
            sLigne = m_oDico.sComposerHypothese(m_aRegles(iNumRegle).aConclusions(C))
            If C = 1 Then sLigne = "alors " & sLigne Else sLigne = "et " & sLigne
            If bLogiqueFloue Then
                If C <= colFiabC.Count Then
                    Dim sFiab$ = colFiabC(C - 1)
                    If sFiab <> "" Then sLigne &= " " & sFiab
                End If
            End If
            AjouterMsg(sLigne)
        Next C

    End Sub

#End Region

End Class