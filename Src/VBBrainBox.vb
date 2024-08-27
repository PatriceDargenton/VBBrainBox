
' Fichier VBBrainBox.vb
' ---------------------

Imports System.IO ' Pour StreamWriter
Imports System.Data.OleDb ' Pour les requêtes Access avec OleDbConnection

Friend Class clsVBBBox : Implements IDisposable

#Region "Déclarations et initialisations"

    ' Ne fonctionne pas bien en mode release, pb d'initialisation de la base de donnée ?
    Private Const bVerifierVersionMdb As Boolean = False

    Private Const rVersionBDMin! = 1.0! ' Version min. de la base de données >= 1.0
    Private Const rVersionBDMax! = 2.0! ' Version max. compatible : < 2.0

    Private m_sConnexion$, m_sRepertoireCourant$
    Private m_sCheminBaseMDB$, m_sProvenanceBR$, m_rVersionBD!

    Private m_oConnexion As OleDbConnection ' Regrouper les connexions ouvertes
    Private m_bModeBD As Boolean ' Mode base de données

    ' Collection d'avertissements
    Private m_colAvert As New Specialized.StringCollection()
    Private m_colCR As New Specialized.StringCollection()
    Private m_colSessions As New Collection() ' Collection de sessions à expertiser
    Private m_oDico As New clsDico(m_colCR)
    Private m_oBR As New clsBR(m_oDico, m_colCR)
    Private m_oBF As New clsBF(m_oBR, m_oDico, m_colCR)

    Private m_iNbAvertissements%

    ' Friend = Public restreint au projet
    Friend Structure TConfig ' Configuration de VBBrainBox
        ' Booléen pour autoriser le changement de valeur
        '  d'un fait défini par défaut (voir documentation)
        Dim bLogiqueNonMonotone As Boolean
        Dim bAutoriserReglesContradictoires As Boolean
        Dim bLogiqueFloue As Boolean
        Dim bLogiqueFloueInterpretee As Boolean
    End Structure

    Private m_config As TConfig

    Friend Const rCodeFiabIndefini! = -2 ' Logique floue désactivée
    Friend Const sFormatFiab$ = "0.##"
    Private Const sFormatFiabRapport$ = "0.####"

    Private Const sSeparation$ =
        "------------------------------------------------------------"
    Private Const sFinFichierTurboExpert$ =
        "============================================================"

    Friend Const sConf_bAutoriserReglesContr$ = "Config_bAutoriserReglesContradictoires"
    Friend Const sConf_bLogiqueNonMonotone$ = "Config_bLogiqueNonMonotone"
    Friend Const sConf_bLogiqueFloue$ = "Config_bLogiqueFloue"
    Friend Const sConf_bLogiqueFloueInterpretee$ = "Config_bLogiqueFloueInterpretee"

    Friend Const sFichierVBBBoxMDB$ = "VBBrainBox.mdb"
    'Friend Const sFichierVBBBoxMDB$ = "VBBrainBox.mde" ' Utile pour la version Runtime de MS-Access

    Friend Const sRepertoireApplications$ = "\Applications"
    Friend Const sRepertoireApplicationsTxt$ = "\Applications\ModeFichiersTxt"

    Private Const sValFaitInitialDefaut$ = clsUtil.sVrai
    ' Autre possibilité : ClsUtil.sFaux 
    Friend Const sValFaitInitialDefautModeFichier$ = ""
    Friend Const sValFaitIntermediaireDefautModeFichier$ = ""
    ' Autre possibilité : ClsUtil.sVrai
    Friend Const sValConfigDefautModeFichier$ = "" ' Indéfini
    Private Const sValHypRegleDef$ = clsUtil.sVrai
    Private Const sOperateurRegleDef$ = "="

    Private typeBooleen As Type = GetType(System.Boolean)

    Private Const sSQLVersion$ = "SELECT Version FROM Version"

    Private Const sSQLApplications$ =
        "SELECT IdApplication, Application FROM Application ORDER BY Application"
    Private Const sChpApplication$ = "Application"
    Private Const sSQLApplicationsDescription$ =
        "SELECT IdApplication, Application, Description, Auteur," &
        " AuteurEMail, AuteurWeb, Date, Version, Remarque FROM Application" &
        " WHERE IdApplication = ?"
    Private Const iColRqAppIdApp% = 0
    Private Const iColRqAppApp% = 1
    Private Const iColRqAppDescr% = 2
    Private Const iColRqAppAuteur% = 3
    Private Const iColRqAppEMail% = 4
    Private Const iColRqAppWeb% = 5
    Private Const iColRqAppDate% = 6
    Private Const iColRqAppVers% = 7
    Private Const iColRqAppRem% = 8

    Private Const sSQLDico$ = "SELECT Variable, ValeurParDef, Constante," &
        " Fiab, bConfiguration as bConfig, bConstante as bConst," &
        " bIntermediaire AS bInterm, Description" &
        " FROM RqVariables WHERE IdApplication = ?"
    Private Const iColRqDicoVar% = 0
    Private Const iColRqDicoValDef% = 1
    Private Const iColRqDicoConst% = 2
    Private Const iColRqDicoFiab% = 3
    Private Const iColRqDicobConfig% = 4
    Private Const iColRqDicobConst% = 5
    Private Const iColRqDicobInterm% = 6
    Private Const iColRqDicoDescr% = 7
    Private Const sChpVar$ = "Variable"
    Private Const sChpValDef$ = "ValeurParDef"
    Private Const sChpConst$ = "Constante"
    Private Const sChpFiab$ = "Fiab"
    Private Const sChpbConfig$ = "bConfig"
    Private Const sChpbConst$ = "bConst"
    Private Const sChpbInterm$ = "bInterm"

    Private Const sSQLRegles$ =
        "SELECT Regle, Fiab, Variable, Op, Val, Variable2, bConcl, bInterm" &
        " FROM RqRegles WHERE IdApplication = ?" &
        " ORDER BY Regle, bConcl DESC, Variable"
    Private Const iColRqReglRegle% = 0
    Private Const iColRqReglFiab% = 1
    Private Const iColRqReglVar% = 2
    Private Const iColRqReglOp% = 3
    Private Const iColRqReglVal% = 4
    Private Const iColRqReglVar2% = 5
    Private Const iColRqReglbConcl% = 6
    Private Const iColRqReglbInterm% = 7
    Private Const sChpRegle$ = "Regle"
    Private Const sChpOp$ = "Op"
    Private Const sChpVal$ = "Val"
    Private Const sChpVar2$ = "Variable2"
    Private Const sChpbConcl$ = "bConcl"

    Private Const sSQLReglesDescription$ =
        "SELECT Regle, Description, Origine, Remarque, Fiabilite, Date" &
        " FROM Regle WHERE IdApplication = ? ORDER BY Regle"
    Private Const iColRqReglDRegle% = 0
    Private Const iColRqReglDDescr% = 1
    Private Const iColRqReglDOrig% = 2
    Private Const iColRqReglDRem% = 3
    Private Const iColRqReglDFiab% = 4
    Private Const iColRqReglDDate% = 5

    Private Const sSQLFaits$ =
        "SELECT NomSession, Variable, Op, Val, Const AS Constante," &
        " Fiab, Rem, IdFait, Description FROM RqFaits WHERE IdApplication = ?" &
        " ORDER BY NomSession, Variable"
    Private Const iColRqFaitsSession% = 0
    Private Const iColRqFaitsVar% = 1
    Private Const iColRqFaitsOp% = 2
    Private Const iColRqFaitsVal% = 3
    Private Const iColRqFaitsConst% = 4
    Private Const iColRqFaitsFiab% = 5
    Private Const iColRqFaitsRem% = 6
    Private Const iColRqFaitsIdFait% = 7
    Private Const iColRqFaitsSessionDescr% = 8
    Private Const sChpSession$ = "NomSession"

    Private Const sChpDebut$ = "Debut"
    Private Const sChpFiabOrig$ = "FiabO"
    Private Const sChpFin$ = "Fin"

    Friend Structure TFait ' Fait initial
        Dim sVar$, sVal$, sOp$
        Dim sSession$, sRemarque$
        Dim rFiab!
    End Structure

    Friend Class ClsSession ' Session à expertiser
        Friend sSession$, sDescription$
        Friend colFaits As New Collection()
    End Class

    Friend Sub New()

        m_sRepertoireCourant = Application.StartupPath
        m_sCheminBaseMDB = m_sRepertoireCourant &
        sRepertoireApplications & "\" & sFichierVBBBoxMDB
        m_sConnexion =
            "Provider=Microsoft.Jet.OLEDB.4.0;Password="""";User ID=Admin;" &
            "Data Source=" & m_sCheminBaseMDB & ";Mode=Share Deny None;"
        m_oConnexion = New OleDbConnection(m_sConnexion)

    End Sub
    Protected Overridable Overloads Sub Dispose(disposing As Boolean)
        If disposing Then ' Dispose managed resources
            If Not m_oConnexion Is Nothing Then m_oConnexion.Dispose()
        End If
        ' Free native resources
    End Sub
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Friend Sub InitialiserApplication()
        m_oBR.Initialiser()
        m_sProvenanceBR = ""
        m_bModeBD = False
    End Sub

    Private Sub InitialiserConfigApp()

        m_config.bLogiqueNonMonotone = False
        m_config.bAutoriserReglesContradictoires = False
        m_config.bLogiqueFloue = False
        m_config.bLogiqueFloueInterpretee = False

    End Sub

#End Region

#Region "Gestion du mode base de données"

    Friend Function bBDVerifierVersion() As Boolean

        InitMessages()
        Dim bConnexion As Boolean
        Try
            m_oConnexion.Open()
            bConnexion = True

        Catch
            ' Vérification si le fichier .mdb existe pour produire un
            '  message d'erreur plus explicite
            If Not File.Exists(m_sCheminBaseMDB) Then
                AjouterMsg("Impossible de trouver la base de données :")
                AjouterMsg(m_sCheminBaseMDB)
            Else
                AjouterMsg("Erreur lors de la connexion à la base :")
                AjouterMsg(m_sCheminBaseMDB)
                AjouterMsg(Err.ToString)
                AjouterMsg("Cause possible : MDAC 2.7 doit être installé (cf. doc. pour le lien)")
                AjouterMsg(" lorsque Visual Studio .NET n'est pas installé sur la machine")
            End If
            Return False

        End Try

        If bVerifierVersionMdb Then m_rVersionBD = 1 : bBDVerifierVersion = True : Exit Function

        Dim sVersion$ = clsUtil.sDLookUp(m_oConnexion, sSQLVersion)
        If sVersion = "" Then sVersion = "?" : GoTo Err
        m_rVersionBD = 0
        Try
            m_rVersionBD = CSng(sVersion)
        Catch
            sVersion = "?"
        End Try
        If m_rVersionBD >= rVersionBDMin And m_rVersionBD < rVersionBDMax Then Return True

Err:
        If bConnexion Then m_oConnexion.Close()
        MsgBox("Version de base de données " & sFichierVBBBoxMDB &
            " incompatible !" & vbCrLf &
            "Version : " & sVersion &
            ", Version min. >= " & rVersionBDMin &
            ", Version max. < " & rVersionBDMax, MsgBoxStyle.Critical)
        Return False

    End Function

    Friend Function bBDDefinirVersion() As Boolean

        InitMessages()
        Dim bConnexion As Boolean
        Try
            m_oConnexion.Open()
            bConnexion = True

        Catch
            ' Vérification si le fichier .mdb existe pour produire un
            '  message d'erreur plus explicite
            If Not File.Exists(m_sCheminBaseMDB) Then
                AjouterMsg("Impossible de trouver la base de données :")
                AjouterMsg(m_sCheminBaseMDB)
            Else
                AjouterMsg("Erreur lors de la connexion à la base :")
                AjouterMsg(m_sCheminBaseMDB)
                AjouterMsg(Err.ToString)
                AjouterMsg("Cause possible : MDAC 2.7 doit être installé (cf. doc. pour le lien)")
                AjouterMsg(" lorsque Visual Studio .NET n'est pas installé sur la machine")
            End If
            Return False

        End Try

        Dim sSQL$ = "INSERT INTO Version (Version) VALUES ('" & rVersionBDMin.ToString("0.0") & "');"
        Dim bSucces As Boolean = clsUtil.bRequeteInsertion(m_oConnexion, sSQL)
        m_oConnexion.Close()
        Return bSucces

Err:
        If bConnexion Then m_oConnexion.Close()
        Return False

    End Function

    Friend Function bBDRemplirApplications(ByRef lbApplications As ListBox) As Boolean

        ' Rechercher les applications dans la base de données
        '  et remplir la ListBox passée en entrée

        Dim dt As New DataTable()
        Try
            Dim adp As New OleDbDataAdapter(sSQLApplications, m_oConnexion)
            adp.Fill(dt) ' Récupérer les applications
            lbApplications.DataSource = dt
            lbApplications.DisplayMember = sChpApplication
            ' Ne pas sélectionner une application
            If lbApplications.SelectedIndex >= 0 Then _
            lbApplications.SetSelected(lbApplications.SelectedIndex, False)
            Return True

        Catch err As Exception
            InitMessages()
            AjouterMsg("Erreur lors de la connexion à la base :")
            AjouterMsg(m_sCheminBaseMDB)
            AjouterMsg("Impossible de se connecter à la table 'Application'")
            AjouterMsg(err.ToString)
            Return False

        Finally
            ' Fermer la connexion pour être sûr d'avoir tjrs les données m.a.j.
            m_oConnexion.Close()

        End Try

    End Function

    Friend Function bBDChargerDico(iIdApp%, ByRef dgVariables As DataGrid) As Boolean

        ' Rechercher les variables dans la base de données 
        '  pour l'application iIdApp
        '  et remplir le DataGrid passé en entrée

        m_bModeBD = True

        InitMessages()
        Dim dtVariables As New DataTable()
        Try
            m_oConnexion.Open()
            Dim sSQL$ = clsUtil.sParametrerRq(sSQLDico, iIdApp)
            Dim adp As New OleDb.OleDbDataAdapter(sSQL, m_oConnexion)
            adp.Fill(dtVariables) ' Récupérer les variables

        Catch err As Exception
            AjouterMsg("Erreur lors de la connexion à la base :")
            AjouterMsg(m_sCheminBaseMDB)
            AjouterMsg("Impossible de se connecter à la requête 'RqVariables'")
            AjouterMsg(err.ToString)
            m_oConnexion.Close()
            Return False

        End Try

        ' Fabriquation d'une collection de variables pour charger le dico
        Dim sItem$
        Dim var As clsDico.TVar
        Dim colVar As New Collection()
        Dim r As DataRow

        InitialiserConfigApp()

        For Each r In dtVariables.Rows
            sItem = CStr(r(iColRqDicoVar))

            If InStr(sItem, " ") > 0 Then
                AjouterMsg("Erreur : les variables doivent être sans espace :")
                AjouterMsg(sItem)
                Return False
            End If

            var.sVariable = sItem
            var.sValeurDef = m_oDico.sTraiterGuillemets(clsUtil.sNonVide(r(iColRqDicoValDef)))
            var.rFiab = clsUtil.rNonVide(r(iColRqDicoFiab), rCodeFiabIndefini)
            var.sConstante = clsUtil.sNonVide(r(iColRqDicoConst))
            var.bConst = clsUtil.bNonVide(r(iColRqDicobConst))
            var.bIntermediaire = clsUtil.bNonVide(r(iColRqDicobInterm))
            var.sDescription = clsUtil.sNonVide(r(iColRqDicoDescr))

            ' Gestion de la configuration
            var.bConfig = m_oBF.bGestionConfig(sItem, var.sValeurDef, m_config)
            m_oBF.m_config = m_config

            colVar.Add(var)
        Next r

        m_oDico.ChargerDico(colVar)

        ' Ajustement de la largeur des colonnes
        FixerStyleTableauDico(dgVariables, bAfficherConst:=True)

        dgVariables.SetDataBinding(dtVariables, "")

        bBDChargerDico = True

    End Function

    Friend Function bBDRemplirRegles(iIdApp%, ByRef dgRegles As DataGrid,
            ByRef lbReglesListe As ListBox) As Boolean

        ' Rechercher les règles de l'application iIdApp dans la base de données
        '  et remplir le DataGrid et la ListBox passés en entrée

        InitMessages()
        Dim dtRegles As New DataTable()
        Try
            Dim sSQL$ = clsUtil.sParametrerRq(sSQLRegles, iIdApp)
            Dim adp As New OleDbDataAdapter(sSQL, m_oConnexion)
            adp.Fill(dtRegles) ' Récupérer les règles
            dgRegles.SetDataBinding(dtRegles, "")
            FixerStyleTableauRegles(dgRegles)

        Catch err As Exception
            AjouterMsg("Erreur lors de la connexion à la base :")
            AjouterMsg(m_sCheminBaseMDB)
            AjouterMsg("Impossible de se connecter à la requête 'RqRegles'")
            AjouterMsg(err.ToString)
            m_oConnexion.Close()
            Return False

        End Try

        ' Fabrication d'une collection de règles pour charger la BR
        Dim colRegles As New Collection()
        Dim bMemConclusion As Boolean
        Dim hyp As clsBR.THypothese
        Dim r As DataRow
        For Each r In dtRegles.Rows

            Dim sRegle$ = CStr(r(iColRqReglRegle))
            Dim rFiab! = clsUtil.rNonVide(r(iColRqReglFiab), rCodeFiabIndefini)
            Dim sNomVar$ = CStr(r(iColRqReglVar))
            Dim sOp$ = clsUtil.sNonVide(r(iColRqReglOp), sOperateurRegleDef)
            Dim sValVar$ = clsUtil.sNonVide(r(iColRqReglVal), sValHypRegleDef)
            Dim sNomVar2$ = clsUtil.sNonVide(r(iColRqReglVar2))
            If Not clsUtil.bEstVide(r(iColRqReglVal)) And sNomVar2 <> "" Then
                lbReglesListe.Items.Add("Erreur dans la règle : " & sRegle)
                lbReglesListe.Items.Add("Deux valeurs sont présentes : " &
                sValVar & " et " & sNomVar2)
                Return False
            End If
            Dim bConclusion As Boolean = CBool(r(iColRqReglbConcl))
            bMemConclusion = bConclusion

            hyp.sRegle = sRegle
            hyp.sVar = sNomVar
            hyp.sOp = sOp
            hyp.sVal = sValVar
            If sNomVar2 <> "" Then hyp.sVal = sNomVar2
            hyp.bConclusion = bConclusion
            hyp.rFiabRegle = rFiab
            colRegles.Add(hyp)

        Next r

        Dim bOk1, bOk2 As Boolean
        bOk1 = m_oBR.bBDChargerBR(colRegles)

        Dim sProvenanceBR$ = m_sCheminBaseMDB & " (version : " & m_rVersionBD & ")"
        bOk2 = bRemplirListeRegles(sProvenanceBR, lbReglesListe)
        m_sProvenanceBR = sProvenanceBR

        If bOk1 And bOk2 Then Return True
        Return False

    End Function

    Friend Function bBDRemplirFaits(iIdApp%,
            ByRef dgFaits As DataGrid, ByRef lbSession As ListBox) As Boolean

        ' Rechercher les variables de sessions (faits initiaux) 
        '  pour l'application iIdApp dans la base de données 
        '  et remplir le DataGrid, la ListBox et la collection passés en entrée

        Dim dtVR As New DataTable()
        Try
            Dim sSQL$ = clsUtil.sParametrerRq(sSQLFaits, iIdApp)
            Dim adp As New OleDbDataAdapter(sSQL, m_oConnexion)
            adp.Fill(dtVR) ' Récupérer les variables des sessions
            dgFaits.SetDataBinding(dtVR, "")
            FixerStyleTableauFaits(dgFaits, bAfficherConstantes:=True)

        Catch err As Exception
            InitMessages()
            AjouterMsg("Erreur lors de la connexion à la base :")
            AjouterMsg(m_sCheminBaseMDB)
            AjouterMsg("Impossible de se connecter à la requête 'RqFaits'")
            AjouterMsg(err.ToString)
            m_oConnexion.Close()
            Return False

        End Try

        Dim sSession$ = "", sSessionDescr$ = ""
        Dim sMemSession$ = ""
        Dim sMemSessionDescr$ = ""
        lbSession.Items.Clear()
        Dim oSession As New ClsSession()
        m_colSessions = New Collection()
        Dim fait As TFait
        Dim r As DataRow
        lbSession.BeginUpdate()
        For Each r In dtVR.Rows

            sSession = CStr(r(iColRqFaitsSession))
            sSessionDescr = clsUtil.sNonVide(r(iColRqFaitsSessionDescr))

            If sSession <> sMemSession And sMemSession <> "" Then
                lbSession.Items.Add(sMemSession)
                oSession.sSession = sMemSession
                oSession.sDescription = sMemSessionDescr
                'm_colSessions.Add(sMemSession, oSession) ' Hashtable
                m_colSessions.Add(oSession, sMemSession)  ' Collection
                oSession = New ClsSession()
            End If
            sMemSession = sSession
            sMemSessionDescr = sSessionDescr

            ' Une session peut n'avoir aucune variable de définie
            If Not clsUtil.bEstVide(r(iColRqFaitsVar)) Then
                fait.sSession = sSession
                fait.sVar = CStr(r(iColRqFaitsVar))
                fait.sOp = "="
                ' Un fait initial qui est défini, est mis à VRAI par défaut
                fait.sVal = clsUtil.sNonVide(r(iColRqFaitsVal), sValFaitInitialDefaut)

                If Not clsUtil.bEstVide(r(iColRqFaitsConst)) Then
                    Dim sConst$ = CStr(r(iColRqFaitsConst))
                    If m_oDico.bVarExiste(sConst) Then _
                        fait.sVal = m_oDico.sValDefVar(sConst)
                End If

                fait.rFiab = clsUtil.rNonVide(r(iColRqFaitsFiab), rCodeFiabIndefini)
                fait.sRemarque = clsUtil.sNonVide(r(iColRqFaitsRem))
                Dim iIdFait% = CInt(clsUtil.rNonVide(r(iColRqFaitsIdFait), -1.0!))
                If iIdFait > -1 Then
                    ' Les champs mémo des requêtes Access un peu complexes sont 
                    '  parfois bogués, solution : lire directement le champ mémo
                    '  dans la table
                    Dim sSQL0$ = "SELECT Remarque FROM Fait WHERE IdFait = " & iIdFait
                    Dim sRem$ = clsUtil.sDLookUp(m_oConnexion, sSQL0)
                    ' Mais est-ce qu'il y a vraiment besoin de faire tout ça ?
                    'If sRem <> fait.sRemarque Then _
                    '    MsgBox("DLookUp : Oui, il y a besoin de faire ça !")
                    fait.sRemarque = sRem
                End If

                oSession.colFaits.Add(fait)
            End If

        Next r
        m_oConnexion.Close()

        If sSession <> "" Then
            lbSession.Items.Add(sSession)
            oSession.sSession = sMemSession
            oSession.sDescription = sMemSessionDescr
            m_colSessions.Add(oSession, sMemSession)
        End If
        lbSession.EndUpdate()

        bBDRemplirFaits = True

    End Function

#End Region

#Region "Gestion du mode fichier"

    Friend Sub RemplirListesFichiers(
            ByRef lbFichiersDico As ListBox,
            ByRef lbFichiersBR As ListBox,
            ByRef lbFichiersBF As ListBox)

        ' Remplir les ListBox de fichiers passés en entrée 
        '  avec les fichiers trouvés correspondants

        lbFichiersDico.Items.Clear()
        lbFichiersBR.Items.Clear()
        lbFichiersBF.Items.Clear()

        Dim sRepertoireAppTxt$ = Application.StartupPath &
            clsVBBBox.sRepertoireApplicationsTxt & "\"

        Dim i%, sCheminFichier$, sFichier$
        ' Liste des fichiers du répertoire courant avec le chemin complet
        Dim aFichiers$()
        Try
            aFichiers = IO.Directory.GetFiles(sRepertoireAppTxt, "*.dic")
        Catch
            lbFichiersDico.Items.Add("Répertoire introuvable !")
            lbFichiersDico.Items.Add(clsVBBBox.sRepertoireApplicationsTxt)
            Exit Sub
        End Try
        Dim iNbFichiersDico% = aFichiers.GetUpperBound(0) + 1

        lbFichiersDico.BeginUpdate()
        For i = 0 To iNbFichiersDico - 1
            sCheminFichier = aFichiers(i)
            ' Liste des fichiers du répertoire courant sans le chemin complet
            sFichier = sCheminFichier.Substring(
            sCheminFichier.LastIndexOf("\") + 1)
            lbFichiersDico.Items.Add(sFichier)
        Next i
        lbFichiersDico.EndUpdate()

        aFichiers = IO.Directory.GetFiles(
        sRepertoireAppTxt, "*.brg")
        Dim iNbFichiersBR% = aFichiers.GetUpperBound(0) + 1

        lbFichiersBR.BeginUpdate()
        For i = 0 To iNbFichiersBR - 1
            sCheminFichier = aFichiers(i)
            sFichier = sCheminFichier.Substring(
            sCheminFichier.LastIndexOf("\") + 1)
            lbFichiersBR.Items.Add(sFichier)
        Next i
        lbFichiersBR.EndUpdate()

        aFichiers = IO.Directory.GetFiles(sRepertoireAppTxt, "*.bfa")
        Dim iNbFichiersBF% = aFichiers.GetUpperBound(0) + 1

        lbFichiersBF.BeginUpdate()
        For i = 0 To iNbFichiersBF - 1
            sCheminFichier = aFichiers(i)
            sFichier = sCheminFichier.Substring(
            sCheminFichier.LastIndexOf("\") + 1)
            lbFichiersBF.Items.Add(sFichier)
        Next i
        lbFichiersBF.EndUpdate()

        ' Sélection du premier fichier
        'If iNbFichiersDico > 0 And lbFichiersDico.Enabled Then _
        '    lbFichiersDico.SelectedIndex = 0

    End Sub

    Friend Function bChargerDico(sCheminFichierDico$, ByRef dgVariables As DataGrid) As Boolean

        ' Remplir le DataGrid avec les variables trouvées dans le dico

        InitMessages()
        InitialiserConfigApp()
        Dim colVar As New Collection()
        If Not m_oDico.bChargerDico(sCheminFichierDico, colVar) Then Return False
        bChargerDico = True

        Dim dtVariables As New DataTable()
        dtVariables.Columns.Add(sChpVar)
        dtVariables.Columns.Add(sChpValDef)
        dtVariables.Columns.Add(sChpbInterm)
        dtVariables.Columns(sChpbInterm).DataType = typeBooleen
        dtVariables.Columns.Add(sChpbConfig)
        dtVariables.Columns(sChpbConfig).DataType = typeBooleen

        Dim sItem$
        For Each sItem In colVar
            Dim row1 As DataRow = dtVariables.NewRow
            row1(sChpVar) = sItem
            row1(sChpValDef) = m_oDico.sValDefVar(sItem)
            row1(sChpbInterm) = m_oDico.bIntermediaire(sItem)
            Dim bConfig As Boolean = False
            If m_oDico.bNomVarConfig(sItem) Then
                bConfig = True
                row1(sChpValDef) = clsUtil.sVrai
            End If
            row1(sChpbConfig) = bConfig

            ' Gestion de la configuration
            If m_oBF.bGestionConfig(sItem, clsUtil.sVrai, m_config) Then
                ' Toujours désactivé en mode fichier :
                m_config.bLogiqueFloue = False
                m_config.bLogiqueFloueInterpretee = False

                m_oBF.m_config = m_config
            End If

            dtVariables.Rows.Add(row1)
        Next sItem

        FixerStyleTableauDico(dgVariables, bAfficherConst:=False)

        dgVariables.SetDataBinding(dtVariables, "")

    End Function

    Friend Function bChargerBR(sCheminFichierBR$,
            ByRef lbReglesListe As ListBox, ByRef dgRegles As DataGrid) As Boolean

        ' Remplir la ListBox et le DataGrid avec les règles trouvées dans la BR

        InitMessages()
        If Not m_oBR.bChargerBR(sCheminFichierBR) Then Return False
        If Not bRemplirListeRegles(sCheminFichierBR, lbReglesListe) Then Return False
        m_sProvenanceBR = sCheminFichierBR
        ' Remplissage du tableau de règles
        RemplirTableauRegles(dgRegles)
        Return True

    End Function

    Private Sub RemplirTableauRegles(ByRef dgRegles As DataGrid)

        ' Remplir le DataGrid contenant le tableau de règles

        Dim dtRegles As New DataTable()
        dtRegles.Columns.Add(sChpRegle)
        dtRegles.Columns.Add(sChpVar)
        dtRegles.Columns.Add(sChpOp)
        dtRegles.Columns.Add(sChpVal)
        dtRegles.Columns.Add(sChpVar2)
        dtRegles.Columns.Add(sChpbConcl)
        dtRegles.Columns(sChpbConcl).DataType = typeBooleen
        dtRegles.Columns.Add(sChpbInterm)
        dtRegles.Columns(sChpbInterm).DataType = typeBooleen
        Dim i%, j%

        For j = 1 To m_oBR.m_iNbRegles
            For i = 1 To m_oBR.m_aRegles(j).aPremisses.GetUpperBound(0)
                Dim row1 As DataRow = dtRegles.NewRow
                Dim prem As clsDico.TPremisse = m_oBR.m_aRegles(j).aPremisses(i)
                row1(sChpRegle) = m_oBR.m_aRegles(j).sRegle
                row1(sChpVar) = prem.sVar
                row1(sChpOp) = m_oDico.sConvOper(prem.oper)
                row1(sChpVal) = prem.sVal
                row1(sChpVar2) = prem.sVar2
                row1(sChpbConcl) = False
                row1(sChpbInterm) = m_oDico.bIntermediaire(prem.sVar)
                dtRegles.Rows.Add(row1)
            Next i
            For i = 1 To m_oBR.m_aRegles(j).aConclusions.GetUpperBound(0)
                Dim row1 As DataRow = dtRegles.NewRow
                Dim conclus As clsDico.TPremisse = m_oBR.m_aRegles(j).aConclusions(i)
                row1(sChpRegle) = m_oBR.m_aRegles(j).sRegle
                row1(sChpVar) = conclus.sVar
                row1(sChpOp) = m_oDico.sConvOper(conclus.oper)
                row1(sChpVal) = conclus.sVal
                row1(sChpVar2) = conclus.sVar2
                row1(sChpbConcl) = True
                row1(sChpbInterm) = m_oDico.bIntermediaire(conclus.sVar)
                dtRegles.Rows.Add(row1)
            Next i
        Next j
        FixerStyleTableauRegles(dgRegles)
        dgRegles.SetDataBinding(dtRegles, "")

    End Sub

    Friend Function bRemplirSessions(sFichierBF$,
            ByRef lbLigneFait As ListBox, ByRef dgFaits As DataGrid) As Boolean

        ' Remplir la collection, la ListBox et le DataGrid avec les sessions
        '  trouvées dans le fichier de base de faits

        Dim sCheminFichierBF$ = m_sRepertoireCourant &
            sRepertoireApplicationsTxt & "\" & sFichierBF
        Dim sr As New StreamReader(sCheminFichierBF, clsUtil.encodageVB6)

        ' Construire une collection de sessions = collection de faits
        m_colSessions = New Collection()
        lbLigneFait.BeginUpdate()

        Do
            Dim sLigne$ = sr.ReadLine
            If sLigne Is Nothing Then Exit Do

            ' Construire une collection de faits

            Dim asFaits$() = Split(sLigne, ";")
            If asFaits.GetUpperBound(0) < 1 Then GoTo LigneSuivante
            ' Extraire le nom de la session
            Dim sSession$ = asFaits(0)

            Dim session0 As New ClsSession()
            Dim fait0 As clsVBBBox.TFait = Nothing

            ' Il faut tenir compte de l'ordre de chargement :
            '  celui du dico
            Dim iNumVar%
            For iNumVar = 1 To m_oDico.m_iNbVarInitiales
                If iNumVar >= asFaits.GetUpperBound(0) Then Exit For
                Dim sVar$ = m_oDico.sNomVar(iNumVar)
                fait0.sSession = sSession
                fait0.sVar = sVar
                fait0.sOp = "="
                fait0.rFiab = rCodeFiabIndefini
                Dim sVal$ = asFaits(iNumVar)
                fait0.sVal = sVal
                If sVal <> "" Then session0.colFaits.Add(fait0)
            Next iNumVar

            m_colSessions.Add(session0, sSession)
            lbLigneFait.Items.Add(sSession)

LigneSuivante:
        Loop While True
        sr.Close()
        lbLigneFait.EndUpdate()

        ' Remplissage du tableau de faits initiaux

        Dim dtVR As New DataTable()
        dtVR.Columns.Add(sChpSession)
        dtVR.Columns.Add(sChpVar)
        dtVR.Columns.Add(sChpOp)
        dtVR.Columns.Add(sChpVal)

        For Each oSession As ClsSession In m_colSessions
            RemplirSession(oSession.colFaits, dtVR)
        Next oSession
        FixerStyleTableauFaits(dgFaits, bAfficherConstantes:=False)
        dgFaits.SetDataBinding(dtVR, "")
        bRemplirSessions = True

    End Function

    Private Sub RemplirSession(col As Collection, dtVR As DataTable)
        For Each fait As TFait In col ' Fabriquation du jeu de données Faits
            Dim row1 As DataRow = dtVR.NewRow
            row1(sChpSession) = fait.sSession
            row1(sChpVar) = fait.sVar
            row1(sChpOp) = fait.sOp
            row1(sChpVal) = fait.sVal
            dtVR.Rows.Add(row1)
        Next fait
    End Sub

#End Region

#Region "Gestion commune aux deux modes"

    Private Function bRemplirListeRegles(sProvenanceBR$,
            ByRef lbReglesListe As ListBox) As Boolean

        ' Remplir la ListBox avec les règles chargées dans l'un des 2 modes

        Dim sItem$
        Dim col As New Specialized.StringCollection()
        bRemplirListeRegles = bTraduireRegles(sProvenanceBR, col)
        lbReglesListe.Items.Clear()
        lbReglesListe.BeginUpdate()
        For Each sItem In col
            lbReglesListe.Items.Add(sItem)
        Next sItem
        lbReglesListe.EndUpdate()

    End Function

    Private Function bTraduireRegles(sProvenanceBR$,
            ByRef col As Specialized.StringCollection,
            Optional bCompatibleTurboExpert As Boolean = False,
            Optional bDetailRegles As Boolean = False,
            Optional iIdApp% = 0, Optional sApplication$ = "") As Boolean

        ' Vérifier la liste des règles et remplir la collection passée en entrée

        Dim bConnexion As Boolean
        Dim dtRegles As DataTable = Nothing
        If bDetailRegles Then
            ' Rechercher les règles de l'application iIdApp dans la base de données
            Try
                Dim sSQL$ = clsUtil.sParametrerRq(sSQLReglesDescription, iIdApp)
                Dim adp As New OleDbDataAdapter(sSQL, m_oConnexion)
                dtRegles = New DataTable()
                m_oConnexion.Open()
                adp.Fill(dtRegles) ' Récupérer les règles
                m_oConnexion.Close()
                bConnexion = True
            Catch
            End Try
        End If

        If bCompatibleTurboExpert Then
            col.Add("* Base de règles convertie : " & sProvenanceBR)
            col.Add("* Date : " & DateTime.Now)
            col.Add("")
        Else
            col.Add("")
            col.Add("Base de règles : " & sProvenanceBR)
            If bDetailRegles Then col.Add("Date : " & DateTime.Now)
            col.Add("")
        End If

        Dim j%
        For j = 1 To m_oBR.m_iNbRegles
            With m_oBR.m_aRegles(j)

                ' Affichage de la règle j
                Dim sTab$ = "    "
                Dim sPrefixe$ = ""
                Dim sRegle$ = .sRegle
                If bCompatibleTurboExpert Then
                    sPrefixe$ = "* * "
                    sTab = ""
                    ' Dans Turbo-Expert, les règles doivent commencées par R
                    Dim sRegleTE$ = sRegle
                    If sRegle.Chars(0) <> "R" Then sRegleTE = "R_" & sRegle
                    col.Add(sRegleTE)
                End If
                Dim sTitre$ = sPrefixe & "Règle n°" & Str(j) & " : " & sRegle

                If bConnexion Then
                    Dim r As DataRow
                    r = dtRegles.Rows(j - 1)
                    Dim sDescription$ = clsUtil.sNonVide(r(iColRqReglDDescr))
                    Dim sOrigine$ = clsUtil.sNonVide(r(iColRqReglDOrig))
                    Dim sRemarque$ = clsUtil.sNonVide(r(iColRqReglDRem))
                    Dim sFiabilite$ = clsUtil.sNonVide(r(iColRqReglDFiab))
                    Dim sDate$ = clsUtil.sNonVide(r(iColRqReglDDate))
                    If sFiabilite <> "" Then sTitre &= " (" & sFiabilite & ")"
                    If sDate <> "" Then sTitre &= " : " & sDate
                    col.Add(sTitre)
                    If sDescription <> "" Then col.Add(sPrefixe & "Description : " & sDescription)
                    If sOrigine <> "" Then col.Add(sPrefixe & "Origine     : " & sOrigine)
                    If sRemarque <> "" Then col.Add(sPrefixe & "Remarque    : " & sRemarque)
                    col.Add("")
                Else
                    If .rFiab <> rCodeFiabIndefini Then sTitre &= " (" & .rFiab & ")"
                    col.Add(sTitre)
                End If

                If .aPremisses.GetUpperBound(0) = 0 Then
                    col.Add("Erreur : la règle ne contient aucune prémisse")
                    Return False
                End If
                If .aConclusions.GetUpperBound(0) = 0 Then
                    col.Add("Erreur : la règle ne contient aucune conclusion")
                    Return False
                End If
                Dim i%, sVar$, sOp$, sVal$, sDebPrem$, sPrem$, sVar2$
                For i = 1 To .aPremisses.GetUpperBound(0)
                    sVar = .aPremisses(i).sVar
                    sOp = " " & m_oDico.sConvOper(.aPremisses(i).oper, bCompatibleTurboExpert) & " "
                    sVal = .aPremisses(i).sVal
                    sVar2 = .aPremisses(i).sVar2
                    If m_oDico.bVarExiste(sVar2) Then sVal = sVar2
                    If m_oDico.bConstante(sVar2) And bCompatibleTurboExpert Then
                        sVal = m_oDico.sValDefVar(sVar2)
                    End If
                    If i > 1 Then sDebPrem = "et " Else sDebPrem = "si "
                    sPrem = sTab & sDebPrem & sVar & sOp & sVal
                    col.Add(sPrem)
                Next i

                Dim sDebConc$, sConc$
                For i = 1 To .aConclusions.GetUpperBound(0)
                    sVar = .aConclusions(i).sVar
                    sOp = " " & m_oDico.sConvOper(.aConclusions(i).oper) & " "
                    sVal = .aConclusions(i).sVal
                    sVar2 = .aConclusions(i).sVar2
                    If m_oDico.bVarExiste(sVar2) Then sVal = sVar2
                    If i > 1 Then sDebConc = "et " Else sDebConc = "alors "
                    sConc = sTab & sDebConc & sVar & sOp & sVal
                    col.Add(sConc)
                Next i
                col.Add("")
                If bConnexion And Not bCompatibleTurboExpert Then col.Add("")
                If bCompatibleTurboExpert Then col.Add(sSeparation)

            End With
        Next j

        If bCompatibleTurboExpert Then col.Add(sFinFichierTurboExpert)

        bTraduireRegles = True

    End Function

    Private Sub InitMessages()

        m_colAvert.Clear()
        m_colCR.Clear()

    End Sub

    Friend Function colLireMessages() As Specialized.StringCollection
        colLireMessages = m_colCR
    End Function

    Private Sub AjouterMsg(sMessage$)
        clsUtil.AjouterMsg(sMessage, m_colCR)
    End Sub

#End Region

#Region "Gestion de l'expertise"

    Friend Function bExpertiser(sApplication$, sSession$,
            ByRef lbFaits As ListBox, ByRef lbFaitsJustes As ListBox,
            ByRef dgBilanSession As DataGrid,
            ByRef bConclusions As Boolean,
            ByRef bAvertissements As Boolean) As Boolean

        ' Faire l'expertise avec la collection de faits de la session et
        '  remplir les ListBox des faits et le DataGrid du Bilan
        '  Retourner bConclusions : si une conclusion à pu être tirée
        '  et bAvertissements : s'il y a des avertissements

        '==========================================================================
        ' MAIN DU PROGRAMME T-EXPERT "one-shot"
        '==========================================================================

        InitialiserExpertise()
        Dim oSession As ClsSession
        oSession = CType(m_colSessions(sSession), ClsSession)

        AjouterMsg("Rapport d'expertise de VBBrainBox")
        AjouterMsg(sSeparation)
        AjouterMsg("Application : " & sApplication)
        AjouterMsg("Session : " & sSession)
        If oSession.sDescription <> "" Then _
        AjouterMsg("Descrip.: " & oSession.sDescription)
        AjouterMsg(sSeparation)

        If Not m_oBF.bChargerFaitsInitiauxSession(oSession.colFaits) Then Return False

        AjouterMsg("Configuration :")
        If m_oBF.m_config.bLogiqueNonMonotone Then
            AjouterMsg("Logique non monotone (les faits peuvent changer)")
        Else
            AjouterMsg("Logique monotone (les faits ne peuvent pas changer)")
        End If
        If m_oBF.m_config.bAutoriserReglesContradictoires Then
            AjouterMsg("Les règles contradictoires sont autorisées")
        Else
            AjouterMsg("Les règles contradictoires ne sont pas autorisées")
        End If
        If m_oBF.m_config.bLogiqueFloue Then
            AjouterMsg("Logique floue activée (les fiabilités sont indiquées entre parenthèses)")
            If m_oBF.m_config.bLogiqueFloueInterpretee Then
                AjouterMsg("Logique floue interprétée (les faits peuvent changer)")
            Else
                AjouterMsg("Logique floue non-interprétée (les faits ne peuvent pas changer)")
            End If
        Else
            AjouterMsg("Logique floue désactivée")
            m_oBF.m_config.bLogiqueFloueInterpretee = False
        End If
        AjouterMsg(sSeparation)
        AjouterMsg("")

        AjouterMsg("Compte-rendu d'expertise")
        AjouterMsg("")

        Dim sErr$ = ""
        bConclusions = bChainageAvant(sErr) ' Expertise

        bExpertiser = True
        If sErr <> "" Then AjouterMsg(sErr) : bExpertiser = False

        AjouterMsg("Nombre d'avertissements : " & m_iNbAvertissements)
        bAvertissements = CBool(m_iNbAvertissements > 0)

        Dim sItem$
        lbFaits.BeginUpdate()
        lbFaitsJustes.BeginUpdate()
        sItem = "Nombre de faits initiaux vrais = " & m_oBF.m_colFaitsIJustes.Count
        lbFaits.Items.Add(sItem)
        lbFaitsJustes.Items.Add(sItem)
        sItem = "Nombre de faits initiaux définis = " & m_oBF.m_iNbFaitsInitiauxDefinis
        lbFaits.Items.Add(sItem)
        lbFaitsJustes.Items.Add(sItem)
        sItem = "Nombre de faits finaux = " & m_oBF.m_colFaits.Count
        lbFaits.Items.Add(sItem)
        lbFaitsJustes.Items.Add(sItem)
        sItem = ""
        lbFaits.Items.Add(sItem)
        lbFaitsJustes.Items.Add(sItem)
        For Each sItem In m_oBF.m_colFaitsI
            lbFaits.Items.Add(sItem)
        Next sItem
        For Each sItem In m_oBF.m_colFaitsIJustes
            lbFaitsJustes.Items.Add(sItem)
        Next sItem
        lbFaits.EndUpdate()
        lbFaitsJustes.EndUpdate()

        ' Remplir le tableau du bilan des variables de la session
        RemplirBilan(dgBilanSession)

    End Function

    Private Sub InitialiserExpertise()

        m_iNbAvertissements = 0
        m_oBF.m_colFaitsI = New Collection()
        m_oBF.m_colFaitsIJustes = New Collection()
        ' Initialisation de la config de la session avec la config de l'application,
        '  la config de la session pourra éventuellement être modifiée dans le
        '  chargement des faits de la session
        m_oBF.m_config = m_config
        m_oBR.InitDeductions()
        InitMessages()

    End Sub

    Private Function bChainageAvant(ByRef sErr$) As Boolean

        ' Chaînage avant proprement dit
        ' Retourner bConclusion = bChainageAvant, et sErr

        Dim bAuMoinsUneConclusion As Boolean
        sErr = ""
        Do
            If Not bDeduction(sErr) Then Exit Do
            If sErr <> "" Then Return False
            bAuMoinsUneConclusion = True
        Loop While True

        If Not bAuMoinsUneConclusion Then
            AjouterMsg("Aucune conclusion n'a pu être trouvée")
            AjouterMsg("")
            Return False
        End If

        Return True

    End Function

    Private Function bDeduction(ByRef sErr$) As Boolean

        ' Moteur principal du chaînage avant
        ' Retourner bConclusion = bDeduction, et sErr

        bDeduction = False
        Dim R%
        sErr = ""
        For R = 1 To m_oBR.m_iNbRegles

            If m_oBR.m_aRegles(R).bDeduction Then GoTo RegleSuivante

            Dim sRegle$ = m_oBR.m_aRegles(R).sRegle ' Pour debug
            Dim P%
            Dim rMinFiab! = rCodeFiabIndefini
            Dim colFiab As New Specialized.StringCollection() ' Pour le compte rendu
            Dim iNbPremisses% = m_oBR.m_aRegles(R).aPremisses.GetUpperBound(0)
            Dim NbPremVraies% = 0
            For P = 1 To iNbPremisses
                Dim sFait$ = "", rFiabFait! = 0
                If Not m_oBF.bTrouverVar(R, P, sFait) Then GoTo PremisseSuivante
                If Not m_oBF.bPremisseVraieDansBF(
                R, P, sFait, rMinFiab, rFiabFait) Then GoTo PremisseSuivante

                NbPremVraies += 1
                If m_oBF.m_config.bLogiqueFloue Then
                    Dim sFiab$ = Format(rFiabFait, sFormatFiab)
                    colFiab.Add(sFiab)
                End If

PremisseSuivante:
            Next P

            If NbPremVraies = iNbPremisses Then
                If bConclusions(R, rMinFiab, colFiab, sErr) Then bDeduction = True
                If sErr <> "" Then Return False
            End If

RegleSuivante:
        Next R

    End Function

    Private Function bConclusions(R%, rMinFiab!,
            colFiab As Specialized.StringCollection, ByRef sErr$) As Boolean

        ' Retourner bConclusions et sErr

        '--------------------------------------------------------------------------
        '                              CHAINAGE AVANT
        '--------------------------------------------------------------------------

        bConclusions = False

        ' La règle ne peut être appliquée qu'une seule fois
        m_oBR.m_aRegles(R).bDeduction = True

        Dim iNbConclusions% = m_oBR.m_aRegles(R).aConclusions.GetUpperBound(0)
        Dim colFiabC As New Specialized.StringCollection() ' Pour le compte rendu
        Dim C%
        For C = 1 To iNbConclusions

            ' En logique floue, on affiche aussi les conclusions portant
            '  sur des faits déjà établis afin de préciser la mise à jour
            '  de leur fiabilité
            If Not m_oBF.m_config.bLogiqueFloue And
            m_oBF.bVerifieeDansBF(m_oBR.m_aRegles(R).aConclusions(C)) Then _
                GoTo ConclusionSuivante

            bConclusions = True
            Dim sFait$ = "", sMajFiab$ = "", rFiab! = 0
            If Not (m_oBF.bExisteDansBF(m_oBR.m_aRegles(R).aConclusions(C), sFait)) Then
                m_oBF.AjouterFait(R, C, rMinFiab, rFiab)
                If m_oBF.m_config.bLogiqueFloue Then
                    Dim sFiab$ = "(" & Format(rFiab, sFormatFiab) & ")"
                    If rFiab = rCodeFiabIndefini Then sFiab = ""
                    colFiabC.Add(sFiab)
                End If
            Else
                If Not m_oBF.bMAJFait(sFait, R, C, rMinFiab, sMajFiab, sErr) Then
                    bConclusions = False
                    ListerRegle(R, rMinFiab, m_oBF.m_config.bLogiqueFloue,
                    colFiab, colFiabC)
                    Exit Function
                End If
                If m_oBF.m_config.bLogiqueFloue Then _
                colFiabC.Add(sMajFiab)
            End If

            ' Affichage des avertissements
            If sErr <> "" Then
                clsUtil.AjouterMsg(sErr, m_colAvert)
                sErr = ""
                m_iNbAvertissements += 1
            End If

            ' Libellé de la conséquence-conclusion
            Dim sConclusion$ = m_oDico.sComposerHypothese(m_oBR.m_aRegles(R).aConclusions(C))
            ' N'afficher la règle qu'à la fin, une fois que toutes les fiab
            '  sont connues
            If C = iNbConclusions Then _
                ListerRegle(R, rMinFiab, m_oBF.m_config.bLogiqueFloue, colFiab, colFiabC)

ConclusionSuivante:
        Next C

    End Function

    Private Sub ListerRegle(R%, rMinFiab!, bLogiqueFloue As Boolean,
        colFiab As Specialized.StringCollection, colFiabC As Specialized.StringCollection)

        Dim sMsg$ = "Selon la règle " & m_oBR.m_aRegles(R).sRegle
        If m_oBF.m_config.bLogiqueFloue Then
            If m_oBR.m_aRegles(R).rFiab <> rCodeFiabIndefini Then _
                sMsg &= " (" & m_oBR.m_aRegles(R).rFiab & ")"
            sMsg &= " :"
            If rMinFiab <> rCodeFiabIndefini Then _
                sMsg &= " (min. fiab. faits : " &
                    Format(rMinFiab, clsVBBBox.sFormatFiab) & ") :"
        End If
        AjouterMsg(sMsg)
        m_oBR.ExprimerRegleOk(R, m_oBF.m_config.bLogiqueFloue, colFiab, colFiabC)
        AjouterMsg("")

    End Sub

#End Region

#Region "Bilan, rapport et exportation"

    Private Sub RemplirBilan(ByRef dgBilanSession As DataGrid)

        ' Remplir le DataGrid contenant le tableau du 
        '  bilan des variables de la session

        Dim dtBilan As New DataTable()
        dtBilan.Columns.Add(sChpVar)
        dtBilan.Columns.Add(sChpDebut)
        dtBilan.Columns.Add(sChpFiabOrig)
        dtBilan.Columns.Add(sChpFin)
        dtBilan.Columns.Add(sChpFiab)
        dtBilan.Columns.Add(sChpRegle)
        dtBilan.Columns.Add(sChpbInterm)
        dtBilan.Columns(sChpbInterm).DataType = typeBooleen
        Dim de As DictionaryEntry
        For Each de In m_oDico.m_colDico
            Dim var As clsDico.TVar = CType(de.Value, clsDico.TVar)
            If m_oDico.bIntermediaire(var.sVariable) Then GoTo VarSuivante
            If m_oDico.bConstante(var.sVariable) Then GoTo VarSuivante
            If Not m_oBF.bVarExisteDansBF(var.sVariable) Then GoTo VarSuivante
            Dim fait As clsDico.TPremisse = m_oBF.fait(var.sVariable)
            Dim row1 As DataRow = dtBilan.NewRow
            row1(sChpVar) = fait.sVar
            row1(sChpDebut) = IIf((fait.sValDebut Is Nothing), "", fait.sValDebut)
            row1(sChpFiabOrig) = ""
            If fait.rFiabOrig <> rCodeFiabIndefini Then row1(sChpFiabOrig) = fait.rFiabOrig
            row1(sChpFin) = fait.sVal
            row1(sChpFiab) = ""
            If fait.rFiab <> rCodeFiabIndefini Then row1(sChpFiab) = fait.rFiab
            row1(sChpRegle) = fait.sReglesApp
            row1(sChpbInterm) = False
            dtBilan.Rows.Add(row1)
VarSuivante:
        Next de

        For Each de In m_oDico.m_colDico
            Dim var As clsDico.TVar = CType(de.Value, clsDico.TVar)
            If Not m_oDico.bIntermediaire(var.sVariable) Then GoTo VarSuivante2
            If Not m_oBF.bVarExisteDansBF(var.sVariable) Then GoTo VarSuivante2
            Dim fait As clsDico.TPremisse = m_oBF.fait(var.sVariable)
            Dim row1 As DataRow = dtBilan.NewRow
            row1(sChpVar) = var.sVariable
            row1(sChpDebut) = ""
            row1(sChpFiabOrig) = ""
            row1(sChpFin) = ""
            row1(sChpFiab) = ""
            row1(sChpRegle) = ""
            row1(sChpbInterm) = True
            row1(sChpDebut) = IIf((fait.sValDebut Is Nothing), "", fait.sValDebut)
            row1(sChpFin) = fait.sVal
            row1(sChpRegle) = fait.sReglesApp
            If fait.rFiab <> rCodeFiabIndefini Then row1(sChpFiab) = fait.rFiab

            dtBilan.Rows.Add(row1)
VarSuivante2:
        Next de

        FixerStyleTableauBilan(dgBilanSession)
        dgBilanSession.SetDataBinding(dtBilan, "")

    End Sub

    ' D'après le fichier d'origine en VB6 :
    ' CRD
    '--------------------------------------
    ' Module Compte Rendu pour Turbo-EXPERT
    '--------------------------------------
    ' version VB6 mai 02
    '--------------------------------------

    Friend Sub CreerCompteRendu(sCheminFichier$, iIdApplication%, sApplication$, sSession$)

        Dim sw As New StreamWriter(sCheminFichier)
        sw.WriteLine("Rapport d'expertise de VBBrainBox")
        sw.WriteLine("Date : " & DateTime.Now)
        sw.WriteLine(sSeparation)
        sw.WriteLine("Application : " & sApplication)
        ExporterDescrApplication(sw, iIdApplication)

        If sSession <> "" Then

            sw.WriteLine("Session     : " & sSession)
            sw.WriteLine(sSeparation)
            sw.WriteLine("")
            If m_iNbAvertissements > 0 Then
                sw.WriteLine("Nombre d'avertissements : " & m_iNbAvertissements)
                sw.WriteLine("")
            End If

            ' Exporter la liste des faits initiaux
            sw.WriteLine("Faits initiaux :")
            sw.WriteLine("")
            Dim sItem$ = "Nombre de faits initiaux = " & m_oBF.m_colFaitsI.Count
            sw.WriteLine(sItem)
            sItem = "Nombre de faits initiaux définis = " & m_oBF.m_colFaitsIJustes.Count
            sw.WriteLine(sItem)
            If Not m_oBF.m_colFaits Is Nothing Then ' Si une err a eu lieu
                sItem = "Nombre de faits finaux = " & m_oBF.m_colFaits.Count
                sw.WriteLine(sItem)
                sw.WriteLine("")
            End If
            For Each sItem In m_oBF.m_colFaitsI
                sw.WriteLine(sItem)
            Next sItem

            ' Exporter les conclusions
            sw.WriteLine("")
            For Each sItem In colLireMessages()
                sw.WriteLine(sItem)
            Next sItem

            ' Exporter les avertissements
            If m_iNbAvertissements > 0 Then
                sw.WriteLine("")
                For Each sItem In m_colAvert
                    sw.WriteLine(sItem)
                Next sItem
            Else
                sw.WriteLine("")
            End If

            If Not m_oBF.m_colFaits Is Nothing Then ExporterBilan(sw)

        Else

            ' Exporter les messages d'erreur s'il y en a
            sw.WriteLine("")
            Dim sItem$
            For Each sItem In colLireMessages()
                sw.WriteLine(sItem)
            Next sItem

            ' Exporter toutes les sessions
            sw.WriteLine(sSeparation)
            sw.WriteLine("Toutes les sessions")
            ExporterFaitsInitiaux(sw, m_colSessions, bCompatibleTurboExpert:=False)
            sw.WriteLine(sSeparation)
        End If

        ExporterRegles(sw, iIdApplication, bCompatibleTurboExpert:=False)

        sw.WriteLine(sSeparation)
        Dim sVersionAppli$ = My.Application.Info.Version.Major &
            "." & My.Application.Info.Version.Minor &
            My.Application.Info.Version.Build
        sw.WriteLine("VBBrainBox " & sVersionAppli)
        sw.WriteLine("")
        sw.WriteLine("d'après Turbo-Expert 1.2 pour Windows")
        sw.WriteLine("(c) Philippe Larvet 1996, 2003")
        sw.WriteLine("")
        sw.WriteLine("https://github.com/PatriceDargenton/VBBrainBox")
        sw.WriteLine(sSeparation)

        sw.Close()

        ' Ancienne méthode :
        'Dim iRet% = Shell("notepad " & sCheminFichier, AppWinStyle.NormalFocus)
        Dim startInfo As New ProcessStartInfo("notepad.exe") With {
        .Arguments = sCheminFichier,
        .WindowStyle = ProcessWindowStyle.Normal
    }
        Process.Start(startInfo)

    End Sub

    Private Sub ExporterDescrApplication(sw As StreamWriter, iIdApp%)

        ' Exporter la description de l'application 

        Dim col As New Specialized.StringCollection()
        ExporterDescrApplication(col, iIdApp)
        Dim sItem$
        For Each sItem In col
            sw.WriteLine(sItem)
        Next sItem

    End Sub

    Private Sub ExporterDescrApplication(col As Specialized.StringCollection,
    iIdApp%, Optional sPrefixe$ = "")

        ' Exporter la description de l'application dans une collection

        If Not m_bModeBD Then Exit Sub

        Dim dtApplication As DataTable
        Try
            Dim sSQL$ = clsUtil.sParametrerRq(sSQLApplicationsDescription, iIdApp)
            Dim adp As New OleDbDataAdapter(sSQL, m_oConnexion)
            dtApplication = New DataTable()
            m_oConnexion.Open()
            adp.Fill(dtApplication) ' Récupérer la description des applications
            m_oConnexion.Close()
        Catch
            Exit Sub
        End Try

        Dim r As DataRow
        For Each r In dtApplication.Rows ' Il n'y en a qu'une

            Dim sApp$ = CStr(r(iColRqAppApp))
            Dim sDescr$ = clsUtil.sNonVide(r(iColRqAppDescr))
            Dim sAuteur$ = clsUtil.sNonVide(r(iColRqAppAuteur))
            Dim sEMail$ = clsUtil.sTraiterHyperlienAccess(
            clsUtil.sNonVide(r(iColRqAppEMail)))
            Dim sWeb$ = clsUtil.sTraiterHyperlienAccess(
            clsUtil.sNonVide(r(iColRqAppWeb)))
            Dim sDate$ = clsUtil.sNonVide(r(iColRqAppDate))
            Dim sVersion$ = clsUtil.sNonVide(r(iColRqAppVers))
            Dim sRem$ = clsUtil.sNonVide(r(iColRqAppRem))

            If sDescr <> "" Then col.Add(sPrefixe &
            "Description : " & sDescr)
            If sAuteur <> "" Then col.Add(sPrefixe &
            "Auteur      : " & sAuteur)
            If sEMail <> "" Then col.Add(sPrefixe &
            "EMail       : " & sEMail)
            If sWeb <> "" Then col.Add(sPrefixe &
            "Web         : " & sWeb)
            If sDate <> "" Then col.Add(sPrefixe &
            "Date        : " & sDate)
            If sVersion <> "" Then col.Add(sPrefixe &
            "Version     : " & sVersion)
            If sRem <> "" Then col.Add(sPrefixe &
            "Remarque    : " & sRem)

        Next r

    End Sub

    Private Sub ExporterBilan(sw As StreamWriter)

        ' Exporter le bilan des variables de la session

        sw.WriteLine(sSeparation)
        sw.WriteLine("")
        sw.WriteLine("Bilan des variables : Avant : Après")
        sw.WriteLine("")
        Dim de As DictionaryEntry
        For Each de In m_oDico.m_colDico
            Dim var As clsDico.TVar = CType(de.Value, clsDico.TVar)
            If m_oDico.bIntermediaire(var.sVariable) Then GoTo VarSuivante
            If m_oDico.bConstante(var.sVariable) Then GoTo VarSuivante
            If Not m_oBF.bVarExisteDansBF(var.sVariable) Then GoTo VarSuivante
            Dim fait As clsDico.TPremisse = m_oBF.fait(var.sVariable)
            Dim sVar$ = fait.sVar
            Dim sDebut$ = fait.sValDebut
            If sDebut = "" Then sDebut = "?"
            Dim sFin$ = fait.sVal
            Dim sRegles$ = fait.sReglesApp
            Dim sLigne$ = sVar & " = " & sDebut
            If fait.rFiabOrig <> rCodeFiabIndefini Then _
            sLigne &= " (" & Format(fait.rFiabOrig, sFormatFiabRapport) & ")"
            If fait.rFiab <> rCodeFiabIndefini Then _
            sFin &= " (" & Format(fait.rFiab, sFormatFiabRapport) & ")"
            If sFin <> "" Then sLigne &= " : " & sFin
            If sRegles <> "" Then sLigne &= " (" & sRegles & ")"
            sw.WriteLine(sLigne)
            If var.sDescription <> "" Then _
            sw.WriteLine("Descrip. : " & var.sDescription)
            If fait.sRemarque <> "" Then
                sw.WriteLine("Remarque : " & fait.sRemarque)
                sw.WriteLine("")
            End If

VarSuivante:
        Next de

        sw.WriteLine("")
        sw.WriteLine("Variables intermédiaires :")
        sw.WriteLine("")
        For Each de In m_oDico.m_colDico
            Dim var As clsDico.TVar = CType(de.Value, clsDico.TVar)
            If Not m_oDico.bIntermediaire(var.sVariable) Then GoTo VarSuivante2
            If Not m_oBF.bVarExisteDansBF(var.sVariable) Then GoTo VarSuivante2
            Dim fait As clsDico.TPremisse = m_oBF.fait(var.sVariable)
            Dim sVar$ = var.sVariable
            Dim sDebut$ = "?"
            Dim sFin$ = ""
            Dim sRegle$ = ""
            sDebut = CStr(IIf((fait.sValDebut Is Nothing),
            "?", fait.sValDebut))
            If fait.rFiabOrig <> rCodeFiabIndefini Then _
            sDebut &= " (" & Format(fait.rFiabOrig, sFormatFiabRapport) & ")"
            sFin = fait.sVal
            If fait.rFiab <> rCodeFiabIndefini Then _
            sFin &= " (" & Format(fait.rFiab, sFormatFiabRapport) & ")"
            sRegle = fait.sReglesApp '.sRegleApp

            If sDebut = "" Then sDebut = "?"
            Dim sLigne$ = sVar & " = " & sDebut
            If sFin <> "" Then sLigne &= " : " & sFin
            If sRegle <> "" Then sLigne &= " (" & sRegle & ")"
            sw.WriteLine(sLigne)
            If var.sDescription <> "" Then _
            sw.WriteLine("Descrip. : " & var.sDescription)
            If fait.sRemarque <> "" Then
                sw.WriteLine("Remarque : " & fait.sRemarque)
                sw.WriteLine("")
            End If

VarSuivante2:
        Next de
        sw.WriteLine("")
        sw.WriteLine(sSeparation)

    End Sub

    Private Sub ExporterRegles(sw As StreamWriter, iIdApp%,
    Optional bCompatibleTurboExpert As Boolean = False,
    Optional sApplication$ = "")

        ' Exporter les règles (compatibles Turbo-Expert en option) 

        Dim col As New Specialized.StringCollection()
        Dim bDetailRegles As Boolean = True
        If Not m_bModeBD Then bDetailRegles = False

        If bCompatibleTurboExpert Then
            col.Add("R")
            col.Add("* Application : " & sApplication)
            Const sPrefixe$ = "* "
            ExporterDescrApplication(col, iIdApp, sPrefixe)
        End If

        bTraduireRegles(m_sProvenanceBR, col,
        bCompatibleTurboExpert, bDetailRegles, iIdApp, sApplication)
        Dim sItem$
        For Each sItem In col
            sw.WriteLine(sItem)
        Next sItem

    End Sub

    Friend Sub ExporterPourTurboExpert12(lbApplications As ListBox)

        Dim drw As DataRowView = CType(lbApplications.Items(
        lbApplications.SelectedIndex), DataRowView)
        Dim iIdApp% = CInt(drw.Item(0))         ' Colonne 0
        Dim sApplication$ = CStr(drw.Item(1))   ' Colonne 1

        Dim sFichierAppli$ = sApplication
        Dim iPosCarInvalide% = sFichierAppli.IndexOfAny(
        New Char() {"."c, ","c, "!"c, "?"c, ":"c, ";"c, "/"c, "\"c})
        If iPosCarInvalide > -1 Then
            sFichierAppli = sFichierAppli.Substring(0, iPosCarInvalide)
            If sFichierAppli = "" Then sFichierAppli = "Application1"
        End If
        sFichierAppli = sFichierAppli.Trim

        Dim sChemin$ = Application.StartupPath &
        sRepertoireApplicationsTxt & "\"
        Dim sCheminFichierDic$ = sChemin & sFichierAppli & ".dic"
        Dim sCheminFichierBR$ = sChemin & sFichierAppli & ".brg"
        Dim sCheminFichierBF$ = sChemin & sFichierAppli & ".bfa"

        ' Vérifier et prévenir de l'écrasement des fichiers précédants
        If File.Exists(sCheminFichierDic) Or
        File.Exists(sCheminFichierBR) Or
        File.Exists(sCheminFichierBF) Then _
        If MsgBoxResult.Yes <> MsgBox("Attention, les fichiers :" & vbLf &
            sFichierAppli & ".dic, .brg et .bfa dans :" & vbLf & sChemin & vbLf &
            "vont être écrasés, voulez-vous continuer ?",
            MsgBoxStyle.Question Or MsgBoxStyle.YesNoCancel) Then Exit Sub

        Const bAjouter As Boolean = False
        Dim sw As New StreamWriter(sCheminFichierDic, bAjouter, clsUtil.encodageVB6)
        ExporterDico(sw)
        sw.Close()

        sw = New StreamWriter(sCheminFichierBR, bAjouter, clsUtil.encodageVB6)
        ExporterRegles(sw, bCompatibleTurboExpert:=True,
        iIdApp:=iIdApp, sApplication:=sApplication)
        sw.Close()

        sw = New StreamWriter(sCheminFichierBF, bAjouter, clsUtil.encodageVB6)
        ExporterFaitsInitiaux(sw, m_colSessions, bCompatibleTurboExpert:=True)
        sw.Close()

        MsgBox("Exportation réussie :" & vbLf & sChemin & sFichierAppli &
        ".dic, .brg et .bfa", MsgBoxStyle.Exclamation)

    End Sub

    Private Sub ExporterDico(sw As StreamWriter)

        ' Exporter le dictionnaire des variables (compatible Turbo-Expert) 

        Dim de As DictionaryEntry
        Dim iPasse%
        For iPasse = 0 To 1
            For Each de In m_oDico.m_colDico
                Dim var As clsDico.TVar = CType(de.Value, clsDico.TVar)
                ' D'abord les variables de configuration, puis les autres
                If (iPasse = 0 And m_oDico.bConfig(var.sVariable)) Or
            (iPasse = 1 And Not m_oDico.bConfig(var.sVariable)) Then _
            If Not m_oDico.bIntermediaire(var.sVariable) Then _
                sw.WriteLine(var.sVariable)
            Next de : Next iPasse

        sw.WriteLine(sSeparation)
        ' Variables intermédiaires
        For Each de In m_oDico.m_colDico
            Dim var As clsDico.TVar = CType(de.Value, clsDico.TVar)
            If m_oDico.bIntermediaire(var.sVariable) Then _
            sw.WriteLine(var.sVariable)
        Next
        sw.WriteLine(sFinFichierTurboExpert)

    End Sub

    Private Sub ExporterFaitsInitiaux(sw As StreamWriter,
    colSessions As Collection, bCompatibleTurboExpert As Boolean)

        ' Exporter le tableau des faits initiaux (compatible Turbo-Expert) 

        Dim oSession As ClsSession
        For Each oSession In colSessions

            Dim sLigneFaits$ = ""

            If bCompatibleTurboExpert Then
                sLigneFaits$ = oSession.sSession
                ' Suppression des éventuels ; : ne marche pas !
                'sLigneFaits = sLigneFaits.Trim(New Char() {";"c})
                'Dim trimChar() As Char = New Char() {";"c}
                'sLigneFaits = sLigneFaits.Trim(trimChar)
                If sLigneFaits.IndexOf(";") > -1 Then
                    MsgBox("Le signe ';' n'est pas permis dans le nom de la session" &
                    " pour l'exportation compatible Turbo-Expert 1.2",
                    MsgBoxStyle.Exclamation)
                    Exit Sub
                End If
                sLigneFaits &= ";"
            Else
                sw.WriteLine("")
                sw.WriteLine(sSeparation)
                sw.WriteLine("Session : " & oSession.sSession)
                If oSession.sDescription <> "" Then _
                sw.WriteLine("Descrip.: " & oSession.sDescription)
            End If

            If Not m_oBF.bChargerFaitsInitiauxSession(oSession.colFaits) Then Exit Sub

            Dim de As DictionaryEntry
            Dim iPasse%
            For iPasse = 0 To 1
                For Each de In m_oDico.m_colDico
                    Dim var As clsDico.TVar = CType(de.Value, clsDico.TVar)
                    Dim sVar$ = var.sVariable
                    If m_oDico.bIntermediaire(sVar) Then GoTo VarSuivante
                    ' D'abord les variables de configuration, puis les autres
                    If Not ((iPasse = 0 And m_oDico.bConfig(sVar)) Or
                (iPasse = 1 And Not m_oDico.bConfig(sVar))) Then GoTo VarSuivante

                    Dim sValFait$ = "", sRem$ = "", rFiabFait! = rCodeFiabIndefini
                    Dim sDescrVar$ = var.sDescription
                    If Not m_oBF.bVarExisteDansBF(sVar) Then
                        If Not m_oDico.bConfig(sVar) Then GoTo Valeur
                        sValFait = m_oDico.sValDefVar(sVar)
                        If sValFait = "" Then sValFait = sValConfigDefautModeFichier
                        rFiabFait = m_oDico.rFiabDef(sVar)
                    Else
                        ' L'ordre du dico n'est pas forcément celui des faits
                        Dim prem As clsDico.TPremisse = m_oBF.fait(sVar)
                        sValFait = prem.sValDebut
                        If prem.sDateOrig <> "" Then
                            sValFait = prem.sDateOrig
                        End If
                        If sValFait <> "" Then
                            ' Suppression des guillemets
                            If sValFait.Chars(0) = sGm Then
                                sValFait = sValFait.Substring(1)
                            End If
                            If sValFait.Chars(sValFait.Length - 1) = sGm Then
                                sValFait = sValFait.Substring(0,
                            sValFait.Length - 1)
                            End If
                        End If
                        sRem = prem.sRemarque
                        rFiabFait = prem.rFiab
                    End If

                    If Not bCompatibleTurboExpert Then
                        If sValFait = "" And m_oDico.bConfig(sVar) Then GoTo VarSuivante
                        If sValFait = "" Then sValFait = "?"

                        Dim sFiab$ = ""
                        If rFiabFait <> rCodeFiabIndefini Then _
                    sFiab = " (" & Format(rFiabFait, sFormatFiabRapport) & ")"

                        sw.WriteLine(sVar & " = " & sValFait & sFiab)
                        If sDescrVar <> "" Then _
                    sw.WriteLine("Descrip. : " & sDescrVar)
                        If sRem <> "" Then _
                    sw.WriteLine("Remarque : " & sRem) : sw.WriteLine("")
                    End If

Valeur:
                    sLigneFaits &= sValFait & ";"

VarSuivante:
                Next de : Next iPasse
            If bCompatibleTurboExpert Then sw.WriteLine(sLigneFaits)
        Next oSession
        sw.WriteLine("")

    End Sub

#End Region

#Region "Présentation des tableaux"

    Private Sub FixerStyleTableauDico(ByRef dgVariables As DataGrid,
    bAfficherConst As Boolean)

        ' Modification du style du DataGrid passé en entrée :
        '  ajustement de la largeur des colonnes

        Dim dgts As New DataGridTableStyle()
        Dim dgtbc As New DataGridTextBoxColumn() With {
        .MappingName = sChpVar,
        .Width = 165,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpValDef,
        .Width = 90,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        If bAfficherConst Then
            dgtbc = New DataGridTextBoxColumn() With {
            .MappingName = sChpConst,
            .Width = 90,
            .HeaderText = .MappingName
        }
            dgts.GridColumnStyles.Add(dgtbc)

            dgtbc = New DataGridTextBoxColumn() With {
            .MappingName = sChpFiab,
            .Width = 0,
            .HeaderText = .MappingName
        }
            If m_config.bLogiqueFloue Then dgtbc.Width = 35
            dgts.GridColumnStyles.Add(dgtbc)
        End If

        Dim dgbc As New DataGridBoolColumn() With {
        .MappingName = sChpbConfig,
        .Width = 40,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgbc)

        dgbc = New DataGridBoolColumn() With {
        .MappingName = sChpbConst,
        .Width = 40,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgbc)

        dgbc = New DataGridBoolColumn() With {
        .MappingName = sChpbInterm,
        .Width = 40,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgbc)

        dgVariables.TableStyles.Clear()
        dgVariables.TableStyles.Add(dgts)

    End Sub

    Private Sub FixerStyleTableauRegles(ByRef dgRegles As DataGrid)

        ' Modification du style du DataGrid passé en entrée :
        '  ajustement de la largeur des colonnes

        Dim dgts As New DataGridTableStyle()
        Dim dgtbc As New DataGridTextBoxColumn() With {
        .MappingName = sChpRegle,
        .Width = 55,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpFiab,
        .Width = 0,
        .HeaderText = .MappingName
    }
        If m_config.bLogiqueFloue Then dgtbc.Width = 35
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpVar,
        .Width = 135,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpOp,
        .Width = 25,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpVal,
        .Width = 75,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpVar2,
        .Width = 90,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        Dim dgbc As New DataGridBoolColumn() With {
        .MappingName = sChpbConcl,
        .Width = 40,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgbc)

        dgbc = New DataGridBoolColumn() With {
        .MappingName = sChpbInterm,
        .Width = 40,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgbc)

        dgRegles.TableStyles.Clear()
        dgRegles.TableStyles.Add(dgts)

    End Sub

    Private Sub FixerStyleTableauFaits(ByRef dgFaits As DataGrid,
    bAfficherConstantes As Boolean)

        ' Modification du style du DataGrid passé en entrée :
        '  ajustement de la largeur des colonnes

        Dim dgts As New DataGridTableStyle()
        Dim dgtbc As New DataGridTextBoxColumn() With {
        .MappingName = sChpSession,
        .HeaderText = "Session",
        .Width = 100
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpVar,
        .Width = 180,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpOp,
        .Width = 25,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpVal,
        .Width = 75,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpVar2,
        .Width = 100,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        If bAfficherConstantes Then
            dgtbc = New DataGridTextBoxColumn() With {
            .MappingName = "Constante",
            .Width = 95,
            .HeaderText = .MappingName
        }
            dgts.GridColumnStyles.Add(dgtbc)

            dgtbc = New DataGridTextBoxColumn() With {
            .MappingName = sChpFiab,
            .Width = 0,
            .HeaderText = .MappingName
        }
            If m_config.bLogiqueFloue Then dgtbc.Width = 35
            dgts.GridColumnStyles.Add(dgtbc)
        End If

        dgFaits.TableStyles.Clear()
        dgFaits.TableStyles.Add(dgts)

    End Sub

    Private Sub FixerStyleTableauBilan(ByRef dgBilanSession As DataGrid)

        ' Modification du style du DataGrid passé en entrée :
        '  ajustement de la largeur des colonnes

        Dim dgts As New DataGridTableStyle()
        Dim dgtbc As New DataGridTextBoxColumn() With {
        .MappingName = sChpVar,
        .Width = 110,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpDebut,
        .Width = 100,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpFiabOrig,
        .Width = 0,
        .HeaderText = .MappingName
    }
        If m_config.bLogiqueFloue Then dgtbc.Width = 35
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpFin,
        .Width = 100,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpFiab,
        .Width = 0,
        .HeaderText = .MappingName
    }
        If m_config.bLogiqueFloue Then dgtbc.Width = 35
        dgts.GridColumnStyles.Add(dgtbc)

        dgtbc = New DataGridTextBoxColumn() With {
        .MappingName = sChpRegle,
        .Width = 75,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgtbc)

        Dim dgbc As New DataGridBoolColumn() With {
        .MappingName = sChpbInterm,
        .Width = 40,
        .HeaderText = .MappingName
    }
        dgts.GridColumnStyles.Add(dgbc)

        dgBilanSession.TableStyles.Clear()
        dgBilanSession.TableStyles.Add(dgts)

    End Sub

#End Region

End Class

'End Namespace