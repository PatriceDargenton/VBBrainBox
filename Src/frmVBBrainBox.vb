
' VBBrainBox : un système expert d'ordre 0+ en VB .NET
' ----------------------------------------------------

' -------------------------------------------------------------------
' Créé à partir de TExpert (Turbo-Expert 1.2 en VB6) :
' (c) Philippe LARVET <ph_larvet@yahoo.fr> Avril 1996.
' Prog "one shot" du 28 mai 96 avec paramètre de ligne de commande
' Version VB6 mai 02, revu en tant que moteur TExpert janvier 03
' -------------------------------------------------------------------

' Fichier frmVBBrainBox.vb
' ------------------------

' Conventions de nommage des variables :
' b pour Boolean (booléen vrai ou faux)
' i pour Integer (%) et Short (System.Int16)
' l pour Long : &
' r pour nombre Réel (Single!, Double# ou Decimal : D)
' a pour Array (tableau) : ()
' o pour Object (objet ou classe)
' m_ pour variable Membre de la classe (mais pas pour les constantes)

Public Class frmVBBrainBox : Inherits Form

#Region "Déclarations"

    Private Const sNomFichierRapport$ = "Rapport.txt" ' Rapport d'expertise

    Private m_bInit As Boolean = False ' Attendre l'initialisation des composants
    Private m_bConnexion As Boolean ' Connexion à la base de données
    ' Afficher l'erreur au lancement de l'application le cas échéant
    Private m_bErr As Boolean = False

    Private m_oSE As New clsVBBBox() ' C'est le système expert

    ' Onglet de sélection des applications
    Private Const iPageBD% = 0
    Private Const iPageFichiers% = 1

    ' Onglet de l'expertise
    Private Const iPageVariables% = 0
    Private Const iPageRegles% = 1
    Private Const iPageReglesListe% = 2
    Private Const iPageFaits% = 3
    Private Const iPageExpertise% = 4
    Private Const iPageBilan% = 5
    Private Const iPageArchivage% = 6
    Private Const iPageAPropos% = 7

    ' Niveaux d'initialisation
    Private Const iTypeInitTout% = 0
    Private Const iTypeInitApplication% = 1
    Private Const iTypeInitDico% = 2
    Private Const iTypeInitBR% = 3
    Private Const iTypeInitBF% = 4
    'Private Const iTypeInitSessions% = 5
    Private Const iTypeInitExpertise% = 6

    Private couleurErr As Color = Color.LightCoral
    Private couleurInit As Color = Color.DarkGray
    Private couleurOk As Color = Color.White

#End Region

#Region "Initialisation"

    Private Sub frmTExpert_Load(eventSender As Object, eventArgs As EventArgs) Handles MyBase.Load

        Dim sVersionAppli$ = My.Application.Info.Version.Major &
                "." & My.Application.Info.Version.Minor &
                My.Application.Info.Version.Build
        Dim sTxt$ = sTitreMsg & " - Version " & sVersionAppli & " (" & sDateVersionAppli & ")"
        If bDebug Then sTxt &= " - Debug"
        Me.Text = sTxt

        Initialiser(iTypeInitTout)
        m_bInit = True
        'Me.tcEntrees.SelectedIndex = iPageFichiers ' Pour générer un év. changed
        Me.tcEntrees.SelectedIndex = iPageBD
        MettreAJourApplications()

        Me.lblEnregistrementOcx.Text = ""

        Me.llblMdb.Links.Add(0, Me.llblMdb.Text.Length,
            "file://" & Application.StartupPath &
            clsVBBBox.sRepertoireApplications & "\" & clsVBBBox.sFichierVBBBoxMDB)
        Me.llblMdb.Text = clsVBBBox.sFichierVBBBoxMDB

        If Not m_bErr Then Me.tcExpertises.SelectedIndex = iPageAPropos

    End Sub

    Private Sub Initialiser(iTypeInit%)

        ' Initialisation d'une expertise
        Me.cmdRapport.Enabled = False
        Me.lbFaits.Items.Clear()
        Me.lbFaitsJustes.Items.Clear()
        Me.lbConclusions.Items.Clear()
        Me.lbConclusions.BackColor = couleurInit
        If iTypeInit = iTypeInitExpertise Then Exit Sub

        ' Initialisation des sessions et des fichiers BF
        Me.lbSessions.Enabled = False
        Me.lbSessions.Items.Clear()
        Me.dgFaits.DataSource = Nothing
        Me.dgBilanSession.DataSource = Nothing
        'If iTypeInit = iTypeInitSessions Then Exit Sub
        If iTypeInit = iTypeInitBF Then Exit Sub

        ' Initialisation des fichiers BR
        Me.lbFichiersBF.Enabled = False
        ' Ne pas sélectionner un fichier BF
        If Me.lbFichiersBF.SelectedIndex >= 0 Then _
            Me.lbFichiersBF.SetSelected(Me.lbFichiersBF.SelectedIndex, False)
        Me.lbReglesListe.Items.Clear()
        Me.lbReglesListe.BackColor = couleurInit
        Me.dgRegles.DataSource = Nothing
        If iTypeInit = iTypeInitBR Then Exit Sub

        ' Initialisation des dictionnaires
        Me.lbFichiersBR.Enabled = False
        ' Ne pas sélectionner un fichier BR
        If Me.lbFichiersBR.SelectedIndex >= 0 Then _
            Me.lbFichiersBR.SetSelected(Me.lbFichiersBR.SelectedIndex, False)
        Me.dgVariables.DataSource = Nothing
        Me.dgVariables.BackColor = couleurInit
        If iTypeInit = iTypeInitDico Then Exit Sub

        ' Initialisation d'une application
        ' Ne pas sélectionner un fichier dico
        If Me.lbFichiersDico.SelectedIndex >= 0 Then _
            Me.lbFichiersDico.SetSelected(Me.lbFichiersDico.SelectedIndex, False)
        m_oSE.InitialiserApplication()
        If iTypeInit = iTypeInitApplication Then Exit Sub

        ' Initialisation des applications : de tout !
        Me.cmdExporter.Enabled = False
        ' Ne pas sélectionner une application
        If Me.lbApplications.SelectedIndex >= 0 Then _
            Me.lbApplications.SetSelected(Me.lbApplications.SelectedIndex, False)
        'If iTypeInit = iTypeInitTout Then Exit Sub

    End Sub

#End Region

#Region "Gestion du mode base de données ou fichier"

    Private Sub tcEntrees_SelectedIndexChanged(sender As Object,
                e As EventArgs) Handles tcEntrees.SelectedIndexChanged

        MettreAJourApplications()

    End Sub

    Private Sub MettreAJourApplications()

        m_bConnexion = False
        If Not m_bInit Then Exit Sub
        Initialiser(iTypeInitTout)

        If Me.tcEntrees.SelectedIndex = iPageFichiers Then

            m_oSE.RemplirListesFichiers(Me.lbFichiersDico, Me.lbFichiersBR, Me.lbFichiersBF)

        Else 'If Me.tcEntrees.SelectedIndex = iPageBD Then 

            Cursor.Current = Cursors.WaitCursor

            If Not m_oSE.bBDVerifierVersion() Then AfficherErreurs() : GoTo Fin
            ' Page base de données selectionnée
            If Not m_oSE.bBDRemplirApplications(Me.lbApplications) Then _
                    AfficherErreurs() : GoTo Fin
            m_bConnexion = True

Fin:
            Cursor.Current = Cursors.Default

        End If

    End Sub

#End Region

#Region "Sélection d'une application"

    Private Sub lbApplications_SelectedIndexChanged(sender As Object,
                e As EventArgs) Handles lbApplications.SelectedIndexChanged

        If Not m_bInit Then Exit Sub
        If Not m_bConnexion Then Exit Sub

        Cursor.Current = Cursors.WaitCursor

        Initialiser(iTypeInitApplication)

        If Me.tcExpertises.SelectedIndex >= iPageExpertise Then
            Me.tcExpertises.SelectedIndex = iPageVariables
        End If

        ' Sélection d'une application
        If Me.lbApplications.SelectedIndex < 0 Then GoTo Fin
        ' Récupération de l'IdApplication
        Dim obj As Object = Me.lbApplications.Items(Me.lbApplications.SelectedIndex)
        Dim iIdApp% = CInt(CType(obj, DataRowView).Item(0))

        If Not m_oSE.bBDChargerDico(iIdApp, Me.dgVariables) Then AfficherErreurs() : GoTo Fin
        Me.dgVariables.BackColor = couleurOk

        If Not m_oSE.bBDRemplirRegles(iIdApp, Me.dgRegles, Me.lbReglesListe) Then
            AfficherErreurs() : GoTo Fin
        End If
        Me.lbReglesListe.BackColor = couleurOk

        If Not m_oSE.bBDRemplirFaits(iIdApp, Me.dgFaits, Me.lbSessions) Then
            AfficherErreurs() : GoTo Fin
        End If
        Me.lbSessions.Enabled = True

        Me.cmdExporter.Enabled = True
        Me.cmdRapport.Enabled = True ' Activation pour faire le rapport des sessions
        m_bErr = False

Fin:
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub cmdExporter_Click(sender As Object,
                e As EventArgs) Handles cmdExporter.Click

        m_oSE.ExporterPourTurboExpert12(Me.lbApplications)

    End Sub

#End Region

#Region "Sélection d'un fichier"

    Private Sub lbFichiersDico_SelectedIndexChanged(sender As Object,
                e As EventArgs) Handles lbFichiersDico.SelectedIndexChanged

        ' Dico des variables de la BR

        If Me.lbFichiersDico.SelectedIndex < 0 Then Exit Sub

        ' Répertoire de l'application : 
        Dim sCheminFichier$ = Application.StartupPath &
            clsVBBBox.sRepertoireApplicationsTxt & "\" &
            Me.lbFichiersDico.Items(Me.lbFichiersDico.SelectedIndex).ToString

        Initialiser(iTypeInitDico)

        Me.tcExpertises.SelectedIndex = iPageVariables
        If Not m_oSE.bChargerDico(sCheminFichier, Me.dgVariables) Then _
            AfficherErreurs() : Exit Sub
        Me.dgVariables.BackColor = couleurOk
        Me.lbFichiersBR.Enabled = True

        ' Essayer de sélectionner automatiquement la BR correspondante
        '  (Ok s'il y a le même nombre de fichiers de chaque type)
        If Me.lbFichiersBR.Items.Count >= Me.lbFichiersDico.SelectedIndex Then _
                Me.lbFichiersBR.SelectedIndex = Me.lbFichiersDico.SelectedIndex

    End Sub

    Private Sub lbFichiersBR_SelectedIndexChanged(sender As Object,
                e As EventArgs) Handles lbFichiersBR.SelectedIndexChanged

        If Me.lbFichiersBR.SelectedIndex < 0 Then Exit Sub

        Dim sCheminFichier$ = Application.StartupPath &
            clsVBBBox.sRepertoireApplicationsTxt & "\" &
            Me.lbFichiersBR.Items(Me.lbFichiersBR.SelectedIndex).ToString

        Initialiser(iTypeInitBR)

        If Not (Me.tcExpertises.SelectedIndex = iPageRegles Or
                Me.tcExpertises.SelectedIndex = iPageReglesListe) Then _
                Me.tcExpertises.SelectedIndex = iPageReglesListe

        If Not m_oSE.bChargerBR(sCheminFichier, Me.lbReglesListe, Me.dgRegles) Then _
                AfficherErreurs() : Exit Sub
        Me.lbReglesListe.BackColor = couleurOk
        Me.lbFichiersBF.Enabled = True

        ' Essayer de sélectionner automatiquement la BF correspondante
        If Me.lbFichiersBF.Items.Count >= Me.lbFichiersBR.SelectedIndex Then _
                Me.lbFichiersBF.SelectedIndex = Me.lbFichiersBR.SelectedIndex

    End Sub

    Private Sub lbFichiersBF_SelectedIndexChanged(sender As Object,
                e As EventArgs) Handles lbFichiersBF.SelectedIndexChanged

        If Me.lbFichiersBF.SelectedIndex < 0 Then Exit Sub

        Dim sFichierBF$ = Me.lbFichiersBF.Items(Me.lbFichiersBF.SelectedIndex).ToString

        Initialiser(iTypeInitBF)

        Me.tcExpertises.SelectedIndex = iPageFaits
        If Not m_oSE.bRemplirSessions(sFichierBF, Me.lbSessions, Me.dgFaits) Then _
                AfficherErreurs() : Exit Sub
        Me.lbSessions.Enabled = True

    End Sub

#End Region

#Region "Expertise"

    Private Sub lbSessions_SelectedIndexChanged(eventSender As Object,
                eventArgs As EventArgs) Handles lbSessions.SelectedIndexChanged
        Expertiser()
    End Sub

    Private Sub Expertiser()

        If Me.tcExpertises.SelectedIndex <> iPageBilan Then _
                Me.tcExpertises.SelectedIndex = iPageExpertise

        Initialiser(iTypeInitExpertise)
        Dim sApplication$ = sApplicationSelectionnee()
        Dim sSession$ = Me.lbSessions.Items(Me.lbSessions.SelectedIndex).ToString
        Dim bConclusions, bAvertissements As Boolean
        If Not m_oSE.bExpertiser(sApplication, sSession,
                Me.lbFaits, Me.lbFaitsJustes, Me.dgBilanSession,
                bConclusions, bAvertissements) Then _
                AfficherErreurs() : Exit Sub
        If bConclusions Then
            Me.lbConclusions.BackColor = Color.Cyan
            If bAvertissements Then Me.lbConclusions.BackColor = Color.Beige
        End If
        AfficherConclusions()
        Me.cmdRapport.Enabled = True ' Tjrs actif pour aff les msg d'err

    End Sub

    Private Sub cmdRapport_Click(eventSender As Object,
                eventArgs As EventArgs) Handles cmdRapport.Click

        Dim iIdApplication% = 0, sApplication$ = "", sSession$ = ""
        If Me.lbApplications.SelectedIndex >= 0 Then
            Dim drw As DataRowView = CType(Me.lbApplications.Items(
                    Me.lbApplications.SelectedIndex), DataRowView)
            iIdApplication = CInt(drw.Item(0)) ' Colonne 0
            sApplication = sApplicationSelectionnee()
        End If
        If Me.lbSessions.SelectedIndex >= 0 Then _
                sSession = Me.lbSessions.Items(Me.lbSessions.SelectedIndex).ToString()
        m_oSE.CreerCompteRendu(sNomFichierRapport, iIdApplication, sApplication, sSession)

    End Sub

    Private Function sApplicationSelectionnee$()

        Dim sApplication$ = ""
        If Me.tcEntrees.SelectedIndex = iPageBD Then
            If Me.lbApplications.SelectedIndex >= 0 Then
                Dim drw As DataRowView = CType(Me.lbApplications.Items(
                        Me.lbApplications.SelectedIndex), DataRowView)
                sApplication = CStr(drw.Item(1)) ' Colonne 1
            End If
        Else
            If Me.lbFichiersDico.SelectedIndex >= 0 Then _
                    sApplication = Me.lbFichiersDico.Items(
                        Me.lbFichiersDico.SelectedIndex).ToString()
        End If
        sApplicationSelectionnee = sApplication

    End Function

    Private Sub chkFaitsJustes_CheckedChanged(sender As Object,
                e As EventArgs) Handles chkFaitsJustes.CheckedChanged
        If Me.chkFaitsJustes.Checked Then
            Me.lbFaits.Visible = False
            Me.lbFaitsJustes.Visible = True
        Else
            Me.lbFaits.Visible = True
            Me.lbFaitsJustes.Visible = False
        End If
    End Sub

    Private Sub AfficherErreurs()
        ' Afficher les messages d'erreur
        Me.lbConclusions.BackColor = couleurErr
        m_bErr = True
        Me.tcExpertises.SelectedIndex = iPageExpertise
        AfficherConclusions()
        Me.cmdRapport.Enabled = True ' Pour imprimer les msg d'err
    End Sub

    Private Sub AfficherConclusions()
        Dim sItem$
        Me.lbConclusions.Items.Clear()
        Me.lbConclusions.BeginUpdate()
        For Each sItem In m_oSE.colLireMessages
            Me.lbConclusions.Items.Add(sItem)
        Next sItem
        Me.lbConclusions.EndUpdate()
    End Sub

#End Region

#Region "Divers"

    Private Sub cmdAPropos_Click(sender As Object, e As EventArgs) Handles cmdAPropos.Click

        Dim sVersionAppli$ = My.Application.Info.Version.Major &
            "." & My.Application.Info.Version.Minor &
            My.Application.Info.Version.Build
        Dim sMsg$
        sMsg = "VBBrainBox " & sVersionAppli & " (" & sDateVersionAppli & ")" & vbLf & vbLf
        sMsg &= "d'après Turbo-Expert 1.2 pour Windows" & vbLf
        sMsg &= "(c) Philippe Larvet 1996, 2003" & vbLf & vbLf
        sMsg &= "Documentation : VBBrainBox.html et README.md" & vbLf & vbLf
        sMsg &= "Base de données : " & vbLf & Application.StartupPath &
            clsVBBBox.sRepertoireApplications & "\" & clsVBBBox.sFichierVBBBoxMDB
        MsgBox(sMsg, MsgBoxStyle.Information)

    End Sub

    Private Sub llblTous_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) _
                Handles llblMdb.LinkClicked

        On Error Resume Next ' Exemple d'erreur : Fichier introuvable
        Process.Start(e.Link.LinkData.ToString)

    End Sub

    Private Sub tcExpertises_SelectedIndexChanged(sender As Object,
                e As EventArgs) Handles tcExpertises.SelectedIndexChanged

        If Me.tcExpertises.SelectedIndex < iPageArchivage Then Exit Sub

        ' Page A Propos
        GestionDBToFile()

    End Sub

#End Region

#Region "Archivage"

    Private Sub ActivationArchivage(bActiver As Boolean)

        If Not bActiver Then
            Me.cmdArchiverAppli.Enabled = False
            Me.cmdArchiverTout.Enabled = False
            Me.cmdChargerAppli.Enabled = False
            Me.cmdChargerTout.Enabled = False
            Exit Sub
        End If

        Me.cmdArchiverTout.Enabled = True
        Me.cmdChargerTout.Enabled = True
        Me.cmdChargerAppli.Enabled = True

        If Me.lbApplications.SelectedIndex >= 0 Then
            Me.cmdArchiverAppli.Enabled = True
        Else
            Me.cmdArchiverAppli.Enabled = False
        End If

    End Sub

    Private Sub GestionDBToFile()

        ' Le composant doit exister
        Dim sCheminOcx$ = Application.StartupPath &
            clsVBBBox.sRepertoireApplications & "\" & sFichierDBToFile
        Me.lblEnregistrementOcx.Text = ""
        If Not IO.File.Exists(sCheminOcx) Then
            Me.cmdDBToFile.Enabled = False
            ActivationArchivage(bActiver:=False)
            Me.lblEnregistrementOcx.Text = "Impossible de trouver le fichier " &
                clsVBBBox.sRepertoireApplications & "\" & sFichierDBToFile
            ' Le composant peut être enregistré ailleurs
            'Exit Sub
        End If

        ' L'application doit être compilée en 32 bits, pour que l'enregistrement du composant fonctionne
        If clsUtil.bCompilation64Bit() Then
            Me.cmdDBToFile.Enabled = False
            ActivationArchivage(bActiver:=False)
            Me.lblEnregistrementOcx.Text = "Le contrôle " & sFichierDBToFile &
                " ne peut être enregistré qu'en mode x86 - 32 bits"
            Me.cmdDBToFile.Enabled = False
            ActivationArchivage(bActiver:=False)
            Exit Sub
        End If

        If Not clsUtil.bEstAdmin() Then
            Me.cmdDBToFile.Enabled = False
            ActivationArchivage(bActiver:=False)
            Me.lblEnregistrementOcx.Text = "Le contrôle " & sFichierDBToFile &
                " ne peut être inscrit ou désinscrit qu'en mode admin" & vbLf &
                "(Exécuter VBBainBox.exe en tant qu'administateur)"
            'Exit Sub
        End If

        If clsUtil.bCleRegistreExiste(sCleRegistreDBToFile) Then
            Me.lblArchivage.Text = "Le contrôle " & sFichierDBToFile & " est inscrit"
            Me.cmdDBToFile.Text = "Désinscrire " & sFichierDBToFile
            ' Il se peut aussi que le composant ait été déplacé depuis
            ' S'il y a une erreur lors de l'archivage alors il faudra ré-enregistrer l'ocx
            ActivationArchivage(bActiver:=True)
        Else
            Me.lblArchivage.Text = "Le contrôle " & sFichierDBToFile & " n'est pas inscrit"
            Me.cmdDBToFile.Text = "Inscrire " & sFichierDBToFile
            'Me.cmdDBToFile.Enabled = True
            ActivationArchivage(bActiver:=False)
        End If

        ' Dans tous les cas, laisser la possibilité de désinscrire le composant
        Me.cmdDBToFile.Enabled = True

    End Sub

    Private Sub cmdDBToFile_Click(sender As Object, e As EventArgs)

        ' L'application doit être compilée en 32 bits pour pouvoir inscrire le composant
        If clsUtil.bCompilation64Bit() Then
            MsgBox("Il faut compiler VBBrainBox en mode x86 - 32 bits pour inscrire ou désinscrire le composant DBToFile.ocx",
                MsgBoxStyle.Exclamation)
            Me.cmdDBToFile.Enabled = False
            Exit Sub
        End If

        ' L'utilisateur doit être admin pour pouvoir inscrire le composant
        If Not clsUtil.bEstAdmin() Then
            MsgBox("Il faut lancer VBBrainBox en mode admin pour inscrire ou désinscrire le composant DBToFile.ocx",
                MsgBoxStyle.Exclamation)
            Me.cmdDBToFile.Enabled = False
            Exit Sub
        End If

        Dim sRepertoireDll$ = Application.StartupPath & clsVBBBox.sRepertoireApplications
        If clsUtil.bCleRegistreExiste(sCleRegistreDBToFile) Then
            clsUtil.bEnregistrerDllActiveX(sFichierDBToFile, sRepertoireDll, bDesenregistrer:=True)
        Else
            clsUtil.bEnregistrerDllActiveX(sFichierDBToFile, sRepertoireDll)
        End If
        GestionDBToFile()

    End Sub

    Private Sub cmdArchiverAppli_Click(sender As Object, e As EventArgs) Handles cmdArchiverAppli.Click

        ' Récupération de l'IdApplication
        Dim obj As Object = Me.lbApplications.Items(Me.lbApplications.SelectedIndex)
        Dim sIdApplication$ = CStr(CType(obj, DataRowView).Item(0))
        Dim sApp$ = Me.lbApplications.Text
        If Not CreerObjetDBToFile(bSauver:=True, sIdApp:=sIdApplication, sApp:=sApp) Then
            'Me.cmdDBToFile.Enabled = True
        End If

    End Sub

    Private Sub cmdChargerAppli_Click(sender As Object, e As EventArgs) Handles cmdChargerAppli.Click

        Dim sInitDir$ = Application.StartupPath & clsVBBBox.sRepertoireApplications
        Dim sCheminFichierBba$ = ""
        If clsUtil.bChoisirFichier(sCheminFichierBba, sFiltreBBA, sExtBBA,
            "Charger une application VBBrainBox (.bba)", sInitDir:=sInitDir) Then
            If Not CreerObjetDBToFile(bSauver:=False, sCheminApp:=sCheminFichierBba) Then
                'Me.cmdDBToFile.Enabled = True
            Else
                MettreAJourApplications()
            End If
        End If

    End Sub

    Private Sub cmdArchiverTout_Click(sender As Object, e As EventArgs) Handles cmdArchiverTout.Click

        If Not CreerObjetDBToFile(bSauver:=True) Then
            'Me.cmdDBToFile.Enabled = True
        End If

    End Sub

    Private Sub cmdChargerTout_Click(sender As Object, e As EventArgs) Handles cmdChargerTout.Click

        Dim bVider As Boolean = False
        If Me.chkViderBaseMdb.Checked Then
            If MsgBoxResult.Yes <> MsgBox(
                "Êtes-vous sûr de vouloir vider entièrement la base de données avant de tout recharger ?",
                MsgBoxStyle.YesNoCancel) Then Exit Sub
            bVider = True
        End If

        Dim sInitDir$ = Application.StartupPath & clsVBBBox.sRepertoireApplications
        Dim sCheminFichierBba$ = sInitDir & "\" & sFichierApplicationsParDefaut
        If clsUtil.bChoisirFichier(sCheminFichierBba, sFiltreBBA, sExtBBA,
            "Charger une application VBBrainBox (.bba)", sInitDir:=sInitDir) Then
            If Not CreerObjetDBToFile(bSauver:=False, bVider:=bVider, sCheminApp:=sCheminFichierBba) Then
                'Me.cmdDBToFile.Enabled = True
            Else
                ' Si on efface tout, alors remettre la version
                If Me.chkViderBaseMdb.Checked Then m_oSE.bBDDefinirVersion()
                MettreAJourApplications()
            End If
        End If

    End Sub

#End Region

End Class