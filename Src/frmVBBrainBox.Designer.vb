<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmVBBrainBox
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then components.Dispose()
            If Not m_oSE Is Nothing Then m_oSE.Dispose()
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents tcEntrees As System.Windows.Forms.TabControl
    Friend WithEvents tpBD As System.Windows.Forms.TabPage
    Friend WithEvents lbApplications As System.Windows.Forms.ListBox
    Friend WithEvents lbFichiersDico As System.Windows.Forms.ListBox
    Friend WithEvents tpFichiers As System.Windows.Forms.TabPage
    Friend WithEvents lbFichiersBR As System.Windows.Forms.ListBox
    Friend WithEvents lbFichiersBF As System.Windows.Forms.ListBox
    Friend WithEvents tcExpertises As System.Windows.Forms.TabControl
    Friend WithEvents tpRegles As System.Windows.Forms.TabPage
    Friend WithEvents tpExpertises As System.Windows.Forms.TabPage
    Friend WithEvents tpBilanSession As System.Windows.Forms.TabPage
    Friend WithEvents dgBilanSession As System.Windows.Forms.DataGrid
    Friend WithEvents lbFaits As System.Windows.Forms.ListBox
    Friend WithEvents lbConclusions As System.Windows.Forms.ListBox
    Friend WithEvents lbFaitsJustes As System.Windows.Forms.ListBox
    Friend WithEvents tpVariables As System.Windows.Forms.TabPage
    Friend WithEvents lbReglesListe As System.Windows.Forms.ListBox
    Friend WithEvents tpReglesListe As System.Windows.Forms.TabPage
    Friend WithEvents tpFaits As System.Windows.Forms.TabPage
    Friend WithEvents dgFaits As System.Windows.Forms.DataGrid
    Friend WithEvents dgRegles As System.Windows.Forms.DataGrid
    Friend WithEvents chkFaitsJustes As System.Windows.Forms.CheckBox
    Friend WithEvents lblFaitsInitiaux As System.Windows.Forms.Label
    Friend WithEvents lblExpertise As System.Windows.Forms.Label
    Friend WithEvents lblSessions As System.Windows.Forms.Label
    Friend WithEvents lbSessions As System.Windows.Forms.ListBox
    Friend WithEvents cmdRapport As System.Windows.Forms.Button
    Friend WithEvents lblDico As System.Windows.Forms.Label
    Friend WithEvents lblBaseFaits As System.Windows.Forms.Label
    Friend WithEvents lblRegles As System.Windows.Forms.Label
    Friend WithEvents dgVariables As System.Windows.Forms.DataGrid
    Friend WithEvents cmdExporter As System.Windows.Forms.Button
    Friend WithEvents tpAPropos As System.Windows.Forms.TabPage
    Friend WithEvents cmdAPropos As System.Windows.Forms.Button
    Friend WithEvents lblVBBrainBox As System.Windows.Forms.Label
    Friend WithEvents llblMdb As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVBBrainBox))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkFaitsJustes = New System.Windows.Forms.CheckBox()
        Me.lbSessions = New System.Windows.Forms.ListBox()
        Me.tcEntrees = New System.Windows.Forms.TabControl()
        Me.tpBD = New System.Windows.Forms.TabPage()
        Me.cmdExporter = New System.Windows.Forms.Button()
        Me.lbApplications = New System.Windows.Forms.ListBox()
        Me.tpFichiers = New System.Windows.Forms.TabPage()
        Me.lblRegles = New System.Windows.Forms.Label()
        Me.lbFichiersBR = New System.Windows.Forms.ListBox()
        Me.lbFichiersBF = New System.Windows.Forms.ListBox()
        Me.lblBaseFaits = New System.Windows.Forms.Label()
        Me.lblDico = New System.Windows.Forms.Label()
        Me.lbFichiersDico = New System.Windows.Forms.ListBox()
        Me.tcExpertises = New System.Windows.Forms.TabControl()
        Me.tpVariables = New System.Windows.Forms.TabPage()
        Me.dgVariables = New System.Windows.Forms.DataGrid()
        Me.tpRegles = New System.Windows.Forms.TabPage()
        Me.dgRegles = New System.Windows.Forms.DataGrid()
        Me.tpReglesListe = New System.Windows.Forms.TabPage()
        Me.lbReglesListe = New System.Windows.Forms.ListBox()
        Me.tpFaits = New System.Windows.Forms.TabPage()
        Me.dgFaits = New System.Windows.Forms.DataGrid()
        Me.tpExpertises = New System.Windows.Forms.TabPage()
        Me.lbFaitsJustes = New System.Windows.Forms.ListBox()
        Me.lbFaits = New System.Windows.Forms.ListBox()
        Me.cmdRapport = New System.Windows.Forms.Button()
        Me.lbConclusions = New System.Windows.Forms.ListBox()
        Me.lblExpertise = New System.Windows.Forms.Label()
        Me.lblFaitsInitiaux = New System.Windows.Forms.Label()
        Me.tpBilanSession = New System.Windows.Forms.TabPage()
        Me.dgBilanSession = New System.Windows.Forms.DataGrid()
        Me.tpArchivage = New System.Windows.Forms.TabPage()
        Me.chkViderBaseMdb = New System.Windows.Forms.CheckBox()
        Me.cmdChargerAppli = New System.Windows.Forms.Button()
        Me.cmdChargerTout = New System.Windows.Forms.Button()
        Me.cmdArchiverAppli = New System.Windows.Forms.Button()
        Me.cmdArchiverTout = New System.Windows.Forms.Button()
        Me.tpAPropos = New System.Windows.Forms.TabPage()
        Me.llblMdb = New System.Windows.Forms.LinkLabel()
        Me.lblVBBrainBox = New System.Windows.Forms.Label()
        Me.cmdAPropos = New System.Windows.Forms.Button()
        Me.lblSessions = New System.Windows.Forms.Label()
        Me.lblEnregistrementOcx = New System.Windows.Forms.Label()
        Me.lblInfoDBToFile = New System.Windows.Forms.Label()
        Me.cmdDBToFile = New System.Windows.Forms.Button()
        Me.lblArchivage = New System.Windows.Forms.Label()
        Me.tcEntrees.SuspendLayout()
        Me.tpBD.SuspendLayout()
        Me.tpFichiers.SuspendLayout()
        Me.tcExpertises.SuspendLayout()
        Me.tpVariables.SuspendLayout()
        CType(Me.dgVariables, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpRegles.SuspendLayout()
        CType(Me.dgRegles, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpReglesListe.SuspendLayout()
        Me.tpFaits.SuspendLayout()
        CType(Me.dgFaits, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpExpertises.SuspendLayout()
        Me.tpBilanSession.SuspendLayout()
        CType(Me.dgBilanSession, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpArchivage.SuspendLayout()
        Me.tpAPropos.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkFaitsJustes
        '
        Me.chkFaitsJustes.Checked = True
        Me.chkFaitsJustes.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFaitsJustes.Location = New System.Drawing.Point(112, 16)
        Me.chkFaitsJustes.Name = "chkFaitsJustes"
        Me.chkFaitsJustes.Size = New System.Drawing.Size(56, 16)
        Me.chkFaitsJustes.TabIndex = 19
        Me.chkFaitsJustes.Text = "Vrais"
        Me.ToolTip1.SetToolTip(Me.chkFaitsJustes, "Afficher seulement les faits initiaux Vrais ou bien définis autre que Faux")
        '
        'lbSessions
        '
        Me.lbSessions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbSessions.BackColor = System.Drawing.SystemColors.Window
        Me.lbSessions.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbSessions.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbSessions.Enabled = False
        Me.lbSessions.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbSessions.Location = New System.Drawing.Point(8, 280)
        Me.lbSessions.Name = "lbSessions"
        Me.lbSessions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbSessions.Size = New System.Drawing.Size(168, 210)
        Me.lbSessions.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.lbSessions, "Liste des sessions expertisables")
        '
        'tcEntrees
        '
        Me.tcEntrees.Controls.Add(Me.tpBD)
        Me.tcEntrees.Controls.Add(Me.tpFichiers)
        Me.tcEntrees.Location = New System.Drawing.Point(8, 8)
        Me.tcEntrees.Name = "tcEntrees"
        Me.tcEntrees.SelectedIndex = 0
        Me.tcEntrees.Size = New System.Drawing.Size(168, 248)
        Me.tcEntrees.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.tcEntrees, "Sélection de la source des applications")
        '
        'tpBD
        '
        Me.tpBD.Controls.Add(Me.cmdExporter)
        Me.tpBD.Controls.Add(Me.lbApplications)
        Me.tpBD.Location = New System.Drawing.Point(4, 22)
        Me.tpBD.Name = "tpBD"
        Me.tpBD.Size = New System.Drawing.Size(160, 222)
        Me.tpBD.TabIndex = 1
        Me.tpBD.Text = "Base de données"
        '
        'cmdExporter
        '
        Me.cmdExporter.Enabled = False
        Me.cmdExporter.Location = New System.Drawing.Point(88, 192)
        Me.cmdExporter.Name = "cmdExporter"
        Me.cmdExporter.Size = New System.Drawing.Size(64, 24)
        Me.cmdExporter.TabIndex = 23
        Me.cmdExporter.Text = "Exporter"
        Me.ToolTip1.SetToolTip(Me.cmdExporter, "Exporter l'application sélectionnée en fichiers Turbo-Expert 1.2")
        '
        'lbApplications
        '
        Me.lbApplications.Location = New System.Drawing.Point(8, 8)
        Me.lbApplications.Name = "lbApplications"
        Me.lbApplications.Size = New System.Drawing.Size(144, 173)
        Me.lbApplications.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.lbApplications, "Liste des applications de la base de données VBBrainBox.mdb")
        '
        'tpFichiers
        '
        Me.tpFichiers.Controls.Add(Me.lblRegles)
        Me.tpFichiers.Controls.Add(Me.lbFichiersBR)
        Me.tpFichiers.Controls.Add(Me.lbFichiersBF)
        Me.tpFichiers.Controls.Add(Me.lblBaseFaits)
        Me.tpFichiers.Controls.Add(Me.lblDico)
        Me.tpFichiers.Controls.Add(Me.lbFichiersDico)
        Me.tpFichiers.Location = New System.Drawing.Point(4, 22)
        Me.tpFichiers.Name = "tpFichiers"
        Me.tpFichiers.Size = New System.Drawing.Size(160, 222)
        Me.tpFichiers.TabIndex = 0
        Me.tpFichiers.Text = "Fichiers"
        '
        'lblRegles
        '
        Me.lblRegles.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblRegles.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRegles.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblRegles.Location = New System.Drawing.Point(8, 80)
        Me.lblRegles.Name = "lblRegles"
        Me.lblRegles.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRegles.Size = New System.Drawing.Size(137, 16)
        Me.lblRegles.TabIndex = 28
        Me.lblRegles.Text = "Base de règles :"
        '
        'lbFichiersBR
        '
        Me.lbFichiersBR.Enabled = False
        Me.lbFichiersBR.Location = New System.Drawing.Point(8, 96)
        Me.lbFichiersBR.Name = "lbFichiersBR"
        Me.lbFichiersBR.Size = New System.Drawing.Size(144, 43)
        Me.lbFichiersBR.Sorted = True
        Me.lbFichiersBR.TabIndex = 27
        '
        'lbFichiersBF
        '
        Me.lbFichiersBF.Enabled = False
        Me.lbFichiersBF.Location = New System.Drawing.Point(8, 168)
        Me.lbFichiersBF.Name = "lbFichiersBF"
        Me.lbFichiersBF.Size = New System.Drawing.Size(144, 43)
        Me.lbFichiersBF.Sorted = True
        Me.lbFichiersBF.TabIndex = 26
        '
        'lblBaseFaits
        '
        Me.lblBaseFaits.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblBaseFaits.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBaseFaits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblBaseFaits.Location = New System.Drawing.Point(8, 152)
        Me.lblBaseFaits.Name = "lblBaseFaits"
        Me.lblBaseFaits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBaseFaits.Size = New System.Drawing.Size(137, 16)
        Me.lblBaseFaits.TabIndex = 25
        Me.lblBaseFaits.Text = "Base de faits :"
        '
        'lblDico
        '
        Me.lblDico.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblDico.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDico.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDico.Location = New System.Drawing.Point(8, 8)
        Me.lblDico.Name = "lblDico"
        Me.lblDico.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDico.Size = New System.Drawing.Size(137, 16)
        Me.lblDico.TabIndex = 24
        Me.lblDico.Text = "Dico des variables :"
        '
        'lbFichiersDico
        '
        Me.lbFichiersDico.Location = New System.Drawing.Point(8, 24)
        Me.lbFichiersDico.Name = "lbFichiersDico"
        Me.lbFichiersDico.Size = New System.Drawing.Size(144, 43)
        Me.lbFichiersDico.Sorted = True
        Me.lbFichiersDico.TabIndex = 23
        '
        'tcExpertises
        '
        Me.tcExpertises.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tcExpertises.Controls.Add(Me.tpVariables)
        Me.tcExpertises.Controls.Add(Me.tpRegles)
        Me.tcExpertises.Controls.Add(Me.tpReglesListe)
        Me.tcExpertises.Controls.Add(Me.tpFaits)
        Me.tcExpertises.Controls.Add(Me.tpExpertises)
        Me.tcExpertises.Controls.Add(Me.tpBilanSession)
        Me.tcExpertises.Controls.Add(Me.tpArchivage)
        Me.tcExpertises.Controls.Add(Me.tpAPropos)
        Me.tcExpertises.Location = New System.Drawing.Point(184, 8)
        Me.tcExpertises.Name = "tcExpertises"
        Me.tcExpertises.SelectedIndex = 0
        Me.tcExpertises.Size = New System.Drawing.Size(582, 490)
        Me.tcExpertises.TabIndex = 28
        Me.ToolTip1.SetToolTip(Me.tcExpertises, "Présentation des données de l'application")
        '
        'tpVariables
        '
        Me.tpVariables.Controls.Add(Me.dgVariables)
        Me.tpVariables.Location = New System.Drawing.Point(4, 22)
        Me.tpVariables.Name = "tpVariables"
        Me.tpVariables.Size = New System.Drawing.Size(574, 464)
        Me.tpVariables.TabIndex = 4
        Me.tpVariables.Text = "Variables"
        '
        'dgVariables
        '
        Me.dgVariables.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgVariables.DataMember = ""
        Me.dgVariables.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgVariables.Location = New System.Drawing.Point(8, 8)
        Me.dgVariables.Name = "dgVariables"
        Me.dgVariables.ReadOnly = True
        Me.dgVariables.Size = New System.Drawing.Size(558, 450)
        Me.dgVariables.TabIndex = 0
        '
        'tpRegles
        '
        Me.tpRegles.Controls.Add(Me.dgRegles)
        Me.tpRegles.Location = New System.Drawing.Point(4, 22)
        Me.tpRegles.Name = "tpRegles"
        Me.tpRegles.Size = New System.Drawing.Size(574, 464)
        Me.tpRegles.TabIndex = 1
        Me.tpRegles.Text = "Règles (tableau)"
        '
        'dgRegles
        '
        Me.dgRegles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgRegles.DataMember = ""
        Me.dgRegles.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgRegles.Location = New System.Drawing.Point(8, 8)
        Me.dgRegles.Name = "dgRegles"
        Me.dgRegles.ReadOnly = True
        Me.dgRegles.Size = New System.Drawing.Size(558, 450)
        Me.dgRegles.TabIndex = 0
        '
        'tpReglesListe
        '
        Me.tpReglesListe.Controls.Add(Me.lbReglesListe)
        Me.tpReglesListe.Location = New System.Drawing.Point(4, 22)
        Me.tpReglesListe.Name = "tpReglesListe"
        Me.tpReglesListe.Size = New System.Drawing.Size(574, 464)
        Me.tpReglesListe.TabIndex = 3
        Me.tpReglesListe.Text = "Règles (liste)"
        '
        'lbReglesListe
        '
        Me.lbReglesListe.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbReglesListe.Location = New System.Drawing.Point(8, 8)
        Me.lbReglesListe.Name = "lbReglesListe"
        Me.lbReglesListe.Size = New System.Drawing.Size(558, 433)
        Me.lbReglesListe.TabIndex = 0
        '
        'tpFaits
        '
        Me.tpFaits.Controls.Add(Me.dgFaits)
        Me.tpFaits.Location = New System.Drawing.Point(4, 22)
        Me.tpFaits.Name = "tpFaits"
        Me.tpFaits.Size = New System.Drawing.Size(574, 464)
        Me.tpFaits.TabIndex = 0
        Me.tpFaits.Text = "Faits"
        '
        'dgFaits
        '
        Me.dgFaits.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgFaits.DataMember = ""
        Me.dgFaits.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgFaits.Location = New System.Drawing.Point(8, 8)
        Me.dgFaits.Name = "dgFaits"
        Me.dgFaits.ReadOnly = True
        Me.dgFaits.Size = New System.Drawing.Size(558, 450)
        Me.dgFaits.TabIndex = 28
        '
        'tpExpertises
        '
        Me.tpExpertises.Controls.Add(Me.chkFaitsJustes)
        Me.tpExpertises.Controls.Add(Me.lbFaitsJustes)
        Me.tpExpertises.Controls.Add(Me.lbFaits)
        Me.tpExpertises.Controls.Add(Me.cmdRapport)
        Me.tpExpertises.Controls.Add(Me.lbConclusions)
        Me.tpExpertises.Controls.Add(Me.lblExpertise)
        Me.tpExpertises.Controls.Add(Me.lblFaitsInitiaux)
        Me.tpExpertises.Location = New System.Drawing.Point(4, 22)
        Me.tpExpertises.Name = "tpExpertises"
        Me.tpExpertises.Size = New System.Drawing.Size(574, 464)
        Me.tpExpertises.TabIndex = 2
        Me.tpExpertises.Text = "Expertises"
        '
        'lbFaitsJustes
        '
        Me.lbFaitsJustes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbFaitsJustes.BackColor = System.Drawing.SystemColors.Window
        Me.lbFaitsJustes.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbFaitsJustes.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbFaitsJustes.Location = New System.Drawing.Point(8, 40)
        Me.lbFaitsJustes.Name = "lbFaitsJustes"
        Me.lbFaitsJustes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbFaitsJustes.Size = New System.Drawing.Size(175, 407)
        Me.lbFaitsJustes.TabIndex = 18
        '
        'lbFaits
        '
        Me.lbFaits.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbFaits.BackColor = System.Drawing.SystemColors.Window
        Me.lbFaits.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbFaits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbFaits.Location = New System.Drawing.Point(8, 40)
        Me.lbFaits.Name = "lbFaits"
        Me.lbFaits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbFaits.Size = New System.Drawing.Size(175, 407)
        Me.lbFaits.TabIndex = 17
        Me.lbFaits.Visible = False
        '
        'cmdRapport
        '
        Me.cmdRapport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRapport.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRapport.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRapport.Enabled = False
        Me.cmdRapport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRapport.Location = New System.Drawing.Point(462, 8)
        Me.cmdRapport.Name = "cmdRapport"
        Me.cmdRapport.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRapport.Size = New System.Drawing.Size(104, 24)
        Me.cmdRapport.TabIndex = 15
        Me.cmdRapport.Text = "Créer Rapport.txt"
        Me.ToolTip1.SetToolTip(Me.cmdRapport, "Créer un fichier Rapport.txt avec les conclusions obtenues")
        Me.cmdRapport.UseVisualStyleBackColor = False
        '
        'lbConclusions
        '
        Me.lbConclusions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbConclusions.BackColor = System.Drawing.Color.LightCoral
        Me.lbConclusions.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbConclusions.ForeColor = System.Drawing.Color.Black
        Me.lbConclusions.Location = New System.Drawing.Point(192, 40)
        Me.lbConclusions.Name = "lbConclusions"
        Me.lbConclusions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbConclusions.Size = New System.Drawing.Size(376, 407)
        Me.lbConclusions.TabIndex = 13
        '
        'lblExpertise
        '
        Me.lblExpertise.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblExpertise.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblExpertise.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblExpertise.Location = New System.Drawing.Point(192, 16)
        Me.lblExpertise.Name = "lblExpertise"
        Me.lblExpertise.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblExpertise.Size = New System.Drawing.Size(200, 16)
        Me.lblExpertise.TabIndex = 14
        Me.lblExpertise.Text = "Rapport d'expertise :"
        '
        'lblFaitsInitiaux
        '
        Me.lblFaitsInitiaux.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblFaitsInitiaux.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFaitsInitiaux.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblFaitsInitiaux.Location = New System.Drawing.Point(8, 16)
        Me.lblFaitsInitiaux.Name = "lblFaitsInitiaux"
        Me.lblFaitsInitiaux.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFaitsInitiaux.Size = New System.Drawing.Size(72, 16)
        Me.lblFaitsInitiaux.TabIndex = 12
        Me.lblFaitsInitiaux.Text = "Faits initiaux"
        '
        'tpBilanSession
        '
        Me.tpBilanSession.Controls.Add(Me.dgBilanSession)
        Me.tpBilanSession.Location = New System.Drawing.Point(4, 22)
        Me.tpBilanSession.Name = "tpBilanSession"
        Me.tpBilanSession.Size = New System.Drawing.Size(574, 464)
        Me.tpBilanSession.TabIndex = 5
        Me.tpBilanSession.Text = "Bilan"
        '
        'dgBilanSession
        '
        Me.dgBilanSession.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgBilanSession.DataMember = ""
        Me.dgBilanSession.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgBilanSession.Location = New System.Drawing.Point(8, 8)
        Me.dgBilanSession.Name = "dgBilanSession"
        Me.dgBilanSession.ReadOnly = True
        Me.dgBilanSession.Size = New System.Drawing.Size(558, 450)
        Me.dgBilanSession.TabIndex = 0
        '
        'tpArchivage
        '
        Me.tpArchivage.BackColor = System.Drawing.SystemColors.Control
        Me.tpArchivage.Controls.Add(Me.lblEnregistrementOcx)
        Me.tpArchivage.Controls.Add(Me.lblInfoDBToFile)
        Me.tpArchivage.Controls.Add(Me.cmdDBToFile)
        Me.tpArchivage.Controls.Add(Me.lblArchivage)
        Me.tpArchivage.Controls.Add(Me.chkViderBaseMdb)
        Me.tpArchivage.Controls.Add(Me.cmdChargerAppli)
        Me.tpArchivage.Controls.Add(Me.cmdChargerTout)
        Me.tpArchivage.Controls.Add(Me.cmdArchiverAppli)
        Me.tpArchivage.Controls.Add(Me.cmdArchiverTout)
        Me.tpArchivage.Location = New System.Drawing.Point(4, 22)
        Me.tpArchivage.Name = "tpArchivage"
        Me.tpArchivage.Padding = New System.Windows.Forms.Padding(3)
        Me.tpArchivage.Size = New System.Drawing.Size(574, 464)
        Me.tpArchivage.TabIndex = 7
        Me.tpArchivage.Text = "Archivage"
        '
        'chkViderBaseMdb
        '
        Me.chkViderBaseMdb.AutoSize = True
        Me.chkViderBaseMdb.Location = New System.Drawing.Point(30, 104)
        Me.chkViderBaseMdb.Name = "chkViderBaseMdb"
        Me.chkViderBaseMdb.Size = New System.Drawing.Size(50, 17)
        Me.chkViderBaseMdb.TabIndex = 4
        Me.chkViderBaseMdb.Text = "Vider"
        Me.ToolTip1.SetToolTip(Me.chkViderBaseMdb, "Vider au préalable toutes les applications de la base mdb avant de recharcher tou" &
        "t le fichier applications.bba")
        Me.chkViderBaseMdb.UseVisualStyleBackColor = True
        '
        'cmdChargerAppli
        '
        Me.cmdChargerAppli.Location = New System.Drawing.Point(131, 75)
        Me.cmdChargerAppli.Name = "cmdChargerAppli"
        Me.cmdChargerAppli.Size = New System.Drawing.Size(82, 23)
        Me.cmdChargerAppli.TabIndex = 3
        Me.cmdChargerAppli.Text = "Charger"
        Me.ToolTip1.SetToolTip(Me.cmdChargerAppli, "Charger une application via un fichier .bba (le nom de l'application trouvée dans" &
        " le fichier sera conservée)")
        Me.cmdChargerAppli.UseVisualStyleBackColor = True
        '
        'cmdChargerTout
        '
        Me.cmdChargerTout.Location = New System.Drawing.Point(30, 75)
        Me.cmdChargerTout.Name = "cmdChargerTout"
        Me.cmdChargerTout.Size = New System.Drawing.Size(82, 23)
        Me.cmdChargerTout.TabIndex = 2
        Me.cmdChargerTout.Text = "Charger tout"
        Me.ToolTip1.SetToolTip(Me.cmdChargerTout, "Charger toutes les applications du fichier applications.bba")
        Me.cmdChargerTout.UseVisualStyleBackColor = True
        '
        'cmdArchiverAppli
        '
        Me.cmdArchiverAppli.Enabled = False
        Me.cmdArchiverAppli.Location = New System.Drawing.Point(131, 28)
        Me.cmdArchiverAppli.Name = "cmdArchiverAppli"
        Me.cmdArchiverAppli.Size = New System.Drawing.Size(82, 23)
        Me.cmdArchiverAppli.TabIndex = 1
        Me.cmdArchiverAppli.Text = "Archiver"
        Me.ToolTip1.SetToolTip(Me.cmdArchiverAppli, "Archiver l'application sélectionnée dans un fichier .bba")
        Me.cmdArchiverAppli.UseVisualStyleBackColor = True
        '
        'cmdArchiverTout
        '
        Me.cmdArchiverTout.Location = New System.Drawing.Point(30, 28)
        Me.cmdArchiverTout.Name = "cmdArchiverTout"
        Me.cmdArchiverTout.Size = New System.Drawing.Size(82, 23)
        Me.cmdArchiverTout.TabIndex = 0
        Me.cmdArchiverTout.Text = "Archiver tout"
        Me.ToolTip1.SetToolTip(Me.cmdArchiverTout, "Archiver toutes les applications dans le fichier applications.bba")
        Me.cmdArchiverTout.UseVisualStyleBackColor = True
        '
        'tpAPropos
        '
        Me.tpAPropos.Controls.Add(Me.llblMdb)
        Me.tpAPropos.Controls.Add(Me.lblVBBrainBox)
        Me.tpAPropos.Controls.Add(Me.cmdAPropos)
        Me.tpAPropos.Location = New System.Drawing.Point(4, 22)
        Me.tpAPropos.Name = "tpAPropos"
        Me.tpAPropos.Size = New System.Drawing.Size(574, 464)
        Me.tpAPropos.TabIndex = 6
        Me.tpAPropos.Text = "A propos"
        '
        'llblMdb
        '
        Me.llblMdb.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llblMdb.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.llblMdb.Location = New System.Drawing.Point(23, 96)
        Me.llblMdb.Name = "llblMdb"
        Me.llblMdb.Size = New System.Drawing.Size(528, 32)
        Me.llblMdb.TabIndex = 30
        Me.llblMdb.TabStop = True
        Me.llblMdb.Text = "VBBrainBox.mdb"
        Me.llblMdb.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.llblMdb, "Base de données de VBBrainBox pour créer ou modifier des applications")
        '
        'lblVBBrainBox
        '
        Me.lblVBBrainBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblVBBrainBox.Font = New System.Drawing.Font("Courier New", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVBBrainBox.Location = New System.Drawing.Point(88, 24)
        Me.lblVBBrainBox.Name = "lblVBBrainBox"
        Me.lblVBBrainBox.Size = New System.Drawing.Size(240, 32)
        Me.lblVBBrainBox.TabIndex = 22
        Me.lblVBBrainBox.Text = "VBBrainBox"
        Me.lblVBBrainBox.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdAPropos
        '
        Me.cmdAPropos.Image = CType(resources.GetObject("cmdAPropos.Image"), System.Drawing.Image)
        Me.cmdAPropos.Location = New System.Drawing.Point(16, 16)
        Me.cmdAPropos.Name = "cmdAPropos"
        Me.cmdAPropos.Size = New System.Drawing.Size(56, 40)
        Me.cmdAPropos.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.cmdAPropos, "A propos...")
        '
        'lblSessions
        '
        Me.lblSessions.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblSessions.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSessions.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSessions.Location = New System.Drawing.Point(8, 264)
        Me.lblSessions.Name = "lblSessions"
        Me.lblSessions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSessions.Size = New System.Drawing.Size(160, 17)
        Me.lblSessions.TabIndex = 5
        Me.lblSessions.Text = "Sessions :"
        '
        'lblEnregistrementOcx
        '
        Me.lblEnregistrementOcx.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEnregistrementOcx.Location = New System.Drawing.Point(215, 196)
        Me.lblEnregistrementOcx.Name = "lblEnregistrementOcx"
        Me.lblEnregistrementOcx.Size = New System.Drawing.Size(336, 51)
        Me.lblEnregistrementOcx.TabIndex = 35
        Me.lblEnregistrementOcx.Text = "Explication"
        '
        'lblInfoDBToFile
        '
        Me.lblInfoDBToFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfoDBToFile.Location = New System.Drawing.Point(32, 231)
        Me.lblInfoDBToFile.Name = "lblInfoDBToFile"
        Me.lblInfoDBToFile.Size = New System.Drawing.Size(160, 16)
        Me.lblInfoDBToFile.TabIndex = 34
        Me.lblInfoDBToFile.Text = "Utilitaire pour l'archivage :"
        '
        'cmdDBToFile
        '
        Me.cmdDBToFile.Location = New System.Drawing.Point(24, 255)
        Me.cmdDBToFile.Name = "cmdDBToFile"
        Me.cmdDBToFile.Size = New System.Drawing.Size(176, 32)
        Me.cmdDBToFile.TabIndex = 33
        Me.cmdDBToFile.Text = "?"
        Me.ToolTip1.SetToolTip(Me.cmdDBToFile, "Inscrire/désinscrire le contrôle DBToFile.ocx dans la base de registre")
        '
        'lblArchivage
        '
        Me.lblArchivage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblArchivage.Location = New System.Drawing.Point(216, 263)
        Me.lblArchivage.Name = "lblArchivage"
        Me.lblArchivage.Size = New System.Drawing.Size(336, 24)
        Me.lblArchivage.TabIndex = 32
        Me.lblArchivage.Text = "Enregistrement de DBTofIle.ocx"
        '
        'frmVBBrainBox
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ClientSize = New System.Drawing.Size(768, 501)
        Me.Controls.Add(Me.tcExpertises)
        Me.Controls.Add(Me.tcEntrees)
        Me.Controls.Add(Me.lbSessions)
        Me.Controls.Add(Me.lblSessions)
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(40, 64)
        Me.Name = "frmVBBrainBox"
        Me.Text = "VBBrainBox"
        Me.tcEntrees.ResumeLayout(False)
        Me.tpBD.ResumeLayout(False)
        Me.tpFichiers.ResumeLayout(False)
        Me.tcExpertises.ResumeLayout(False)
        Me.tpVariables.ResumeLayout(False)
        CType(Me.dgVariables, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpRegles.ResumeLayout(False)
        CType(Me.dgRegles, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpReglesListe.ResumeLayout(False)
        Me.tpFaits.ResumeLayout(False)
        CType(Me.dgFaits, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpExpertises.ResumeLayout(False)
        Me.tpBilanSession.ResumeLayout(False)
        CType(Me.dgBilanSession, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpArchivage.ResumeLayout(False)
        Me.tpArchivage.PerformLayout()
        Me.tpAPropos.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tpArchivage As TabPage
    Friend WithEvents cmdChargerAppli As Button
    Friend WithEvents cmdChargerTout As Button
    Friend WithEvents cmdArchiverAppli As Button
    Friend WithEvents cmdArchiverTout As Button
    Friend WithEvents chkViderBaseMdb As CheckBox
    Friend WithEvents lblEnregistrementOcx As Label
    Friend WithEvents lblInfoDBToFile As Label
    Friend WithEvents cmdDBToFile As Button
    Friend WithEvents lblArchivage As Label
End Class