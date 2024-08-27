
' Fichier modUtil.vb
' ------------------

'Option Strict On
'Option Explicit On 

Imports System.Data.OleDb
Imports System.Security.Principal

Public Class clsUtil

    Public Const sVrai$ = sGm & "VRAI" & sGm
    Public Const sFaux$ = sGm & "FAUX" & sGm

    ' Exportation Turbo-Expert 1.2 en VB6
    '  impossible de définir une constante, donc variable globale
    ' Le code page 1252 correspond à FileOpen de VB .NET, l'équivalent en VB6 de
    '  Open sCheminFichier For Input As #1
    Public Shared encodageVB6 As System.Text.Encoding =
            System.Text.Encoding.GetEncoding(1252) ' Code page : windows-1252

#Region "Base de données"

    Public Shared Function sDLookUp$(sConnexion$, sSQL$)

        ' DLookUp sur une base de données : 
        '  c'est une requête qui renvoie la valeur d'un champ
        Try
            Dim connexion As New OleDb.OleDbConnection(sConnexion)
            connexion.Open()
            sDLookUp = sDLookUp(connexion, sSQL)
            connexion.Close()
        Catch
            sDLookUp = ""
        End Try

    End Function

    Public Shared Function sDLookUp$(connexion As OleDb.OleDbConnection, sSQL$)

        sDLookUp = ""
        Try
            Dim adp0 As New OleDb.OleDbDataAdapter(sSQL, connexion)
            Dim dt As New DataTable()
            adp0.Fill(dt)
            sDLookUp = sNonVide(CType(dt.Rows(0), DataRow)(0))
        Catch
            MsgBox("Erreur lors du LookUp :" & vbCrLf &
                    sSQL & vbCrLf & Err.ToString, MsgBoxStyle.Critical)
        End Try

    End Function

    Public Shared Function bRequeteInsertion(connexion As OleDb.OleDbConnection, sSQL$) As Boolean

        Try
            Using command As New OleDbCommand(sSQL, connexion)
                'command.Parameters.AddWithValue("@Value1", "Valeur 1")
                Dim rowsAffected As Integer = command.ExecuteNonQuery()
                If rowsAffected > 0 Then
                    Return True
                Else
                    Return False
                End If
            End Using

        Catch ex As Exception
            MsgBox("Erreur lors de bRequete :" & vbCrLf &
                    sSQL & vbCrLf & Err.ToString, MsgBoxStyle.Critical)
            Return False
        End Try

    End Function

    Public Shared Function bEstVide(oChamp As Object) As Boolean

        ' oChamp est un String sauf si c'est un DBNull
        'If oChamp.GetType Is GetType(System.String) Then MsgBox("String")
        If IsDBNull(oChamp) Then Return True
        If CType(oChamp, String) Is Nothing Then Return True
        If CType(oChamp, String) = "" Then Return True
        Return False

    End Function

    Public Shared Function sNonVide$(oChamp As Object, Optional sDefaut$ = "")
        If bEstVide(oChamp) Then sNonVide = sDefaut : Exit Function
        sNonVide = CStr(oChamp)
    End Function

    Public Shared Function bNonVide(oChamp As Object, Optional bDefaut As Boolean = False) As Boolean
        If bEstVide(oChamp) Then bNonVide = bDefaut : Exit Function
        bNonVide = CBool(oChamp)
    End Function

    Public Shared Function rNonVide!(oChamp As Object, rDefaut!)
        If bEstVide(oChamp) Then rNonVide = rDefaut : Exit Function
        rNonVide = CSng(oChamp)
    End Function

    Public Shared Function bValeurNulleOuVrai(sValeur$) As Boolean
        If (sValeur Is Nothing) Or sValeur = "" Or sValeur = sVrai Then Return True
        Return False
    End Function

    Public Shared Function sParametrerRq$(sSQL$, iPrm%)

        Dim iIndexPrm% = sSQL.IndexOf("?")
        sParametrerRq = sSQL.Substring(0, iIndexPrm) & iPrm
        If iIndexPrm < sSQL.Length Then sParametrerRq &= sSQL.Substring(iIndexPrm + 1)

    End Function

#End Region

    Public Shared Sub AjouterMsg(sMessage$, ByRef colMsg As Specialized.StringCollection)

        ' Ajouter le message à la collection passée en entrée

        If sMessage = "" Then colMsg.Add(sMessage) : Exit Sub

        Dim iMultiligne%, iMemIndex%
        iMemIndex = 0
        Do
            iMultiligne = sMessage.IndexOf(vbCrLf, iMemIndex + 1)
            If iMultiligne = -1 And iMemIndex = 0 Then colMsg.Add(sMessage) : Exit Sub

            If iMemIndex = 0 Then iMemIndex = -2

            If iMultiligne = -1 Then
                colMsg.Add(sMessage.Substring(iMemIndex + 2))
                Exit Sub
            End If

            colMsg.Add(sMessage.Substring(iMemIndex + 2, iMultiligne - iMemIndex - 2))
            iMemIndex = iMultiligne
        Loop While iMultiligne > -1

    End Sub

    Public Shared Function bInverserDate(ByRef sDate$) As Boolean

        ' Inverser la date passée en entrée afin de transformer un champ date
        '  en un entier numérique qui peut alors être comparé dans VBBrainBox

        ' La date "3/12/97" devient "19971203"
        ' La date "1/1/92"  devient "19920101"

        Dim sMois, sJour, sMoisAnnee, sAnnee As String

        Dim j% = InStr(sDate, "/")
        sJour = Left(sDate, j - 1)
        If Len(sJour) = 1 Then sJour = "0" & sJour
        sMoisAnnee = Right(sDate, Len(sDate) - j)

        j = InStr(sMoisAnnee, "/")
        If j = 0 Then
            sMois = sMoisAnnee
            sDate = "2000" & sMois & sJour
            Return True
        Else
            sMois = Left(sMoisAnnee, j - 1)
        End If
        If Len(sMois) = 1 Then sMois = "0" & sMois
        sAnnee = sMoisAnnee.Substring(j)

        Select Case sAnnee.Length
            Case 1
                sAnnee = "200" & sAnnee

            Case 2
                If Val(sAnnee) < 50 Then
                    sAnnee = "20" & sAnnee
                Else
                    sAnnee = "19" & sAnnee
                End If
            Case 3
                sAnnee = "0" & sAnnee
            Case 4

            Case Else
                Return False
        End Select

        sDate = sAnnee & sMois & sJour
        Return True

    End Function

    Public Shared Function sTraiterHyperlienAccess$(sLien$)

        ' Simplification de l'affichage des champs Access dans le rapport
        ' Les champs hyperliens d'Access sont stockés avec 2 représentations :
        '  une forme de présentation et une URL valide
        ' ex. : http://www.web.com#http://www.web.com#
        '       patrice.dargenton@free.fr#mailto:patrice.dargenton@free.fr#
        '       lien1#lien2# ou lien1#lien2

        Dim iPosDiese% = sLien.IndexOf("#")
        Dim iPos2Diese% = sLien.LastIndexOf("#")
        If Not (iPosDiese > -1 And iPos2Diese > -1) Then GoTo Fin

        Dim sLien1$ = sLien.Substring(0, iPosDiese)

        Dim iLenLien2% = sLien.Length - iPosDiese - 1
        If iPos2Diese < sLien.Length Then iLenLien2 = sLien.Length - iPosDiese - 1
        ' Les hyperliens Access ne gère pas bien les signets dans les URL, 
        '  d'où un petit bug : il peut manquer un # à la fin !
        If Right(sLien, 1) = "#" Then iLenLien2 -= 1

        Dim bMail As Boolean
        Const sMailto$ = "mailto:"
        Dim iPosMailto% = sLien.IndexOf(sMailto)
        Dim iPosLien2% = iPosDiese + 1
        If iPosMailto > -1 Then
            iPosLien2 = iPosMailto + sMailto.Length
            iLenLien2 -= sMailto.Length
            bMail = True
        End If
        Dim sLien2$ = sLien.Substring(iPosLien2, iLenLien2)
        If sLien2 = sLien1 Then sTraiterHyperlienAccess = sLien1 : Exit Function
        ' Pour les URL, mieux vaux ne conserver que l'URL réelle
        If Not bMail Then sTraiterHyperlienAccess = sLien2 : Exit Function

Fin:
        sTraiterHyperlienAccess = sLien : Exit Function

    End Function

    Public Shared Function bCompilation64Bit() As Boolean
        Return (IntPtr.Size = 8)
    End Function

    Public Shared Function bEstAdmin() As Boolean

        ' Récupérer l'identité de l'utilisateur en cours
        Dim identity As WindowsIdentity = WindowsIdentity.GetCurrent()

        ' Créer un principal Windows avec cette identité
        Dim principal As New WindowsPrincipal(identity)

        ' Vérifier si l'utilisateur fait partie du groupe "Administrateurs"
        Return principal.IsInRole(WindowsBuiltInRole.Administrator)

    End Function

    Public Shared Function bCleRegistreExiste(sCle$) As Boolean

        ' Vérifier si une clé existe dans la base de registre
        '  c'est utile pour savoir si un contrôle ActiveX est enregistré

        Dim sValCle$
        Try
            '  This call goes to the Catch block if the registry key is not set.
            Dim myRegKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot
            myRegKey = myRegKey.OpenSubKey(sCle)
            If IsNothing(myRegKey) Then Return False ' 22/08/2024
            Dim sSousCle$ = "" ' Sous-clé par défaut
            Dim oValue As Object = myRegKey.GetValue(sSousCle)
            sValCle = CStr(oValue)
            'MsgBox(sValCle)
            Return True  ' On peut lire cette clé, donc elle existe
        Catch
            sValCle = ""
            Return False
        End Try

    End Function

    Public Shared Function bEnregistrerDllActiveX(sDllActiveX$, sRepertoireDll$,
                Optional bDesenregistrer As Boolean = False,
                Optional bConfirmer As Boolean = True) As Boolean

        ' Enregistrer (inscrire) la Dll ActiveX dans la base de registre :
        ' Une Dll ActiveX requiert la commande :
        '  C:\Windows\System\Regsvr32.exe MaDll.dll (system32 pour les NT/XP)
        ' Le désenregistrement se fait avec -u
        '  C:\Windows\System\Regsvr32.exe -u MaDll.dll
        '  (s'il y a un package d'installation, il prend alors en charge
        '   automatiquement l'enregistrement et le désenregistrement)

        Dim sCheminDll$ = sRepertoireDll & "\" & sDllActiveX
        If Not IO.File.Exists(sCheminDll) Then
            MsgBox("Impossible de trouver le fichier :" & vbCrLf &
                    sCheminDll, MsgBoxStyle.Critical)
            Return False
        End If

        Dim iReponse%
        If bConfirmer Then
            Dim sInfo$ = "Voulez-vous inscrire le contrôle " &
                sDllActiveX & " dans la base de registre ?"
            If bDesenregistrer Then sInfo$ = "Voulez-vous désinscrire le contrôle " &
                sDllActiveX & " dans la base de registre ?"
            sInfo &= vbLf & "L'application doit être lancée en mode admin pour cela."
            iReponse = MsgBox(sInfo,
                    MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question)
            If iReponse <> MsgBoxResult.Yes Then Return False
        End If

        Dim sOption$ = ""
        If bDesenregistrer Then sOption = "-u"
        ' Autre possibilité : Environment.GetFolderPath(SpecialFolder.System)
        Dim sExe$ = Environment.SystemDirectory & "\regsvr32.exe"
        Dim startInfo As New ProcessStartInfo(sExe) With {
                .Arguments = sOption & " " & sGm & sCheminDll & sGm,
                .WindowStyle = ProcessWindowStyle.Normal
            }
        Process.Start(startInfo)
        iReponse = MsgBox("L'opération a-t-elle réussie ?",
            MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question)
        If iReponse <> MsgBoxResult.Yes Then Return False
        Return True

    End Function

    Public Shared Function bChoisirFichier(ByRef sCheminFichier$, sFiltre$, sExtDef$,
            sTitre$, Optional sInitDir$ = "",
            Optional bDoitExister As Boolean = True,
            Optional bMultiselect As Boolean = False) As Boolean

        ' Afficher une boite de dialogue pour choisir un fichier
        ' Exemple de filtre : "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
        ' On peut indiquer le dossier initial via InitDir, ou bien via le chemin du fichier

        ' 1ère utilisation : initialiser, puis conserver le dossier courant
        Static bInit As Boolean = False

        Dim ofd As New OpenFileDialog
        With ofd
            If Not bInit Then
                bInit = True
                If String.IsNullOrEmpty(sInitDir) Then
                    If String.IsNullOrEmpty(sCheminFichier) Then
                        .InitialDirectory = Application.StartupPath
                    Else
                        .InitialDirectory = IO.Path.GetDirectoryName(sCheminFichier)
                    End If
                Else
                    .InitialDirectory = sInitDir
                End If
            End If
            If Not String.IsNullOrEmpty(sCheminFichier) Then
                ' Proposer le fichier par défaut, sans le chemin complet
                .FileName = IO.Path.GetFileName(sCheminFichier)
            End If
            .CheckFileExists = bDoitExister
            .DefaultExt = sExtDef
            .Filter = sFiltre
            .Multiselect = bMultiselect
            .Title = sTitre
            Dim retour As DialogResult = .ShowDialog()
            If retour = DialogResult.Cancel Then Return False

            If .FileName <> "" Then sCheminFichier = .FileName : Return True
            Return False

        End With

    End Function

    Public Shared Function sConvNomDos$(sChaine$,
            Optional bLimit8Car As Boolean = False,
            Optional bConserverSignePlus As Boolean = False)

        ' Remplacer les caractères interdits pour les noms de fichiers DOS
        '  et retourner un nom de fichier correct

        Dim iSel%, sBuffer$, sCar$, iCode%, sCarConv2$, sCarDest$
        Dim bOk As Boolean, bMaj As Boolean
        sBuffer = Trim$(sChaine)
        If bLimit8Car Then sBuffer = Left$(sBuffer, 8)
        Const sCarConv$ = " .«»/[]:;|=,*-" ' Caractères à convertir en souligné
        sCarConv2 = sCarConv
        If Not bConserverSignePlus Then sCarConv2 &= "+"
        For iSel = 1 To Len(sBuffer)
            sCar = Mid$(sBuffer, iSel, 1)
            iCode = Asc(sCar)
            bMaj = False
            If iCode >= 65 AndAlso iCode <= 90 Then bMaj = True
            If iCode >= 192 AndAlso iCode <= 221 Then bMaj = True
            If InStr(sCarConv2, sCar) > 0 Then _
                Mid$(sBuffer, iSel, 1) = "_" : GoTo Suite
            If InStr("èéêë", sCar) > 0 Then
                If bMaj Then sCarDest = "E" Else sCarDest = "e"
                Mid$(sBuffer, iSel, 1) = sCarDest
                GoTo Suite
            End If
            If InStr("àáâä", sCar) > 0 Then
                If bMaj Then sCarDest = "A" Else sCarDest = "a"
                Mid$(sBuffer, iSel, 1) = sCarDest
                GoTo Suite
            End If
            If InStr("ìíîï", sCar) > 0 Then
                If bMaj Then sCarDest = "I" Else sCarDest = "i"
                Mid$(sBuffer, iSel, 1) = sCarDest
                GoTo Suite
            End If
            If InStr("ùúûü", sCar) > 0 Then
                If bMaj Then sCarDest = "U" Else sCarDest = "u"
                Mid$(sBuffer, iSel, 1) = sCarDest
                GoTo Suite
            End If
            If InStr("òóôõö", sCar) > 0 Then ' 08/05/2013
                If bMaj Then sCarDest = "O" Else sCarDest = "o"
                Mid$(sBuffer, iSel, 1) = sCarDest
                GoTo Suite
            End If
            If InStr("ç", sCar) > 0 Then ' 12/06/2015
                If bMaj Then sCarDest = "C" Else sCarDest = "c"
                Mid$(sBuffer, iSel, 1) = sCarDest
                GoTo Suite
            End If
            If bConserverSignePlus AndAlso iCode = 43 Then GoTo Suite
            'de 65 à 90  maj
            'de 97 à 122 min
            'de 48 à 57 Chiff
            bOk = False
            If (iCode >= 65 AndAlso iCode <= 90) Then bOk = True
            If (iCode >= 97 AndAlso iCode <= 122) Then bOk = True
            If (iCode >= 48 AndAlso iCode <= 57) Then bOk = True
            If Not bOk Then Mid$(sBuffer, iSel, 1) = "_"
Suite:
        Next iSel
        sConvNomDos = sBuffer

    End Function

End Class