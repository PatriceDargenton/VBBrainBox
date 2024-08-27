
Option Strict Off ' Mode liaison tardive, si le composant DBToFile.ocx n'est pas inscrit dans la base de registre

Imports System.IO

Module modDBToFile

    Public Const sFichierDBToFile$ = "DBToFile.ocx"
    Public Const sCleRegistreDBToFile$ = "CLSID\{A8EEB80D-9749-11D3-8214-E23042430D34}"
    Public Const sFiltreBBA$ = "Fichiers bba (*.bba)|*.bba|Tous les fichiers (*.*)|*.*"
    Public Const sExtBBA$ = "*.bba"
    Public Const sFichierApplicationsParDefaut$ = "Applications.bba"
    Public Const sFichierApplicationParDefaut$ = "Application.bba"
    Const sCleInstanciationDBToFile$ = "DBToFile.UserControl1"

    Public Function CreerObjetDBToFile(bSauver As Boolean,
            Optional bVider As Boolean = False,
            Optional sCheminApp$ = "", Optional sIdApp$ = "", Optional sApp$ = "") As Boolean

        Dim oDBToFile As Object = Nothing

        Dim sDossierCourant$ = Application.StartupPath & "\Applications\"

        Dim bSucces = False

        Try
            oDBToFile = CreateObject(sCleInstanciationDBToFile)

            Dim sFichierDos$ = ""
            With oDBToFile
                .iLanguage = 1 ' Messages en Français
                .strDBToFile = sDossierCourant & "DBToFile2.mdb" 'sFichierDBToFileMDB
                .strDatabaseFile = sDossierCourant & clsVBBBox.sFichierVBBBoxMDB
                sFichierDos = sFichierApplicationsParDefaut
                .strBackupFile = sDossierCourant & sFichierDos
                If Not String.IsNullOrEmpty(sCheminApp) Then
                    .strBackupFile = sCheminApp
                    sFichierDos = Path.GetFileName(sCheminApp)
                End If
                .strSaveFileArg = ""
                .bIDsFieldNameEndsByID = False
                .bIDsFieldNameStartsByID = True
                .strSaveFileTypeMn = "VBBRAINBOXALL"
                .bNoConfirm = False ' Possibilité de ne pas confirmer l'opération
                If Not String.IsNullOrEmpty(sIdApp) Then
                    .strSaveFileTypeMn = "VBBRAINBOXAPP"
                    .strSaveFileArg = "|Application:IdApplication=" & sIdApp & "|"
                    sFichierDos = sFichierApplicationParDefaut
                    .strBackupFile = sDossierCourant & sFichierDos
                    If Not String.IsNullOrEmpty(sApp) Then
                        sFichierDos = clsUtil.sConvNomDos(sApp) & ".bba"
                        .strBackupFile = sDossierCourant & sFichierDos
                    End If
                End If

                If bSauver Then
                    bSucces = .bCreateArchivFile
                    If bSucces Then MsgBox("Le fichier a été créé avec succès : " & sFichierDos, MsgBoxStyle.Information)
                Else
                    If bVider Then .bDeleteRecords

                    If Not String.IsNullOrEmpty(sCheminApp) Then
                        .strSaveFileTypeMn = "VBBRAINBOXAPP"
                        .strSaveFileArg = "|Application:IdApplication|"
                        .strBackupFile = sCheminApp
                    End If

                    bSucces = .bLoadArchivFile
                    ' L'annulation via cancel fonctionne bien, mais le booléen reste à False ici
                    ' Donc le message ne sera pas correct
                    Dim bCancel As Boolean = .bCancel

                    If bSucces Then MsgBox(
                        "Le fichier a été chargé avec succès : " & sFichierDos,
                        MsgBoxStyle.Information)

                End If

                ' Note : on peut récup. les événements via oDBToFile.strProgress avec un timer

            End With

            Return bSucces

        Catch ex As Exception
            MsgBox("Erreur : " & vbLf & ex.Message & vbLf &
                "Cause possible : Le chemin vers DBToFile.ocx a été modifié" & vbLf &
                "(refaire l'enregistrement de DBToFile.ocx en mode admin.)",
                MsgBoxStyle.Critical, "CreerObjetDBToFile")
            Return False
        End Try

    End Function

End Module