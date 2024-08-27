
' Fichier modConst.vb
' -------------------

Module _modConst

    Public ReadOnly sNomAppli$ = My.Application.Info.Title ' VBBrainBox
    Public Const sTitreMsg$ = "VBBrainBox : un système expert d'ordre 0+"
    Public Const sDateVersionAppli$ = "27/08/2024"

#If DEBUG Then
    Public Const bDebug As Boolean = True
    Public Const bRelease As Boolean = False
#Else
    Public Const bDebug As Boolean = False
    Public Const bRelease As Boolean = True
#End If

    Public Const sGm$ = """" ' Guillemets : sGm correspond à un seul "

End Module