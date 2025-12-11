Attribute VB_Name = "RemoteCore"
Option Explicit

' =====================================================================
'      CONFIG DISTANTE (LOCK GITHUB)
' =====================================================================
Private Const GITHUB_LOCK As String = _
    "https://raw.githubusercontent.com/Nick-O-lay/ANAVEO/main/lock.txt"


' =====================================================================
'      ALIAS : LA MACRO QUE TON BOUTON APPELLE
' =====================================================================
Public Sub GenerateListe()
    MainEntry
End Sub


' =====================================================================
'      POINT D'ENTRÉE PRINCIPAL — EXÉCUTÉ APRÈS MISE À JOUR DU MODULE
' =====================================================================
Public Sub MainEntry()
    
    ' 1) Vérifier le verrou distant
    If Not CheckRemoteLock() Then
        MsgBox "Exécution bloquée par le verrou distant (lock.txt <> ALLOW).", vbCritical
        Exit Sub
    End If
    
    ' 2) Lancement du vrai code métier
    RunBusinessLogic

End Sub


' =====================================================================
'      VÉRIFICATION DU LOCK GITHUB
' =====================================================================
Private Function CheckRemoteLock() As Boolean

    Dim resp As String
    resp = RemoteDownloadText(GITHUB_LOCK)
    
    resp = Trim$(UCase$(resp))
    
    CheckRemoteLock = (resp = "ALLOW")

End Function


' =====================================================================
'      TÉLÉCHARGEMENT TEXTE (HTTP GET)
' =====================================================================
Private Function RemoteDownloadText(ByVal url As String) As String
    
    On Error GoTo Fail
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", url, False
    http.send
    
    If http.readyState = 4 And http.Status = 200 Then
        RemoteDownloadText = CStr(http.responseText)
    Else
        RemoteDownloadText = ""
    End If
    
    Exit Function
    
Fail:
    RemoteDownloadText = ""

End Function


' =====================================================================
'      CODE MÉTIER PRINCIPAL (À PERSONNALISER)
' =====================================================================
Private Sub RunBusinessLogic()

    ' Ici tu mets ton vrai traitement
    ' Exemple provisoire :
    
    MsgBox "Code distant autorisé et exécuté depuis GitHub.", vbInformation
    
    ' =============================================
    ' ton vrai code sera ici :
    '
    ' Call Fetch_CA_Chunk_Start
    ' Call Fetch_Dirigeants_Start
    ' Call Export_From_N8N
    ' =============================================
    
End Sub
