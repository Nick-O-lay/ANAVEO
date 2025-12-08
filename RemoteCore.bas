Attribute VB_Name = "RemoteCore"
Option Explicit

' ============================================================
' CONFIG DISTANTE
' ============================================================
Private Const GITHUB_LOCK As String = _
    "https://raw.githubusercontent.com/Nick-O-lay/ANAVEO/main/lock.txt"


' ============================================================
' POINT D'ENTRÉE GLOBAL — APPELÉ DEPUIS EXCEL
' ============================================================
Public Sub MainEntry()
    ' 1) Vérifier le lock DANS le code GitHub
    If Not CheckRemoteLock() Then
        MsgBox "Exécution bloquée par le verrou distant (lock.txt <> ALLOW).", vbCritical
        Exit Sub
    End If
    
    ' 2) Si on passe ici, ALLOW → on exécute le vrai code
    RunBusinessLogic
End Sub


' ============================================================
' LOGIQUE D'AUTORISATION CÔTÉ GITHUB
' ============================================================
Private Function CheckRemoteLock() As Boolean
    Dim resp As String
    resp = RemoteDownloadText(GITHUB_LOCK)
    
    resp = Trim$(UCase$(resp))
    CheckRemoteLock = (resp = "ALLOW")
End Function


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


' ============================================================
' TON VRAI CODE MÉTIER
' ============================================================
Private Sub RunBusinessLogic()
    ' TODO : ici tu mets tout ton vrai code.
    ' Exemple :
    MsgBox "Code distant autorisé et exécuté depuis GitHub.", vbInformation
    
    ' Appel ex :
    ' Call Fetch_CA_Chunk_Start
    ' Call Module1.Traitement...
End Sub
