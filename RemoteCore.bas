Option Explicit

' =========================================================================
'   API Sleep (non bloquant via SleepStop)
' =========================================================================
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


' =========================================================================
'   MODULE STOP — Gestion arrêt utilisateur
' =========================================================================
Public GlobalStopFlag As Boolean

' BOUTON STOP : n’appelle que ceci
Public Sub StopProcess()
    GlobalStopFlag = True
End Sub

' Remise à zéro au début du traitement
Public Sub ResetStopFlag()
    GlobalStopFlag = False
End Sub

' Interrogation STOP
Public Function ShouldStop() As Boolean
    ShouldStop = GlobalStopFlag
End Function

' Sleep non bloquant qui respecte STOP
Public Sub SleepStop(ms As Long)
    Dim t As Double: t = Timer
    Dim d As Double: d = ms / 1000#

    Do While Timer - t < d
        If GlobalStopFlag Then Exit Sub
        DoEvents
    Loop
End Sub


' =========================================================================
'   TEST INTERNET (WinHTTP FIABILISÉ)
' =========================================================================
Public Function HasInternetConnection() As Boolean
    On Error GoTo FailSafe

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.SetTimeouts 3000, 3000, 3000, 3000
    http.Option(6) = &H800000

    http.Open "GET", "https://raw.githubusercontent.com", False

    On Error Resume Next
    http.send
    If Err.Number <> 0 Then GoTo FailSafe
    On Error GoTo FailSafe

    If http.readyState <> 4 Then GoTo FailSafe

    HasInternetConnection = (http.status >= 200 And http.status < 400)
    Exit Function

FailSafe:
    HasInternetConnection = False
End Function


' =========================================================================
'   CHECKREMOTELOCK — Vérification ALLOW / DENY
' =========================================================================
Public Function CheckRemoteLock() As Boolean
    Dim http As Object
    Dim url As String
    Dim resp As String

    ' 1) Vérification Internet
    If Not HasInternetConnection() Then
        MsgBox "Aucune connexion Internet détectée." & vbCrLf & _
               "Vérifiez votre réseau WiFi/4G.", vbCritical
        CheckRemoteLock = False
        Exit Function
    End If

    ' 2) Lecture lock.txt
    On Error GoTo FailSafe

    url = "https://raw.githubusercontent.com/Nick-O-lay/ANAVEO/main/lock.txt"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 3000, 3000, 3000, 3000
    http.Option(6) = &H800000

    http.Open "GET", url, False

    On Error Resume Next
    http.send
    If Err.Number <> 0 Then GoTo FailSafe
    On Error GoTo FailSafe

    If http.readyState <> 4 Then GoTo FailSafe
    If http.status <> 200 Then GoTo FailSafe

    resp = Trim(Replace(Replace(http.responseText, vbCr, ""), vbLf, ""))

    If UCase(resp) = "ALLOW" Then
        CheckRemoteLock = True
        Exit Function
    Else
        MsgBox "Accès refusé par le verrou distant." & vbCrLf & _
               "Contact : +33 (0)6 42 12 50 12", vbCritical
        CheckRemoteLock = False
        Exit Function
    End If

FailSafe:
    MsgBox "Impossible d'accéder au verrou distant." & vbCrLf & _
           "Vérifiez votre connexion.", vbCritical
    CheckRemoteLock = False
End Function


' =========================================================================
'   POINT D’ENTRÉE APPELÉ PAR LE LOADER
' =========================================================================
Public Sub GenerateListe()
    If Not CheckRemoteLock() Then Exit Sub
    Export_From_N8N
End Sub


' =========================================================================
'   EXPORT PRINCIPAL — AVEC STOP, ETA, RETRY, JSON
' =========================================================================
Sub Export_From_N8N()

    ResetStopFlag ' IMPORTANT

    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, dstRow As Long
    Dim jsonRequest As String, jsonResponse As String
    Dim http As Object
    Dim url As String

    Dim startTime As Double, elapsed As Double, avgTime As Double, remaining As Double

    Set wsSrc = Sheets("etablissements")
    Set wsDst = Sheets("MiseEnPage")

    url = "https://n8n.srv933744.hstgr.cloud/webhook/42402c7f-7d45-42be-8706-80d104efe948"

    ' RESET DESTINATION
    wsDst.Range("A3:W100000").ClearContents

    wsDst.Range("A2:W2").Value = Array( _
        "Société", "Enseigne SalesForce", "Siège social", "", _
        "Création établissement", "Effectifs", "Genre représentant", "Nom représentant", _
        "Prénom représentant", "Téléphone", "Email", "Commentaire", _
        "ESS", "Famille NAF", "Catégorie entreprise", _
        "Longitude", "Latitude", "Adresse complète", "Code postal", "Ville", _
        "Siren", "Siret", "CA")

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    dstRow = 3
    startTime = Timer


    ' =========================================================================
    '   BOUCLE LIGNE PAR LIGNE — AVEC STOP
    ' =========================================================================
    For r = 2 To lastRow

        If ShouldStop() Then GoTo EndProcess
        DoEvents

        jsonRequest = BuildJsonFromRow(wsSrc, r, lastCol)

        Dim success As Boolean: success = False
        Dim attempts As Long
        Dim waitList As Variant
        waitList = Array(300, 800, 1500, 2500, 4000)


        ' ---- 5 TENTATIVES ----
        For attempts = 0 To 4

            If ShouldStop() Then GoTo EndProcess
            DoEvents

            Set http = CreateObject("MSXML2.XMLHTTP")
            http.Open "POST", url, False
            http.setRequestHeader "Content-Type", "application/json"

            On Error Resume Next
                http.send jsonRequest
            On Error GoTo 0

            Dim t As Double: t = Timer
            Do While http.readyState <> 4 And (Timer - t) < 5
                If ShouldStop() Then GoTo EndProcess
                DoEvents
                SleepStop 50
            Loop

            jsonResponse = Trim(http.responseText)

            If Len(jsonResponse) > 0 _
               And InStr(jsonResponse, "{") > 0 _
               And InStr(jsonResponse, "}") > 0 Then

                success = True
                Exit For
            End If

            SleepStop CLng(waitList(attempts))
        Next attempts


        ' ---- SUCCÈS ----
        If success Then
            WriteJsonToSheet wsDst, dstRow, jsonResponse

        ' ---- ÉCHEC : copie brute ----
        Else
            Dim col As Long
            For col = 1 To lastCol
                wsDst.Cells(dstRow, col).Value = wsSrc.Cells(r, col).Value
            Next col
        End If

        dstRow = dstRow + 1


        ' =========================================================================
        '   BARRE DE PROGRESSION + ETA
        ' =========================================================================
        elapsed = Timer - startTime
        avgTime = elapsed / (r - 1 + 0.0001)
        remaining = avgTime * (lastRow - r + 1)

        Application.StatusBar = _
            "Progression : " & Format((r - 1) / (lastRow - 1), "0.0%") & _
            " | Temps restant estimé : " & Format(remaining / 60, "0.0") & " min"

    Next r

    MsgBox "Traitement terminé !", vbInformation
    GoTo EndProcessOK


' =========================================================================
'   SORTIES PROPRES
' =========================================================================
EndProcess:
    Application.StatusBar = False
    MsgBox "Traitement interrompu.", vbExclamation
    Exit Sub

EndProcessOK:
    Application.StatusBar = False
    Exit Sub

End Sub



' =========================================================================
'   JSON CONSTRUCTION
' =========================================================================
Function BuildJsonFromRow(ws As Worksheet, row As Long, lastCol As Long) As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim col As Long

    For col = 1 To lastCol
        dict(ws.Cells(1, col).Value) = CStr(ws.Cells(row, col).Value)
    Next col

    BuildJsonFromRow = DictToJson(dict)
End Function


' =========================================================================
'   OBJET → JSON
' =========================================================================
Function DictToJson(dict As Object) As String
    Dim key As Variant, s As String
    s = "{"

    For Each key In dict.Keys
        s = s & """" & key & """:""" & Replace(CStr(dict(key)), """", "'") & ""","
    Next key

    If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)
    s = s & "}"
    DictToJson = s
End Function


' =========================================================================
'   PARSEUR JSON SIMPLE
' =========================================================================
Function ParseSimpleJsonObject(json As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim cleaned As String: cleaned = Trim(json)

    If Left(cleaned, 1) = "[" Then cleaned = Mid(cleaned, 2)
    If Right(cleaned, 1) = "]" Then cleaned = Left(cleaned, Len(cleaned) - 1)

    cleaned = Trim(cleaned)
    If Left(cleaned, 1) = "{" Then cleaned = Mid(cleaned, 2)
    If Right(cleaned, 1) = "}" Then cleaned = Left(cleaned, Len(cleaned) - 1)

    Dim parts() As String
    parts = Split(cleaned, ",")

    Dim i As Long
    Dim kv() As String, k As String, v As String

    For i = LBound(parts) To UBound(parts)

        kv = Split(parts(i), ":")

        If UBound(kv) >= 1 Then
            k = Replace(Replace(Trim(kv(0)), """", ""), "'", "")
            v = Replace(Replace(Trim(kv(1)), """", ""), "'", "")
            dict(k) = v
        End If
    Next i

    Set ParseSimpleJsonObject = dict
End Function
