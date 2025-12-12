Attribute VB_Name = "RemoteCore"
Option Explicit

' =========================================================================
'   API Sleep (non bloquant grâce à SleepStop) - PROUT 2 PROUT
' =========================================================================
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


' =========================================================================
'   STOP PROCESS — Interruption propre
' =========================================================================
Public GlobalStopFlag As Boolean

Public Sub StopProcess()
    GlobalStopFlag = True
End Sub

Public Sub ResetStopFlag()
    GlobalStopFlag = False
End Sub

Public Function ShouldStop() As Boolean
    ShouldStop = GlobalStopFlag
End Function

Public Sub SleepStop(ms As Long)
    Dim t As Double: t = Timer
    Dim limit As Double: limit = ms / 1000#
    Do While Timer - t < limit
        If GlobalStopFlag Then Exit Sub
        DoEvents
    Loop
End Sub


' =========================================================================
'   POINT D’ENTRÉE PRINCIPAL (appelé par le Loader Excel)
' =========================================================================
Public Sub GenerateListe()
    Export_From_N8N
End Sub


' =========================================================================
'   EXPORT PRINCIPAL — HTTP POST + RETRY + ETA + STOP
' =========================================================================
Public Sub Export_From_N8N()

    ResetStopFlag

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

    ' RESET destination
    wsDst.Range("A3:W100000").ClearContents

    wsDst.Range("A2:W2").Value = Array( _
        "Société", "Enseigne SalesForce", "Siège social", "", _
        "Création établissement", "Effectifs", "Genre représentant", "Nom représentant", _
        "Prénom représentant", "Téléphone", "Email", "Commentaire", _
        "ESS", "Famille NAF", "Catégorie entreprise", _
        "Longitude", "Latitude", "Adresse complète", "Code postal", "Ville", _
        "Siren", "Siret", "CA")

    wsDst.Columns("F").NumberFormat = "@"
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    
    dstRow = 3
    startTime = Timer


    ' =========================================================================
    '   BOUCLE PRINCIPALE
    ' =========================================================================
    For r = 2 To lastRow

        If ShouldStop() Then GoTo EndProcess
        DoEvents

        jsonRequest = BuildJsonFromRow(wsSrc, r, lastCol)

        Dim success As Boolean: success = False
        Dim attempts As Long

        ' Liste des temps d'attente entre tentatives (ms)
        Dim waitList As Variant: waitList = Array(300, 800, 1500, 2500, 4000)

        ' ----------------------------
        '   RETRY HTTP (5 tentatives)
        ' ----------------------------
        For attempts = 0 To 4

            If ShouldStop() Then GoTo EndProcess

            Set http = CreateObject("MSXML2.XMLHTTP")
            http.Open "POST", url, False
            http.setRequestHeader "Content-Type", "application/json"

            On Error Resume Next
                http.send jsonRequest
            On Error GoTo 0

            ' attendre la réponse max 5s
            Dim t As Double: t = Timer
            Do While http.readyState <> 4 And (Timer - t) < 5
                If ShouldStop() Then GoTo EndProcess
                DoEvents
                SleepStop 50
            Loop

            jsonResponse = Trim(http.responseText)

            If Len(jsonResponse) > 0 And InStr(jsonResponse, "{") > 0 Then
                success = True
                Exit For
            End If

            ' attendre avant nouvelle tentative
            SleepStop CLng(waitList(attempts))

        Next attempts


        ' -------------------------
        '   ÉCRITURE RESULTATS
        ' -------------------------
        If success Then
            WriteJsonToSheet wsDst, dstRow, jsonResponse
        Else
            ' copie brute si échec
            Dim c As Long
            For c = 1 To lastCol
                wsDst.Cells(dstRow, c).Value = wsSrc.Cells(r, c).Value
            Next c
        End If

        dstRow = dstRow + 1


        ' -------------------------
        '   ETAT / ETA / PROGRESS
        ' -------------------------
        elapsed = Timer - startTime
        avgTime = elapsed / (r - 1 + 0.0001)
        remaining = avgTime * (lastRow - r + 1)

        Application.StatusBar = _
            "Progression : " & Format((r - 1) / (lastRow - 1), "0.0%") & _
            " | Temps restant : " & Format(remaining / 60, "0.0") & " min"

    Next r

    MsgBox "Traitement terminé !", vbInformation
    GoTo EndOK


' =========================================================================
'   FIN
' =========================================================================
EndProcess:
    Application.StatusBar = False
    MsgBox "Traitement interrompu.", vbExclamation
    Exit Sub

EndOK:
    Application.StatusBar = False
End Sub



' =========================================================================
'   JSON → FEUILLE
' =========================================================================
Public Sub WriteJsonToSheet(ws As Worksheet, row As Long, jsonText As String)
    Dim obj As Object
    Set obj = ParseSimpleJsonObject(jsonText)
    If obj Is Nothing Then Exit Sub

    Dim key As Variant
    Dim col As Long: col = 1
    For Each key In obj.keys
        ws.Cells(row, col).value = obj(key)
        col = col + 1
    Next key
End Sub

' =========================================================================
'   JSON PARSER SIMPLE
' =========================================================================
Public Function ParseSimpleJsonObject(json As String) As Object

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim cleaned As String: cleaned = Trim(json)

    If cleaned = "" Then
        Set ParseSimpleJsonObject = dict
        Exit Function
    End If

    ' Nettoyage []
    If Left(cleaned, 1) = "[" Then cleaned = Mid(cleaned, 2)
    If Right(cleaned, 1) = "]" Then cleaned = Left(cleaned, Len(cleaned) - 1)

    cleaned = Trim(cleaned)

    ' Nettoyage {}
    If Left(cleaned, 1) = "{" Then cleaned = Mid(cleaned, 2)
    If Right(cleaned, 1) = "}" Then cleaned = Left(cleaned, Len(cleaned) - 1)

    cleaned = Trim(cleaned)

    Dim parts() As String: parts = Split(cleaned, ",")
    Dim i As Long

    For i = LBound(parts) To UBound(parts)
        Dim kv() As String
        kv = Split(parts(i), ":")

        If UBound(kv) >= 1 Then
            Dim k As String, v As String
            k = Replace(Replace(Trim(kv(0)), """", ""), "'", "")
            v = Replace(Replace(Trim(kv(1)), """", ""), "'", "")
            dict(k) = v
        End If
    Next i

    Set ParseSimpleJsonObject = dict
End Function



' =========================================================================
'   CONSTRUCTION JSON
' =========================================================================
Public Function BuildJsonFromRow(ws As Worksheet, row As Long, lastCol As Long) As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim c As Long

    For c = 1 To lastCol
        dict(ws.Cells(1, c).Value) = CStr(ws.Cells(row, c).Value)
    Next c

    BuildJsonFromRow = DictToJson(dict)
End Function



' =========================================================================
'   DICTIONNAIRE → JSON
' =========================================================================
Public Function DictToJson(dict As Object) As String

    Dim key As Variant
    Dim s As String: s = "{"

    For Each key In dict.Keys
        s = s & """" & key & """:""" & Replace(CStr(dict(key)), """", "'") & ""","
    Next key

    If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)
    s = s & "}"

    DictToJson = s
End Function
