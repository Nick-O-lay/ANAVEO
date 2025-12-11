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

' Bouton STOP → N'appelle que ceci
Public Sub StopProcess()
    GlobalStopFlag = True
End Sub

' Reset automatique au début du traitement
Public Sub ResetStopFlag()
    GlobalStopFlag = False
End Sub

' Vérifie si STOP demandé
Public Function ShouldStop() As Boolean
    ShouldStop = GlobalStopFlag
End Function

' Sleep non bloquant : s'interrompt immédiatement si STOP
Public Sub SleepStop(ms As Long)
    Dim t As Double: t = Timer
    Dim limit As Double: limit = ms / 1000#

    Do While Timer - t < limit
        If GlobalStopFlag Then Exit Sub
        DoEvents
    Loop
End Sub



' =========================================================================
'   POINT D’ENTRÉE APPELÉ PAR LE LOADER
' =========================================================================
' → Plus de test Internet
' → Plus de lock distant
' → Exécution immédiate
' =========================================================================
Public Sub GenerateListe()
    Export_From_N8N
End Sub



' =========================================================================
'   EXPORT PRINCIPAL — AVEC STOP, RETRY, ETA
' =========================================================================
Public Sub Export_From_N8N()

    ResetStopFlag

    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, dstRow As Long
    Dim jsonRequest As String, jsonResponse As String
    Dim http As Object, url As String

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
    '   BOUCLE LIGNE PAR LIGNE
    ' =========================================================================
    For r = 2 To lastRow

        If ShouldStop() Then GoTo EndProcess
        DoEvents

        jsonRequest = BuildJsonFromRow(wsSrc, r, lastCol)

        Dim success As Boolean: success = False
        Dim attempts As Long
        Dim waitList As Variant: waitList = Array(300, 800, 1500, 2500, 4000)


        ' -----------------------------------------------------------------
        '   5 TENTATIVES HTTP
        ' -----------------------------------------------------------------
        For attempts = 0 To 4

            If ShouldStop() Then GoTo EndProcess

            Set http = CreateObject("MSXML2.XMLHTTP")
            http.Open "POST", url, False
            http.setRequestHeader "Content-Type", "application/json"

            On Error Resume Next
                http.send jsonRequest
            On Error GoTo 0

            ' Timeout réception 5s
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


        ' -----------------------------------------------------------------
        '   SUCCÈS → écriture JSON
        ' -----------------------------------------------------------------
        If success Then
            WriteJsonToSheet wsDst, dstRow, jsonResponse

        ' -----------------------------------------------------------------
        '   ÉCHEC → copie brute
        ' -----------------------------------------------------------------
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
            " | Temps restant : " & Format(remaining / 60, "0.0") & " min"

    Next r

    MsgBox "Traitement terminé !", vbInformation
    GoTo EndProcessOK



' =========================================================================
'   SORTIE PROPRE
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
'   CONSTRUCTION JSON
' =========================================================================
Public Function BuildJsonFromRow(ws As Worksheet, row As Long, lastCol As Long) As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim col As Long

    For col = 1 To lastCol
        dict(ws.Cells(1, col).Value) = CStr(ws.Cells(row, col).Value)
    Next col

    BuildJsonFromRow = DictToJson(dict)
End Function



' =========================================================================
'   DICTIONNAIRE → JSON
' =========================================================================
Public Function DictToJson(dict As Object) As String
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
'   JSON PARSER SIMPLE
' =========================================================================
Public Function ParseSimpleJsonObject(json As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim cleaned As String: cleaned = Trim(json)
    If cleaned = "" Then Set ParseSimpleJsonObject = dict: Exit Function

    ' Supprime crochets []
    If Left(cleaned, 1) = "[" Then cleaned = Mid(cleaned, 2)
    If Right(cleaned, 1) = "]" Then cleaned = Left(cleaned, Len(cleaned) - 1)

    cleaned = Trim(cleaned)
    If cleaned = "" Then Set ParseSimpleJsonObject = dict: Exit Function

    ' Supprime accolades {}
    If Left(cleaned, 1) = "{" Then cleaned = Mid(cleaned, 2)
    If Right(cleaned, 1) = "}" Then cleaned = Left(cleaned, Len(cleaned) - 1)

    cleaned = Trim(cleaned)
    If cleaned = "" Then Set ParseSimpleJsonObject = dict: Exit Function

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
