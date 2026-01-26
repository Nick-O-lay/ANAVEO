Attribute VB_Name = "RemoteCore"
Option Explicit

' =========================================================================
'   API Sleep (non bloquant grâce à SleepStop)
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

    ' =========================================================================
    '   RESET destination (NOUVEAU FORMAT A:X)
    ' =========================================================================
    wsDst.Range("A3:X100000").ClearContents

    wsDst.Range("A2:X2").Value = Array( _
        "Société", "Origine", "Marché", "Enseigne SalesForce", _
        "Siège social", "Création établissement", "Effectifs", "Genre", _
        "Représentant", "Score", "Téléphone", "Email", "Commentaire", _
        "ESS", "Métier", "Catégorie entreprise", _
        "Longitude", "Latitude", "Adresse", "Code postal", "Ville", _
        "Siren", "Siret", "CA" _
    )

    ' Formats texte (évite perte de 0 / scientific)
    wsDst.Columns("T").NumberFormat = "@"  ' Code postal (col 20)
    wsDst.Columns("V").NumberFormat = "@"  ' Siren (col 22)
    wsDst.Columns("W").NumberFormat = "@"  ' Siret (col 23)

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

            jsonResponse = Trim$(http.responseText)

            If Len(jsonResponse) > 0 And InStr(1, jsonResponse, "{", vbTextCompare) > 0 Then
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
            ' copie brute si échec (sur 24 colonnes max pour rester cohérent)
            Dim c As Long
            For c = 1 To Application.Min(lastCol, 24)
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
'   JSON → FEUILLE (écrit dans l’ordre des entêtes en ligne 2)
' =========================================================================
Public Sub WriteJsonToSheet(ws As Worksheet, row As Long, jsonText As String)

    Dim obj As Object
    Set obj = ParseSimpleJsonObject(jsonText)
    If obj Is Nothing Then Exit Sub

    Dim lastCol As Long, c As Long, header As String
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        header = CStr(ws.Cells(2, c).Value)
        If header <> "" Then
            If obj.Exists(header) Then
                ws.Cells(row, c).Value = obj(header)
            Else
                ws.Cells(row, c).Value = ""
            End If
        End If
    Next c
End Sub

' =========================================================================
'   JSON PARSER (quote-safe : ne casse pas sur virgules dans les valeurs)
' =========================================================================
Public Function ParseSimpleJsonObject(json As String) As Object

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim s As String: s = Trim$(json)

    If s = "" Then
        Set ParseSimpleJsonObject = dict
        Exit Function
    End If

    ' Retire [] si besoin
    If Left$(s, 1) = "[" Then s = Mid$(s, 2)
    If Right$(s, 1) = "]" Then s = Left$(s, Len(s) - 1)
    s = Trim$(s)

    ' Retire {}
    If Left$(s, 1) = "{" Then s = Mid$(s, 2)
    If Right$(s, 1) = "}" Then s = Left$(s, Len(s) - 1)
    s = Trim$(s)

    Dim pairs As Collection
    Set pairs = SplitJsonPairs(s)

    Dim i As Long, p As String, k As String, v As String
    For i = 1 To pairs.Count
        p = pairs(i)
        If SplitJsonKV(p, k, v) Then
            dict(k) = v
        End If
    Next i

    Set ParseSimpleJsonObject = dict
End Function

Private Function SplitJsonPairs(ByVal s As String) As Collection
    Dim col As New Collection
    Dim i As Long, ch As String
    Dim inQ As Boolean: inQ = False
    Dim esc As Boolean: esc = False
    Dim startPos As Long: startPos = 1

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If esc Then
            esc = False
        ElseIf ch = "\" Then
            esc = True
        ElseIf ch = """" Then
            inQ = Not inQ
        ElseIf ch = "," And Not inQ Then
            col.Add Trim$(Mid$(s, startPos, i - startPos))
            startPos = i + 1
        End If
    Next i

    If startPos <= Len(s) Then col.Add Trim$(Mid$(s, startPos))
    Set SplitJsonPairs = col
End Function

Private Function SplitJsonKV(ByVal pair As String, ByRef k As String, ByRef v As String) As Boolean
    Dim i As Long, ch As String
    Dim inQ As Boolean: inQ = False
    Dim esc As Boolean: esc = False

    For i = 1 To Len(pair)
        ch = Mid$(pair, i, 1)

        If esc Then
            esc = False
        ElseIf ch = "\" Then
            esc = True
        ElseIf ch = """" Then
            inQ = Not inQ
        ElseIf ch = ":" And Not inQ Then
            k = Trim$(Left$(pair, i - 1))
            v = Trim$(Mid$(pair, i + 1))
            k = UnquoteJson(k)
            v = UnquoteJson(v)
            SplitJsonKV = True
            Exit Function
        End If
    Next i

    SplitJsonKV = False
End Function

Private Function UnquoteJson(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) >= 2 Then
        If Left$(s, 1) = """" And Right$(s, 1) = """" Then
            s = Mid$(s, 2, Len(s) - 2)
        End If
    End If

    ' Unescape minimal
    s = Replace(s, "\\", "\")
    s = Replace(s, "\""",""")

    UnquoteJson = s
End Function

' =========================================================================
'   CONSTRUCTION JSON (envoi vers n8n)
' =========================================================================
Public Function BuildJsonFromRow(ws As Worksheet, row As Long, lastCol As Long) As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To lastCol
        dict(CStr(ws.Cells(1, c).Value)) = CStr(ws.Cells(row, c).Value)
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
