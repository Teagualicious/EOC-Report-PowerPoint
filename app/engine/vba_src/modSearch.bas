Attribute VB_Name = "modSearch"
Option Explicit

Private Const RESULTS_ROW As Long = 12
Private Const SUGGEST_ROW As Long = 9

' Cached _SearchIndex / _Config (bulk-loaded once — per-cell COM reads on
' every keystroke were the slowness culprit)
Private mIdx As Variant          ' 2D: term|kind|canonical|display
Private mIdxN As Long
Private mCfg As Variant          ' 2D: metric|enabled|order|agg
Private mCfgN As Long

' Extent of the currently displayed results table (set by WriteResults)
Private mResHeaderRow As Long
Private mDimSeq(1 To 3) As String   ' dimension order (subset of mColSeq)
Private mDimN As Long
' Unified column sequence — the search is a TYPED PIVOT TABLE: columns
' render in the exact order the user typed them, dimensions and metrics
' interleaved ("client"/"camp"/"level" tokens, or "m:<metric name>").
Private mColSeq(1 To 67) As String
Private mColN As Long
Private mResLastRow As Long
Private mResLastCol As Long

Private Sub EnsureCaches()
    Dim ws As Worksheet, lastRow As Long
    If mIdxN = 0 Then
        Set ws = ThisWorkbook.Worksheets("_SearchIndex")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then
            mIdx = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 4)).Value
            mIdxN = UBound(mIdx, 1)
        End If
    End If
    If mCfgN = 0 Then
        Set ws = ThisWorkbook.Worksheets("_Config")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then
            mCfg = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 4)).Value
            mCfgN = UBound(mCfg, 1)
        End If
    End If
End Sub

Public Sub InvalidateCaches()
    mIdxN = 0
    mCfgN = 0
End Sub

' ============================== ENTRY POINTS ==============================

Public Sub RunSearch()
    On Error GoTo fail
    EnsureCaches
    Dim wsS As Worksheet, wsU As Worksheet
    Set wsS = ThisWorkbook.Worksheets("Search")
    Set wsU = ThisWorkbook.Worksheets("Unified Data")

    Dim raw As String
    raw = Trim$(CStr(wsS.Range("SearchCell").Value))

    ' ---- Parse ----
    Dim metrics() As String, metricCount As Long
    Dim levelType As String, campGroup As Boolean, clientGroup As Boolean
    Dim levelValues As Object, camps As Object
    Set levelValues = CreateObject("Scripting.Dictionary")
    Set camps = CreateObject("Scripting.Dictionary")
    levelValues.CompareMode = vbTextCompare
    camps.CompareMode = vbTextCompare
    Dim unknownTerms As String
    Dim colOrder As String
    colOrder = "|"
    ParseTerms raw, metrics, metricCount, levelType, levelValues, camps, _
               campGroup, clientGroup, unknownTerms, colOrder

    ' No metric typed -> the _Config defaults fill in, appended AFTER the
    ' typed columns so the typed layout is untouched
    Dim typedMetrics As Boolean
    typedMetrics = (metricCount > 0)
    If metricCount = 0 Then DefaultMetrics metrics, metricCount

    ' Rate metrics derive as 100 * completions / impressions — accumulate
    ' both even when not displayed (approved: sum/sum, never avg of avgs)
    Dim rateNeeded As Boolean, numName As String
    rateNeeded = False
    Dim jr As Long
    For jr = 1 To metricCount
        If MetricAgg(metrics(jr)) = "rate" Then rateNeeded = True
    Next jr
    If rateNeeded Then numName = PickNumerator()

    ' ---- Row dimensions ----
    ' NEW RULE (approved): Campaign appears ONLY when explicitly typed —
    ' the word "campaign"/"client" or an actual campaign name. Searching
    ' a level VALUE (e.g. a zip) now shows THAT level as the rows,
    ' aggregated across campaigns, instead of forcing campaign rows in.
    Dim useLevel As Boolean, useCamp As Boolean, useClient As Boolean
    If levelType = "" And levelValues.Count > 0 Then
        Dim v0 As Variant
        For Each v0 In levelValues.keys
            levelType = Left$(CStr(v0), InStr(CStr(v0), ":") - 1)
            Exit For
        Next v0
    End If
    useLevel = (levelType <> "")
    useCamp = campGroup Or (camps.Count > 0) Or (Not useLevel)
    useClient = clientGroup

    ' Column layout = EXACT typed order (this is a typed pivot table).
    ' Implied-but-untyped dimensions (default campaign rows, campaign-name
    ' filters) go FIRST as traditional label columns; then every typed
    ' token in the order written; then default metrics when none were typed.
    mColN = 0
    Dim ord As Variant, tok As Variant
    If useClient And InStr(colOrder, "|client|") = 0 Then
        mColN = mColN + 1: mColSeq(mColN) = "client"
    End If
    If useCamp And InStr(colOrder, "|camp|") = 0 Then
        mColN = mColN + 1: mColSeq(mColN) = "camp"
    End If
    If useLevel And InStr(colOrder, "|level|") = 0 Then
        mColN = mColN + 1: mColSeq(mColN) = "level"
    End If
    ord = Split(Mid$(colOrder, 2), "|")
    For Each tok In ord
        If CStr(tok) = "client" Then
            If useClient Then mColN = mColN + 1: mColSeq(mColN) = "client"
        ElseIf CStr(tok) = "camp" Then
            If useCamp Then mColN = mColN + 1: mColSeq(mColN) = "camp"
        ElseIf CStr(tok) = "level" Then
            If useLevel Then mColN = mColN + 1: mColSeq(mColN) = "level"
        ElseIf Left$(CStr(tok), 2) = "m:" Then
            mColN = mColN + 1: mColSeq(mColN) = CStr(tok)
        End If
    Next tok
    If Not typedMetrics Then
        Dim jm As Long
        For jm = 1 To metricCount
            mColN = mColN + 1: mColSeq(mColN) = "m:" & metrics(jm)
        Next jm
    End If

    ' Dimension subsequence (row-key build order)
    mDimN = 0
    Dim ci As Long
    For ci = 1 To mColN
        If Left$(mColSeq(ci), 2) <> "m:" Then
            mDimN = mDimN + 1
            mDimSeq(mDimN) = mColSeq(ci)
        End If
    Next ci

    ' ---- Load unified data ----
    Dim data As Variant, lastRow As Long
    lastRow = wsU.Cells(wsU.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    data = wsU.Range(wsU.Cells(1, 1), wsU.Cells(lastRow, 9)).Value

    Dim cClient As Long, cCamp As Long, cLevel As Long, cMetric As Long
    Dim cValue As Long, cStart As Long, cEnd As Long, j As Long
    For j = 1 To UBound(data, 2)
        Select Case LCase$(CStr(data(1, j)))
            Case "client": cClient = j
            Case "campaign": cCamp = j
            Case "metric_level": cLevel = j
            Case "metric_name": cMetric = j
            Case "metric_value": cValue = j
            Case "start_date": cStart = j
            Case "end_date": cEnd = j
        End Select
    Next j

    Dim dFrom As Date, dTo As Date, hasFrom As Boolean, hasTo As Boolean
    GetDateFilter wsS.Range("DateStart").Value, dFrom, hasFrom
    GetDateFilter wsS.Range("DateEnd").Value, dTo, hasTo

    ' ---- Aggregate ----
    Dim sums As Object, counts As Object, rowKeys As Object
    Set sums = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    Set rowKeys = CreateObject("Scripting.Dictionary")
    Dim collapseSources As Boolean
    collapseSources = (Not useLevel) And (levelValues.Count = 0)

    Dim i As Long
    For i = 2 To UBound(data, 1)
        Dim camp As String, lvl As String, met As String
        camp = CStr(data(i, cCamp))
        lvl = CStr(data(i, cLevel))
        met = CStr(data(i, cMetric))
        If met = "" Then GoTo nextRow
        If camps.Count > 0 And Not camps.Exists(camp) Then GoTo nextRow
        Dim forRate As Boolean
        forRate = rateNeeded And (StrComp(met, "Impressions", vbTextCompare) = 0 _
                                  Or StrComp(met, numName, vbTextCompare) = 0)
        If Not MetricWanted(met, metrics, metricCount) And Not forRate Then GoTo nextRow
        If hasFrom Or hasTo Then
            If Not RowInDateRange(data(i, cStart), data(i, cEnd), _
                                  dFrom, hasFrom, dTo, hasTo, lvl) Then GoTo nextRow
        End If
        If levelValues.Count > 0 Then
            If Not levelValues.Exists(lvl) Then GoTo nextRow
        End If
        If useLevel Then
            If Left$(lvl, Len(levelType) + 1) <> levelType & ":" Then GoTo nextRow
        End If

        ' Composite row key: client | campaign | level value (as selected)
        Dim keyParts As String, grp As String
        keyParts = ""
        Dim di As Long
        For di = 1 To mDimN
            Select Case mDimSeq(di)
                Case "client": keyParts = keyParts & Chr$(2) & CStr(data(i, cClient))
                Case "camp":   keyParts = keyParts & Chr$(2) & camp
                Case "level":  keyParts = keyParts & Chr$(2) & Mid$(lvl, Len(levelType) + 2)
            End Select
        Next di
        If Left$(keyParts, 1) = Chr$(2) Then keyParts = Mid$(keyParts, 2)
        If keyParts = "" Then GoTo nextRow

        grp = ""
        If collapseSources Then
            If InStr(lvl, ":") > 0 Then
                grp = Left$(lvl, InStr(lvl, ":") - 1)
            Else
                grp = "campaign"
            End If
        End If

        Dim k As String
        If MetricWanted(met, metrics, metricCount) Then
            k = keyParts & Chr$(1) & met & Chr$(1) & grp
            If Not sums.Exists(k) Then
                sums(k) = 0#
                counts(k) = 0&
            End If
            sums(k) = sums(k) + Val(CStr(data(i, cValue)))
            counts(k) = counts(k) + 1
            If Not rowKeys.Exists(keyParts) Then rowKeys(keyParts) = rowKeys.Count
        End If
        If forRate Then
            Dim rtag As String
            If StrComp(met, "Impressions", vbTextCompare) = 0 Then
                rtag = Chr$(4) & "IMP"
            Else
                rtag = Chr$(4) & "NUM"
            End If
            k = keyParts & Chr$(1) & rtag & Chr$(1) & grp
            If Not sums.Exists(k) Then
                sums(k) = 0#
                counts(k) = 0&
            End If
            sums(k) = sums(k) + Val(CStr(data(i, cValue)))
            counts(k) = counts(k) + 1
        End If
nextRow:
    Next i

    ' Collapse source groups (totals view): best source per row+metric,
    ' matching the app KPI rule
    If collapseSources Then
        Dim collapsed As Object, colCounts As Object, kk As Variant
        Set collapsed = CreateObject("Scripting.Dictionary")
        Set colCounts = CreateObject("Scripting.Dictionary")
        For Each kk In sums.keys
            Dim parts() As String, base As String
            parts = Split(CStr(kk), Chr$(1))
            base = parts(0) & Chr$(1) & parts(1)
            If Not collapsed.Exists(base) Then
                collapsed(base) = sums(kk)
                colCounts(base) = counts(kk)
            ElseIf sums(kk) > collapsed(base) Then
                collapsed(base) = sums(kk)
                colCounts(base) = counts(kk)
            End If
        Next kk
        Set sums = collapsed
        Set counts = colCounts
    End If

    WriteResults wsS, rowKeys, sums, counts, metrics, metricCount, _
                 levelType, useClient, useCamp, useLevel, unknownTerms
    Exit Sub
fail:
    wsS.Cells(RESULTS_ROW, 2).Value = "Search error: " & Err.Description
End Sub

' ============================== PARSING ==============================

Private Sub ParseTerms(raw As String, metrics() As String, metricCount As Long, _
                       levelType As String, levelValues As Object, _
                       camps As Object, campGroup As Boolean, _
                       clientGroup As Boolean, unknownTerms As String, _
                       ByRef colOrder As String)
    ' colOrder records the EXACT typed token order ("m:<name>", "level",
    ' "camp", "client") — WriteResults renders columns in this order.
    ReDim metrics(1 To 64)
    metricCount = 0
    levelType = ""
    unknownTerms = ""
    campGroup = False
    clientGroup = False
    If Trim$(raw) = "" Then Exit Sub
    Dim pcs() As String, p As Variant, term As String
    pcs = Split(raw, ",")
    For Each p In pcs
        term = Trim$(CStr(p))
        If term <> "" Then
            Dim kind As String, canonical As String
            ResolveTerm term, kind, canonical
            Select Case kind
                Case "metric"
                    If Not MetricWanted(canonical, metrics, metricCount) Then
                        metricCount = metricCount + 1
                        metrics(metricCount) = canonical
                        colOrder = colOrder & "m:" & canonical & "|"
                    End If
                Case "level_type"
                    levelType = canonical
                    If InStr(colOrder, "|level|") = 0 Then colOrder = colOrder & "level|"
                Case "level_value"
                    Dim vType As String, ok As Boolean
                    vType = Left$(canonical, InStr(canonical, ":") - 1)
                    ok = True
                    If levelValues.Count > 0 Then
                        Dim k0 As Variant, t0 As String
                        For Each k0 In levelValues.keys
                            t0 = Left$(CStr(k0), InStr(CStr(k0), ":") - 1)
                            Exit For
                        Next k0
                        If StrComp(t0, vType, vbTextCompare) <> 0 Then
                            unknownTerms = AppendTerm(unknownTerms, _
                                term & " (mixed level types)")
                            ok = False
                        End If
                    End If
                    If ok Then
                        levelValues(canonical) = True
                        If InStr(colOrder, "|level|") = 0 Then colOrder = colOrder & "level|"
                    End If
                Case "campaign"
                    camps(canonical) = True
                Case "row_group"
                    If canonical = "campaign" Then
                        campGroup = True
                        If InStr(colOrder, "|camp|") = 0 Then colOrder = colOrder & "camp|"
                    End If
                    If canonical = "client" Then
                        clientGroup = True
                        If InStr(colOrder, "|client|") = 0 Then colOrder = colOrder & "client|"
                    End If
                Case Else
                    unknownTerms = AppendTerm(unknownTerms, term)
            End Select
        End If
    Next p
End Sub

Private Function AppendTerm(agg As String, term As String) As String
    If agg <> "" Then agg = agg & ", "
    AppendTerm = agg & term
End Function

Public Sub ResolveTerm(term As String, ByRef kind As String, ByRef canonical As String)
    ' In-memory over the cached index.
    ' exact/prefix passes: level_value > metric > level_type > row_group >
    '   campaign (typing "281" should hit zip values first).
    ' contains pass: values are the noisiest, so they drop to LAST —
    '   "work" should find the network TYPE, not "ACC Network".
    EnsureCaches
    Dim t As String
    t = LCase$(Trim$(term))
    kind = ""
    canonical = ""
    If mIdxN = 0 Or t = "" Then Exit Sub
    Dim passesA As Variant, passesC As Variant, passes As Variant
    Dim pass As Variant, mode As Long, i As Long
    passesA = Array("level_value", "metric", "level_type", "row_group", "campaign")
    passesC = Array("metric", "level_type", "row_group", "campaign", "level_value")
    For mode = 1 To 3
        If mode = 3 Then passes = passesC Else passes = passesA
        For Each pass In passes
            For i = 1 To mIdxN
                If CStr(mIdx(i, 2)) = pass Then
                    Dim hay As String
                    hay = LCase$(CStr(mIdx(i, 1)))
                    If (mode = 1 And hay = t) _
                       Or (mode = 2 And InStr(1, hay, t) = 1) _
                       Or (mode = 3 And InStr(1, hay, t) > 0) Then
                        kind = CStr(pass)
                        canonical = CStr(mIdx(i, 3))
                        Exit Sub
                    End If
                End If
            Next i
        Next pass
    Next mode
End Sub

Private Sub DefaultMetrics(metrics() As String, ByRef metricCount As Long)
    EnsureCaches
    Dim i As Long
    ReDim metrics(1 To mCfgN + 1)
    For i = 1 To mCfgN
        If UCase$(CStr(mCfg(i, 2))) = "Y" Then
            metricCount = metricCount + 1
            metrics(metricCount) = CStr(mCfg(i, 1))
        End If
    Next i
End Sub

Private Function MetricWanted(met As String, metrics() As String, _
                              metricCount As Long) As Boolean
    Dim i As Long
    For i = 1 To metricCount
        If StrComp(met, metrics(i), vbTextCompare) = 0 Then
            MetricWanted = True
            Exit Function
        End If
    Next i
    MetricWanted = False
End Function

Private Function PickNumerator() As String
    EnsureCaches
    Dim prefs As Variant, pn As Variant, i As Long
    prefs = Array("Contributions", "100% Completions", "Completions")
    For Each pn In prefs
        For i = 1 To mCfgN
            If StrComp(CStr(mCfg(i, 1)), CStr(pn), vbTextCompare) = 0 Then
                PickNumerator = CStr(pn)
                Exit Function
            End If
        Next i
    Next pn
    PickNumerator = "Contributions"
End Function

Private Function MetricAgg(met As String) As String
    EnsureCaches
    Dim i As Long
    For i = 1 To mCfgN
        If StrComp(CStr(mCfg(i, 1)), met, vbTextCompare) = 0 Then
            MetricAgg = CStr(mCfg(i, 4))
            Exit Function
        End If
    Next i
    MetricAgg = "sum"
End Function

Private Sub GetDateFilter(v As Variant, ByRef d As Date, ByRef hasIt As Boolean)
    hasIt = False
    On Error Resume Next
    If Not IsEmpty(v) And CStr(v) <> "" Then
        d = CDate(v)
        hasIt = (Err.Number = 0)
    End If
    Err.Clear
End Sub

Private Function RowInDateRange(vStart As Variant, vEnd As Variant, _
        dFrom As Date, hasFrom As Boolean, dTo As Date, hasTo As Boolean, _
        lvl As String) As Boolean
    On Error GoTo lenient
    Dim d1 As Date, d2 As Date
    If Left$(lvl, 5) = "date:" Then
        d1 = CDate(Mid$(lvl, 6))
        d2 = d1
    Else
        d1 = CDate(vStart)
        d2 = CDate(vEnd)
    End If
    If hasFrom And d2 < dFrom Then Exit Function
    If hasTo And d1 > dTo Then Exit Function
    RowInDateRange = True
    Exit Function
lenient:
    RowInDateRange = True
End Function

' ============================== OUTPUT ==============================

Private Sub WriteResults(wsS As Worksheet, rowKeys As Object, sums As Object, _
        counts As Object, metrics() As String, metricCount As Long, _
        levelType As String, useClient As Boolean, useCamp As Boolean, _
        useLevel As Boolean, unknownTerms As String)
    ' Typed pivot table: ONLY the results table renders (no auto KPI strip —
    ' it added summary boxes and aggregations the user never asked for),
    ' and columns follow mColSeq: the exact typed order, dimensions and
    ' metrics interleaved as written.
    Application.ScreenUpdating = False
    wsS.Range(wsS.Cells(RESULTS_ROW, 2), wsS.Cells(RESULTS_ROW + 5000, 60)).Clear

    Dim r As Long
    r = RESULTS_ROW
    If unknownTerms <> "" Then
        wsS.Cells(r, 2).Value = "Ignored (no match): " & unknownTerms
        wsS.Cells(r, 2).Font.Italic = True
        wsS.Cells(r, 2).Font.Color = RGB(180, 100, 0)
        r = r + 1
    End If

    ' Header in unified column order
    Dim c As Long, tok As String
    For c = 1 To mColN
        tok = mColSeq(c)
        If tok = "client" Then
            wsS.Cells(r, 1 + c).Value = "Client"
        ElseIf tok = "camp" Then
            wsS.Cells(r, 1 + c).Value = "Campaign"
        ElseIf tok = "level" Then
            wsS.Cells(r, 1 + c).Value = _
                Application.WorksheetFunction.Proper(levelType)
        Else
            wsS.Cells(r, 1 + c).Value = Mid$(tok, 3)
        End If
    Next c
    With wsS.Range(wsS.Cells(r, 2), wsS.Cells(r, 1 + mColN))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 48, 87)
    End With
    mResHeaderRow = r
    mResLastCol = 1 + mColN
    r = r + 1

    Dim n As Long
    n = rowKeys.Count
    If n = 0 Then
        wsS.Cells(r, 2).Value = "No matching data."
        mResHeaderRow = 0   ' nothing copyable
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Dim keys() As Variant
    keys = rowKeys.keys
    SortStrings keys

    Dim ki As Long
    For ki = LBound(keys) To UBound(keys)
        Dim dimParts() As String, di As Long
        dimParts = Split(CStr(keys(ki)), Chr$(2))
        di = 0
        For c = 1 To mColN
            tok = mColSeq(c)
            If Left$(tok, 2) <> "m:" Then
                If di <= UBound(dimParts) Then
                    wsS.Cells(r, 1 + c).Value = dimParts(di)
                End If
                di = di + 1
            Else
                Dim met As String, k As String
                met = Mid$(tok, 3)
                k = CStr(keys(ki)) & Chr$(1) & met
                If Not sums.Exists(k) Then k = k & Chr$(1)
                If MetricAgg(met) = "rate" Then
                    Dim kImp As String, kNum As String
                    kImp = CStr(keys(ki)) & Chr$(1) & Chr$(4) & "IMP"
                    kNum = CStr(keys(ki)) & Chr$(1) & Chr$(4) & "NUM"
                    If Not sums.Exists(kImp) Then kImp = kImp & Chr$(1)
                    If Not sums.Exists(kNum) Then kNum = kNum & Chr$(1)
                    If sums.Exists(kImp) And sums.Exists(kNum) Then
                        If sums(kImp) > 0 Then
                            wsS.Cells(r, 1 + c).NumberFormat = "0.00""%"""
                            wsS.Cells(r, 1 + c).Value = _
                                100# * sums(kNum) / sums(kImp)
                            GoTo nextCol
                        End If
                    End If
                End If
                If sums.Exists(k) Then
                    Dim v As Double
                    If MetricAgg(met) = "avg" Or MetricAgg(met) = "rate" Then
                        v = sums(k) / counts(k)
                        wsS.Cells(r, 1 + c).NumberFormat = "#,##0.00"
                    Else
                        v = sums(k)
                        If v = Int(v) Then
                            wsS.Cells(r, 1 + c).NumberFormat = "#,##0"
                        Else
                            wsS.Cells(r, 1 + c).NumberFormat = "#,##0.00"
                        End If
                    End If
                    wsS.Cells(r, 1 + c).Value = v
                End If
            End If
nextCol:
        Next c
        If (r - RESULTS_ROW) Mod 2 = 0 Then
            wsS.Range(wsS.Cells(r, 2), wsS.Cells(r, 1 + mColN)) _
               .Interior.Color = RGB(238, 242, 247)
        End If
        r = r + 1
    Next ki
    On Error Resume Next
    wsS.Range(wsS.Cells(mResHeaderRow, 2), wsS.Cells(r, 1 + mColN)) _
       .Columns.AutoFit
    mResLastRow = r - 1
    If mResLastRow > mResHeaderRow Then
        Dim dcol As Range, db As Object
        For c = 1 To mColN
            If Left$(mColSeq(c), 2) = "m:" Then
                Set dcol = wsS.Range(wsS.Cells(mResHeaderRow + 1, 1 + c), _
                                     wsS.Cells(mResLastRow, 1 + c))
                Set db = dcol.FormatConditions.AddDatabar
                db.BarColor.Color = RGB(187, 214, 242)
                db.ShowValue = True
            End If
        Next c
    End If
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

' ============================== COPY ==============================

Public Sub CopyResults()
    ' One click: put the currently displayed results table (header +
    ' rows, formatting included) on the clipboard.
    On Error GoTo nothingToCopy
    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Worksheets("Search")
    If mResHeaderRow < RESULTS_ROW Or mResLastRow < mResHeaderRow _
       Or mResLastCol < 2 Then GoTo nothingToCopy
    wsS.Range(wsS.Cells(mResHeaderRow, 2), _
              wsS.Cells(mResLastRow, mResLastCol)).Copy
    FlashCopyChip ChrW$(&H2713) & " Copied " & _
                  (mResLastRow - mResHeaderRow) & " rows"
    Exit Sub
nothingToCopy:
    FlashCopyChip "Run a search first"
End Sub

Private Sub FlashCopyChip(msg As String)
    On Error Resume Next
    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Worksheets("Search")
    wsS.Range("D11").Value = msg
    Application.OnTime Now + TimeSerial(0, 0, 2), "modSearch.RestoreCopyChip"
End Sub

Public Sub RestoreCopyChip()
    On Error Resume Next
    ' U+1F4CB is above the BMP — VBA needs the surrogate pair
    ThisWorkbook.Worksheets("Search").Range("D11").Value = _
        ChrW$(&HD83D) & ChrW$(&HDCCB) & " Copy results"
End Sub

Private Sub SortStrings(arr() As Variant)
    Dim i As Long, jj As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For jj = i + 1 To UBound(arr)
            If StrComp(CStr(arr(i)), CStr(arr(jj)), vbTextCompare) > 0 Then
                tmp = arr(i)
                arr(i) = arr(jj)
                arr(jj) = tmp
            End If
        Next jj
    Next i
End Sub

' ============================== SUGGESTIONS ==============================

Public Sub UpdateSuggestions()
    UpdateSuggestionsFor CStr(ThisWorkbook.Worksheets("Search") _
                              .Range("SearchCell").Value)
End Sub

Public Sub UpdateSuggestionsFor(raw As String)
    On Error Resume Next
    EnsureCaches
    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Worksheets("Search")

    Dim frag As String, pcs() As String
    If InStr(raw, ",") > 0 Then
        pcs = Split(raw, ",")
        frag = Trim$(pcs(UBound(pcs)))
    Else
        frag = Trim$(raw)
    End If

    wsS.Range(wsS.Cells(SUGGEST_ROW, 2), wsS.Cells(SUGGEST_ROW, 40)).Clear
    If Len(frag) < 1 Or mIdxN = 0 Then Exit Sub

    Dim i As Long, hits As Long, seen As Object, mode As Long
    Set seen = CreateObject("Scripting.Dictionary")
    ' prefix matches first, then contains
    For mode = 1 To 2
        For i = 1 To mIdxN
            Dim hay As String, disp As String
            hay = LCase$(CStr(mIdx(i, 1)))
            disp = CStr(mIdx(i, 4))
            If Not seen.Exists(disp) Then
                If (mode = 1 And InStr(1, hay, LCase$(frag)) = 1) _
                   Or (mode = 2 And InStr(1, hay, LCase$(frag)) > 1) Then
                    seen(disp) = True
                    hits = hits + 1
                    ' Alternating columns = visible gaps between chips
                    Dim chipCols As Variant
                    chipCols = Array(2, 4, 6, 8, 10, 13)
                    If hits > 6 Then Exit Sub
                    With wsS.Cells(SUGGEST_ROW, chipCols(hits - 1))
                        .Value = disp
                        .Interior.Color = RGB(222, 232, 245)
                        .Font.Size = 9
                        .Font.Color = RGB(0, 48, 87)
                    End With
                    If hits >= 6 Then Exit Sub
                End If
            End If
        Next i
    Next mode
End Sub

Public Sub ApplySuggestion(chipText As String)
    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Worksheets("Search")
    Dim clean As String
    clean = chipText
    If InStr(clean, " (") > 0 Then clean = Left$(clean, InStr(clean, " (") - 1)
    ' Read the LIVE draft from the overlay textbox — reading SearchCell
    ' here resurrected the PREVIOUS committed search and appended to it
    ' (the "clicking a chip adds a ton of things" bug). The chip must
    ' only complete the last fragment of what the user is typing NOW.
    Dim raw As String
    raw = ""
    On Error Resume Next
    raw = wsS.OLEObjects("txtSearch").Object.Text
    On Error GoTo 0
    If raw = "" Then raw = CStr(wsS.Range("SearchCell").Value)
    If InStr(raw, ",") > 0 Then
        raw = Left$(raw, InStrRev(raw, ",")) & " " & clean
    Else
        raw = clean
    End If
    On Error Resume Next
    wsS.OLEObjects("txtSearch").Object.Text = raw
    On Error GoTo 0
    wsS.Range("SearchCell").Value = raw
    wsS.Range(wsS.Cells(SUGGEST_ROW, 2), wsS.Cells(SUGGEST_ROW, 40)).Clear
    RunSearch
End Sub

Public Sub ShowCalendarFor(targetName As String)
    On Error Resume Next
    frmCalendar.TargetRangeName = targetName
    frmCalendar.Show
End Sub
