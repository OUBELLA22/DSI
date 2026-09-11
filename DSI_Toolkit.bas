Attribute VB_Name = "DSI_Toolkit"
'==============================================================================
' DSI Toolkit -- VBA edition
'
' Same job as the Python tool: read one DSI and show chosen elements, or read
' two and compare them.
'
' Import into any workbook: Alt+F11 -> File -> Import File -> DSI_Toolkit.bas
' Run: Alt+F8 -> DSI_Run
'
' Why this is not the old macro rewritten
' ---------------------------------------
' The original comparison subs scanned file 2 from the top for every row of
' file 1. On 1215 terminals that is 1.4 million comparisons through the Excel
' object model. Here rows are indexed once into a Dictionary, so matching is
' one pass. VBA was never the reason that was slow.
'
' Three other deliberate differences:
'
'  1. No "On Error Resume Next". A failure stops and says what happened.
'     grapheCAPH silently produced an empty Y column for years because an
'     error was swallowed.
'  2. Section headers are read for BOTH control characters. The exporter writes
'     seven headers with "!" instead of "%" -- terminals, seals, plug, clips,
'     grommets, other components, extra node components -- and they carry real
'     data. Splitting on "%" alone fuses them into one 1707-row block.
'  3. Sheets are written from a Variant array in a single assignment, not cell
'     by cell. This is the difference between seconds and minutes.
'
' Match keys live in RuleTable(). One line per section. That replaces the
' knowledge previously spread across ~60 near-identical subs.
'
' UTF-8 is read through ADODB.Stream because Open/Input assumes ANSI and would
' corrupt accented descriptions.
'==============================================================================
Option Explicit

Private Const KEYSEP As String = vbTab           ' joins key fields, absent from DSI data
Private Const V_OK As String = "OK"
Private Const V_MOD As String = "MODIFY"
Private Const V_ADD As String = "ADD"
Private Const V_REM As String = "REMOVE"

' Excel colour constants (interior)
Private Const C_YELLOW As Long = 10087423        ' modified
Private Const C_RED As Long = 13551615           ' removed
Private Const C_GREEN As Long = 13561798         ' added
Private Const C_GREY As Long = 15921906          ' not compared
Private Const C_HEADER As Long = 6299648         ' dark blue

'==============================================================================
' ENTRY POINT
'==============================================================================
Public Sub DSI_Run()
    Dim answer As String

    answer = InputBox( _
        "How many DSI files?" & vbCrLf & vbCrLf & _
        "  1  =  one file, pick the elements you want to see" & vbCrLf & _
        "  2  =  two files, compare them", _
        "DSI Toolkit", "1")

    Select Case Trim$(answer)
        Case "1": ViewOneFile
        Case "2": CompareTwoFiles
        Case "": ' cancelled
        Case Else: MsgBox "Enter 1 or 2.", vbExclamation, "DSI Toolkit"
    End Select
End Sub

'==============================================================================
' MODE 1 -- one file, chosen elements
'==============================================================================
Public Sub ViewOneFile()
    Dim path As String, dsi As Object, views_ As Variant
    Dim pick As String, chosen As Variant, i As Long, n As Long
    Dim menu As String, avail As Collection, idxList As Collection

    path = PickFile("Select the DSI file")
    If Len(path) = 0 Then Exit Sub

    Set dsi = ParseDSI(path)
    views_ = ViewTable()

    ' Only offer views that actually have rows in this file.
    Set avail = New Collection
    Set idxList = New Collection
    menu = "What do you want to see?" & vbCrLf & vbCrLf
    For i = LBound(views_) To UBound(views_)
        n = ViewRowCount(dsi, CStr(views_(i)))
        If n > 0 Then
            avail.Add views_(i)
            idxList.Add i
            menu = menu & Format$(avail.Count, "@@") & "  " & _
                   PadRight(ViewField(CStr(views_(i)), 0), 30) & _
                   Format$(n, "@@@@@@") & " rows" & vbCrLf
        End If
    Next i
    menu = menu & vbCrLf & "a  = everything" & vbCrLf & vbCrLf & _
           "Enter numbers separated by commas (e.g. 1,2,5) or a:"

    pick = InputBox(menu, "DSI Toolkit -- elements", "a")
    If Len(Trim$(pick)) = 0 Then Exit Sub

    chosen = ParseSelection(pick, avail.Count)
    If Not IsArray(chosen) Then
        MsgBox "Could not read that selection.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo Fail
    For i = LBound(chosen) To UBound(chosen)
        WriteViewSheet dsi, CStr(avail(CLng(chosen(i))))
    Next i
    Application.ScreenUpdating = True

    MsgBox "Done. " & (UBound(chosen) - LBound(chosen) + 1) & _
           " sheet(s) written." & vbCrLf & vbCrLf & _
           "File: " & path, vbInformation, "DSI Toolkit"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Failed while writing sheets:" & vbCrLf & Err.Description, vbCritical
End Sub

'==============================================================================
' MODE 2 -- two files, compare
'==============================================================================
Public Sub CompareTwoFiles()
    Dim pathA As String, pathB As String
    Dim dsiA As Object, dsiB As Object
    Dim names As Collection, nm As Variant
    Dim rules As Object, rule As Object
    Dim synth As Collection, diffs As Collection
    Dim res As Object

    pathA = PickFile("Select FILE 1 (the reference / older one)")
    If Len(pathA) = 0 Then Exit Sub
    pathB = PickFile("Select FILE 2 (the one to check)")
    If Len(pathB) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo Fail

    Set dsiA = ParseDSI(pathA)
    Set dsiB = ParseDSI(pathB)
    Set rules = BuildRules()

    Set names = UnionSectionNames(dsiA, dsiB)
    Set synth = New Collection
    Set diffs = New Collection

    For Each nm In names
        If CStr(nm) <> "End of file marker." Then
            If SectionRowCount(dsiA, CStr(nm)) > 0 Or SectionRowCount(dsiB, CStr(nm)) > 0 Then
                Set res = CompareSection(dsiA, dsiB, CStr(nm), rules)
                synth.Add res
                CollectDiffs res, diffs
            End If
        End If
    Next nm

    WriteSynthesisSheet synth, pathA, pathB
    WriteDifferencesSheet diffs
    WriteSectionSheets synth

    Application.ScreenUpdating = True
    MsgBox "Comparison finished." & vbCrLf & vbCrLf & _
           "SYNTHESIS   per-section counts and the match key used" & vbCrLf & _
           "DIFFERENCES every delta, one row each" & vbCrLf & _
           "<section>   side by side, file 1 then file 2", _
           vbInformation, "DSI Toolkit"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Comparison failed:" & vbCrLf & Err.Description, vbCritical
End Sub

'==============================================================================
' PARSER
'
' Returns a Dictionary:
'   "order" -> Collection of section names, in file order
'   "rows"  -> Dictionary: section name -> Collection of Variant field arrays
'   "delim" -> field delimiter
'   "meta"  -> Collection of "Key<tab>Value" from the banner
'==============================================================================
Public Function ParseDSI(ByVal path As String) As Object
    Dim st As Object, raw As String, lines_ As Variant
    Dim cmt As String, sep As String, dlm As String
    Dim i As Long, ln As String, payload As String
    Dim dsi As Object, rowsD As Object, order As Collection, meta As Collection
    Dim curName As String, curRows As Collection
    Dim isHeader As Boolean

    If Len(Dir$(path)) = 0 Then Err.Raise 53, , "File not found: " & path

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    raw = st.ReadText(-1)
    st.Close

    ' Normalise line endings so Split gives clean lines either way.
    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)
    lines_ = Split(raw, vbLf)

    cmt = "!": sep = "%": dlm = ":"
    ' The banner declares the control characters -- read them, don't assume.
    For i = LBound(lines_) To UBound(lines_)
        If i > 40 Then Exit For
        ln = CStr(lines_(i))
        If InStr(1, ln, "Comment Marker", vbTextCompare) > 0 Then cmt = LastCharAfterEquals(ln, cmt)
        If InStr(1, ln, "Section Separator", vbTextCompare) > 0 Then sep = LastCharAfterEquals(ln, sep)
        If InStr(1, ln, "Field Delimiter", vbTextCompare) > 0 Then dlm = LastCharAfterEquals(ln, dlm)
    Next i

    Set dsi = CreateObject("Scripting.Dictionary")
    Set rowsD = CreateObject("Scripting.Dictionary")
    Set order = New Collection
    Set meta = New Collection
    curName = ""
    Set curRows = Nothing

    For i = LBound(lines_) To UBound(lines_)
        ln = CStr(lines_(i))
        If Len(ln) = 0 Then GoTo NextLine

        If Left$(ln, Len(sep)) = sep Then
            curName = Trim$(Mid$(ln, Len(sep) + 1))
            Set curRows = NewSection(rowsD, order, curName)

        ElseIf Left$(ln, Len(cmt)) = cmt Then
            payload = Trim$(Mid$(ln, Len(cmt) + 1))
            ' A comment line is a HEADER unless it is banner text, blank, or
            ' the empty-section marker. This is the load-bearing rule.
            isHeader = (Len(payload) > 0) And (Left$(payload, 1) <> "*") _
                       And (StrComp(payload, "None", vbTextCompare) <> 0)
            If isHeader Then
                curName = payload
                Set curRows = NewSection(rowsD, order, curName)
            ElseIf Len(payload) > 0 And Left$(payload, 1) = "*" Then
                If InStr(payload, ":") > 0 Then meta.Add BannerPair(payload)
            End If

        Else
            If curRows Is Nothing Then
                Err.Raise 5, , "Data on line " & (i + 1) & " before any section header."
            End If
            curRows.Add Split(ln, dlm)
        End If
NextLine:
    Next i

    dsi.Add "order", order
    dsi.Add "rows", rowsD
    dsi.Add "delim", dlm
    dsi.Add "meta", meta
    dsi.Add "path", path
    Set ParseDSI = dsi
End Function

Private Function NewSection(ByVal rowsD As Object, ByVal order As Collection, _
                            ByVal name As String) As Collection
    Dim c As Collection
    If rowsD.Exists(name) Then
        Set NewSection = rowsD(name)          ' duplicate header: keep appending
    Else
        Set c = New Collection
        rowsD.Add name, c
        order.Add name
        Set NewSection = c
    End If
End Function

Private Function LastCharAfterEquals(ByVal ln As String, ByVal fallback As String) As String
    Dim p As Long, tail As String
    p = InStr(ln, "=")
    If p = 0 Then LastCharAfterEquals = fallback: Exit Function
    tail = Trim$(Mid$(ln, p + 1))
    If Len(tail) = 0 Then LastCharAfterEquals = fallback Else LastCharAfterEquals = Left$(tail, 1)
End Function

Private Function BannerPair(ByVal payload As String) As String
    Dim s As String, p As Long
    s = Trim$(Mid$(payload, 2))
    p = InStr(s, ":")
    If p = 0 Then BannerPair = s & KEYSEP: Exit Function
    BannerPair = Trim$(Left$(s, p - 1)) & KEYSEP & Trim$(Mid$(s, p + 1))
End Function

'------------------------------------------------------------------ accessors
Public Function SectionRows(ByVal dsi As Object, ByVal name As String) As Collection
    Dim rowsD As Object
    Set rowsD = dsi("rows")
    If rowsD.Exists(name) Then
        Set SectionRows = rowsD(name)
    Else
        Set SectionRows = New Collection
    End If
End Function

Public Function SectionRowCount(ByVal dsi As Object, ByVal name As String) As Long
    SectionRowCount = SectionRows(dsi, name).Count
End Function

' Ragged rows are normal -- branch configuration carries 21 to 25 fields row to
' row -- so positional access needs a floor rather than an error.
Public Function Fld(ByVal row As Variant, ByVal idx As Long) As String
    If idx < LBound(row) Or idx > UBound(row) Then
        Fld = ""
    Else
        Fld = CStr(row(idx))
    End If
End Function

Public Function SectionWidth(ByVal dsi As Object, ByVal name As String) As Long
    Dim r As Variant, w As Long, rows_ As Collection
    Set rows_ = SectionRows(dsi, name)
    w = 0
    For Each r In rows_
        If UBound(r) - LBound(r) + 1 > w Then w = UBound(r) - LBound(r) + 1
    Next r
    SectionWidth = w
End Function

Private Function UnionSectionNames(ByVal a As Object, ByVal b As Object) As Collection
    Dim out As Collection, seen As Object, nm As Variant
    Set out = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    For Each nm In a("order")
        If Not seen.Exists(CStr(nm)) Then seen.Add CStr(nm), 1: out.Add CStr(nm)
    Next nm
    For Each nm In b("order")
        If Not seen.Exists(CStr(nm)) Then seen.Add CStr(nm), 1: out.Add CStr(nm)
    Next nm
    Set UnionSectionNames = out
End Function

'==============================================================================
' MATCH KEY TABLE  -- the domain knowledge, one line per section
'
'   name | keys | swap | ignore
'
'   keys   "0+3"        match on fields 0 and 3
'          "0+8;0"      try 0+8, then fall back to 0
'   swap   "0-3,2-5"    field pairs that trade places when a segment is stored
'                       in the opposite direction. Applies to the key AND to
'                       the field comparison, so a branch stored A->B in one
'                       file and B->A in the other comes out OK instead of
'                       reporting its own node fields as differences.
'   ignore "2,5,19"     fields excluded from comparison
'==============================================================================
Public Function RuleTable() As Variant
    RuleTable = Array( _
        "Harness name information|0||", _
        "Harness circuit information|0||", _
        "Harness branch configuration|0+3|0-3,2-5|2,5,19,20,21,22,23", _
        "Harness wire specification|0||", _
        "Harness main node components|0||", _
        "terminals|0+1||", _
        "Harness terminals|0+1||", _
        "seals|0+1||", _
        "Harness cavity seals|0+1||", _
        "plug|0+1||", _
        "Harness cavity plugs|0+1||", _
        "clips|0||", _
        "grommets|0||", _
        "other components|0||", _
        "Harness extra node components|0+8;0||", _
        "Extra_Node_Component|0+8;0||", _
        "Harness branch insulations|0+2|0-2|4", _
        "Branch_insulation|0+2|0-2|4", _
        "Harness multicores|0||", _
        "Harness center strips|0+1||", _
        "Module child details|7||4,9", _
        "Module compatibility details|0+1|0-1|", _
        "Manual BOM quantities|0||", _
        "Harness wire through nodes|0+2||", _
        "Harness branch insulation through nodes|0+1||", _
        "Harness mid wire components|0+1||", _
        "Harness wire / multicore markers|0+1||", _
        "Harness pin mappings|0+1||", _
        "Harness Scope|0+1+2+4;0+1+2||", _
        "Composite Option Codes|0||", _
        "Harness note information|0||", _
        "Assemblies|8+26;8||", _
        "Harness assembly items|8+26;8||", _
        "Multi Location Components|0+3+8;0+3||", _
        "Harness multiple location components|0+3+8;0+3||", _
        "property|0+1+2+3;0+1+2||", _
        "tape|0||", _
        "extracavity|0+1||", _
        "WIREend|0+1||" _
    )
End Function

Private Function BuildRules() As Object
    Dim t As Variant, i As Long, parts As Variant, d As Object, r As Object
    Set d = CreateObject("Scripting.Dictionary")
    t = RuleTable()
    For i = LBound(t) To UBound(t)
        parts = Split(CStr(t(i)), "|")
        Set r = CreateObject("Scripting.Dictionary")
        r.Add "keys", parts(1)
        r.Add "swap", parts(2)
        r.Add "ignore", parts(3)
        If Not d.Exists(CStr(parts(0))) Then d.Add CStr(parts(0)), r
    Next i
    Set BuildRules = d
End Function

Private Function RuleFor(ByVal rules As Object, ByVal name As String) As Object
    Dim r As Object
    If rules.Exists(name) Then
        Set RuleFor = rules(name)
    Else
        Set r = CreateObject("Scripting.Dictionary")
        r.Add "keys", "0"
        r.Add "swap", ""
        r.Add "ignore", ""
        Set RuleFor = r
    End If
End Function

Private Function ParseIntList(ByVal s As String) As Variant
    Dim parts As Variant, out() As Long, i As Long, n As Long
    If Len(Trim$(s)) = 0 Then ParseIntList = Array(): Exit Function
    parts = Split(s, ",")
    ReDim out(LBound(parts) To UBound(parts))
    n = LBound(parts)
    For i = LBound(parts) To UBound(parts)
        out(n) = CLng(Trim$(parts(i)))
        n = n + 1
    Next i
    ParseIntList = out
End Function

Private Function SwapMap(ByVal s As String) As Object
    Dim d As Object, pairs As Variant, i As Long, ab As Variant
    Set d = CreateObject("Scripting.Dictionary")
    If Len(Trim$(s)) > 0 Then
        pairs = Split(s, ",")
        For i = LBound(pairs) To UBound(pairs)
            ab = Split(Trim$(pairs(i)), "-")
            d(CLng(ab(0))) = CLng(ab(1))
            d(CLng(ab(1))) = CLng(ab(0))
        Next i
    End If
    Set SwapMap = d
End Function

'==============================================================================
' COMPARISON
'
' Result Dictionary: name, keyDesc, ignoreDesc, rowsA, rowsB, width,
'                    nOK, nMod, nAdd, nRem, nRev, pairs (Collection)
' Each pair: verdict, key, rowA, rowB, rev, diffs (Collection of "ua|ub|va|vb")
'==============================================================================
Private Function CompareSection(ByVal dsiA As Object, ByVal dsiB As Object, _
                               ByVal name As String, ByVal rules As Object) As Object
    Dim rule As Object, res As Object, pairs As Collection
    Dim rowsA As Collection, rowsB As Collection
    Dim keySets As Variant, swapD As Object, ignoreD As Object
    Dim width As Long, k As Long, u As Long, ub As Long
    Dim matchedB() As Long, claimed() As Boolean, revd() As Boolean
    Dim i As Long, j As Long, keyFields As Variant
    Dim idx As Object, kk As String, bucket As Collection, bi As Long
    Dim doRev As Long, useRev As Boolean
    Dim pr As Object, diffs As Collection
    Dim va As String, vb As String
    Dim nOK As Long, nMod As Long, nAdd As Long, nRem As Long, nRev As Long

    Set rule = RuleFor(rules, name)
    Set rowsA = SectionRows(dsiA, name)
    Set rowsB = SectionRows(dsiB, name)
    Set swapD = SwapMap(CStr(rule("swap")))
    Set ignoreD = ToSet(ParseIntList(CStr(rule("ignore"))))
    keySets = Split(CStr(rule("keys")), ";")

    width = SectionWidth(dsiA, name)
    If SectionWidth(dsiB, name) > width Then width = SectionWidth(dsiB, name)

    If rowsA.Count > 0 Then ReDim matchedB(1 To rowsA.Count) Else ReDim matchedB(1 To 1)
    If rowsA.Count > 0 Then ReDim revd(1 To rowsA.Count) Else ReDim revd(1 To 1)
    If rowsB.Count > 0 Then ReDim claimed(1 To rowsB.Count) Else ReDim claimed(1 To 1)
    For i = 1 To rowsA.Count
        matchedB(i) = 0
    Next i

    ' Ordered passes: each key direct, then the same key reversed. Full passes,
    ' so a direct match always beats a reversed or fallback one.
    For k = LBound(keySets) To UBound(keySets)
        keyFields = ParseIntList(Replace(CStr(keySets(k)), "+", ","))
        For doRev = 0 To 1
            useRev = (doRev = 1)
            If useRev And swapD.Count = 0 Then GoTo NextRev
            If rowsA.Count = 0 Or rowsB.Count = 0 Then GoTo NextRev

            Set idx = CreateObject("Scripting.Dictionary")
            For j = 1 To rowsB.Count
                If Not claimed(j) Then
                    kk = KeyOf(rowsB(j), keyFields, Nothing)
                    If Not idx.Exists(kk) Then
                        Set bucket = New Collection
                        idx.Add kk, bucket
                    End If
                    idx(kk).Add j
                End If
            Next j

            For i = 1 To rowsA.Count
                If matchedB(i) = 0 Then
                    If useRev Then
                        kk = KeyOf(rowsA(i), keyFields, swapD)
                    Else
                        kk = KeyOf(rowsA(i), keyFields, Nothing)
                    End If
                    If idx.Exists(kk) Then
                        Set bucket = idx(kk)
                        Do While bucket.Count > 0
                            bi = bucket(1)
                            bucket.Remove 1
                            If Not claimed(bi) Then
                                matchedB(i) = bi
                                claimed(bi) = True
                                revd(i) = useRev
                                Exit Do
                            End If
                        Loop
                    End If
                End If
            Next i
NextRev:
        Next doRev
    Next k

    Set pairs = New Collection

    For i = 1 To rowsA.Count
        Set pr = CreateObject("Scripting.Dictionary")
        pr.Add "key", KeyOf(rowsA(i), ParseIntList(Replace(CStr(keySets(LBound(keySets))), "+", ",")), Nothing)
        pr.Add "rowA", rowsA(i)
        pr.Add "rev", revd(i)
        If matchedB(i) = 0 Then
            pr.Add "verdict", V_REM
            pr.Add "rowB", Empty
            pr.Add "diffs", New Collection
            nRem = nRem + 1
        Else
            pr.Add "rowB", rowsB(matchedB(i))
            Set diffs = New Collection
            For u = 0 To width - 1
                If Not ignoreD.Exists(u) Then
                    ' A reversed row has its paired fields the other way round.
                    If revd(i) And swapD.Exists(u) Then ub = swapD(u) Else ub = u
                    va = Fld(rowsA(i), u)
                    vb = Fld(rowsB(matchedB(i)), ub)
                    If va <> vb Then diffs.Add u & "|" & ub & "|" & va & "|" & vb
                End If
            Next u
            pr.Add "diffs", diffs
            If diffs.Count > 0 Then
                pr.Add "verdict", V_MOD
                nMod = nMod + 1
            Else
                pr.Add "verdict", V_OK
                nOK = nOK + 1
            End If
            If revd(i) Then nRev = nRev + 1
        End If
        pairs.Add pr
    Next i

    For j = 1 To rowsB.Count
        If Not claimed(j) Then
            Set pr = CreateObject("Scripting.Dictionary")
            pr.Add "verdict", V_ADD
            pr.Add "key", KeyOf(rowsB(j), ParseIntList(Replace(CStr(keySets(LBound(keySets))), "+", ",")), Nothing)
            pr.Add "rowA", Empty
            pr.Add "rowB", rowsB(j)
            pr.Add "rev", False
            pr.Add "diffs", New Collection
            pairs.Add pr
            nAdd = nAdd + 1
        End If
    Next j

    Set res = CreateObject("Scripting.Dictionary")
    res.Add "name", name
    res.Add "keyDesc", CStr(rule("keys")) & IIf(swapD.Count > 0, "  (may reverse: " & CStr(rule("swap")) & ")", "")
    res.Add "ignoreDesc", IIf(Len(CStr(rule("ignore"))) = 0, "-", CStr(rule("ignore")))
    res.Add "rowsA", rowsA.Count
    res.Add "rowsB", rowsB.Count
    res.Add "width", width
    res.Add "nOK", nOK
    res.Add "nMod", nMod
    res.Add "nAdd", nAdd
    res.Add "nRem", nRem
    res.Add "nRev", nRev
    res.Add "ignoreSet", ignoreD
    res.Add "pairs", pairs
    Set CompareSection = res
End Function

Private Function KeyOf(ByVal row As Variant, ByVal fields As Variant, _
                       ByVal swapD As Object) As String
    Dim i As Long, f As Long, s As String
    If Not IsArray(fields) Then KeyOf = "": Exit Function
    On Error GoTo Empt
    If UBound(fields) < LBound(fields) Then KeyOf = "": Exit Function
    On Error GoTo 0
    For i = LBound(fields) To UBound(fields)
        f = CLng(fields(i))
        If Not swapD Is Nothing Then
            If swapD.Exists(f) Then f = swapD(f)
        End If
        If Len(s) > 0 Then s = s & KEYSEP
        s = s & Fld(row, f)
    Next i
    KeyOf = s
    Exit Function
Empt:
    KeyOf = ""
End Function

Private Function ToSet(ByVal arr As Variant) As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    If IsArray(arr) Then
        On Error Resume Next        ' empty Array() has no bounds; nothing else can fail here
        For i = LBound(arr) To UBound(arr)
            d(CLng(arr(i))) = True
        Next i
        On Error GoTo 0
    End If
    Set ToSet = d
End Function

Private Sub CollectDiffs(ByVal res As Object, ByVal diffs As Collection)
    Dim pr As Variant, d As Variant, p As Variant
    For Each pr In res("pairs")
        If pr("verdict") <> V_OK Or pr("rev") Then
            If pr("verdict") = V_MOD Then
                For Each d In pr("diffs")
                    p = Split(CStr(d), "|")
                    diffs.Add Array(res("name"), V_MOD, IIf(pr("rev"), "reversed", "direct"), _
                                    Replace(CStr(pr("key")), KEYSEP, " | "), _
                                    "f" & p(0) & IIf(p(0) <> p(1), " vs f" & p(1), ""), p(2), p(3))
                Next d
            ElseIf pr("verdict") = V_ADD Then
                diffs.Add Array(res("name"), V_ADD, "", Replace(CStr(pr("key")), KEYSEP, " | "), _
                                "(whole row)", "", JoinRow(pr("rowB")))
            ElseIf pr("verdict") = V_REM Then
                diffs.Add Array(res("name"), V_REM, "", Replace(CStr(pr("key")), KEYSEP, " | "), _
                                "(whole row)", JoinRow(pr("rowA")), "")
            Else
                ' Reversed but otherwise identical: the file still changed.
                diffs.Add Array(res("name"), V_OK, "reversed", Replace(CStr(pr("key")), KEYSEP, " | "), _
                                "(direction reversed, no other change)", "", "")
            End If
        End If
    Next pr
End Sub

Private Function JoinRow(ByVal row As Variant) As String
    If IsEmpty(row) Then JoinRow = "": Exit Function
    JoinRow = Join(row, ":")
End Function

'==============================================================================
' OUTPUT -- comparison
'==============================================================================
Private Sub WriteSynthesisSheet(ByVal synth As Collection, ByVal pathA As String, _
                                ByVal pathB As String)
    Dim ws As Worksheet, data() As Variant, r As Long, res As Variant
    Set ws = FreshSheet("SYNTHESIS")

    ws.Range("A1").Value = "File 1"
    ws.Range("B1").Value = pathA
    ws.Range("A2").Value = "File 2"
    ws.Range("B2").Value = pathB
    ws.Range("A1:A2").Font.Bold = True

    ReDim data(1 To synth.Count + 1, 1 To 11)
    data(1, 1) = "Section": data(1, 2) = "Rows file 1": data(1, 3) = "Rows file 2"
    data(1, 4) = "OK": data(1, 5) = "MODIFY": data(1, 6) = "ADD"
    data(1, 7) = "REMOVE": data(1, 8) = "REVERSED": data(1, 9) = "Match key"
    data(1, 10) = "Ignored fields": data(1, 11) = "Rows compared"

    r = 2
    For Each res In synth
        data(r, 1) = res("name")
        data(r, 2) = res("rowsA")
        data(r, 3) = res("rowsB")
        data(r, 4) = res("nOK")
        data(r, 5) = res("nMod")
        data(r, 6) = res("nAdd")
        data(r, 7) = res("nRem")
        data(r, 8) = IIf(res("nRev") = 0, "", res("nRev"))
        data(r, 9) = res("keyDesc")
        data(r, 10) = res("ignoreDesc")
        data(r, 11) = res("width")
        r = r + 1
    Next res

    ws.Range("A4").Resize(UBound(data, 1), UBound(data, 2)).Value = data
    StyleHeader ws.Range("A4").Resize(1, 11)

    ' Colour only the non-zero change cells, so the eye goes to them.
    For r = 2 To UBound(data, 1)
        If data(r, 5) > 0 Then ws.Cells(r + 3, 5).Interior.Color = C_YELLOW
        If data(r, 6) > 0 Then ws.Cells(r + 3, 6).Interior.Color = C_GREEN
        If data(r, 7) > 0 Then ws.Cells(r + 3, 7).Interior.Color = C_RED
    Next r

    ws.Columns("A:K").AutoFit
    ws.Rows(5).Select
    ActiveWindow.FreezePanes = False
End Sub

Private Sub WriteDifferencesSheet(ByVal diffs As Collection)
    Dim ws As Worksheet, data() As Variant, r As Long, item As Variant, c As Long
    Set ws = FreshSheet("DIFFERENCES")

    ReDim data(1 To diffs.Count + 1, 1 To 7)
    data(1, 1) = "Section": data(1, 2) = "Check": data(1, 3) = "Match"
    data(1, 4) = "Key": data(1, 5) = "Field"
    data(1, 6) = "Value in file 1": data(1, 7) = "Value in file 2"

    r = 2
    For Each item In diffs
        For c = 1 To 7
            data(r, c) = item(c - 1)
        Next c
        r = r + 1
    Next item

    ws.Range("A1").Resize(UBound(data, 1), 7).Value = data
    StyleHeader ws.Range("A1:G1")

    For r = 2 To UBound(data, 1)
        Select Case data(r, 2)
            Case V_MOD: ws.Cells(r, 6).Resize(1, 2).Interior.Color = C_YELLOW
            Case V_ADD: ws.Cells(r, 7).Interior.Color = C_GREEN
            Case V_REM: ws.Cells(r, 6).Interior.Color = C_RED
        End Select
    Next r

    ws.Columns("A:G").AutoFit
    If diffs.Count > 0 Then ws.Range("A1:G1").AutoFilter
End Sub

Private Sub WriteSectionSheets(ByVal synth As Collection)
    Dim res As Variant, ws As Worksheet, pairs As Collection, pr As Variant
    Dim shown As Collection, data() As Variant, w As Long, r As Long, u As Long
    Dim offA As Long, offB As Long, d As Variant, p As Variant, ig As Object

    For Each res In synth
        Set pairs = res("pairs")
        Set shown = New Collection
        For Each pr In pairs
            If pr("verdict") <> V_OK Or pr("rev") Then shown.Add pr
        Next pr
        If shown.Count > 0 Then
            w = res("width")
            Set ig = res("ignoreSet")
            Set ws = FreshSheet(SafeSheetName(CStr(res("name"))))
            offA = 4
            offB = 4 + w

            ReDim data(1 To shown.Count + 1, 1 To 3 + 2 * w)
            data(1, 1) = "CHECK": data(1, 2) = "MATCH": data(1, 3) = "KEY"
            For u = 0 To w - 1
                data(1, offA + u) = "f" & u & " (1)"
                data(1, offB + u) = "f" & u & " (2)"
            Next u

            r = 2
            For Each pr In shown
                data(r, 1) = pr("verdict")
                data(r, 2) = IIf(pr("rev"), "reversed", "")
                data(r, 3) = Replace(CStr(pr("key")), KEYSEP, " | ")
                For u = 0 To w - 1
                    If Not IsEmpty(pr("rowA")) Then data(r, offA + u) = Fld(pr("rowA"), u)
                    If Not IsEmpty(pr("rowB")) Then data(r, offB + u) = Fld(pr("rowB"), u)
                Next u
                r = r + 1
            Next pr

            ws.Range("A1").Resize(UBound(data, 1), UBound(data, 2)).Value = data
            StyleHeader ws.Range("A1").Resize(1, UBound(data, 2))

            r = 2
            For Each pr In shown
                Select Case pr("verdict")
                    Case V_MOD
                        ws.Cells(r, 1).Interior.Color = C_YELLOW
                        For Each d In pr("diffs")
                            p = Split(CStr(d), "|")
                            ws.Cells(r, offA + CLng(p(0))).Interior.Color = C_YELLOW
                            ws.Cells(r, offB + CLng(p(1))).Interior.Color = C_YELLOW
                        Next d
                    Case V_ADD
                        ws.Cells(r, 1).Interior.Color = C_GREEN
                    Case V_REM
                        ws.Cells(r, 1).Interior.Color = C_RED
                End Select
                For u = 0 To w - 1
                    If ig.Exists(u) Then
                        ws.Cells(r, offA + u).Interior.Color = C_GREY
                        ws.Cells(r, offB + u).Interior.Color = C_GREY
                    End If
                Next u
                r = r + 1
            Next pr

            ws.Columns.AutoFit
        End If
    Next res
End Sub

'==============================================================================
' VIEW TABLE -- one file, chosen elements
'
'   SheetName | section | filter | col:spec, col:spec, ...
'
'   filter  "4=CONNECTOR"  keep rows whose field 4 equals CONNECTOR
'           ""             keep everything
'   spec    12             field index
'           x2 / y2 / z2   one axis of the coordinate packed in field 2
'                          (they look like x1820.00y-905.00z0.00)
'           #terminals     how many rows of section "terminals" share this
'                          row's field 0 -- i.e. this connector's terminal count
'           @36            join every field from 36 to the end with spaces
'                          (circuit option codes are a variable-length tail)
'==============================================================================
Public Function ViewTable() As Variant
    ViewTable = Array( _
      "Connectors|Harness main node components|4=CONNECTOR|Reference:0,Type:4,Description:6,Part Name:8,Part Number:12,Family:19,Colour:27,Cavities:28,Terminals:#terminals,Seals:#seals,Plugs:#plug,Wires:#wires,Route:7", _
      "Splices|Harness main node components|4=SPLICE|Reference:0,Type:4,Description:6,Part Name:8,Part Number:12,Family:19,Colour:27,Cavities:28,Terminals:#terminals,Seals:#seals,Plugs:#plug,Wires:#wires,Route:7", _
      "IDC|Harness main node components|4=IDC|Reference:0,Type:4,Description:6,Part Name:8,Part Number:12,Family:19,Colour:27,Cavities:28,Terminals:#terminals,Seals:#seals,Plugs:#plug,Wires:#wires,Route:7", _
      "All Node Components|Harness main node components||Reference:0,Type:4,Description:6,Part Name:8,Part Number:12,Family:19,Colour:27,Cavities:28,Terminals:#terminals,Seals:#seals,Plugs:#plug,Wires:#wires,Route:7", _
      "Terminals|terminals||Connector:0,Cavity:1,Type:4,Part Name:8,Qty:9,Part Number:12,Sealed:13,Selection:17,Plating:18", _
      "Cavity Seals|seals||Connector:0,Cavity:1,Type:4,Part Name:8,Qty:9,Part Number:12", _
      "Cavity Plugs|plug||Connector:0,Cavity:1,Type:4,Part Name:8,Qty:9,Part Number:12", _
      "Clips|clips||Reference:0,Type:4,Option Expression:5,Parent Node:6,Part Name:8,Qty:9,Part Number:12,Family:19,Colour:27", _
      "Grommets|grommets||Reference:0,Type:4,Option Expression:5,Parent Node:6,Part Name:8,Qty:9,Part Number:12,Family:19,Colour:27", _
      "Other Components|other components||Reference:0,Type:4,Option Expression:5,Parent Node:6,Part Name:8,Qty:9,Part Number:12,Family:19,Colour:27", _
      "Extra Node Components|Harness extra node components||Reference:0,Type:4,Option Expression:5,Parent Node:6,Part Name:8,Qty:9,Part Number:12,Family:19,Colour:27", _
      "Wires|Harness wire specification||Wire:0,Option Expression:1,Wire Spec:2,Colour:3,Size mm2:4,Class:5,Multicore:7,From Connector:8,From Cavity:10,From Plating:11,To Connector:12,To Cavity:14,To Plating:15,Length Min:24,Length Max:25,Family:28,Part Number:30", _
      "Multicores|Harness multicores||Multicore:0,Flag A:1,Flag B:2,Lay Length:13,Twisted:14,Length Min:20,Length Max:21,Group:22,Part Number:24", _
      "Wire Through Nodes|Harness wire through nodes||Wire:0,Option Expression:1,Sequence:2,Node:3,Through:5", _
      "Branches|Harness branch configuration||From Node:0,From X:x2,From Y:y2,From Z:z2,To Node:3,To X:x5,To Y:y5,To Z:z5,Length mm:6,Option Expression:7,Diameter mm:8", _
      "Insulations|Harness branch insulations||From Node:0,To Node:2,Route:3,Sequence:4,Diameter mm:6,Option Expression:7,Source:12,Order:17,Type:18,Material Spec:19,Part Number:21,Colour:22,Family:23,Coverage:25", _
      "Insulation Through Nodes|Harness branch insulation through nodes||Insulation From:0,Sequence:1,Node:2,Through:4", _
      "Circuits|Harness circuit information||Circuit ID:0,Rev:1,Date:2,Source:3,Site:11,Description:26,Metric:31,Option Codes:@36", _
      "Option Codes|Composite Option Codes||Option Code:0,Description:1", _
      "Harness Scope|Harness Scope||Harness ID:0,Rev:1,Harness ID 2:2,Rev 2:3,Attribute:4,Value:5,Value 2:6", _
      "Harness Identity|Harness name information||Harness ID:0,Rev:1,Date:2,Field 4:3,Source:4,Harness ID 2:5,Rev 2:6,Date 2:7,Site:11,Drawing:25" _
    )
End Function

Private Function ViewField(ByVal spec As String, ByVal idx As Long) As String
    Dim p As Variant
    p = Split(spec, "|")
    ViewField = CStr(p(idx))
End Function

Private Function ViewRowCount(ByVal dsi As Object, ByVal spec As String) As Long
    Dim p As Variant, rows_ As Collection, r As Variant, n As Long
    p = Split(spec, "|")
    Set rows_ = SectionRows(dsi, CStr(p(1)))
    If Len(CStr(p(2))) = 0 Then ViewRowCount = rows_.Count: Exit Function
    For Each r In rows_
        If RowPassesFilter(r, CStr(p(2))) Then n = n + 1
    Next r
    ViewRowCount = n
End Function

Private Function RowPassesFilter(ByVal row As Variant, ByVal filter As String) As Boolean
    Dim p As Variant
    If Len(filter) = 0 Then RowPassesFilter = True: Exit Function
    p = Split(filter, "=")
    RowPassesFilter = (StrComp(Fld(row, CLng(p(0))), CStr(p(1)), vbTextCompare) = 0)
End Function

Private Sub WriteViewSheet(ByVal dsi As Object, ByVal spec As String)
    Dim p As Variant, cols As Variant, i As Long, r As Variant
    Dim rows_ As Collection, ws As Worksheet, data() As Variant
    Dim nRows As Long, rr As Long, titleSpec As Variant
    Dim tallies As Object

    p = Split(spec, "|")
    cols = Split(CStr(p(3)), ",")
    Set rows_ = SectionRows(dsi, CStr(p(1)))
    Set tallies = CreateObject("Scripting.Dictionary")

    nRows = ViewRowCount(dsi, spec)
    If nRows = 0 Then Exit Sub

    Set ws = FreshSheet(SafeSheetName(CStr(p(0))))
    ReDim data(1 To nRows + 1, 1 To UBound(cols) - LBound(cols) + 1)

    For i = LBound(cols) To UBound(cols)
        titleSpec = Split(CStr(cols(i)), ":")
        data(1, i - LBound(cols) + 1) = CStr(titleSpec(0))
    Next i

    rr = 2
    For Each r In rows_
        If RowPassesFilter(r, CStr(p(2))) Then
            For i = LBound(cols) To UBound(cols)
                titleSpec = Split(CStr(cols(i)), ":")
                data(rr, i - LBound(cols) + 1) = CellValue(r, CStr(titleSpec(1)), dsi, tallies)
            Next i
            rr = rr + 1
        End If
    Next r

    ws.Range("A1").Resize(UBound(data, 1), UBound(data, 2)).Value = data
    StyleHeader ws.Range("A1").Resize(1, UBound(data, 2))
    ws.Range("A1").Resize(1, UBound(data, 2)).AutoFilter
    ws.Columns.AutoFit
End Sub

Private Function CellValue(ByVal row As Variant, ByVal spec As String, _
                           ByVal dsi As Object, ByVal tallies As Object) As String
    Dim axis As String, fieldIdx As Long, n As Long, tal As Object
    Dim i As Long, s As String, v As String

    If IsNumeric(spec) Then
        CellValue = Fld(row, CLng(spec))

    ElseIf Left$(spec, 1) = "#" Then
        ' Count of related rows, e.g. how many terminals this connector has.
        Set tal = TallyFor(dsi, tallies, Mid$(spec, 2))
        If tal.Exists(Fld(row, 0)) Then n = tal(Fld(row, 0)) Else n = 0
        If n = 0 Then CellValue = "" Else CellValue = CStr(n)

    ElseIf Left$(spec, 1) = "@" Then
        ' Variable-length tail, e.g. circuit option codes from field 36 on.
        fieldIdx = CLng(Mid$(spec, 2))
        For i = fieldIdx To UBound(row)
            v = Trim$(Fld(row, i))
            If Len(v) > 0 Then
                If Len(s) > 0 Then s = s & " "
                s = s & v
            End If
        Next i
        CellValue = s

    Else
        axis = LCase$(Left$(spec, 1))
        fieldIdx = CLng(Mid$(spec, 2))
        CellValue = CoordPart(Fld(row, fieldIdx), axis)
    End If
End Function

' Cached count of rows per reference. "wires" is special: a wire lands on two
' nodes, so both ends are tallied.
Private Function TallyFor(ByVal dsi As Object, ByVal tallies As Object, _
                          ByVal what As String) As Object
    Dim d As Object, r As Variant, ref As String, rows_ As Collection

    If tallies.Exists(what) Then
        Set TallyFor = tallies(what)
        Exit Function
    End If

    Set d = CreateObject("Scripting.Dictionary")
    If what = "wires" Then
        Set rows_ = SectionRows(dsi, "Harness wire specification")
        For Each r In rows_
            ref = Fld(r, 8)
            If Len(ref) > 0 Then d(ref) = d(ref) + 1
            ref = Fld(r, 12)
            If Len(ref) > 0 Then d(ref) = d(ref) + 1
        Next r
    Else
        Set rows_ = SectionRows(dsi, what)
        For Each r In rows_
            ref = Fld(r, 0)
            If Len(ref) > 0 Then d(ref) = d(ref) + 1
        Next r
    End If

    tallies.Add what, d
    Set TallyFor = d
End Function

' Coordinates are packed 3D: x1820.00y-905.00z0.00. Splitting on "y" alone --
' which grapheCAPH does -- leaves Y as "-905.00z0.00", CSng throws, and the
' error was swallowed, so the column came out silently blank. Parse all three.
Public Function CoordPart(ByVal raw As String, ByVal axis As String) As String
    Dim px As Long, py As Long, pz As Long
    If Len(raw) = 0 Then CoordPart = "": Exit Function
    px = InStr(1, raw, "x", vbTextCompare)
    py = InStr(1, raw, "y", vbTextCompare)
    pz = InStr(1, raw, "z", vbTextCompare)
    If px = 0 Or py = 0 Or pz = 0 Or Not (px < py And py < pz) Then
        CoordPart = "?? " & raw          ' loud, not blank
        Exit Function
    End If
    Select Case axis
        Case "x": CoordPart = Mid$(raw, px + 1, py - px - 1)
        Case "y": CoordPart = Mid$(raw, py + 1, pz - py - 1)
        Case "z": CoordPart = Mid$(raw, pz + 1)
        Case Else: CoordPart = ""
    End Select
End Function

'==============================================================================
' HELPERS
'==============================================================================
Private Function PickFile(ByVal title As String) As String
    Dim f As Variant
    f = Application.GetOpenFilename( _
        "DSI files (*.dsi),*.dsi,All files (*.*),*.*", 1, title, , False)
    If VarType(f) = vbBoolean Then PickFile = "" Else PickFile = CStr(f)
End Function

Private Function FreshSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next            ' the only tolerated swallow: sheet may not exist
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = name
    Else
        ws.Cells.Clear
        ws.Cells.Interior.ColorIndex = xlColorIndexNone
    End If
    Set FreshSheet = ws
End Function

Private Function SafeSheetName(ByVal name As String) As String
    Dim s As String, bad As Variant, b As Variant
    s = name
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each b In bad
        s = Replace(s, CStr(b), " ")
    Next b
    s = Trim$(s)
    If Len(s) > 31 Then s = Left$(s, 31)
    If Len(s) = 0 Then s = "Sheet"
    SafeSheetName = s
End Function

Private Sub StyleHeader(ByVal rng As Range)
    rng.Font.Bold = True
    rng.Font.Color = RGB(255, 255, 255)
    rng.Interior.Color = C_HEADER
    rng.HorizontalAlignment = xlCenter
End Sub

Private Function PadRight(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then PadRight = Left$(s, n) Else PadRight = s & Space$(n - Len(s))
End Function

Private Function ParseSelection(ByVal text As String, ByVal count As Long) As Variant
    Dim t As String, parts As Variant, i As Long
    Dim out() As Long, n As Long, v As Long, j As Long, dup As Boolean

    t = LCase$(Trim$(text))
    If t = "a" Or t = "all" Then
        ReDim out(0 To count - 1)
        For i = 1 To count
            out(i - 1) = i
        Next i
        ParseSelection = out
        Exit Function
    End If

    parts = Split(Replace(t, " ", ""), ",")
    ReDim out(0 To UBound(parts) - LBound(parts))
    n = 0
    For i = LBound(parts) To UBound(parts)
        If Len(CStr(parts(i))) > 0 Then
            If Not IsNumeric(parts(i)) Then ParseSelection = Empty: Exit Function
            v = CLng(parts(i))
            If v >= 1 And v <= count Then
                dup = False
                For j = 0 To n - 1
                    If out(j) = v Then dup = True
                Next j
                If Not dup Then
                    out(n) = v
                    n = n + 1
                End If
            End If
        End If
    Next i
    If n = 0 Then ParseSelection = Empty: Exit Function
    ReDim Preserve out(0 To n - 1)
    ParseSelection = out
End Function

'==============================================================================
' Optional: quick census, to confirm the parser sees the file correctly.
'==============================================================================
Public Sub DSI_Census()
    Dim path As String, dsi As Object, ws As Worksheet
    Dim nm As Variant, r As Long, data() As Variant, order As Collection

    path = PickFile("Select a DSI file to inspect")
    If Len(path) = 0 Then Exit Sub

    Set dsi = ParseDSI(path)
    Set order = dsi("order")
    Set ws = FreshSheet("CENSUS")

    ReDim data(1 To order.Count + 1, 1 To 3)
    data(1, 1) = "Section": data(1, 2) = "Rows": data(1, 3) = "Max fields"
    r = 2
    For Each nm In order
        data(r, 1) = CStr(nm)
        data(r, 2) = SectionRowCount(dsi, CStr(nm))
        data(r, 3) = SectionWidth(dsi, CStr(nm))
        r = r + 1
    Next nm

    ws.Range("A1").Resize(UBound(data, 1), 3).Value = data
    StyleHeader ws.Range("A1:C1")
    ws.Columns("A:C").AutoFit

    MsgBox "Sections found: " & order.Count & vbCrLf & _
           "Expected for the sample file: 27 sections, 3435 rows." & vbCrLf & vbCrLf & _
           "If you see 20 sections instead of 27, the sub-blocks under " & _
           "'Harness main node components' were fused and the parser is wrong.", _
           vbInformation, "DSI Census"
End Sub
