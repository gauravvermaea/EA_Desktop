Attribute VB_Name = "modPaint"
' =====================================================================
' Painter for EA_Entities type sheets (T_<TypeCode>). Mirrors
' build/builder.py paint_type_sheet():
'   row 1 hidden AttrCodes, row 2 headers, data from row 3
'   columns A:C hidden (EntityID / RelID, RowKind E|R, PaintHash)
'   relationship rows grouped (+) under their entity, linking to the
'   other entity's row on its own type sheet
' =====================================================================
Option Explicit

Private Const FIXED_COLS As Long = 8      ' EntityID RowKind PaintHash DisplayID Name Description Status RepositoryArea

Private mMaster As Variant                ' master data array (1-based, table columns)
Private mCols As Object                   ' AttrCode -> table column index (1-based within table)
Private mRel As Variant                   ' relationship data array
Private mByType As Object                 ' TypeCode -> Collection of master row indexes (sorted by name)
Private mRelsOf As Object                 ' EntityID -> Collection of Array(relIndex, "out"|"in")
Private mRowOfID As Object                ' EntityID -> master row index

' ---------------------------------------------------------------- context
Public Sub BuildContext()
    Dim lo As ListObject, cols As Object, k As Variant, i As Long, base As Long, t As String, id As String
    modMeta.LoadMeta
    Set lo = modData.MasterLO()
    Set cols = modData.MasterCols()
    base = lo.Range.Column - 1
    Set mCols = modCore.NewDict()
    For Each k In cols.Keys: mCols(k) = cols(k) - base: Next k
    mMaster = modCore.TableData(lo)
    mRel = modCore.TableData(modData.RelLO())
    Set mByType = modCore.NewDict()
    Set mRowOfID = modCore.NewDict()
    Set mRelsOf = modCore.NewDict()
    If IsArray(mMaster) Then
        For i = 1 To UBound(mMaster, 1)
            id = CStr(mMaster(i, mCols("EntityID")))
            If id <> "" Then
                t = CStr(mMaster(i, mCols("TypeCode")))
                mRowOfID(id) = i
                If Not mByType.Exists(t) Then mByType.Add t, New Collection
                InsertByName mByType(t), i
            End If
        Next i
    End If
    If IsArray(mRel) Then
        For i = 1 To UBound(mRel, 1)
            If CStr(mRel(i, 1)) <> "" And CStr(mRel(i, 11)) <> "Retired" Then
                AddRel CStr(mRel(i, 3)), i, "out"
                AddRel CStr(mRel(i, 4)), i, "in"
            End If
        Next i
    End If
End Sub

Private Sub AddRel(ByVal id As String, ByVal i As Long, ByVal dirn As String)
    If Not mRelsOf.Exists(id) Then mRelsOf.Add id, New Collection
    mRelsOf(id).Add Array(i, dirn)
End Sub

Private Sub InsertByName(c As Collection, ByVal i As Long)
    Dim j As Long, nm As String
    nm = LCase$(CStr(mMaster(i, mCols("Name"))))
    For j = 1 To c.Count
        If nm < LCase$(CStr(mMaster(c(j), mCols("Name")))) Then c.Add i, Before:=j: Exit Sub
    Next j
    c.Add i
End Sub

' ---------------------------------------------------------------- context accessors (used by modSync)
Public Function CtxHas(ByVal entityID As String) As Boolean
    If mRowOfID Is Nothing Then BuildContext
    CtxHas = mRowOfID.Exists(entityID)
End Function

Public Function CtxValue(ByVal entityID As String, ByVal key As String) As Variant
    If mRowOfID Is Nothing Then BuildContext
    If Not mRowOfID.Exists(entityID) Then Exit Function
    If Not mCols.Exists(key) Then Exit Function
    CtxValue = mMaster(mRowOfID(entityID), mCols(key))
End Function

' Hash of the master values in the order of the given sheet keys (from index 5 = Name onward):
' equals the PaintHash the painter would write for this entity.
Public Function CtxProjectedHash(ByVal entityID As String, keys As Variant) As String
    Dim vals As New Collection, j As Long
    For j = 5 To UBound(keys)
        vals.Add CtxValue(entityID, CStr(keys(j)))
    Next j
    CtxProjectedHash = modCore.RowHashOf(vals)
End Function

' IDs in the order the painter writes them (entity, then its relationship rows)
Public Function CtxSequence(ByVal typeCode As String) As Collection
    Dim c As New Collection, e As Variant, x As Variant, eid As String
    If mByType Is Nothing Then BuildContext
    Set CtxSequence = c
    If Not mByType.Exists(typeCode) Then Exit Function
    For Each e In mByType(typeCode)
        eid = CStr(mMaster(e, mCols("EntityID")))
        c.Add eid
        If mRelsOf.Exists(eid) Then
            For Each x In mRelsOf(eid)
                c.Add CStr(mRel(x(0), 1))
            Next x
        End If
    Next e
End Function

Public Function EntityCountOf(ByVal typeCode As String) As Long
    If mByType Is Nothing Then BuildContext
    If mByType.Exists(typeCode) Then EntityCountOf = mByType(typeCode).Count
End Function

Public Function TypesInSheetOrder() As Collection
    Dim c As New Collection, dom As Variant, t As Variant, names As Object, sorted As Collection
    modMeta.LoadMeta
    For Each dom In Array("Strategy", "Motivation", "Business", "Application", "Technology", "Common", _
                          "Implementation & Migration", "")
        Set sorted = New Collection
        For Each t In modMeta.gTypes.Keys
            If modMeta.gTypes(t)("Status") = "Active" Then
                If modMeta.gTypes(t)("Domain") = dom Or (dom = "" And Not KnownDomain(modMeta.gTypes(t)("Domain"))) Then
                    InsertTypeByName sorted, CStr(t)
                End If
            End If
        Next t
        For Each t In sorted: c.Add t: Next t
    Next dom
    Set TypesInSheetOrder = c
End Function

Private Function KnownDomain(ByVal d As String) As Boolean
    Select Case d
        Case "Strategy", "Motivation", "Business", "Application", "Technology", "Common", "Implementation & Migration"
            KnownDomain = True
    End Select
End Function

Private Sub InsertTypeByName(c As Collection, ByVal t As String)
    Dim j As Long
    For j = 1 To c.Count
        If LCase$(modMeta.gTypes(t)("Name")) < LCase$(modMeta.gTypes(c(j))("Name")) Then c.Add t, Before:=j: Exit Sub
    Next j
    c.Add t
End Sub

' ---------------------------------------------------------------- paint
Public Sub PaintAll(wbE As Workbook)
    Dim t As Variant, ws As Worksheet, n As Long, total As Long, pend As Object, skipped As String
    BuildContext
    Set ws = Nothing
    Set pend = modSync.BlockedSheets()
    total = modMeta.gTypes.Count
    For Each t In TypesInSheetOrder()
        n = n + 1
        If pend.Exists(modMeta.gTypes(t)("Sheet")) Then
            ' never overwrite edits that have not reached EA_Core yet
            skipped = skipped & ", " & modMeta.gTypes(t)("Sheet")
        Else
            modCore.Status "painting " & modMeta.gTypes(t)("Name") & " (" & n & " of " & total & ")"
            PaintTypeSheet wbE, CStr(t)
        End If
    Next t
    If skipped <> "" Then
        MsgBox "These type sheets have edits not yet saved to EA_Core, or open items on SYNC_Review, and were " & _
               "not refreshed: " & Mid$(skipped, 3) & vbLf & _
               "Save EA_Entities and settle SYNC_Review, then run Sync Now.", vbExclamation, "EA Workbench"
    End If
    ' hide sheets of retired types
    For Each t In modMeta.gTypes.Keys
        If modMeta.gTypes(t)("Status") <> "Active" Then
            If modCore.SheetExists(wbE, modMeta.gTypes(t)("Sheet")) Then wbE.Worksheets(modMeta.gTypes(t)("Sheet")).Visible = xlSheetHidden
        End If
    Next t
    OrderSheets wbE
End Sub

' Repaints the sheets of the given types (and rebuilds the context first)
Public Sub PaintTypes(types As Object)
    Dim wbE As Workbook, t As Variant
    Set wbE = modCore.EntitiesWB(False)
    If wbE Is Nothing Then Exit Sub
    modCore.BeginBusy
    On Error GoTo Done
    BuildContext
    Dim pend As Object
    Set pend = modSync.BlockedSheets()
    For Each t In types.Keys
        If modMeta.TypeExists(CStr(t)) Then
            If Not pend.Exists(modMeta.gTypes(t)("Sheet")) Then PaintTypeSheet wbE, CStr(t)
        End If
    Next t
Done:
    modCore.EndBusy
End Sub

Public Sub OrderSheets(wbE As Workbook)
    Dim t As Variant, prev As Worksheet, ws As Worksheet
    Set prev = wbE.Worksheets("Home")
    For Each t In TypesInSheetOrder()
        If modCore.SheetExists(wbE, modMeta.gTypes(t)("Sheet")) Then
            Set ws = wbE.Worksheets(modMeta.gTypes(t)("Sheet"))
            If ws.Index <> prev.Index + 1 Then ws.Move After:=prev
            Set prev = ws
        End If
    Next t
End Sub

Public Sub PaintTypeSheet(wbE As Workbook, ByVal typeCode As String)
    Dim ws As Worksheet, sheetName As String, attrs As Collection, nCols As Long, a As Variant
    Dim keys() As String, i As Long, j As Long, isNew As Boolean
    If mByType Is Nothing Then BuildContext

    sheetName = modMeta.gTypes(typeCode)("Sheet")
    If sheetName = "" Then sheetName = Left$("T_" & typeCode, 31)
    If modCore.SheetExists(wbE, sheetName) Then
        Set ws = wbE.Worksheets(sheetName)
    Else
        Set ws = wbE.Worksheets.Add(After:=wbE.Worksheets(1))
        ws.Name = sheetName
        isNew = True
    End If
    ws.Visible = xlSheetVisible
    modCore.UnprotectSheet ws

    ' ---- reset
    ws.Cells.ClearOutline
    ws.Cells.Hyperlinks.Delete
    ws.Cells.ClearComments
    ws.Cells.Validation.Delete
    ws.Cells.Clear
    ws.Rows.Hidden = False
    ws.Outline.SummaryRow = xlSummaryAbove

    ' ---- columns
    Set attrs = modMeta.AttrsForType(typeCode)
    nCols = FIXED_COLS + attrs.Count
    ReDim keys(1 To nCols)
    keys(1) = "EntityID": keys(2) = "RowKind": keys(3) = "PaintHash": keys(4) = "DisplayID"
    keys(5) = "Name": keys(6) = "Description": keys(7) = "Status": keys(8) = "RepositoryArea"
    i = FIXED_COLS
    For Each a In attrs
        i = i + 1
        keys(i) = a
    Next a

    ' ---- key row + header row
    Dim hdr() As Variant, kr() As Variant, fill As Long
    ReDim hdr(1 To 1, 1 To nCols): ReDim kr(1 To 1, 1 To nCols)
    For j = 1 To nCols
        kr(1, j) = keys(j)
        hdr(1, j) = HeaderText(typeCode, keys(j))
    Next j
    ws.Range(ws.Cells(modCore.TS_KEY_ROW, 1), ws.Cells(modCore.TS_KEY_ROW, nCols)).Value = kr
    ws.Range(ws.Cells(modCore.TS_HDR_ROW, 1), ws.Cells(modCore.TS_HDR_ROW, nCols)).Value = hdr
    ws.Rows(modCore.TS_KEY_ROW).Hidden = True
    ws.Rows(modCore.TS_KEY_ROW).Font.Size = 8
    fill = HexToRGB(modMeta.gTypes(typeCode)("Fill"))
    With ws.Range(ws.Cells(modCore.TS_HDR_ROW, 1), ws.Cells(modCore.TS_HDR_ROW, nCols))
        .Font.Bold = True
        .Interior.Color = fill
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.Rows(modCore.TS_HDR_ROW).RowHeight = 30
    For j = FIXED_COLS + 1 To nCols
        If modMeta.gAttrs(keys(j))("Desc") <> "" Then
            ws.Cells(modCore.TS_HDR_ROW, j).AddComment modMeta.gAttrs(keys(j))("Desc")
        End If
    Next j

    ' ---- body
    Dim nRows As Long, ents As Collection, e As Variant, rels As Collection, x As Variant
    Set ents = New Collection
    If mByType.Exists(typeCode) Then Set ents = mByType(typeCode)
    nRows = 0
    For Each e In ents
        nRows = nRows + 1
        Dim eid As String
        eid = CStr(mMaster(e, mCols("EntityID")))
        If mRelsOf.Exists(eid) Then nRows = nRows + mRelsOf(eid).Count
    Next e

    ' number formats go on before values, so text such as "19.30" is not turned into a number
    Dim lastRow As Long, rr As Long
    lastRow = modCore.TS_FIRST_ROW + nRows - 1 + modCore.TS_SPARE_ROWS
    ws.Range(ws.Cells(modCore.TS_FIRST_ROW, 1), ws.Cells(lastRow, 4)).NumberFormat = "@"   ' IDs, kind, hash, Display ID
    ws.Range(ws.Cells(modCore.TS_FIRST_ROW, 6), ws.Cells(lastRow, 8)).NumberFormat = "@"
    For j = FIXED_COLS + 1 To nCols
        ws.Range(ws.Cells(modCore.TS_FIRST_ROW, j), ws.Cells(lastRow, j)).NumberFormat = modGenerator.NumberFormatFor(keys(j))
    Next j

    Dim relRows As New Collection, entRows As New Collection, retiredRows As New Collection
    If nRows > 0 Then
        Dim out() As Variant, r As Long, hv As Collection
        ReDim out(1 To nRows, 1 To nCols)
        r = 0
        For Each e In ents
            r = r + 1
            eid = CStr(mMaster(e, mCols("EntityID")))
            Set hv = New Collection
            out(r, 1) = eid
            out(r, 2) = "E"
            For j = 4 To nCols
                If mCols.Exists(keys(j)) Then
                    out(r, j) = modCore.SafeFor(mMaster(e, mCols(keys(j))), ColFormat(j, keys(j)))
                    If j >= 5 Then hv.Add mMaster(e, mCols(keys(j)))
                ElseIf j >= 5 Then
                    hv.Add Empty                               ' same inputs as CtxProjectedHash
                End If
            Next j
            out(r, 3) = modCore.RowHashOf(hv)
            entRows.Add r
            If CStr(mMaster(e, mCols("Status"))) = "Retired" Then retiredRows.Add r
            If mRelsOf.Exists(eid) Then
                For Each x In mRelsOf(eid)
                    r = r + 1
                    FillRelRow out, r, CLng(x(0)), CStr(x(1))
                    relRows.Add r
                Next x
            End If
        Next e
        ws.Range(ws.Cells(modCore.TS_FIRST_ROW, 1), ws.Cells(modCore.TS_FIRST_ROW + nRows - 1, nCols)).Value = out
    End If

    ' ---- validation
    For j = FIXED_COLS + 1 To nCols
        modGenerator.ApplyValidation ws.Range(ws.Cells(modCore.TS_FIRST_ROW, j), ws.Cells(lastRow, j)), keys(j)
    Next j
    modGenerator.ApplyListValidation ws.Range(ws.Cells(modCore.TS_FIRST_ROW, 7), ws.Cells(lastRow, 7)), "Status"
    modGenerator.ApplyListValidation ws.Range(ws.Cells(modCore.TS_FIRST_ROW, 8), ws.Cells(lastRow, 8)), "RepositoryArea"
    For Each x In entRows
        ws.Cells(modCore.TS_FIRST_ROW + x - 1, 5).Font.Bold = True
    Next x
    For Each x In retiredRows
        ws.Rows(modCore.TS_FIRST_ROW + x - 1).Font.Color = RGB(128, 128, 128)
    Next x
    ' relationship rows: style, read-only, grouped and collapsed
    Dim blockStart As Long, prevRow As Long
    blockStart = 0: prevRow = -10
    For Each x In relRows
        rr = modCore.TS_FIRST_ROW + x - 1
        With ws.Range(ws.Cells(rr, 4), ws.Cells(rr, nCols))
            .Font.Italic = True
            .Font.Color = RGB(89, 89, 89)
            .Validation.Delete
        End With
        ws.Cells(rr, 5).Font.Color = RGB(5, 99, 193)
        ws.Cells(rr, 5).Font.Underline = xlUnderlineStyleSingle
        If rr <> prevRow + 1 Then
            If blockStart > 0 Then ws.Rows(blockStart & ":" & prevRow).Group
            blockStart = rr
        End If
        prevRow = rr
    Next x
    If blockStart > 0 Then ws.Rows(blockStart & ":" & prevRow).Group
    If relRows.Count > 0 Then ws.Outline.ShowLevels RowLevels:=1

    ' ---- column layout
    ws.Columns("A:C").Hidden = True
    ws.Columns(4).ColumnWidth = 18
    ws.Columns(5).ColumnWidth = 34
    ws.Columns(6).ColumnWidth = 40
    ws.Columns(7).ColumnWidth = 10
    ws.Columns(8).ColumnWidth = 24
    For j = FIXED_COLS + 1 To nCols
        ws.Columns(j).ColumnWidth = 16
    Next j
    ws.Tab.Color = fill

    ' ---- locks: entity rows and spare rows editable from Name onward
    ws.Cells.Locked = True
    For Each x In entRows
        rr = modCore.TS_FIRST_ROW + x - 1
        ws.Range(ws.Cells(rr, 5), ws.Cells(rr, nCols)).Locked = False
    Next x
    ws.Range(ws.Cells(modCore.TS_FIRST_ROW + nRows, 5), ws.Cells(lastRow, nCols)).Locked = False

    If isNew Then
        Dim curWb As Workbook
        Set curWb = ActiveWorkbook
        wbE.Activate
        ws.Activate
        ActiveWindow.FreezePanes = False
        ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
        ws.Range("F3").Select
        ActiveWindow.FreezePanes = True
        If Not curWb Is Nothing Then curWb.Activate
    End If
    modCore.ProtectSheet ws
End Sub

Private Sub FillRelRow(out() As Variant, ByVal r As Long, ByVal relIdx As Long, ByVal dirn As String)
    Dim otherID As String, verb As String, oi As Variant, otherType As String, otherName As String, sh As String
    If dirn = "out" Then otherID = CStr(mRel(relIdx, 4)) Else otherID = CStr(mRel(relIdx, 3))
    verb = modMeta.RelVerb(CStr(mRel(relIdx, 2)))
    If dirn = "in" Then verb = modMeta.InverseVerb(verb)
    out(r, 1) = CStr(mRel(relIdx, 1))
    out(r, 2) = "R"
    out(r, 4) = "  " & IIf(dirn = "out", ChrW$(8594), ChrW$(8592)) & " " & verb
    If mRowOfID.Exists(otherID) Then
        oi = mRowOfID(otherID)
        otherType = CStr(mMaster(oi, mCols("TypeCode")))
    End If
    If otherType <> "" And Not modMeta.gTypes.Exists(otherType) Then
        out(r, 5) = "(unknown type " & otherType & ")"
    ElseIf otherType <> "" Then
        otherName = CStr(mMaster(oi, mCols("Name")))
        sh = modMeta.gTypes(otherType)("Sheet")
        out(r, 5) = "=HYPERLINK(""#'" & sh & "'!E""&MATCH(""" & otherID & """,'" & sh & "'!$A:$A,0),""" & _
                    Replace(otherName, """", """""") & """)"
        out(r, 6) = modMeta.TypeName_(otherType)
    Else
        out(r, 5) = "(missing entity)"
    End If
    out(r, 7) = modCore.SafeFor(CStr(mRel(relIdx, 9)), "@")
End Sub

' Number format the painter gives a type-sheet column (D and F:H text, E general)
Private Function ColFormat(ByVal j As Long, ByVal key As String) As String
    If j = 5 Then
        ColFormat = "General"
    ElseIf j <= FIXED_COLS Then
        ColFormat = "@"
    Else
        ColFormat = modGenerator.NumberFormatFor(key)
    End If
End Function

Private Function HeaderText(ByVal typeCode As String, ByVal key As String) As String
    Select Case key
        Case "EntityID": HeaderText = "Entity ID"
        Case "RowKind": HeaderText = "Row"
        Case "PaintHash": HeaderText = "Paint hash"
        Case "DisplayID": HeaderText = "Display ID"
        Case "RepositoryArea": HeaderText = "Repository Area"
        Case "Name", "Description", "Status": HeaderText = key
        Case Else
            HeaderText = modMeta.AttrName(key) & IIf(modMeta.IsMandatory(typeCode, key), " *", "")
    End Select
End Function

Private Function HexToRGB(ByVal hexColour As String) As Long
    If Len(hexColour) <> 6 Then HexToRGB = RGB(238, 238, 238): Exit Function
    HexToRGB = RGB(CLng("&H" & Mid$(hexColour, 1, 2)), CLng("&H" & Mid$(hexColour, 3, 2)), CLng("&H" & Mid$(hexColour, 5, 2)))
End Function

' Hash of a type-sheet entity row as currently shown (Name onward)
Public Function SheetRowHash(ws As Worksheet, ByVal r As Long) As String
    Dim lastCol As Long, c As Long, vals As New Collection
    lastCol = ws.Cells(modCore.TS_KEY_ROW, ws.Columns.Count).End(xlToLeft).Column
    For c = 5 To lastCol
        vals.Add ws.Cells(r, c).Value
    Next c
    SheetRowHash = modCore.RowHashOf(vals)
End Function
