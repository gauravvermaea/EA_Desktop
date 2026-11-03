Attribute VB_Name = "modData"
' =====================================================================
' The single write path for repository data. Every change - from a form,
' a sheet edit, an EA_Entities push, an import, a merge or a harvest -
' goes through these procedures, so validation, stamps, row hashes and
' history are always applied the same way.
' =====================================================================
Option Explicit

Public Const HASH_FIELDS_SYS As String = "Name|Description|Status|RepositoryArea"

' ---------------------------------------------------------------- master layout
Public Function MasterSheet() As Worksheet
    Set MasterSheet = ThisWorkbook.Worksheets(modCore.SH_MASTER)
End Function

Public Function MasterLO() As ListObject
    Set MasterLO = MasterSheet().ListObjects("tblMaster")
End Function

' AttrCode -> sheet column, from the hidden key row
Public Function MasterCols() As Object
    Dim d As Object, ws As Worksheet, lo As ListObject, c As Long, k As String
    Set d = modCore.NewDict()
    Set ws = MasterSheet(): Set lo = MasterLO()
    For c = lo.Range.Column To lo.Range.Column + lo.ListColumns.Count - 1
        k = CStr(ws.Cells(modCore.MASTER_KEY_ROW, c).Value)
        If k <> "" Then d(k) = c
    Next c
    Set MasterCols = d
End Function

' Business attribute codes in master column order (for hashing)
Public Function MasterAttrOrder() As Collection
    Dim c As New Collection, ws As Worksheet, lo As ListObject, col As Long, k As String
    Set ws = MasterSheet(): Set lo = MasterLO()
    For col = lo.Range.Column To lo.Range.Column + lo.ListColumns.Count - 1
        k = CStr(ws.Cells(modCore.MASTER_KEY_ROW, col).Value)
        If k <> "" And Not modMeta.IsSystemAttr(k) Then c.Add k
    Next col
    Set MasterAttrOrder = c
End Function

Public Function FindEntityRow(ByVal entityID As String) As Long
    If entityID = "" Then Exit Function
    Dim r As Variant
    r = Application.Match(entityID, MasterSheet().Columns(1), 0)
    If Not IsError(r) Then
        If r >= modCore.MASTER_FIRST_ROW Then FindEntityRow = CLng(r)
    End If
End Function

Public Function ReadEntityRow(ByVal r As Long, Optional cols As Object = Nothing) As Object
    Dim d As Object, k As Variant, ws As Worksheet, vals As Variant, lo As ListObject
    If cols Is Nothing Then Set cols = MasterCols()
    Set ws = MasterSheet(): Set lo = MasterLO()
    vals = ws.Range(ws.Cells(r, lo.Range.Column), ws.Cells(r, lo.Range.Column + lo.ListColumns.Count - 1)).Value
    Set d = modCore.NewDict()
    For Each k In cols.Keys
        d(k) = vals(1, cols(k) - lo.Range.Column + 1)
    Next k
    Set ReadEntityRow = d
End Function

Public Function EntityTypeOf(ByVal entityID As String) As String
    Dim r As Long
    r = FindEntityRow(entityID)
    If r > 0 Then EntityTypeOf = CStr(MasterSheet().Cells(r, MasterCols()("TypeCode")).Value)
End Function

Public Function EntityNameOf(ByVal entityID As String) As String
    Dim r As Long
    r = FindEntityRow(entityID)
    If r > 0 Then EntityNameOf = CStr(MasterSheet().Cells(r, MasterCols()("Name")).Value)
End Function

' EntityRef rule: value is an EntityID or an exact Name, of one of the allowed types
Public Function EntityRefOK(ByVal value As String, ByVal allowedTypes As String) As Boolean
    Dim r As Long, t As String, ws As Worksheet, cols As Object, found As Variant
    Set ws = MasterSheet(): Set cols = MasterCols()
    r = FindEntityRow(value)
    If r = 0 Then
        found = Application.Match(value, ws.Columns(cols("Name")), 0)
        If IsError(found) Then Exit Function
        r = CLng(found)
    End If
    t = CStr(ws.Cells(r, cols("TypeCode")).Value)
    If Trim$(allowedTypes) = "" Then
        EntityRefOK = True
    Else
        EntityRefOK = (InStr(1, "," & Replace(allowedTypes, " ", "") & ",", "," & t & ",", vbTextCompare) > 0)
    End If
End Function

' ---------------------------------------------------------------- name index for Unique rules
Public Function BuildNameIndex() As Object
    Dim idx As Object, d As Variant, i As Long, cols As Object, lo As ListObject, k As String
    Dim cId As Long, cT As Long, cN As Long, base As Long
    Set idx = modCore.NewDict()
    Set lo = MasterLO(): Set cols = MasterCols()
    d = modCore.TableData(lo)
    If Not IsArray(d) Then Set BuildNameIndex = idx: Exit Function
    base = lo.Range.Column - 1
    cId = cols("EntityID") - base: cT = cols("TypeCode") - base: cN = cols("Name") - base
    For i = 1 To UBound(d, 1)
        If CStr(d(i, cId)) <> "" Then
            k = CStr(d(i, cT)) & "|" & LCase$(Trim$(CStr(d(i, cN))))
            If Not idx.Exists(k) Then idx.Add k, modCore.NewDict()
            idx(k)(CStr(d(i, cId))) = True
            k = "*|" & LCase$(Trim$(CStr(d(i, cN))))
            If Not idx.Exists(k) Then idx.Add k, modCore.NewDict()
            idx(k)(CStr(d(i, cId))) = True
        End If
    Next i
    Set BuildNameIndex = idx
End Function

' ---------------------------------------------------------------- display IDs
Public Function DisplayInUse(ByVal disp As String) As Boolean
    Dim m As Variant
    m = Application.Match(disp, MasterSheet().Columns(MasterCols()("DisplayID")), 0)
    DisplayInUse = Not IsError(m)
End Function

' After keeping a Display ID such as APP-0012, make sure the counter for APP never hands it out again
Public Sub KeepCounterAbove(ByVal typeCode As String, ByVal disp As String)
    Dim lo As ListObject, pfx As String, r As Variant, n As Long, p As Long
    pfx = modMeta.gTypes(typeCode)("Prefix")
    If pfx = "" Then pfx = UCase$(Left$(typeCode, 3))
    p = InStrRev(disp, "-")
    If p = 0 Then Exit Sub
    If UCase$(Left$(disp, p - 1)) <> UCase$(pfx) Then Exit Sub
    n = Val(Mid$(disp, p + 1))
    Set lo = modCore.Tbl(modCore.SH_COUNTERS, "tblCounters")
    r = Application.Match(pfx, lo.ListColumns("IDPrefix").DataBodyRange, 0)
    modCore.UnprotectSheet lo.Parent
    If IsError(r) Then
        modCore.AppendRow lo, Array(pfx, n + 1)
    ElseIf Val(CStr(lo.ListColumns("NextNo").DataBodyRange.Cells(r, 1).Value)) <= n Then
        lo.ListColumns("NextNo").DataBodyRange.Cells(r, 1).Value = n + 1
    End If
    modCore.ProtectSheet lo.Parent
End Sub

Public Function NextDisplayID(ByVal typeCode As String) As String
    Dim lo As ListObject, pfx As String, r As Variant, n As Long
    modMeta.LoadMeta
    pfx = modMeta.gTypes(typeCode)("Prefix")
    If pfx = "" Then pfx = UCase$(Left$(typeCode, 3))
    Set lo = modCore.Tbl(modCore.SH_COUNTERS, "tblCounters")
    r = Application.Match(pfx, lo.ListColumns("IDPrefix").DataBodyRange, 0)
    If IsError(r) Then
        n = 1
        modCore.UnprotectSheet lo.Parent              ' table rows cannot be added on a protected sheet
        modCore.AppendRow lo, Array(pfx, 2)
        modCore.ProtectSheet lo.Parent
    Else
        n = Val(CStr(lo.ListColumns("NextNo").DataBodyRange.Cells(r, 1).Value))
        If n < 1 Then n = 1
        lo.ListColumns("NextNo").DataBodyRange.Cells(r, 1).Value = n + 1
    End If
    NextDisplayID = pfx & "-" & Format$(n, "0000")
End Function

' ---------------------------------------------------------------- hashing and stamps
Public Function MasterRowHash(ByVal r As Long, Optional cols As Object = Nothing) As String
    Dim vals As New Collection, keys As New Collection, k As Variant, ws As Worksheet
    If cols Is Nothing Then Set cols = MasterCols()
    Set ws = MasterSheet()
    For Each k In Split(HASH_FIELDS_SYS, "|")
        keys.Add CStr(k)
        vals.Add ws.Cells(r, cols(CStr(k))).Value
    Next k
    For Each k In MasterAttrOrder()
        keys.Add CStr(k)
        vals.Add ws.Cells(r, cols(CStr(k))).Value
    Next k
    MasterRowHash = KeyedHash(keys, vals)
End Function

' Keyed hash of a master row (mirrors model.master_hash): the 4 system fields always count,
' attributes only when not blank - so adding an attribute column leaves existing hashes alone.
Public Function KeyedHash(keys As Collection, vals As Collection) As String
    Dim s As String, i As Long, t As String, first As Boolean
    first = True
    For i = 1 To keys.Count
        t = modCore.CellText(vals(i))
        If i <= 4 Or t <> "" Then
            If Not first Then s = s & ChrW$(30)
            s = s & keys(i) & "=" & t
            first = False
        End If
    Next i
    KeyedHash = modCore.HashText(s)
End Function

Public Sub StampRow(ByVal r As Long, Optional ByVal isNew As Boolean = False, Optional cols As Object = Nothing)
    Dim ws As Worksheet, ts As String, u As String
    If cols Is Nothing Then Set cols = MasterCols()
    Set ws = MasterSheet()
    ts = modCore.UtcNow(): u = modCore.CurrentUser()
    If isNew Then
        ws.Cells(r, cols("Created")).Value = ts
        ws.Cells(r, cols("CreatedBy")).Value = u
    End If
    ws.Cells(r, cols("Modified")).Value = ts
    ws.Cells(r, cols("ModifiedBy")).Value = u
    ws.Cells(r, cols("RowHash")).Value = MasterRowHash(r, cols)
End Sub

' ---------------------------------------------------------------- create / update entity
' values: Dictionary AttrCode -> value. Returns new EntityID, or "" with issues filled.
' withID / withDisplay keep an entity's identity when it comes from a child model (merge)
Public Function CreateEntity(ByVal typeCode As String, values As Object, ByVal source As String, _
                             Optional ByVal batchID As String = "", Optional issues As Collection, _
                             Optional ByVal withID As String = "", Optional ByVal withDisplay As String = "") As String
    modMeta.LoadMeta
    If issues Is Nothing Then Set issues = New Collection
    If Not modMeta.TypeExists(typeCode) Then
        issues.Add "Error|Type|TypeCode|Unknown entity type " & typeCode
        Exit Function
    End If
    If Not values.Exists("Status") Then values("Status") = "Active"
    If modCore.IsBlankValue(values("Status")) Then values("Status") = "Active"
    If Not values.Exists("RepositoryArea") Then values("RepositoryArea") = "Architecture Landscape"
    If modCore.IsBlankValue(values("RepositoryArea")) Then values("RepositoryArea") = "Architecture Landscape"

    Dim found As Collection, it As Variant
    Set found = modRules.ValidateRow(typeCode, values, BuildNameIndex(), "")
    For Each it In found: issues.Add it: Next it
    If modRules.HasErrors(issues) Then Exit Function

    Dim ws As Worksheet, lo As ListObject, cols As Object, rng As Range, r As Long, id As String, k As Variant
    Dim dispID As String
    If batchID = "" Then batchID = modCore.NewGUID()
    Set ws = MasterSheet(): Set lo = MasterLO(): Set cols = MasterCols()
    modCore.BeginBusy
    On Error GoTo Fail
    If withDisplay <> "" And Not DisplayInUse(withDisplay) Then
        dispID = withDisplay
        KeepCounterAbove typeCode, withDisplay
    Else
        dispID = NextDisplayID(typeCode)
    End If
    modCore.UnprotectSheet ws
    If withID <> "" Then id = withID Else id = modCore.NewGUID()
    Set rng = modCore.AppendRow(lo, Array(id))
    r = rng.Row
    ws.Cells(r, cols("TypeCode")).Value = typeCode
    ws.Cells(r, cols("DisplayID")).Value = dispID
    For Each k In values.Keys
        If cols.Exists(k) And Not IsStampCol(CStr(k)) Then
            ws.Cells(r, cols(k)).Value = modCore.SafeFor(ToCellValue(CStr(k), values(k)), CStr(ws.Cells(r, cols(k)).NumberFormat))
        End If
    Next k
    StampRow r, True, cols
    modCore.ProtectSheet ws
    WriteHistory typeCode, "Create", id, "Name", "", values("Name"), source, batchID
    For Each k In values.Keys
        If Not modMeta.IsSystemAttr(CStr(k)) And Not modCore.IsBlankValue(values(k)) Then
            WriteHistory typeCode, "Create", id, CStr(k), "", values(k), source, batchID
        End If
    Next k
    modCore.EndBusy
    CreateEntity = id
    Exit Function
Fail:
    Dim failMsg As String
    failMsg = Err.Description
    On Error Resume Next
    ' remove a half-written row so no orphan is left in DATA_Master
    If r > 0 Then
        modCore.UnprotectSheet ws
        lo.ListRows(r - lo.HeaderRowRange.Row).Delete
    End If
    modCore.ProtectSheet ws
    modCore.ResetBusy
    issues.Add "Error|System||Could not create entity: " & failMsg
End Function

' Applies the given values to an existing entity. Unchanged values are ignored.
' Returns True when something changed. Validation runs on the merged row.
Public Function UpdateEntity(ByVal entityID As String, values As Object, ByVal source As String, _
                             Optional ByVal batchID As String = "", Optional issues As Collection, _
                             Optional ByRef blocked As Boolean = False) As Boolean
    modMeta.LoadMeta
    If issues Is Nothing Then Set issues = New Collection
    blocked = False
    Dim r As Long, cols As Object, cur As Object, merged As Object, k As Variant, typeCode As String
    r = FindEntityRow(entityID)
    If r = 0 Then issues.Add "Error|Missing|EntityID|Entity not found in EA_Core": blocked = True: Exit Function
    Set cols = MasterCols()
    Set cur = ReadEntityRow(r, cols)
    typeCode = CStr(cur("TypeCode"))

    Set merged = modCore.NewDict()
    For Each k In cur.Keys: merged(k) = cur(k): Next k
    Dim changed As Object
    Set changed = modCore.NewDict()
    For Each k In values.Keys
        If cols.Exists(k) And Not IsStampCol(CStr(k)) And k <> "EntityID" And k <> "TypeCode" And k <> "DisplayID" Then
            If modCore.CellText(values(k)) <> modCore.CellText(cur(k)) Then
                changed(k) = True
                merged(k) = values(k)
            End If
        End If
    Next k
    If changed.Count = 0 Then Exit Function

    Dim found As Collection, it As Variant
    Set found = modRules.ValidateRow(typeCode, merged, BuildNameIndex(), entityID)
    ' only block on errors that involve changed fields or row-level rules
    For Each it In found: issues.Add it: Next it
    If modRules.HasErrors(issues) Then
        Dim blocking As Boolean, parts() As String
        For Each it In issues
            parts = Split(it, "|", 4)
            If parts(0) = "Error" Then
                If changed.Exists(parts(2)) Then blocking = True
            End If
        Next it
        If blocking Then blocked = True: Exit Function
    End If

    Dim ws As Worksheet
    If batchID = "" Then batchID = modCore.NewGUID()
    Set ws = MasterSheet()
    modCore.BeginBusy
    On Error GoTo Fail
    modCore.ProtectSheet ws          ' re-applies UserInterfaceOnly so code can write
    For Each k In changed.Keys
        ws.Cells(r, cols(k)).Value = modCore.SafeFor(ToCellValue(CStr(k), merged(k)), CStr(ws.Cells(r, cols(k)).NumberFormat))
        WriteHistory typeCode, "Update", entityID, CStr(k), cur(k), merged(k), source, batchID
    Next k
    StampRow r, False, cols
    modCore.EndBusy
    UpdateEntity = True
    Exit Function
Fail:
    modCore.ResetBusy
    blocked = True
    issues.Add "Error|System||Could not update entity: " & Err.Description
End Function

Public Function RetireEntity(ByVal entityID As String, ByVal source As String) As Boolean
    Dim v As Object
    Set v = modCore.NewDict()
    v("Status") = "Retired"
    RetireEntity = UpdateEntity(entityID, v, source)
End Function

Private Function IsStampCol(ByVal k As String) As Boolean
    Select Case k
        Case "Created", "CreatedBy", "Modified", "ModifiedBy", "RowHash": IsStampCol = True
    End Select
End Function

' Converts form or sheet input to the value stored for the attribute's data type
Public Function ToCellValue(ByVal attrCode As String, ByVal v As Variant) As Variant
    ToCellValue = v
    If modCore.IsBlankValue(v) Then ToCellValue = Empty: Exit Function
    If Not modMeta.gAttrs.Exists(attrCode) Then Exit Function
    Select Case modMeta.gAttrs(attrCode)("DataType")
        Case "Date"
            If modRules.IsDateValue(v) Then ToCellValue = modRules.ToDate(v)
        Case "Integer"
            If modRules.IsWhole(v) Then ToCellValue = CLng(v)
        Case "Decimal", "Currency", "Percent"
            If IsNumeric(v) Then ToCellValue = CDbl(v)
    End Select
End Function

' ---------------------------------------------------------------- relationships
Public Function RelLO() As ListObject
    Set RelLO = ThisWorkbook.Worksheets(modCore.SH_REL).ListObjects("tblRel")
End Function

Public Function CreateRelationship(ByVal relCode As String, ByVal sourceID As String, ByVal targetID As String, _
                                   ByVal label As String, ByVal descr As String, ByVal source As String, _
                                   Optional ByVal batchID As String = "", Optional issues As Collection, _
                                   Optional ByVal withID As String = "") As String
    modMeta.LoadMeta
    If issues Is Nothing Then Set issues = New Collection
    Dim st As String, tt As String
    st = EntityTypeOf(sourceID): tt = EntityTypeOf(targetID)
    If st = "" Then issues.Add "Error|Missing|SourceID|Source entity not found"
    If tt = "" Then issues.Add "Error|Missing|TargetID|Target entity not found"
    If Not modMeta.gRelTypes.Exists(relCode) Then issues.Add "Error|Type|RelCode|Unknown relationship type " & relCode
    If issues.Count > 0 Then Exit Function
    If Not modMeta.IsRelAllowed(st, relCode, tt) Then
        issues.Add "Error|NotAllowed|RelCode|" & modMeta.TypeName_(st) & " " & modMeta.RelVerb(relCode) & " " & _
                   modMeta.TypeName_(tt) & " is not allowed by the metamodel"
        Exit Function
    End If
    If RelationshipExists(relCode, sourceID, targetID) Then
        issues.Add "Error|Duplicate|RelCode|This relationship already exists"
        Exit Function
    End If

    Dim lo As ListObject, ws As Worksheet, rng As Range, r As Long, id As String, ts As String, u As String
    If batchID = "" Then batchID = modCore.NewGUID()
    Set lo = RelLO(): Set ws = lo.Parent
    modCore.BeginBusy
    On Error GoTo Fail
    modCore.UnprotectSheet ws
    If withID <> "" Then id = withID Else id = modCore.NewGUID()
    ts = modCore.UtcNow(): u = modCore.CurrentUser()
    Set rng = modCore.AppendRow(lo, Array(id, relCode, sourceID, targetID, Empty, Empty, Empty, Empty, _
                                         modCore.SafeValue(label), modCore.SafeValue(descr), "Active", ts, u, ts, u, ""))
    r = rng.Row
    WriteRelFormulas ws, r
    ws.Cells(r, ColOfRel("RowHash")).Value = RelRowHash(relCode, sourceID, targetID, label, descr, "Active")
    modCore.ProtectSheet ws
    WriteHistory "", "Create", id, "RelCode", "", relCode & ": " & EntityNameOf(sourceID) & " -> " & _
                 EntityNameOf(targetID), source, batchID, modCore.SH_HREL
    modCore.EndBusy
    CreateRelationship = id
    Exit Function
Fail:
    modCore.ProtectSheet ws
    modCore.ResetBusy
    issues.Add "Error|System||Could not create relationship: " & Err.Description
End Function

' Label, Description, Status of a relationship, with history and a fresh RowHash
Public Function UpdateRelationship(ByVal relID As String, upd As Object, ByVal src As String, ByVal batch As String) As Boolean
    Dim lo As ListObject, ws As Worksheet, m As Variant, r As Long, k As Variant, c As Long, oldV As Variant
    On Error GoTo Fail
    Set lo = RelLO(): Set ws = lo.Parent
    m = Application.Match(relID, lo.ListColumns(1).DataBodyRange, 0)
    If IsError(m) Then Exit Function
    r = lo.DataBodyRange.Row + CLng(m) - 1
    modCore.UnprotectSheet ws
    For Each k In upd.Keys
        c = ColOfRel(CStr(k))
        oldV = ws.Cells(r, c).Value
        ws.Cells(r, c).Value = modCore.SafeValue(upd(k))
        WriteHistory "", "Merge update", relID, CStr(k), oldV, upd(k), src, batch, modCore.SH_HREL
    Next k
    ws.Cells(r, ColOfRel("Modified")).Value = modCore.UtcNow()
    ws.Cells(r, ColOfRel("ModifiedBy")).Value = modCore.CurrentUser()
    ws.Cells(r, ColOfRel("RowHash")).Value = RelRowHash(CStr(ws.Cells(r, ColOfRel("RelCode")).Value), _
        CStr(ws.Cells(r, ColOfRel("SourceID")).Value), CStr(ws.Cells(r, ColOfRel("TargetID")).Value), _
        CStr(ws.Cells(r, ColOfRel("Label")).Value), CStr(ws.Cells(r, ColOfRel("Description")).Value), _
        CStr(ws.Cells(r, ColOfRel("Status")).Value))
    modCore.ProtectSheet ws
    UpdateRelationship = True
    Exit Function
Fail:
    modCore.ProtectSheet ws
End Function

Public Function RelationshipExists(ByVal relCode As String, ByVal sourceID As String, ByVal targetID As String) As Boolean
    Dim d As Variant, i As Long, lo As ListObject
    Set lo = RelLO()
    d = modCore.TableData(lo)
    If Not IsArray(d) Then Exit Function
    For i = 1 To UBound(d, 1)
        If d(i, 2) = relCode And d(i, 3) = sourceID And d(i, 4) = targetID And d(i, 11) <> "Retired" Then
            RelationshipExists = True: Exit Function
        End If
    Next i
End Function

Public Sub WriteRelFormulas(ws As Worksheet, ByVal r As Long)
    ws.Cells(r, 5).Formula = "=IFERROR(INDEX(DATA_Master!$D:$D,MATCH(C" & r & ",DATA_Master!$A:$A,0)),""(missing)"")"
    ws.Cells(r, 6).Formula = "=IFERROR(INDEX(DATA_Master!$B:$B,MATCH(C" & r & ",DATA_Master!$A:$A,0)),""(missing)"")"
    ws.Cells(r, 7).Formula = "=IFERROR(INDEX(DATA_Master!$D:$D,MATCH(D" & r & ",DATA_Master!$A:$A,0)),""(missing)"")"
    ws.Cells(r, 8).Formula = "=IFERROR(INDEX(DATA_Master!$B:$B,MATCH(D" & r & ",DATA_Master!$A:$A,0)),""(missing)"")"
End Sub

Public Function ColOfRel(ByVal header As String) As Long
    ColOfRel = RelLO().ListColumns(header).Range.Column
End Function

Public Function RelRowHash(ByVal relCode As String, ByVal s As String, ByVal t As String, ByVal label As String, _
                           ByVal descr As String, ByVal status As String) As String
    Dim c As New Collection
    c.Add relCode: c.Add s: c.Add t: c.Add label: c.Add descr: c.Add status
    RelRowHash = modCore.RowHashOf(c)
End Function

' ---------------------------------------------------------------- history
' typeCode = "" with explicit sheetName for relationship / metamodel history
Public Sub WriteHistory(ByVal typeCode As String, ByVal action As String, ByVal recordID As String, _
                        ByVal attrCode As String, ByVal oldV As Variant, ByVal newV As Variant, _
                        ByVal source As String, ByVal batchID As String, Optional ByVal sheetName As String = "")
    Dim ws As Worksheet, lo As ListObject
    If sheetName = "" Then sheetName = EnsureHistorySheet(typeCode)
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set lo = ws.ListObjects(1)
    modCore.UnprotectSheet ws
    modCore.AppendRow lo, Array(modCore.NewGUID(), modCore.UtcNow(), modCore.CurrentUser(), action, recordID, _
                                attrCode, "'" & modCore.CellText(oldV), "'" & modCore.CellText(newV), source, batchID)
    modCore.ProtectSheet ws
End Sub

Public Function EnsureHistorySheet(ByVal typeCode As String) As String
    Dim nm As String, ws As Worksheet, lo As ListObject, hdr As Variant, i As Long
    nm = Left$("H_" & typeCode, 31)
    EnsureHistorySheet = nm
    If modCore.SheetExists(ThisWorkbook, nm) Then Exit Function
    Dim prevSheet As Object, prevBook As Workbook
    Set prevBook = ActiveWorkbook
    If Not prevBook Is Nothing Then Set prevSheet = prevBook.ActiveSheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(modCore.SH_HREL))
    ws.Name = nm
    hdr = Array("ChangeID", "Timestamp", "User", "Action", "RecordID", "AttrCode", "OldValue", "NewValue", _
                "Source", "BatchID")
    For i = 0 To UBound(hdr)
        ws.Cells(1, i + 1).Value = hdr(i)
    Next i
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, UBound(hdr) + 1), , xlYes)
    lo.Name = "tblH_" & Left$(typeCode, 25)
    lo.TableStyle = "TableStyleLight1"
    ws.Columns("A").ColumnWidth = 38: ws.Columns("E").ColumnWidth = 38: ws.Columns("J").ColumnWidth = 38
    ws.Columns("G:H").ColumnWidth = 30
    ws.Tab.Color = RGB(127, 127, 127)
    modCore.ProtectSheet ws
    ' Worksheets.Add switched the active sheet; put the user back where they were
    On Error Resume Next
    If Not prevBook Is Nothing Then prevBook.Activate
    If Not prevSheet Is Nothing Then prevSheet.Activate
End Function
