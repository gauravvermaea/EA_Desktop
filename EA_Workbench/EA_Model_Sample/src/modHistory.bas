Attribute VB_Name = "modHistory"
' =====================================================================
' Versions (Phase 2).
'   Snapshots  - values-only copies of EA_Core (MM_, DATA_, H_ sheets)
'                and of EA_Entities' filled type sheets, saved as a pair
'                in History\ and registered in SYS_Versions.
'   Compare    - current model against a snapshot, or two snapshots,
'                listed on VERSION_Compare.
'   Restore    - one entity, all entities of a type, or the whole model,
'                always after an automatic snapshot, always logged.
'   History    - the change history of one entity on VERSION_Compare.
'   Archiving  - history sheets past 100,000 rows roll into an archive.
' =====================================================================
Option Explicit

Private Const SH_VERSIONS As String = "SYS_Versions"
Private Const SH_FORM As String = "FORM_Versions"
Private Const SH_COMPARE As String = "VERSION_Compare"
Private Const LIST_FIRST_ROW As Long = 9
Private Const ARCHIVE_AT As Long = 100000
Private Const ARCHIVE_KEEP As Long = 50000

' ================================================================ snapshots
Private Function HistoryFolder() As String
    Dim p As String
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then Err.Raise vbObjectError + 301, , _
        "The model folder is on a web location; move it to a local folder to take snapshots."
    p = ThisWorkbook.Path & Application.PathSeparator & "History"
    If Dir(p, vbDirectory) = "" Then MkDir p
    HistoryFolder = p & Application.PathSeparator
End Function

Private Function VersionsLO() As ListObject
    Set VersionsLO = ThisWorkbook.Worksheets(SH_VERSIONS).ListObjects("tblVersions")
End Function

Private Function NextVersionNo() As Long
    Dim d As Variant, i As Long, m As Long
    d = modCore.TableData(VersionsLO())
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If Val(CStr(d(i, 1))) > m Then m = Val(CStr(d(i, 1)))
        Next i
    End If
    NextVersionNo = m + 1
End Function

Private Function SafeName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_-]" Then
            out = out & ch
        ElseIf ch = " " Then
            out = out & "_"
        End If
    Next i
    SafeName = Left$(out, 40)
End Function

' Takes a snapshot pair. Returns the version number (0 if it failed).
Public Function TakeSnapshot(ByVal label As String, ByVal isAuto As Boolean, ByVal reason As String) As Long
    Dim folder As String, ver As Long, base As String, coreFile As String, entFile As String
    Dim wbS As Workbook, ws As Worksheet, wbE As Workbook, nE As Long, nR As Long
    On Error GoTo Fail
    folder = HistoryFolder()
    ver = NextVersionNo()
    base = SafeName(CStr(modCore.Setting("ModelName"))) & "-v" & Format$(ver, "000") & "-" & Format$(Now, "yyyymmdd-hhnnss")
    If label <> "" Then base = base & "-" & SafeName(label)
    coreFile = base & "-core.xlsx"
    entFile = base & "-entities.xlsx"
    modCore.BeginBusy
    modCore.Status "snapshot " & ver & ": EA_Core"

    ' ---- EA_Core: metamodel, data, history
    Set wbS = Workbooks.Add(xlWBATWorksheet)
    InfoSheet wbS.Worksheets(1), ver, label, reason
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "MM_*" Or ws.Name Like "DATA_*" Or ws.Name Like "H_*" Then CopyValues ws, wbS
    Next ws
    wbS.SaveAs Filename:=folder & coreFile, FileFormat:=51
    wbS.Close SaveChanges:=False

    ' ---- EA_Entities: type sheets that hold entities
    Set wbE = modCore.EntitiesWB(False)
    If Not wbE Is Nothing Then
        modCore.Status "snapshot " & ver & ": EA_Entities"
        Set wbS = Workbooks.Add(xlWBATWorksheet)
        InfoSheet wbS.Worksheets(1), ver, label, reason
        For Each ws In wbE.Worksheets
            If ws.Name Like "T_*" Then
                If Application.WorksheetFunction.CountIf(ws.Columns(2), "E") > 0 Then CopyValues ws, wbS
            End If
        Next ws
        wbS.SaveAs Filename:=folder & entFile, FileFormat:=51
        wbS.Close SaveChanges:=False
    Else
        entFile = ""
    End If

    nE = CountRows(modData.MasterLO())
    nR = CountRows(modData.RelLO())
    modCore.UnprotectSheet VersionsLO().Parent
    modCore.AppendRow VersionsLO(), Array(ver, label, modCore.UtcNow(), modCore.CurrentUser(), IIf(isAuto, "Y", "N"), _
                                          reason, coreFile, entFile, nE, nR, modMeta.SchemaVersion())
    modCore.ProtectSheet VersionsLO().Parent
    modCore.EndBusy
    modCore.Status "snapshot " & ver & " saved"
    TakeSnapshot = ver
    Exit Function
Fail:
    Dim msg As String
    msg = Err.Description
    On Error Resume Next
    If Not wbS Is Nothing Then
        If Not wbS Is ThisWorkbook Then wbS.Close SaveChanges:=False
    End If
    modCore.ResetBusy
    MsgBox "Snapshot failed: " & msg, vbExclamation, "EA Workbench"
End Function

' Automatic snapshot before a risky operation. Returns False if the user cancels after a failure.
Public Function AutoSnapshot(ByVal reason As String) As Boolean
    If TakeSnapshot("auto", True, reason) > 0 Then
        AutoSnapshot = True
    Else
        AutoSnapshot = (MsgBox("The automatic snapshot before """ & reason & """ failed. Continue without it?", _
                               vbYesNo + vbExclamation, "EA Workbench") = vbYes)
    End If
End Function

Private Function CountRows(lo As ListObject) As Long
    Dim d As Variant
    d = modCore.TableData(lo)
    If IsArray(d) Then CountRows = UBound(d, 1)
End Function

Private Sub InfoSheet(ws As Worksheet, ByVal ver As Long, ByVal label As String, ByVal reason As String)
    Dim info As Variant, i As Long
    ws.Name = "Snapshot"
    info = Array("Model", CStr(modCore.Setting("ModelName")), "ModelID", CStr(modCore.Setting("ModelID")), _
                 "Version", ver, "Label", label, "Reason", reason, "Taken (UTC)", modCore.UtcNow(), _
                 "Taken by", modCore.CurrentUser(), "SchemaVersion", modMeta.SchemaVersion())
    For i = 0 To UBound(info) Step 2
        ws.Cells(i \ 2 + 1, 1).Value = info(i)
        ws.Cells(i \ 2 + 1, 2).Value = modCore.SafeValue(info(i + 1))
    Next i
    ws.Columns("A:B").AutoFit
End Sub

' Values-only copy of a sheet's used area (from A1) into a new sheet of the same name
Private Sub CopyValues(src As Worksheet, wbS As Workbook)
    Dim lastR As Long, lastC As Long, dst As Worksheet, v As Variant, i As Long, j As Long
    Set dst = wbS.Worksheets.Add(After:=wbS.Worksheets(wbS.Worksheets.Count))
    dst.Name = src.Name
    lastR = modCore.LastRowOf(src): lastC = modCore.LastColOf(src)
    If lastR = 0 Or lastC = 0 Then Exit Sub
    v = src.Range(src.Cells(1, 1), src.Cells(lastR, lastC)).Value
    If Not IsArray(v) Then
        dst.Range("A1").Value = modCore.SafeValue(v)
        Exit Sub
    End If
    For i = 1 To UBound(v, 1)
        For j = 1 To UBound(v, 2)
            If IsError(v(i, j)) Then
                v(i, j) = Empty
            ElseIf VarType(v(i, j)) = vbString Then
                v(i, j) = "'" & v(i, j)                   ' text stays text under any regional settings
            End If
        Next j
    Next i
    dst.Range("A1").Resize(lastR, lastC).Value = v
End Sub

' ================================================================ Version history sheet
Public Sub ShowVersions()
    Dim ws As Worksheet, d As Variant, i As Long, r As Long, core As String, j As Long, b As Variant, n As Long
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    core = "'" & ThisWorkbook.Name & "'!modHistory."
    modCore.BeginBusy
    On Error GoTo Done
    modCore.UnprotectSheet ws
    modCore.DeleteButtons ws
    ws.Range("A3:Z5000").Clear
    ws.Range("A1").Value = "Version history"
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = RGB(31, 56, 100)
    ws.Range("A2").Value = "Select a row in the list, then use a button. Snapshots are kept in the History folder."
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Color = RGB(89, 89, 89)
    For Each b In Array( _
        Array("Take snapshot", "TakeSnapshotButton"), _
        Array("Compare selected with current", "CompareSelectedWithCurrent"), _
        Array("Compare two versions", "CompareTwoVersions"), _
        Array("Entity history", "ShowEntityHistory"), _
        Array("Restore entity or type", "RestoreEntityOrType"), _
        Array("Restore whole model", "RestoreWholeModel"), _
        Array("Back to Home", "GoHomeFromVersions"))
        modCore.AddButton ws, CStr(b(0)), core & CStr(b(1)), ws.Cells(3 + (n \ 4) * 2, 1 + (n Mod 4) * 3), 190
        n = n + 1
    Next b
    r = LIST_FIRST_ROW - 1
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Value = Array("Version", "Label", "Taken (UTC)", "By", "Auto", "Reason", _
                                                           "Entities", "Relationships")
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Interior.Color = RGB(221, 235, 247)
    d = modCore.TableData(VersionsLO())
    If IsArray(d) Then
        For i = UBound(d, 1) To 1 Step -1                ' newest first
            r = r + 1
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Value = Array(d(i, 1), modCore.SafeValue(CStr(d(i, 2))), d(i, 3), _
                                                                   d(i, 4), d(i, 5), modCore.SafeValue(CStr(d(i, 6))), d(i, 9), d(i, 10))
        Next i
    Else
        ws.Cells(r + 1, 1).Value = "No snapshots yet."
    End If
    For j = 1 To 8
        ws.Columns(j).ColumnWidth = Choose(j, 9, 28, 22, 16, 7, 28, 10, 13)
    Next j
    ThisWorkbook.Activate
    ws.Activate
    ws.Cells(LIST_FIRST_ROW, 1).Select
Done:
    modCore.ProtectSheet ws
    modCore.EndBusy
End Sub

Public Sub GoHomeFromVersions()
    modForms.GoHome
End Sub

Public Sub TakeSnapshotButton()
    Dim label As String, v As Long
    modCore.Heal
    label = InputBox("Label for this snapshot (optional), e.g. ""before workshop"":", "Take snapshot")
    If StrPtr(label) = 0 Then Exit Sub
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    v = TakeSnapshot(label, False, "Manual")
    If v > 0 Then
        ShowVersions
        MsgBox "Snapshot " & v & " saved in the History folder.", vbInformation, "EA Workbench"
    End If
End Sub

' Version number of the selected row on FORM_Versions, or asked for
Private Function SelectedVersion(ByVal prompt As String) As Long
    Dim s As String
    If ActiveSheet.Name = SH_FORM And ActiveWorkbook Is ThisWorkbook Then
        If ActiveCell.Row >= LIST_FIRST_ROW Then SelectedVersion = Val(CStr(ActiveSheet.Cells(ActiveCell.Row, 1).Value))
    End If
    If SelectedVersion = 0 Then
        s = InputBox(prompt & vbLf & "Version number:", "EA Workbench")
        SelectedVersion = Val(s)
    End If
End Function

Private Function VersionFile(ByVal ver As Long, ByVal which As String) As String
    Dim lo As ListObject, r As Variant
    Set lo = VersionsLO()
    If lo.DataBodyRange Is Nothing Then Exit Function
    r = Application.Match(ver, lo.ListColumns("VersionNo").DataBodyRange, 0)
    If IsError(r) Then Exit Function
    VersionFile = CStr(lo.ListColumns(IIf(which = "core", "CoreFile", "EntitiesFile")).DataBodyRange.Cells(r, 1).Value)
    If VersionFile <> "" Then VersionFile = HistoryFolder() & VersionFile
    If Dir(VersionFile) = "" Then VersionFile = ""
End Function

Private Function OpenSnapshot(ByVal ver As Long) As Workbook
    Dim p As String
    p = VersionFile(ver, "core")
    If p = "" Then
        MsgBox "The snapshot file for version " & ver & " was not found in the History folder.", vbExclamation, "EA Workbench"
        Exit Function
    End If
    Set OpenSnapshot = Workbooks.Open(p, UpdateLinks:=0, ReadOnly:=True)
    ThisWorkbook.Activate
End Function

' ================================================================ model loading and diff
' Model as Dictionary: "E|id" / "R|id" / "M|sheet|key" -> Dictionary(field -> value)
Private Function LoadModel(wb As Workbook) As Object
    Dim m As Object, ws As Worksheet
    Set m = modCore.NewDict()
    For Each ws In wb.Worksheets
        Select Case True
            Case ws.Name = "DATA_Master": LoadBlock m, ws, "E", 1, 3, "EntityID", ""
            Case ws.Name = "DATA_Relationships": LoadBlock m, ws, "R", 1, 2, "RelID", ""
            Case ws.Name = "MM_TypeGroups", ws.Name = "MM_Lists": LoadBlock m, ws, "M|" & ws.Name, 1, 2, "", "pair"
            Case ws.Name Like "MM_*": LoadBlock m, ws, "M|" & ws.Name, 1, 2, "", "first"
        End Select
    Next ws
    Set LoadModel = m
End Function

Private Sub LoadBlock(m As Object, ws As Worksheet, ByVal prefix As String, ByVal keyRow As Long, _
                      ByVal firstRow As Long, ByVal idKey As String, ByVal keyMode As String)
    Dim lastR As Long, lastC As Long, v As Variant, i As Long, j As Long, keys() As String, idCol As Long
    Dim rec As Object, id As String
    lastR = modCore.LastRowOf(ws): lastC = modCore.LastColOf(ws)
    If lastR < firstRow Then Exit Sub
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    ReDim keys(1 To lastC)
    For j = 1 To lastC
        keys(j) = CStr(modCore.Nz(v(keyRow, j), ""))
        If keys(j) = idKey And idKey <> "" Then idCol = j
    Next j
    If idKey <> "" And idCol = 0 Then Exit Sub
    For i = firstRow To lastR
        Select Case keyMode
            Case "pair": id = CStr(modCore.Nz(v(i, 1), "")) & "|" & CStr(modCore.Nz(v(i, 2), ""))
            Case "first": id = CStr(modCore.Nz(v(i, 1), ""))
            Case Else: id = CStr(modCore.Nz(v(i, idCol), ""))
        End Select
        If id <> "" And id <> "|" Then
            Set rec = modCore.NewDict()
            For j = 1 To lastC
                If keys(j) <> "" Then rec(keys(j)) = v(i, j)
            Next j
            Set m(prefix & "|" & id) = rec
        End If
    Next i
End Sub

Private Function IgnoredField(ByVal f As String) As Boolean
    Select Case f
        Case "Created", "CreatedBy", "Modified", "ModifiedBy", "RowHash", "SourceName", "SourceType", _
             "TargetName", "TargetType"
            IgnoredField = True
    End Select
End Function

' Differences old -> new as rows: Area, Change, ID, Display ID, Name, Type, Field, Old, New
Private Function DiffModels(oldM As Object, newM As Object) As Collection
    Dim out As New Collection, k As Variant, f As Variant, o As Object, n As Object
    For Each k In newM.Keys
        Set n = newM(k)
        If Not oldM.Exists(k) Then
            out.Add DiffRow(CStr(k), "Added", n, "", Empty, Empty)
        Else
            Set o = oldM(k)
            For Each f In n.Keys
                If Not IgnoredField(CStr(f)) Then
                    If o.Exists(f) Then
                        If modCore.CellText(o(f)) <> modCore.CellText(n(f)) Then out.Add DiffRow(CStr(k), "Changed", n, CStr(f), o(f), n(f))
                    ElseIf Not modCore.IsBlankValue(n(f)) Then
                        out.Add DiffRow(CStr(k), "Changed", n, CStr(f), Empty, n(f))
                    End If
                End If
            Next f
            For Each f In o.Keys
                If Not IgnoredField(CStr(f)) And Not n.Exists(f) Then
                    If Not modCore.IsBlankValue(o(f)) Then out.Add DiffRow(CStr(k), "Changed", n, CStr(f), o(f), Empty)
                End If
            Next f
        End If
    Next k
    For Each k In oldM.Keys
        If Not newM.Exists(k) Then out.Add DiffRow(CStr(k), "Removed", oldM(k), "", Empty, Empty)
    Next k
    Set DiffModels = out
End Function

Private Function DiffRow(ByVal k As String, ByVal change As String, rec As Object, ByVal field As String, _
                         ByVal oldV As Variant, ByVal newV As Variant) As Variant
    Dim parts() As String, area As String, id As String, disp As String, nm As String, tp As String
    parts = Split(k, "|")
    Select Case parts(0)
        Case "E"
            area = "Entity": id = parts(1)
            disp = CStr(modCore.Nz(RecVal(rec, "DisplayID"), "")): nm = CStr(modCore.Nz(RecVal(rec, "Name"), ""))
            tp = CStr(modCore.Nz(RecVal(rec, "TypeCode"), ""))
        Case "R"
            area = "Relationship": id = parts(1)
            nm = CStr(modCore.Nz(RecVal(rec, "SourceName"), "")) & " -> " & CStr(modCore.Nz(RecVal(rec, "TargetName"), ""))
            tp = CStr(modCore.Nz(RecVal(rec, "RelCode"), ""))
        Case Else
            area = "Metamodel " & parts(1): id = Mid$(k, Len(parts(0)) + Len(parts(1)) + 3)
    End Select
    DiffRow = Array(area, change, id, disp, nm, tp, field, oldV, newV)
End Function

Private Function RecVal(rec As Object, ByVal f As String) As Variant
    If rec.Exists(f) Then RecVal = rec(f)
End Function

Private Sub WriteCompareSheet(ByVal title As String, ByVal subtitle As String, rows As Collection, hdr As Variant)
    Dim ws As Worksheet, r As Long, x As Variant, j As Long, out() As Variant, n As Long
    Set ws = ThisWorkbook.Worksheets(SH_COMPARE)
    modCore.UnprotectSheet ws
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    ws.Range("A1").Value = title
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = RGB(31, 56, 100)
    ws.Range("A2").Value = subtitle
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Color = RGB(89, 89, 89)
    n = UBound(hdr) - LBound(hdr) + 1
    For j = 1 To n
        ws.Cells(4, j).Value = hdr(LBound(hdr) + j - 1)
    Next j
    With ws.Range(ws.Cells(4, 1), ws.Cells(4, n))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    If rows.Count > 0 Then
        ReDim out(1 To rows.Count, 1 To n)
        r = 0
        For Each x In rows
            r = r + 1
            For j = 1 To n
                If VarType(x(j - 1)) = vbString Then out(r, j) = modCore.SafeValue(x(j - 1)) Else out(r, j) = x(j - 1)
            Next j
        Next x
        ws.Range(ws.Cells(5, 1), ws.Cells(4 + rows.Count, n)).Value = out
        ws.Range(ws.Cells(4, 1), ws.Cells(4 + rows.Count, n)).AutoFilter
    Else
        ws.Range("A5").Value = "No differences."
    End If
    For j = 1 To n
        ws.Columns(j).ColumnWidth = 18
    Next j
    ThisWorkbook.Activate
    ws.Activate
    ws.Range("A5").Select
End Sub

Public Sub CompareSelectedWithCurrent()
    Dim ver As Long, wbS As Workbook, rows As Collection
    modCore.Heal
    ver = SelectedVersion("Compare which version with the current model?")
    If ver = 0 Then Exit Sub
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    Application.ScreenUpdating = False
    Set wbS = OpenSnapshot(ver)
    If wbS Is Nothing Then Application.ScreenUpdating = True: Exit Sub
    Set rows = DiffModels(LoadModel(wbS), LoadModel(ThisWorkbook))
    wbS.Close SaveChanges:=False
    Application.ScreenUpdating = True
    WriteCompareSheet "Version " & ver & " compared with the current model", _
        rows.Count & " difference(s). Old = version " & ver & ", New = current.", rows, _
        Array("Area", "Change", "ID", "Display ID", "Name", "Type / Rel", "Field", "Old", "New")
End Sub

Public Sub CompareTwoVersions()
    Dim s As String, a As Long, b As Long, wbA As Workbook, wbB As Workbook, rows As Collection
    modCore.Heal
    s = InputBox("Enter the two version numbers to compare, older first, e.g. 3,5", "Compare two versions")
    If InStr(s, ",") = 0 Then Exit Sub
    a = Val(Split(s, ",")(0)): b = Val(Split(s, ",")(1))
    If a = 0 Or b = 0 Or a = b Then Exit Sub
    Application.ScreenUpdating = False
    Set wbA = OpenSnapshot(a)
    If wbA Is Nothing Then Application.ScreenUpdating = True: Exit Sub
    Set wbB = OpenSnapshot(b)
    If wbB Is Nothing Then wbA.Close False: Application.ScreenUpdating = True: Exit Sub
    Set rows = DiffModels(LoadModel(wbA), LoadModel(wbB))
    wbA.Close SaveChanges:=False
    wbB.Close SaveChanges:=False
    Application.ScreenUpdating = True
    WriteCompareSheet "Version " & a & " compared with version " & b, _
        rows.Count & " difference(s). Old = version " & a & ", New = version " & b & ".", rows, _
        Array("Area", "Change", "ID", "Display ID", "Name", "Type / Rel", "Field", "Old", "New")
End Sub

' ================================================================ entity history
' From EA_Entities (selected row) or asked for by Display ID
Public Sub ShowEntityHistory()
    Dim id As String, disp As String, r As Long, typeCode As String, rows As New Collection, nm As String
    Dim ws As Worksheet, relIDs As Object, d As Variant, i As Long, sh As String
    modCore.Heal
    If ActiveSheet.Name Like "T_*" Then
        If CStr(ActiveSheet.Cells(ActiveCell.Row, 2).Value) = "E" Then id = CStr(ActiveSheet.Cells(ActiveCell.Row, 1).Value)
    End If
    If id = "" Then
        disp = InputBox("Display ID of the entity (e.g. APP-0002):", "Entity history")
        If disp = "" Then Exit Sub
        id = EntityIDByDisplay(disp)
        If id = "" Then MsgBox "No entity with Display ID " & disp & ".", vbExclamation, "EA Workbench": Exit Sub
    End If
    r = modData.FindEntityRow(id)
    If r = 0 Then MsgBox "That entity is not in EA_Core.", vbExclamation, "EA Workbench": Exit Sub
    typeCode = modData.EntityTypeOf(id)
    nm = modData.EntityNameOf(id)
    disp = CStr(modData.MasterSheet().Cells(r, modData.MasterCols()("DisplayID")).Value)

    sh = Left$("H_" & typeCode, 31)
    If modCore.SheetExists(ThisWorkbook, sh) Then CollectHistory ThisWorkbook.Worksheets(sh), id, Nothing, rows
    ' relationships that involve this entity
    Set relIDs = modCore.NewDict()
    d = modCore.TableData(modData.RelLO())
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If d(i, 3) = id Or d(i, 4) = id Then relIDs(CStr(d(i, 1))) = True
        Next i
    End If
    If relIDs.Count > 0 Then CollectHistory ThisWorkbook.Worksheets(modCore.SH_HREL), "", relIDs, rows
    WriteCompareSheet "History of " & disp & " " & nm, rows.Count & " change(s), newest first. Includes its relationships.", _
        SortNewestFirst(rows), Array("Timestamp (UTC)", "User", "Action", "Field", "Old", "New", "Source", "Batch")
End Sub

Private Function EntityIDByDisplay(ByVal disp As String) As String
    Dim m As Variant, ws As Worksheet
    Set ws = modData.MasterSheet()
    m = Application.Match(Trim$(disp), ws.Columns(modData.MasterCols()("DisplayID")), 0)
    If Not IsError(m) Then EntityIDByDisplay = CStr(ws.Cells(CLng(m), modData.MasterCols()("EntityID")).Value)
End Function

Private Sub CollectHistory(ws As Worksheet, ByVal id As String, ids As Object, rows As Collection)
    Dim d As Variant, i As Long, rid As String
    d = modCore.TableData(ws.ListObjects(1))
    If Not IsArray(d) Then Exit Sub
    For i = 1 To UBound(d, 1)
        rid = CStr(d(i, 5))
        Dim hit As Boolean
        hit = (id <> "" And rid = id)
        If Not hit Then
            If Not ids Is Nothing Then hit = ids.Exists(rid)
        End If
        If hit Then rows.Add Array(d(i, 2), d(i, 3), d(i, 4), d(i, 6), d(i, 7), d(i, 8), d(i, 9), d(i, 10))
    Next i
End Sub

Private Function SortNewestFirst(rows As Collection) As Collection
    ' timestamps are ISO text: insertion sort descending (history lists are small per entity)
    Dim out As New Collection, x As Variant, j As Long, placed As Boolean
    For Each x In rows
        placed = False
        For j = 1 To out.Count
            If CStr(x(0)) > CStr(out(j)(0)) Then out.Add x, Before:=j: placed = True: Exit For
        Next j
        If Not placed Then out.Add x
    Next x
    Set SortNewestFirst = out
End Function

' ================================================================ restore
Public Sub RestoreEntityOrType()
    Dim ver As Long, s As String, wbS As Workbook, snap As Object, k As Variant, rec As Object, hits As New Collection
    Dim values As Object, f As Variant, issues As Collection, blocked As Boolean, nOK As Long, nSkip As Long
    Dim msgs As String, touched As Object, batch As String, id As String, typeCode As String
    modCore.Heal
    ver = SelectedVersion("Restore from which version?")
    If ver = 0 Then Exit Sub
    s = Trim$(InputBox("Restore from version " & ver & ":" & vbLf & vbLf & _
                       "Enter a Display ID (e.g. APP-0002) for one entity, or a type code (e.g. ApplicationComponent) " & _
                       "for every entity of that type.", "Restore entity or type"))
    If s = "" Then Exit Sub
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    Application.ScreenUpdating = False
    Set wbS = OpenSnapshot(ver)
    If wbS Is Nothing Then Application.ScreenUpdating = True: Exit Sub
    Set snap = LoadModel(wbS)
    wbS.Close SaveChanges:=False
    Application.ScreenUpdating = True
    For Each k In snap.Keys
        If Left$(k, 2) = "E|" Then
            Set rec = snap(k)
            If StrComp(CStr(RecVal(rec, "DisplayID")), s, vbTextCompare) = 0 Or _
               StrComp(CStr(RecVal(rec, "TypeCode")), s, vbTextCompare) = 0 Then hits.Add rec
        End If
    Next k
    If hits.Count = 0 Then MsgBox "Nothing in version " & ver & " matches """ & s & """.", vbInformation, "EA Workbench": Exit Sub
    If MsgBox("Restore " & hits.Count & " entit" & IIf(hits.Count = 1, "y", "ies") & " to their values in version " & ver & "?" & _
              vbLf & "A snapshot of the current model is taken first.", vbOKCancel + vbQuestion, "EA Workbench") = vbCancel Then Exit Sub
    If Not AutoSnapshot("Before restore of " & s & " from version " & ver) Then Exit Sub

    modMeta.LoadMeta
    Set touched = modCore.NewDict()
    batch = modCore.NewGUID()
    For Each rec In hits
        id = CStr(RecVal(rec, "EntityID"))
        typeCode = CStr(RecVal(rec, "TypeCode"))
        If modData.FindEntityRow(id) = 0 Then
            nSkip = nSkip + 1
            msgs = msgs & vbLf & RecVal(rec, "DisplayID") & ": no longer in EA_Core"
        ElseIf modData.EntityTypeOf(id) <> typeCode Then
            nSkip = nSkip + 1
            msgs = msgs & vbLf & RecVal(rec, "DisplayID") & ": its type has changed since that version"
        Else
            Set values = modCore.NewDict()
            For Each f In rec.Keys
                Select Case CStr(f)
                    Case "Name", "Description", "Status", "RepositoryArea"
                        values(f) = rec(f)
                    Case Else
                        If modMeta.IsMapped(typeCode, CStr(f)) Then values(f) = rec(f)
                End Select
            Next f
            Set issues = New Collection
            blocked = False
            modData.UpdateEntity id, values, "Restore from version " & ver, batch, issues, blocked
            If blocked Then
                nSkip = nSkip + 1
                msgs = msgs & vbLf & RecVal(rec, "DisplayID") & ": " & Replace(modRules.IssueText(issues, 2), vbLf, "; ")
            Else
                nOK = nOK + 1
                touched(typeCode) = True
            End If
        End If
    Next rec
    If touched.Count > 0 Then modPaint.PaintTypes touched
    modSync.SaveWithStamp ThisWorkbook
    MsgBox nOK & " entit" & IIf(nOK = 1, "y", "ies") & " restored." & _
           IIf(nSkip > 0, vbLf & nSkip & " not restored:" & Left$(msgs, 1200), ""), _
           IIf(nSkip > 0, vbExclamation, vbInformation), "EA Workbench"
End Sub

' Replaces the metamodel, entities and relationships with a snapshot. History is kept and logged.
Public Sub RestoreWholeModel()
    Dim ver As Long, wbS As Workbook, snap As Object, cur As Object, rows As Collection, x As Variant
    Dim batch As String, ws As Worksheet, lo As ListObject
    modCore.Heal
    ver = SelectedVersion("Restore the whole model from which version?")
    If ver = 0 Then Exit Sub
    If MsgBox("Replace the whole model (metamodel, entities and relationships) with version " & ver & "?" & vbLf & vbLf & _
              "A snapshot of the current model is taken first, and every change is logged in history.", _
              vbOKCancel + vbExclamation, "Restore whole model") = vbCancel Then Exit Sub
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    If modEvents.PendingCount() > 0 Then
        MsgBox "Some EA_Entities edits could not be saved to EA_Core. Fix or discard them first.", vbExclamation, "EA Workbench"
        Exit Sub
    End If
    If modSync.OpenReviewCount() > 0 Then
        MsgBox "SYNC_Review has open items. Settle them first.", vbExclamation, "EA Workbench"
        Exit Sub
    End If
    If Not AutoSnapshot("Before whole-model restore from version " & ver) Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set wbS = OpenSnapshot(ver)
    If wbS Is Nothing Then Application.ScreenUpdating = True: Exit Sub
    Set snap = LoadModel(wbS)
    Set cur = LoadModel(ThisWorkbook)
    Set rows = DiffModels(cur, snap)
    modCore.BeginBusy
    batch = modCore.NewGUID()

    ' ---- log every change first (history is never replaced)
    For Each x In rows
        LogRestoreRow x, ver, batch
    Next x

    ' ---- metamodel tables
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "MM_*" And modCore.SheetExists(wbS, ws.Name) Then
            If ws.ListObjects.Count > 0 Then ReplaceTableData ws.ListObjects(1), wbS.Worksheets(ws.Name), 1, 2, False
        End If
    Next ws
    modMeta.LoadMeta True
    modGenerator.RebuildCoreCaches
    modGenerator.EnsureMasterColumns

    ' ---- data
    ReplaceTableData modData.MasterLO(), wbS.Worksheets("DATA_Master"), 1, 3, True
    ReplaceTableData modData.RelLO(), wbS.Worksheets("DATA_Relationships"), 1, 2, False
    Set lo = modData.RelLO()
    If Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        modCore.UnprotectSheet lo.Parent
        For r = lo.DataBodyRange.Row To lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1
            modData.WriteRelFormulas lo.Parent, r
        Next r
    End If
    wbS.Close SaveChanges:=False
    Set wbS = Nothing
    modCore.EndBusy
    Application.ScreenUpdating = True

    ' ---- rebuild everything derived from the model
    modMeta.BumpSchemaVersion
    modGenerator.RegenerateAll
    modSync.SaveWithStamp modCore.EntitiesWB(False)
    modSync.SaveWithStamp ThisWorkbook
    MsgBox "Model restored from version " & ver & ". " & rows.Count & " change(s) logged.", vbInformation, "EA Workbench"
    Exit Sub
Fail:
    Dim msg As String
    msg = Err.Description
    On Error Resume Next
    If Not wbS Is Nothing Then wbS.Close SaveChanges:=False
    modCore.ResetBusy
    modCore.ProtectCoreSheets
    MsgBox "Restore stopped: " & msg & vbLf & "Use the automatic snapshot taken just before to roll back if needed.", _
           vbCritical, "EA Workbench"
End Sub

Private Sub LogRestoreRow(x As Variant, ByVal ver As Long, ByVal batch As String)
    Dim src As String
    src = "Restore from version " & ver
    Select Case x(0)
        Case "Entity"
            modData.WriteHistory CStr(x(5)), "Restore", CStr(x(2)), IIf(x(6) = "", "*", CStr(x(6))), _
                                 IIf(x(1) = "Added", "(not present)", x(7)), IIf(x(1) = "Removed", "(removed)", x(8)), src, batch
        Case "Relationship"
            modData.WriteHistory "", "Restore", CStr(x(2)), IIf(x(6) = "", "*", CStr(x(6))), _
                                 IIf(x(1) = "Added", "(not present)", x(7)), IIf(x(1) = "Removed", "(removed)", x(8)), src, batch, _
                                 modCore.SH_HREL
        Case Else
            modData.WriteHistory "", "Restore", CStr(x(2)), CStr(x(0)) & IIf(x(6) = "", "", "." & x(6)), _
                                 IIf(x(1) = "Added", "(not present)", x(7)), IIf(x(1) = "Removed", "(removed)", x(8)), src, batch, _
                                 modCore.SH_HMETA
    End Select
End Sub

' Replaces a table's rows with a snapshot sheet's rows, matching columns by name
' (byKeyRow: the target's hidden key row 1, as on DATA_Master; otherwise its header text).
Private Sub ReplaceTableData(lo As ListObject, src As Worksheet, ByVal srcKeyRow As Long, ByVal srcFirstRow As Long, _
                             ByVal byKeyRow As Boolean)
    Dim ws As Worksheet, sv As Variant, lastR As Long, lastC As Long, srcCol As Object, j As Long, i As Long
    Dim n As Long, nCols As Long, out() As Variant, key As String, firstBody As Long, rowsUsed As Long
    Set ws = lo.Parent
    lastR = modCore.LastRowOf(src): lastC = modCore.LastColOf(src)
    Set srcCol = modCore.NewDict()
    If lastR >= srcKeyRow And lastC > 0 Then
        sv = src.Range(src.Cells(1, 1), src.Cells(lastR, lastC)).Value
        If Not IsArray(sv) Then
            Dim one(1 To 1, 1 To 1) As Variant
            one(1, 1) = sv
            sv = one
        End If
        For j = 1 To lastC
            key = CStr(modCore.Nz(sv(srcKeyRow, j), ""))
            If key <> "" And Not srcCol.Exists(key) Then srcCol(key) = j
        Next j
    End If
    ' rows with a value in the first column
    If lastR >= srcFirstRow Then
        For i = srcFirstRow To lastR
            If Not modCore.IsBlankValue(sv(i, 1)) Then n = n + 1
        Next i
    End If
    nCols = lo.ListColumns.Count
    Dim tKeys() As String, tFmt() As String
    ReDim tKeys(1 To nCols): ReDim tFmt(1 To nCols)
    modCore.UnprotectSheet ws
    modCore.ClearTable lo
    For j = 1 To nCols
        If byKeyRow Then
            tKeys(j) = CStr(ws.Cells(modCore.MASTER_KEY_ROW, lo.Range.Column + j - 1).Value)
        Else
            tKeys(j) = CStr(lo.HeaderRowRange.Cells(1, j).Value)
        End If
        tFmt(j) = CStr(lo.HeaderRowRange.Cells(2, j).NumberFormat)
    Next j
    If n = 0 Then Exit Sub
    ReDim out(1 To n, 1 To nCols)
    rowsUsed = 0
    For i = srcFirstRow To lastR
        If Not modCore.IsBlankValue(sv(i, 1)) Then
            rowsUsed = rowsUsed + 1
            For j = 1 To nCols
                key = tKeys(j)
                If srcCol.Exists(key) Then
                    If VarType(sv(i, srcCol(key))) = vbString Then
                        out(rowsUsed, j) = modCore.SafeFor(sv(i, srcCol(key)), tFmt(j))
                    Else
                        out(rowsUsed, j) = sv(i, srcCol(key))
                    End If
                End If
            Next j
        End If
    Next i
    firstBody = lo.HeaderRowRange.Row + 1
    ws.Cells(firstBody, lo.Range.Column).Resize(n, nCols).Value = out
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1, 1), ws.Cells(firstBody + n - 1, lo.Range.Column + nCols - 1))
End Sub

' ================================================================ history archiving
Public Sub ArchiveHistoryIfLarge()
    Dim ws As Worksheet, lo As ListObject, n As Long, m As Long, wbA As Workbook, p As String, aws As Worksheet
    Dim v As Variant, startRow As Long
    On Error GoTo Fail
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "H_*" And ws.ListObjects.Count > 0 Then
            Set lo = ws.ListObjects(1)
            n = 0
            If Not lo.DataBodyRange Is Nothing Then n = lo.ListRows.Count
            If n > ARCHIVE_AT Then
                m = n - ARCHIVE_KEEP
                If wbA Is Nothing Then
                    p = HistoryFolder() & "Archive_" & Format$(Now, "yyyy") & ".xlsx"
                    If Dir(p) <> "" Then
                        Set wbA = Workbooks.Open(p, UpdateLinks:=0)
                    Else
                        Set wbA = Workbooks.Add(xlWBATWorksheet)
                        wbA.Worksheets(1).Name = "About"
                        wbA.Worksheets(1).Range("A1").Value = "Archived history rows from " & ThisWorkbook.Name
                        wbA.SaveAs Filename:=p, FileFormat:=51
                    End If
                End If
                If modCore.SheetExists(wbA, ws.Name) Then
                    Set aws = wbA.Worksheets(ws.Name)
                Else
                    Set aws = wbA.Worksheets.Add(After:=wbA.Worksheets(wbA.Worksheets.Count))
                    aws.Name = ws.Name
                    aws.Range("A1").Resize(1, lo.ListColumns.Count).Value = lo.HeaderRowRange.Value
                End If
                startRow = modCore.LastRowOf(aws) + 1
                modCore.UnprotectSheet ws
                If lo.ShowAutoFilter Then
                    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
                End If
                v = lo.DataBodyRange.Resize(m).Value
                Dim ai As Long, aj As Long
                For ai = 1 To UBound(v, 1)
                    For aj = 1 To UBound(v, 2)
                        If VarType(v(ai, aj)) = vbString Then v(ai, aj) = "'" & v(ai, aj)
                    Next aj
                Next ai
                aws.Cells(startRow, 1).Resize(m, lo.ListColumns.Count).Value = v
                lo.DataBodyRange.Resize(m).Rows.Delete
                modCore.ProtectSheet ws
                modData.WriteHistory "", "Archive", "", "", "", m & " older rows moved to " & Mid$(p, InStrRev(p, Application.PathSeparator) + 1), _
                                     "History housekeeping", modCore.NewGUID(), ws.Name
            End If
        End If
    Next ws
    If Not wbA Is Nothing Then
        wbA.Save
        wbA.Close SaveChanges:=False
    End If
    Exit Sub
Fail:
    On Error Resume Next
    If Not wbA Is Nothing Then wbA.Close SaveChanges:=False
End Sub
