Attribute VB_Name = "modSync"
' =====================================================================
' Sync between EA_Entities and EA_Core (Phase 2).
'
'   Quick sync  - runs when either workbook opens. Pushes queued edits,
'                 repaints rows changed in EA_Core since the last sync,
'                 fixes structure and row order. Rows are only hashed
'                 when a workbook was saved without the Workbench macros.
'   Full sync   - Sync Now. Hashes every row in both workbooks, checks
'                 relationships (Validate All) and archives large history.
'
' EA_Core wins every true conflict by default. Differences that need a
' decision go to SYNC_Review; sheets with open items there are not
' repainted until the items are settled. Mirrors model.decide().
' =====================================================================
Option Explicit

Public Const SH_REVIEW As String = "SYNC_Review"
Public Const KEEP_CORE As String = "EA_Core"
Public Const KEEP_SHEET As String = "EA_Entities"
Public Const KEEP_SKIP As String = "Skip"

Private Const SIT_SHEET As String = "Edited in EA_Entities outside the Workbench"
Private Const SIT_BOTH As String = "Changed in both EA_Core and EA_Entities"
Private Const SIT_NEW As String = "New row in EA_Entities"
Private Const SIT_UNKNOWN As String = "Row in EA_Entities not found in EA_Core"

Private mSessionSynced As Boolean
Public gCloseAt As Date              ' scheduled CloseCoreIfAlone (cancelled if EA_Core closes first)
Private mCoreOpenedByEntities As Boolean

' ================================================================ entry points
' Button: full sync from either workbook
Public Sub SyncNow()
    modCore.Heal
    RunSync True, True, False
End Sub

' Scheduled by StartupCore (Application.OnTime) after EA_Core opens
Public Sub SyncOnOpenFromCore()
    If mSessionSynced Then Exit Sub
    RunSync False, False, True
End Sub

' Called by EA_Entities' Workbook_Open (through modEntitiesThin.ConnectCore)
Public Sub EntitiesOpened(Optional ByVal openedCore As Boolean = False)
    Dim wbE As Workbook, ws As Worksheet
    modCore.Heal
    mCoreOpenedByEntities = openedCore
    Set wbE = modCore.EntitiesWB(False)
    If wbE Is Nothing Then Exit Sub
    For Each ws In wbE.Worksheets
        If ws.Name Like "T_*" Then modCore.ProtectSheet ws
    Next ws
    RunSync False, False, False
End Sub

' Called by EA_Entities' Workbook_BeforeClose
Public Sub EntitiesClosing()
    If mCoreOpenedByEntities Then
        gCloseAt = Now + TimeSerial(0, 0, 2)
        Application.OnTime gCloseAt, "'" & ThisWorkbook.Name & "'!modSync.CloseCoreIfAlone"
    End If
End Sub

' Closes EA_Core once EA_Entities has really closed, if EA_Entities opened it
Public Sub CloseCoreIfAlone()
    If Not modCore.EntitiesWB(False) Is Nothing Then Exit Sub
    mCoreOpenedByEntities = False
    gCloseAt = 0
    SaveWithStamp ThisWorkbook
    ThisWorkbook.Close SaveChanges:=False
End Sub

' EA_Core Workbook_BeforeSave: remember that the Workbench saved it
Public Sub CoreBeforeSave()
    If modCore.IsBusy() Then Exit Sub
    On Error Resume Next
    modCore.SetSetting "WBSaveStamp", CDbl(Now)
End Sub

' ================================================================ the sync run
Public Sub RunSync(ByVal full As Boolean, ByVal interactive As Boolean, ByVal closeEntitiesAfter As Boolean)
    Dim wbE As Workbook, wasOpen As Boolean, coreOutside As Boolean, entOutside As Boolean
    Dim schemaChanged As Boolean, nRehash As Long, items As Collection, toPaint As Object
    Dim nRepaint As Long, summary As String, valSummary As String, lastSync As String
    On Error GoTo Fail
    modMeta.LoadMeta True

    ' ---- 1. companion
    Set wbE = modCore.EntitiesWB(False)
    wasOpen = Not wbE Is Nothing
    If Not wasOpen Then Set wbE = modCore.EntitiesWB(True)
    If wbE Is Nothing Then
        BannerBoth Nothing, "EA_Entities was not found next to EA_Core - sync skipped."
        If interactive Then MsgBox "EA_Entities was not found next to EA_Core.", vbExclamation, "EA Workbench"
        Exit Sub
    End If
    If CStr(modGenerator.GetSyncState(wbE, "ModelID")) <> CStr(modCore.Setting("ModelID")) Then
        BannerBoth wbE, "EA_Entities belongs to a different model - sync stopped."
        If interactive Then MsgBox "EA_Entities in this folder belongs to a different model (model IDs differ). " & _
                                   "Sync stopped.", vbExclamation, "EA Workbench"
        If closeEntitiesAfter And Not wasOpen Then CloseQuietly wbE
        Exit Sub
    End If
    If ThisWorkbook.ReadOnly Or wbE.ReadOnly Then
        BannerBoth wbE, "A workbook is open read-only - sync skipped."
        If closeEntitiesAfter And Not wasOpen Then CloseQuietly wbE
        Exit Sub
    End If
    ' code-friendly protection (UserInterfaceOnly is not saved with the file)
    Dim tws As Worksheet
    For Each tws In wbE.Worksheets
        If tws.Name Like "T_*" Then modCore.ProtectSheet tws
    Next tws

    ' ---- 2. saved outside the Workbench? (before anything here saves either file)
    coreOutside = SavedOutside(ThisWorkbook, modCore.Setting("WBSaveStamp"))
    entOutside = SavedOutside(wbE, modGenerator.GetSyncState(wbE, "WBSaveStamp"))
    lastSync = CStr(modGenerator.GetSyncState(wbE, "LastSync"))
    If lastSync = "" Then full = True

    ' ---- 3. queued edits, without repainting: the comparison below decides what to repaint
    modCore.Status "sync: saving queued edits"
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True, True

    modCore.BeginBusy
    ' ---- 4. structure
    schemaChanged = (Val(CStr(modGenerator.GetSyncState(wbE, "SchemaVersion"))) <> modMeta.SchemaVersion())
    If schemaChanged Then
        modCore.Status "sync: metamodel changed - refreshing caches"
        modGenerator.RebuildCoreCaches
        modGenerator.EnsureMasterColumns
        modGenerator.ApplyMasterFormatting
        modGenerator.RefreshEntitiesCaches wbE
    End If

    ' ---- 5. master rows edited without the Workbench
    If full Or coreOutside Then
        modCore.Status "sync: checking EA_Core rows"
        nRehash = RehashMaster(wbE)
    End If

    ' ---- 6. compare every type sheet
    modCore.Status "sync: comparing type sheets"
    modPaint.BuildContext
    Set items = New Collection
    Set toPaint = modCore.NewDict()
    CompareSheets wbE, (full Or entOutside), lastSync, items, toPaint

    ' ---- 7. review sheet (keeps earlier Keep choices)
    WriteReview items
    modCore.EndBusy

    ' ---- 8. repaint what changed (blocked sheets are skipped by the painter)
    If toPaint.Count > 0 Then
        modCore.Status "sync: repainting " & toPaint.Count & " type sheet(s)"
        nRepaint = toPaint.Count
        modPaint.PaintTypes toPaint
        modCore.BeginBusy
        modPaint.OrderSheets wbE
        modCore.EndBusy
    End If
    If schemaChanged Then
        modCore.BeginBusy
        modGenerator.UpdateEntitiesHome wbE
        modGenerator.SetSyncState wbE, "SchemaVersion", modMeta.SchemaVersion()
        modCore.EndBusy
    End If

    ' ---- 9. full only: relationships and data quality, history housekeeping
    If full Then
        modCore.Status "sync: validating"
        valSummary = modValidate.RunValidation(True)
        modHistory.ArchiveHistoryIfLarge
        modCore.Status "sync: draw.io libraries and views"
        modDiagram.AfterFullSync
        modCore.SetSetting "LastFullSync", modCore.UtcNow()
    End If

    ' ---- 10. stamp and save
    modCore.BeginBusy
    modGenerator.SetSyncState wbE, "LastSync", modCore.UtcNow()
    modCore.SetSetting "LastSync", modCore.UtcNow()
    modCore.EndBusy
    JsonSet "lastSync", """" & modCore.UtcNow() & """"
    JsonSet "schemaVersion", CStr(modMeta.SchemaVersion())
    SaveWithStamp wbE
    SaveWithStamp ThisWorkbook
    mSessionSynced = True

    summary = OpenReviewCount() & " difference(s) need a decision"
    If OpenReviewCount() > 0 Then
        BannerBoth wbE, summary & " - see SYNC_Review in EA_Core."
    ElseIf valSummary <> "" Then
        BannerBoth wbE, valSummary
    Else
        BannerBoth wbE, ""
    End If

    If closeEntitiesAfter And Not wasOpen Then CloseQuietly wbE
    modCore.Status IIf(full, "full", "quick") & " sync done"
    If interactive Then
        MsgBox IIf(full, "Full", "Quick") & " sync finished." & vbLf & vbLf & _
               "Rows checked for edits made outside the Workbench: " & IIf(full Or coreOutside, "EA_Core ", "") & _
               IIf(full Or entOutside, "EA_Entities", "") & vbLf & _
               "EA_Core rows logged as edited outside the Workbench: " & nRehash & vbLf & _
               "Type sheets refreshed: " & nRepaint & vbLf & _
               summary & IIf(valSummary <> "", vbLf & valSummary, ""), _
               IIf(OpenReviewCount() > 0, vbExclamation, vbInformation), "EA Workbench"
        If OpenReviewCount() > 0 Then ShowReview
    End If
    Exit Sub
Fail:
    modCore.ResetBusy
    BannerBoth wbE, "Sync stopped: " & Err.Description
    If interactive Then MsgBox "Sync stopped: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

Private Sub CloseQuietly(wb As Workbook)
    On Error Resume Next
    modCore.BeginBusy
    wb.Close SaveChanges:=False
    modCore.EndBusy
End Sub

' ================================================================ decision rule (mirrors model.decide)
Public Function Decide(ByVal paintHash As String, ByVal sheetHash As String, ByVal masterHash As String) As String
    Dim sheetEdited As Boolean, masterChanged As Boolean
    If paintHash = "" Then
        Decide = IIf(sheetHash = masterHash, "repaint", "both")
        Exit Function
    End If
    sheetEdited = (sheetHash <> paintHash)
    masterChanged = (masterHash <> paintHash)
    If Not sheetEdited And Not masterChanged Then
        Decide = "none"
    ElseIf masterChanged And Not sheetEdited Then
        Decide = "repaint"
    ElseIf sheetEdited And Not masterChanged Then
        Decide = "sheet"
    ElseIf sheetHash = masterHash Then
        Decide = "repaint"
    Else
        Decide = "both"
    End If
End Function

' ================================================================ saved-outside detection
' The Workbench writes a stamp (as a date serial) just before every save it makes or sees.
' If the file's own Last Save Time is clearly later, someone saved it without the macros.
Private Function SavedOutside(wb As Workbook, ByVal stamp As Variant) As Boolean
    Dim lastSave As Date
    If IsBlankStamp(stamp) Then SavedOutside = True: Exit Function
    On Error GoTo Unknown
    lastSave = wb.BuiltinDocumentProperties("Last Save Time").Value
    SavedOutside = (DateDiff("s", CDate(CDbl(stamp)), lastSave) > 90)
    Exit Function
Unknown:
    SavedOutside = True
End Function

Private Function IsBlankStamp(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then IsBlankStamp = True: Exit Function
    If Not IsNumeric(v) Then IsBlankStamp = True: Exit Function
    IsBlankStamp = (CDbl(v) <= 0)
End Function

Public Sub SaveWithStamp(wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    On Error GoTo Done
    modCore.BeginBusy
    If wb Is ThisWorkbook Then
        modCore.SetSetting "WBSaveStamp", CDbl(Now)
    Else
        modGenerator.SetSyncState wb, "WBSaveStamp", CDbl(Now)
    End If
    wb.Save
Done:
    modCore.EndBusy
End Sub

' ================================================================ EA_Core rows edited outside the Workbench
' Recomputes every row hash. Rows whose stored hash differs were edited without the
' macros: they are validated, stamped and logged. Old values come from the painted
' EA_Entities row when that row is untouched; otherwise the change is logged generally.
Private Function RehashMaster(wbE As Workbook) As Long
    Dim ws As Worksheet, lo As ListObject, cols As Object, d As Variant, i As Long, base As Long
    Dim order As Collection, k As Variant, vals As Collection, h As String, r As Long, id As String
    Dim batch As String, n As Long, typeCode As String
    Set ws = modData.MasterSheet(): Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    d = modCore.TableData(lo)
    If Not IsArray(d) Then Exit Function
    base = lo.Range.Column - 1
    Set order = New Collection
    For Each k In Split(modData.HASH_FIELDS_SYS, "|"): order.Add CStr(k): Next k
    For Each k In modData.MasterAttrOrder(): order.Add CStr(k): Next k
    batch = modCore.NewGUID()
    For i = 1 To UBound(d, 1)
        id = CStr(d(i, cols("EntityID") - base))
        If id <> "" Then
            Set vals = New Collection
            For Each k In order
                vals.Add d(i, cols(k) - base)
            Next k
            h = modData.KeyedHash(order, vals)
            If h <> CStr(d(i, cols("RowHash") - base)) Then
                r = lo.DataBodyRange.Row + i - 1
                typeCode = CStr(d(i, cols("TypeCode") - base))
                LogOutsideEdit wbE, ws, cols, r, id, typeCode, batch
                ws.Cells(r, cols("Modified")).Value = modCore.UtcNow()
                ws.Cells(r, cols("ModifiedBy")).Value = modCore.CurrentUser() & " (outside Workbench)"
                ws.Cells(r, cols("RowHash")).Value = h
                n = n + 1
            End If
        End If
    Next i
    RehashMaster = n
End Function

Private Sub LogOutsideEdit(wbE As Workbook, ws As Worksheet, cols As Object, ByVal r As Long, ByVal id As String, _
                           ByVal typeCode As String, ByVal batch As String)
    Dim sh As String, tws As Worksheet, m As Variant, tr As Long, lastCol As Long, c As Long, key As String
    Dim logged As Boolean
    If modMeta.gTypes.Exists(typeCode) Then sh = modMeta.gTypes(typeCode)("Sheet")
    If sh <> "" Then
        If modCore.SheetExists(wbE, sh) Then
            Set tws = wbE.Worksheets(sh)
            m = Application.Match(id, tws.Columns(1), 0)
            If Not IsError(m) Then
                tr = CLng(m)
                ' the painted row still holds the old values only if nobody edited it since
                If modPaint.SheetRowHash(tws, tr) = CStr(tws.Cells(tr, 3).Value) Then
                    lastCol = tws.Cells(modCore.TS_KEY_ROW, tws.Columns.Count).End(xlToLeft).Column
                    For c = 5 To lastCol
                        key = CStr(tws.Cells(modCore.TS_KEY_ROW, c).Value)
                        If cols.Exists(key) Then
                            If modCore.CellText(tws.Cells(tr, c).Value) <> modCore.CellText(ws.Cells(r, cols(key)).Value) Then
                                modData.WriteHistory typeCode, "Sync", id, key, tws.Cells(tr, c).Value, _
                                                     ws.Cells(r, cols(key)).Value, "Edited in EA_Core outside the Workbench", batch
                                logged = True
                            End If
                        End If
                    Next c
                End If
            End If
        End If
    End If
    If Not logged Then
        modData.WriteHistory typeCode, "Sync", id, "*", "(not recorded)", "(edited outside the Workbench)", _
                             "Edited in EA_Core outside the Workbench", batch
    End If
End Sub

' ================================================================ compare type sheets
Private Sub CompareSheets(wbE As Workbook, ByVal hashRows As Boolean, ByVal lastSync As String, _
                          items As Collection, toPaint As Object)
    Dim t As Variant, sh As String, ws As Worksheet, keys As Variant, pend As Object, blocked As Object
    Set pend = modEvents.PendingKeys()
    Set blocked = BlockedSheets()          ' sheets with open review items keep their edits: always hash them
    For Each t In modMeta.gTypes.Keys
        sh = modMeta.gTypes(t)("Sheet")
        If modMeta.gTypes(t)("Status") <> "Active" Then
            If modCore.SheetExists(wbE, sh) Then
                If wbE.Worksheets(sh).Visible = xlSheetVisible Then wbE.Worksheets(sh).Visible = xlSheetHidden
            End If
        ElseIf Not modCore.SheetExists(wbE, sh) Then
            toPaint(CStr(t)) = True
        Else
            Set ws = wbE.Worksheets(sh)
            keys = SheetKeys(ws)
            If Not KeysMatch(keys, CStr(t)) Then toPaint(CStr(t)) = True
            CompareOneSheet ws, CStr(t), keys, hashRows Or blocked.Exists(ws.Name), lastSync, pend, items, toPaint
        End If
    Next t
End Sub

' Key row of a type sheet as a 1-based array
Private Function SheetKeys(ws As Worksheet) As Variant
    Dim lastCol As Long, k() As String, c As Long
    lastCol = ws.Cells(modCore.TS_KEY_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 8 Then lastCol = 8
    ReDim k(1 To lastCol)
    For c = 1 To lastCol
        k(c) = CStr(ws.Cells(modCore.TS_KEY_ROW, c).Value)
    Next c
    SheetKeys = k
End Function

Private Function KeysMatch(keys As Variant, ByVal typeCode As String) As Boolean
    Dim want As Variant, a As Variant, i As Long
    want = Array("EntityID", "RowKind", "PaintHash", "DisplayID", "Name", "Description", "Status", "RepositoryArea")
    If UBound(keys) <> 8 + modMeta.AttrsForType(typeCode).Count Then Exit Function
    For i = 0 To 7
        If keys(i + 1) <> want(i) Then Exit Function
    Next i
    i = 8
    For Each a In modMeta.AttrsForType(typeCode)
        i = i + 1
        If keys(i) <> a Then Exit Function
    Next a
    KeysMatch = True
End Function

Private Sub CompareOneSheet(ws As Worksheet, ByVal typeCode As String, keys As Variant, ByVal hashRows As Boolean, _
                            ByVal lastSync As String, pend As Object, items As Collection, toPaint As Object)
    Dim lastRow As Long, nCols As Long, arr As Variant, i As Long, j As Long, id As String, kind As String
    Dim seqSheet As New Collection, seqWant As Collection, sheetRow As Long, vals As Collection
    Dim decision As String, paintH As String, sheetH As String, masterH As String, modified As String
    nCols = UBound(keys)
    lastRow = modCore.LastRowOf(ws)
    If lastRow >= modCore.TS_FIRST_ROW Then
        arr = ws.Range(ws.Cells(modCore.TS_FIRST_ROW, 1), ws.Cells(lastRow, nCols)).Value
        If Not IsArray(arr) Then
            Dim one(1 To 1, 1 To 1) As Variant
            one(1, 1) = arr
            arr = one
        End If
        For i = 1 To UBound(arr, 1)
            sheetRow = modCore.TS_FIRST_ROW + i - 1
            id = CStr(modCore.Nz(arr(i, 1), ""))
            kind = CStr(modCore.Nz(arr(i, 2), ""))
            If id <> "" Then seqSheet.Add id
            If kind = "R" Then
                ' relationship rows are read-only; order is checked below
            ElseIf id = "" Then
                If RowHasValues(arr, i, nCols) And Not pend.Exists(ws.Name & "|" & sheetRow) Then
                    items.Add NewItem(SIT_NEW, ws.Name, sheetRow, "", "", CStr(modCore.Nz(arr(i, 5), "")), typeCode, _
                                      "", "(new entity)", "", CStr(modCore.Nz(arr(i, 5), "")), KEEP_SHEET)
                End If
            ElseIf pend.Exists("id:" & id) Then
                ' queued edit that failed validation: PushPending already reported it
            ElseIf Not modPaint.CtxHas(id) Then
                items.Add NewItem(SIT_UNKNOWN, ws.Name, sheetRow, id, CStr(modCore.Nz(arr(i, 4), "")), _
                                  CStr(modCore.Nz(arr(i, 5), "")), typeCode, "", "(whole row)", "", "", KEEP_CORE)
            ElseIf CStr(modPaint.CtxValue(id, "TypeCode")) <> typeCode Then
                toPaint(typeCode) = True                       ' type changed in EA_Core: row moves sheet
                toPaint(CStr(modPaint.CtxValue(id, "TypeCode"))) = True
            Else
                paintH = CStr(modCore.Nz(arr(i, 3), ""))
                If hashRows Then
                    Set vals = New Collection
                    For j = 5 To nCols: vals.Add arr(i, j): Next j
                    sheetH = modCore.RowHashOf(vals)
                    masterH = modPaint.CtxProjectedHash(id, keys)
                    decision = Decide(paintH, sheetH, masterH)
                Else
                    ' quick: trust that Workbench-made sheet edits were queued; use EA_Core's stamp
                    modified = CStr(modPaint.CtxValue(id, "Modified"))
                    If paintH = "" Or modified > lastSync Then decision = "repaint" Else decision = "none"
                End If
                Select Case decision
                    Case "repaint"
                        toPaint(typeCode) = True
                    Case "sheet"
                        If AddDiffItems(items, ws, sheetRow, arr, i, keys, id, typeCode, SIT_SHEET, KEEP_SHEET) = 0 Then toPaint(typeCode) = True
                    Case "both"
                        If AddDiffItems(items, ws, sheetRow, arr, i, keys, id, typeCode, SIT_BOTH, KEEP_CORE) = 0 Then toPaint(typeCode) = True
                End Select
            End If
        Next i
    End If
    ' row order and completeness (deleted, sorted or missing rows, new relationships)
    Set seqWant = modPaint.CtxSequence(typeCode)
    If Not SameSequence(seqSheet, seqWant) Then toPaint(typeCode) = True
End Sub

Private Function RowHasValues(arr As Variant, ByVal i As Long, ByVal nCols As Long) As Boolean
    Dim j As Long
    For j = 5 To nCols
        If Not modCore.IsBlankValue(arr(i, j)) Then
            If j <> 7 And j <> 8 Then RowHasValues = True: Exit Function   ' Status / Repository Area alone do not count
        End If
    Next j
End Function

Private Function SameSequence(a As Collection, b As Collection) As Boolean
    Dim i As Long
    If a.Count <> b.Count Then Exit Function
    For i = 1 To a.Count
        If a(i) <> b(i) Then Exit Function
    Next i
    SameSequence = True
End Function

' Adds one review item per differing value; returns how many (0 = values equal, only formatting differs)
Private Function AddDiffItems(items As Collection, ws As Worksheet, ByVal sheetRow As Long, arr As Variant, ByVal i As Long, _
                              keys As Variant, ByVal id As String, ByVal typeCode As String, ByVal situation As String, _
                              ByVal dflt As String) As Long
    Dim j As Long, mv As Variant, n As Long
    For j = 5 To UBound(keys)
        mv = modPaint.CtxValue(id, CStr(keys(j)))
        If modCore.CellText(arr(i, j)) <> modCore.CellText(mv) Then
            items.Add NewItem(situation, ws.Name, sheetRow, id, CStr(modPaint.CtxValue(id, "DisplayID")), _
                              CStr(modPaint.CtxValue(id, "Name")), typeCode, CStr(keys(j)), _
                              modMeta.AttrName(CStr(keys(j))), mv, arr(i, j), dflt)
            n = n + 1
        End If
    Next j
    AddDiffItems = n
End Function

Private Function NewItem(ByVal situation As String, ByVal sheetName As String, ByVal sheetRow As Long, _
                         ByVal id As String, ByVal displayID As String, ByVal nm As String, ByVal typeCode As String, _
                         ByVal attrCode As String, ByVal attrName As String, ByVal coreV As Variant, _
                         ByVal sheetV As Variant, ByVal dflt As String) As Variant
    NewItem = Array(situation, sheetName, sheetRow, id, displayID, nm, typeCode, attrCode, attrName, coreV, sheetV, dflt)
End Function

' ================================================================ SYNC_Review
Private Function ReviewLO() As ListObject
    Set ReviewLO = ThisWorkbook.Worksheets(SH_REVIEW).ListObjects("tblSyncReview")
End Function

Private Function ReviewKey(ByVal situation As String, ByVal sheetName As String, ByVal sheetRow As Variant, _
                           ByVal id As String, ByVal attrCode As String) As String
    If id = "" Then
        ReviewKey = situation & "|" & sheetName & "|" & CStr(sheetRow)
    Else
        ReviewKey = id & "|" & attrCode          ' a choice survives a change of situation
    End If
End Function

Private Sub WriteReview(items As Collection)
    Dim lo As ListObject, d As Variant, i As Long, prev As Object, it As Variant, k As String, keep As String
    Dim ws As Worksheet, rng As Range
    Set lo = ReviewLO(): Set ws = lo.Parent
    Set prev = modCore.NewDict()
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, 14)) = "Open" Or Left$(CStr(d(i, 14)), 5) = "Error" Then
                prev(ReviewKey(CStr(d(i, 2)), CStr(d(i, 3)), d(i, 4), CStr(d(i, 5)), CStr(d(i, 9)))) = CStr(d(i, 13))
            End If
        Next i
    End If
    modCore.UnprotectSheet ws
    modCore.ClearTable lo
    For Each it In items
        k = ReviewKey(CStr(it(0)), CStr(it(1)), it(2), CStr(it(3)), CStr(it(7)))
        keep = CStr(it(11))
        If prev.Exists(k) Then keep = prev(k)
        Set rng = modCore.AppendRow(lo, Array(modCore.UtcNow(), it(0), it(1), it(2), it(3), it(4), it(5), it(6), _
                                              it(7), it(8), it(9), it(10), keep, "Open", ""))
    Next it
    If Not lo.DataBodyRange Is Nothing Then
        lo.ListColumns("Keep").DataBodyRange.Locked = False
        lo.ListColumns("SheetValue").DataBodyRange.Locked = False
    End If
    modCore.ProtectSheet ws
End Sub

Public Function OpenReviewCount() As Long
    Dim d As Variant, i As Long
    d = modCore.TableData(ReviewLO())
    If Not IsArray(d) Then Exit Function
    For i = 1 To UBound(d, 1)
        If CStr(d(i, 14)) = "Open" Or Left$(CStr(d(i, 14)), 5) = "Error" Then OpenReviewCount = OpenReviewCount + 1
    Next i
End Function

' Sheets that must not be repainted: queued edits, or open review items
Public Function BlockedSheets() As Object
    Dim b As Object, d As Variant, i As Long, k As Variant, p As Object
    Set b = modCore.NewDict()
    Set p = modEvents.PendingSheets()
    For Each k In p.Keys: b(k) = True: Next k
    d = modCore.TableData(ReviewLO())
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, 14)) = "Open" Or Left$(CStr(d(i, 14)), 5) = "Error" Then b(CStr(d(i, 3))) = True
        Next i
    End If
    Set BlockedSheets = b
End Function

Public Sub ShowReview()
    Dim ws As Worksheet, core As String
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_REVIEW)
    core = "'" & ThisWorkbook.Name & "'!modSync."
    modCore.UnprotectSheet ws
    modCore.DeleteButtons ws
    modCore.AddButton ws, "Apply decisions", core & "ApplyReview", ws.Range("H1"), 150
    modCore.AddButton ws, "Set all to EA_Core", core & "SetAllToCore", ws.Range("J1"), 150
    modCore.AddButton ws, "Sync Now", core & "SyncNow", ws.Range("L1"), 120
    modCore.AddButton ws, "Back to Home", "'" & ThisWorkbook.Name & "'!modForms.GoHome", ws.Range("N1"), 120
    ws.Range("A3").Value = OpenReviewCount() & " open item(s)."
    ws.Range("A3").Font.Bold = True
    modCore.ProtectSheet ws
    ThisWorkbook.Activate
    ws.Activate
End Sub

Public Sub SetAllToCore()
    Dim lo As ListObject, i As Long, d As Variant
    Set lo = ReviewLO()
    d = modCore.TableData(lo)
    If Not IsArray(d) Then Exit Sub
    modCore.UnprotectSheet lo.Parent
    For i = 1 To UBound(d, 1)
        If CStr(d(i, 14)) = "Open" Or Left$(CStr(d(i, 14)), 5) = "Error" Then
            lo.ListColumns("Keep").DataBodyRange.Cells(i, 1).Value = KEEP_CORE
        End If
    Next i
    modCore.ProtectSheet lo.Parent
End Sub

' Applies the Keep choices. EA_Entities values go through modData (validation, stamps, history).
Public Sub ApplyReview()
    Dim lo As ListObject, d As Variant, i As Long, wbE As Workbook, keep As String, sit As String, st As String
    Dim perEntity As Object, rowsOf As Object, id As String, k As Variant, issues As Collection, blocked As Boolean
    Dim touched As Object, batch As String, nApplied As Long, nErr As Long, newID As String, values As Object
    Dim settled As Object       ' EntityID -> sheet name, rows settled in this run
    Dim coreCells As New Collection  ' Array(sheet, EntityID, AttrCode) to take EA_Core's value now
    modCore.Heal
    On Error GoTo Fail
    Set lo = ReviewLO()
    d = modCore.TableData(lo)
    If Not IsArray(d) Then MsgBox "Nothing to apply.", vbInformation, "EA Workbench": Exit Sub
    Set wbE = modCore.EntitiesWB(True)
    If wbE Is Nothing Then MsgBox "EA_Entities was not found.", vbExclamation, "EA Workbench": Exit Sub
    Set perEntity = modCore.NewDict(): Set rowsOf = modCore.NewDict(): Set touched = modCore.NewDict()
    Set settled = modCore.NewDict()
    batch = modCore.NewGUID()
    modMeta.LoadMeta

    ' first pass: settle EA_Core choices, new rows and unknown rows; collect EA_Entities values per entity
    For i = 1 To UBound(d, 1)
        st = CStr(d(i, 14))
        If st = "Open" Or Left$(st, 5) = "Error" Then
            keep = CStr(d(i, 13)): sit = CStr(d(i, 2)): id = CStr(d(i, 5))
            Select Case keep
                Case KEEP_SKIP
                    ' stays open
                Case KEEP_CORE
                    If sit = SIT_NEW Or sit = SIT_UNKNOWN Then
                        ClearSheetRow wbE, CStr(d(i, 3)), CLng(d(i, 4))
                    ElseIf id <> "" And CStr(d(i, 9)) <> "" Then
                        coreCells.Add Array(CStr(d(i, 3)), id, CStr(d(i, 9)))
                    End If
                    SetStatus lo, i, "Applied: kept EA_Core", ""
                    touched(CStr(d(i, 8))) = True
                    If id <> "" And sit <> SIT_UNKNOWN Then settled(id) = CStr(d(i, 3))
                    nApplied = nApplied + 1
                Case KEEP_SHEET
                    If sit = SIT_NEW Then
                        Set issues = New Collection
                        Set values = SheetRowValues(wbE, CStr(d(i, 3)), CLng(d(i, 4)))
                        newID = modData.CreateEntity(CStr(d(i, 8)), values, "Sync review", batch, issues)
                        If newID = "" Then
                            SetStatus lo, i, "Error", Replace(modRules.IssueText(issues, 3), vbLf, "; ")
                            nErr = nErr + 1
                        Else
                            MarkNewRow wbE, CStr(d(i, 3)), CLng(d(i, 4)), newID
                            settled(newID) = CStr(d(i, 3))
                            SetStatus lo, i, "Applied: added to EA_Core", ""
                            touched(CStr(d(i, 8))) = True
                            nApplied = nApplied + 1
                        End If
                    ElseIf sit = SIT_UNKNOWN Then
                        SetStatus lo, i, "Error", "This row has no entity in EA_Core; choose EA_Core to remove it from the sheet"
                        nErr = nErr + 1
                    Else
                        If Not perEntity.Exists(id) Then
                            perEntity.Add id, modCore.NewDict()
                            rowsOf.Add id, New Collection
                        End If
                        If CStr(d(i, 9)) <> "" Then perEntity(id)(CStr(d(i, 9))) = modData.ToCellValue(CStr(d(i, 9)), d(i, 12))
                        rowsOf(id).Add i
                    End If
                Case Else
                    SetStatus lo, i, "Error", "Keep must be EA_Core, EA_Entities or Skip"
                    nErr = nErr + 1
            End Select
        End If
    Next i

    ' second pass: one update per entity
    Dim x As Variant
    For Each k In perEntity.Keys
        Set issues = New Collection
        blocked = False
        If perEntity(k).Count > 0 Then modData.UpdateEntity CStr(k), perEntity(k), "Sync review", batch, issues, blocked
        For Each x In rowsOf(k)
            If blocked Then
                SetStatus lo, CLng(x), "Error", Replace(modRules.IssueText(issues, 3), vbLf, "; ")
                nErr = nErr + 1
            Else
                SetStatus lo, CLng(x), "Applied: kept EA_Entities", ""
                settled(CStr(k)) = CStr(d(x, 3))
                nApplied = nApplied + 1
            End If
        Next x
        touched(modData.EntityTypeOf(CStr(k))) = True
    Next k

    ' rows with nothing left open are rewritten from EA_Core now, even on sheets that stay blocked
    Dim stillOpen As Object, d2 As Variant
    Set stillOpen = modCore.NewDict()
    d2 = modCore.TableData(lo)
    For i = 1 To UBound(d2, 1)
        If CStr(d2(i, 14)) = "Open" Or Left$(CStr(d2(i, 14)), 5) = "Error" Then stillOpen(CStr(d2(i, 5))) = True
    Next i
    modPaint.BuildContext
    For Each k In settled.Keys
        If Not stillOpen.Exists(k) Then RefreshRow wbE, CStr(settled(k)), CStr(k)
    Next k
    For Each x In coreCells
        If stillOpen.Exists(x(1)) Then RefreshCell wbE, CStr(x(0)), CStr(x(1)), CStr(x(2))
    Next x

    ' repaint sheets that no longer have open items (the painter skips blocked ones)
    If touched.Count > 0 Then modPaint.PaintTypes touched
    SaveWithStamp wbE
    SaveWithStamp ThisWorkbook
    If OpenReviewCount() = 0 Then BannerBoth wbE, "" Else BannerBoth wbE, OpenReviewCount() & " difference(s) still need a decision."
    ShowReview
    MsgBox nApplied & " item(s) applied." & IIf(nErr > 0, vbLf & nErr & " item(s) could not be applied - see Status and Note.", ""), _
           IIf(nErr > 0, vbExclamation, vbInformation), "EA Workbench"
    Exit Sub
Fail:
    modCore.ResetBusy
    MsgBox "Apply stopped: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

' A new sheet row just added to EA_Core: give it its ID so it is not offered again
Private Sub MarkNewRow(wbE As Workbook, ByVal sheetName As String, ByVal r As Long, ByVal newID As String)
    Dim ws As Worksheet
    If Not modCore.SheetExists(wbE, sheetName) Then Exit Sub
    Set ws = wbE.Worksheets(sheetName)
    modCore.UnprotectSheet ws
    ws.Cells(r, 1).Value = newID
    ws.Cells(r, 2).Value = "E"
    modCore.ProtectSheet ws
End Sub

Private Sub ClearSheetRow(wbE As Workbook, ByVal sheetName As String, ByVal r As Long)
    Dim ws As Worksheet, lastCol As Long
    If Not modCore.SheetExists(wbE, sheetName) Then Exit Sub
    Set ws = wbE.Worksheets(sheetName)
    lastCol = ws.Cells(modCore.TS_KEY_ROW, ws.Columns.Count).End(xlToLeft).Column
    modCore.BeginBusy
    modCore.UnprotectSheet ws
    ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).ClearContents
    modCore.ProtectSheet ws
    modCore.EndBusy
End Sub

' One cell back to EA_Core's value, for a row that still has other open items
Private Sub RefreshCell(wbE As Workbook, ByVal sheetName As String, ByVal entityID As String, ByVal attrCode As String)
    Dim ws As Worksheet, m As Variant, keys As Variant, j As Long
    If Not modCore.SheetExists(wbE, sheetName) Then Exit Sub
    If Not modPaint.CtxHas(entityID) Then Exit Sub
    Set ws = wbE.Worksheets(sheetName)
    m = Application.Match(entityID, ws.Columns(1), 0)
    If IsError(m) Then Exit Sub
    keys = SheetKeys(ws)
    For j = 5 To UBound(keys)
        If keys(j) = attrCode Then
            modCore.BeginBusy
            modCore.UnprotectSheet ws
            ws.Cells(CLng(m), j).Value = modCore.SafeFor(modPaint.CtxValue(entityID, attrCode), CStr(ws.Cells(CLng(m), j).NumberFormat))
            ws.Cells(CLng(m), j).Interior.ColorIndex = xlColorIndexNone
            modCore.ProtectSheet ws
            modCore.EndBusy
            Exit For
        End If
    Next j
End Sub

' Rewrites one entity row on its type sheet from EA_Core and resets its PaintHash
Private Sub RefreshRow(wbE As Workbook, ByVal sheetName As String, ByVal entityID As String)
    Dim ws As Worksheet, m As Variant, r As Long, keys As Variant, j As Long
    If Not modCore.SheetExists(wbE, sheetName) Then Exit Sub
    If Not modPaint.CtxHas(entityID) Then Exit Sub
    Set ws = wbE.Worksheets(sheetName)
    m = Application.Match(entityID, ws.Columns(1), 0)
    If IsError(m) Then Exit Sub
    r = CLng(m)
    keys = SheetKeys(ws)
    modCore.BeginBusy
    modCore.UnprotectSheet ws
    For j = 4 To UBound(keys)
        ws.Cells(r, j).Value = modCore.SafeFor(modPaint.CtxValue(entityID, CStr(keys(j))), CStr(ws.Cells(r, j).NumberFormat))
        ws.Cells(r, j).Interior.ColorIndex = xlColorIndexNone
        If Not ws.Cells(r, j).Comment Is Nothing Then ws.Cells(r, j).Comment.Delete
    Next j
    ws.Cells(r, 3).Value = modPaint.CtxProjectedHash(entityID, keys)
    modCore.ProtectSheet ws
    modCore.EndBusy
End Sub

Private Sub SetStatus(lo As ListObject, ByVal i As Long, ByVal st As String, ByVal note As String)
    modCore.UnprotectSheet lo.Parent
    lo.ListColumns("Status").DataBodyRange.Cells(i, 1).Value = st
    lo.ListColumns("Note").DataBodyRange.Cells(i, 1).Value = modCore.SafeValue(note)
    modCore.ProtectSheet lo.Parent
End Sub

' Values of a type-sheet row keyed by AttrCode (Name onward)
Private Function SheetRowValues(wbE As Workbook, ByVal sheetName As String, ByVal r As Long) As Object
    Dim ws As Worksheet, d As Object, c As Long, lastCol As Long, k As String
    Set d = modCore.NewDict()
    Set SheetRowValues = d
    If Not modCore.SheetExists(wbE, sheetName) Then Exit Function
    Set ws = wbE.Worksheets(sheetName)
    lastCol = ws.Cells(modCore.TS_KEY_ROW, ws.Columns.Count).End(xlToLeft).Column
    For c = 5 To lastCol
        k = CStr(ws.Cells(modCore.TS_KEY_ROW, c).Value)
        If k <> "" Then d(k) = modData.ToCellValue(k, ws.Cells(r, c).Value)
    Next c
End Function

' ================================================================ banners and model.json
Public Sub BannerBoth(wbE As Workbook, ByVal text As String)
    modCore.Banner text
    On Error Resume Next
    If wbE Is Nothing Then Set wbE = modCore.EntitiesWB(False)
    If wbE Is Nothing Then Exit Sub
    With wbE.Worksheets("Home")
        modCore.UnprotectSheet wbE.Worksheets("Home")
        .Range("B4").Value = text
        .Range("B4").Font.Bold = True
        .Range("B4").Font.Color = IIf(text = "", vbBlack, RGB(192, 0, 0))
    End With
End Sub

' Sets one top-level value in model.json (jsonValue is already JSON: "text", number or null)
Public Sub JsonSet(ByVal key As String, ByVal jsonValue As String)
    Dim p As String, txt As String, ff As Integer, i As Long, j As Long, inQ As Boolean, ch As String
    On Error GoTo Done
    p = ThisWorkbook.Path & Application.PathSeparator & "model.json"
    If Dir(p) = "" Then Exit Sub
    ff = FreeFile
    Open p For Input As #ff
    txt = Input$(LOF(ff), ff)
    Close #ff
    i = InStr(txt, """" & key & """:")
    If i = 0 Then Exit Sub
    i = i + Len(key) + 3
    j = i
    Do While j <= Len(txt)
        ch = Mid$(txt, j, 1)
        If ch = """" Then inQ = Not inQ
        If Not inQ And (ch = "," Or ch = vbCr Or ch = vbLf Or ch = "}") Then Exit Do
        j = j + 1
    Loop
    txt = Left$(txt, i - 1) & " " & jsonValue & Mid$(txt, j)
    ff = FreeFile
    Open p For Output As #ff
    Print #ff, txt;
    Close #ff
Done:
End Sub
