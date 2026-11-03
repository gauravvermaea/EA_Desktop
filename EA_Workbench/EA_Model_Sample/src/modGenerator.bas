Attribute VB_Name = "modGenerator"
' =====================================================================
' Generator. Keeps everything derived from the metamodel in step:
'   - list and mapping caches (named ranges lst_*, MapKeys)
'   - DATA_Master columns, headers, number formats, validation and the
'     not-applicable shading governed by MM_EntityAttributes
'   - EA_Entities caches and type sheets (via modPaint)
' =====================================================================
Option Explicit

Public gRegenFailed As Boolean     ' set by RegenerateAll; read by callers that must not go on after a failure

Public Sub RegenerateAll()
    modCore.Heal
    Dim wbE As Workbook
    gRegenFailed = False
    On Error GoTo Fail
    modCore.BeginBusy
    modCore.Status "loading metamodel"
    modMeta.LoadMeta True
    modCore.Status "rebuilding caches"
    RebuildCoreCaches
    modCore.Status "updating DATA_Master columns"
    EnsureMasterColumns
    ApplyMasterFormatting
    FillBlankHashes
    Set wbE = modCore.EntitiesWB(True)
    If Not wbE Is Nothing Then
        modCore.Status "refreshing EA_Entities caches"
        RefreshEntitiesCaches wbE
        modPaint.PaintAll wbE
        UpdateEntitiesHome wbE
    End If
    modCore.ProtectCoreSheets
    modCore.EndBusy
    ThisWorkbook.Activate
    If Not wbE Is Nothing Then
        modCore.Banner ""
    Else
        gRegenFailed = True
        modCore.Banner "EA_Entities not found: type sheets were not regenerated."
    End If
    Exit Sub
Fail:
    gRegenFailed = True
    modCore.ResetBusy
    MsgBox "Regenerate stopped: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

' Menu entry: called after metamodel edits
Public Sub ApplyMetamodelChanges()
    If MsgBox("Apply metamodel changes now? DATA_Master columns, validation and every EA_Entities type sheet " & _
              "will be regenerated.", vbOKCancel + vbQuestion, "EA Workbench") = vbCancel Then Exit Sub
    modMeta.BumpSchemaVersion
    RegenerateAll
    MsgBox "Metamodel applied (schema version " & modMeta.SchemaVersion() & ").", vbInformation, "EA Workbench"
End Sub

' ---------------------------------------------------------------- Home buttons (both workbooks)
Public Sub SetupHomeButtons()
    Dim ws As Worksheet, wbE As Workbook, core As String, i As Long, b As Variant
    core = "'" & ThisWorkbook.Name & "'!"
    Set ws = ThisWorkbook.Worksheets(modCore.SH_HOME)
    modCore.UnprotectSheet ws
    modCore.DeleteButtons ws
    ws.Range("E6").ClearContents
    i = 0
    For Each b In Array( _
        Array("Add / edit entity", "modForms.ShowEntityFormBlank"), _
        Array("Add relationship", "modForms.ShowRelFormBlank"), _
        Array("Add / edit type", "modForms.ShowTypeFormBlank"), _
        Array("Define rule", "modForms.ShowRuleFormBlank"), _
        Array("Validate All", "modValidate.ValidateAll"), _
        Array("Sync Now", "modSync.SyncNow"), _
        Array("Review differences", "modSync.ShowReview"), _
        Array("Version history", "modHistory.ShowVersions"), _
        Array("Templates, import, export", "modIO.ShowIO"), _
        Array("Refresh reports", "modReports.RefreshReports"), _
        Array("Entity detail", "modReports.EntityDetail"), _
        Array("Impact", "modReports.Impact"), _
        Array("Diagrams (draw.io)", "modDiagram.ShowDiagrams"), _
        Array("Parent and child models", "modMerge.ShowMerge"), _
        Array("Apply metamodel changes", "modGenerator.ApplyMetamodelChanges"), _
        Array("Push EA_Entities edits", "modEvents.PushPending"), _
        Array("Open EA_Entities", "modCore.OpenEntitiesWorkbook"))
        modCore.AddButton ws, CStr(b(0)), core & CStr(b(1)), ws.Cells(6 + i * 2, 5), 190
        i = i + 1
    Next b
    ws.Range("B42").Value = "Tip: on any EA_Entities type sheet, right-click a row to edit it, add a relationship, " & _
                            "draw it, or see its history, detail or impact."
    ws.Range("B42").Font.Italic = True
    ws.Range("B42").Font.Color = RGB(89, 89, 89)

    Set wbE = modCore.EntitiesWB(False)
    If wbE Is Nothing Then Exit Sub
    Set ws = wbE.Worksheets("Home")
    modCore.UnprotectSheet ws
    modCore.DeleteButtons ws
    ws.Range("E6").ClearContents
    ws.Columns("G").ColumnWidth = 30
    i = 0
    For Each b In Array( _
        Array("Add entity (form)", "Btn_AddEntity"), _
        Array("Add relationship", "Btn_AddRelationship"), _
        Array("Save edits to EA_Core", "Btn_Push"), _
        Array("Sync Now", "Btn_SyncNow"), _
        Array("Version history", "Btn_Versions"), _
        Array("Entity detail", "Btn_Detail"), _
        Array("Impact", "Btn_Impact"), _
        Array("Import / export", "Btn_IO"), _
        Array("Refresh reports", "Btn_Reports"), _
        Array("Diagrams (draw.io)", "Btn_Diagrams"), _
        Array("Parent and child models", "Btn_Merge"), _
        Array("Validate All", "Btn_ValidateAll"), _
        Array("Refresh type sheets", "Btn_Regenerate"), _
        Array("Open EA_Core", "Btn_OpenCore"))
        modCore.AddButton ws, CStr(b(0)), "'" & wbE.Name & "'!modEntitiesThin." & CStr(b(1)), ws.Cells(6 + i * 2, 7), 190
        i = i + 1
    Next b
    ws.Range("G32:G35").ClearContents
    ws.Range("G36").Value = "Edit directly on the type sheets; edits go to EA_Core when you save."
    ws.Range("G37").Value = "Right-click a row to edit it, add a relationship, draw it, or see its history, detail or impact."
    ws.Range("G36:G37").Font.Italic = True
    ws.Range("G36:G37").Font.Color = RGB(89, 89, 89)
End Sub

' ---------------------------------------------------------------- caches
Public Sub RebuildCoreCaches()
    WriteListCache ThisWorkbook, modCore.SH_LISTCACHE
    WriteMapCache
End Sub

Public Function ListSources() As Object
    Dim d As Object, k As Variant, c As Collection, t As Variant, a As Variant
    Set d = modCore.NewDict()
    For Each k In modMeta.gLists.Keys
        Set d(k) = modMeta.gLists(k)
    Next k
    Set c = New Collection
    For Each t In modMeta.gTypes.Keys
        If modMeta.gTypes(t)("Status") = "Active" Then c.Add CStr(t)
    Next t
    Set d("TypeCodes") = c
    Set c = New Collection
    For Each t In modMeta.gRelTypes.Keys
        If modMeta.gRelTypes(t)("Status") = "Active" Then c.Add CStr(t)
    Next t
    Set d("RelCodes") = c
    Set c = New Collection
    For Each a In modMeta.gAttrs.Keys
        If modMeta.gAttrs(a)("Status") = "Active" And Not modMeta.gAttrs(a)("IsSystem") Then c.Add CStr(a)
    Next a
    Set d("AttrCodes") = c
    Set c = New Collection
    c.Add "Name": c.Add "Description": c.Add "Status": c.Add "RepositoryArea"
    For Each a In modMeta.gAttrs.Keys
        If modMeta.gAttrs(a)("Status") = "Active" And Not modMeta.gAttrs(a)("IsSystem") Then c.Add CStr(a)
    Next a
    Set d("RuleAttrCodes") = c
    Set c = New Collection
    For Each k In modMeta.gLists.Keys: c.Add CStr(k): Next k
    Set d("ListNames") = c
    Set c = New Collection
    c.Add "EA_Core": c.Add "EA_Entities": c.Add "Skip"
    Set d("SyncKeep") = c
    Set c = New Collection
    c.Add "All entities": c.Add "One entity type": c.Add "Around one entity": c.Add "Full model with metamodel"
    Set d("ExportScope") = c
    Set c = New Collection
    c.Add "Excel": c.Add "CSV"
    Set d("ExportFormat") = c
    Set c = New Collection
    c.Add "Around one entity": c.Add "One entity type": c.Add "All entities"
    Set d("DiagramScope") = c
    Set c = New Collection
    c.Add "png": c.Add "svg": c.Add "pdf"
    Set d("ImageFormat") = c
    Set c = New Collection
    c.Add "Around entities": c.Add "One entity type": c.Add "Repository area": c.Add "Whole model"
    Set d("MergeScope") = c
    Set c = New Collection
    c.Add "Parent": c.Add "Child": c.Add "Skip"
    Set d("MergeKeep") = c
    Set c = New Collection
    c.Add "Entity type": c.Add "Attribute type": c.Add "Relationship type"
    Set d("TypeKinds") = c
    Set c = New Collection
    For Each k In Array("Internal Active Structure", "External Active Structure", "Active Structure", _
                        "Internal Behaviour", "External Behaviour", "Behaviour", "Event", "Passive Structure", _
                        "Structure", "Motivation", "Composite")
        c.Add CStr(k)
    Next k
    Set d("Aspects") = c
    Set ListSources = d
End Function

Public Sub WriteListCache(wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet, src As Object, k As Variant, c As Long, r As Long, v As Variant, n As Long
    Set src = ListSources()
    If Not modCore.SheetExists(wb, sheetName) Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    Else
        Set ws = wb.Worksheets(sheetName)
    End If
    modCore.UnprotectSheet ws
    ws.Cells.Clear
    c = 0
    For Each k In src.Keys
        c = c + 1
        ws.Cells(1, c).Value = k
        r = 1
        For Each v In src(k)
            r = r + 1
            ws.Cells(r, c).Value = "'" & v
        Next v
        n = IIf(r < 2, 2, r)
        On Error Resume Next
        wb.Names("lst_" & k).Delete
        On Error GoTo 0
        wb.Names.Add Name:="lst_" & k, _
            RefersTo:="='" & sheetName & "'!" & ws.Range(ws.Cells(2, c), ws.Cells(n, c)).Address
    Next k
    ws.Visible = xlSheetHidden
End Sub

Private Sub WriteMapCache()
    Dim ws As Worksheet, k As Variant, r As Long
    Set ws = ThisWorkbook.Worksheets(modCore.SH_MAPCACHE)
    modCore.UnprotectSheet ws
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "MapKey"
    r = 1
    For Each k In modMeta.gMapInfo.Keys
        r = r + 1
        ws.Cells(r, 1).Value = "'" & k
    Next k
    On Error Resume Next
    ThisWorkbook.Names("MapKeys").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="MapKeys", RefersTo:="='" & modCore.SH_MAPCACHE & "'!$A$2:$A$" & IIf(r < 2, 2, r)
    ws.Visible = xlSheetHidden
End Sub

' ---------------------------------------------------------------- DATA_Master columns
Public Sub EnsureMasterColumns()
    Dim ws As Worksheet, lo As ListObject, cols As Object, a As Variant, pos As Long, col As Long
    Set ws = modData.MasterSheet(): Set lo = modData.MasterLO()
    modCore.UnprotectSheet ws
    Set cols = modData.MasterCols()
    For Each a In modMeta.gAttrs.Keys
        If Not modMeta.gAttrs(a)("IsSystem") Then
            If Not cols.Exists(a) Then
                If modMeta.gAttrs(a)("Status") = "Active" Then
                    ' insert a whole sheet column so the hidden key row shifts with the table
                    col = cols("Created")
                    ws.Columns(col).Insert Shift:=xlToRight
                    pos = col - lo.Range.Column + 1
                    ws.Cells(modCore.MASTER_KEY_ROW, col).Value = a
                    lo.HeaderRowRange.Cells(1, pos).Value = UniqueHeader(lo, modMeta.gAttrs(a)("Name"), pos)
                    Set cols = modData.MasterCols()
                End If
            Else
                col = cols(a)
                lo.HeaderRowRange.Cells(1, col - lo.Range.Column + 1).Value = _
                    UniqueHeader(lo, modMeta.gAttrs(a)("Name"), col - lo.Range.Column + 1)
                ws.Columns(col).Hidden = (modMeta.gAttrs(a)("Status") <> "Active")
            End If
        End If
    Next a
    ws.Rows(modCore.MASTER_KEY_ROW).Hidden = True
    ws.Columns(1).Hidden = True
End Sub

Private Function UniqueHeader(lo As ListObject, ByVal wanted As String, ByVal pos As Long) As String
    Dim i As Long, nm As String, n As Long
    nm = wanted
    Do
        For i = 1 To lo.ListColumns.Count
            If i <> pos Then
                If StrComp(CStr(lo.HeaderRowRange.Cells(1, i).Value), nm, vbTextCompare) = 0 Then Exit For
            End If
        Next i
        If i > lo.ListColumns.Count Then Exit Do
        n = n + 1
        nm = wanted & " (" & n & ")"
    Loop
    UniqueHeader = nm
End Function

Public Sub ApplyMasterFormatting()
    Dim ws As Worksheet, lo As ListObject, cols As Object, a As Variant, rng As Range
    Dim firstAttr As Long, lastAttr As Long, lastRow As Long
    Set ws = modData.MasterSheet(): Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    modCore.UnprotectSheet ws
    lastRow = Application.WorksheetFunction.Max(5000, lo.Range.Row + lo.Range.Rows.Count + 500)
    firstAttr = 100000: lastAttr = 0
    For Each a In cols.Keys
        If Not modMeta.IsSystemAttr(CStr(a)) Then
            If cols(a) < firstAttr Then firstAttr = cols(a)
            If cols(a) > lastAttr Then lastAttr = cols(a)
            Set rng = ws.Range(ws.Cells(modCore.MASTER_FIRST_ROW, cols(a)), ws.Cells(lastRow, cols(a)))
            ApplyValidation rng, CStr(a)
            rng.NumberFormat = NumberFormatFor(CStr(a))
        End If
    Next a
    ' hashes and IDs are text (an 8-hex hash such as 00012345 would otherwise become a number)
    ws.Range(ws.Cells(modCore.MASTER_FIRST_ROW, cols("RowHash")), ws.Cells(lastRow, cols("RowHash"))).NumberFormat = "@"
    ws.Range(ws.Cells(modCore.MASTER_FIRST_ROW, cols("EntityID")), ws.Cells(lastRow, cols("EntityID"))).NumberFormat = "@"
    ApplyListValidation ws.Range(ws.Cells(modCore.MASTER_FIRST_ROW, cols("TypeCode")), ws.Cells(lastRow, cols("TypeCode"))), "TypeCodes"
    ApplyListValidation ws.Range(ws.Cells(modCore.MASTER_FIRST_ROW, cols("Status")), ws.Cells(lastRow, cols("Status"))), "Status"
    ApplyListValidation ws.Range(ws.Cells(modCore.MASTER_FIRST_ROW, cols("RepositoryArea")), ws.Cells(lastRow, cols("RepositoryArea"))), "RepositoryArea"

    ' not-applicable shading, governed by MM_EntityAttributes via MapKeys.
    ' No relative references, so the rule is independent of the active cell.
    If lastAttr >= firstAttr Then
        Set rng = ws.Range(ws.Cells(modCore.MASTER_FIRST_ROW, firstAttr), ws.Cells(lastRow, lastAttr))
        rng.FormatConditions.Delete
        Dim tcol As String, f As String
        tcol = Split(ws.Cells(1, cols("TypeCode")).Address(True, True), "$")(1)
        f = "=AND(INDEX($" & tcol & ":$" & tcol & ",ROW())<>"""",ISNA(MATCH(INDEX($" & tcol & ":$" & tcol & _
            ",ROW())&""|""&INDEX($1:$1,COLUMN()),MapKeys,0)))"
        ' FormatConditions reads the user's local formula syntax (e.g. ";" separators):
        ' translate the English formula through a scratch cell first
        With ThisWorkbook.Worksheets(modCore.SH_MAPCACHE).Cells(2, 3)
            .Formula = f
            f = .FormulaLocal
            .ClearContents
        End With
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=f)
            .Interior.Color = RGB(231, 230, 230)
            .StopIfTrue = True
        End With
    End If
End Sub

Public Sub FillBlankHashes()
    Dim ws As Worksheet, lo As ListObject, cols As Object, r As Long, last As Long
    Set ws = modData.MasterSheet(): Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    last = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1
    For r = lo.DataBodyRange.Row To last
        ' only blank hashes: recomputing all would hide rows edited outside the Workbench
        If CStr(ws.Cells(r, cols("EntityID")).Value) <> "" And CStr(ws.Cells(r, cols("RowHash")).Value) = "" Then
            ws.Cells(r, cols("RowHash")).Value = modData.MasterRowHash(r, cols)
        End If
    Next r
End Sub

' ---------------------------------------------------------------- validation helpers (shared with modPaint / forms)
Public Function NumberFormatFor(ByVal attrCode As String) As String
    NumberFormatFor = "General"
    If Not modMeta.gAttrs.Exists(attrCode) Then Exit Function
    Select Case modMeta.gAttrs(attrCode)("DataType")
        Case "Date": NumberFormatFor = "yyyy-mm-dd"
        Case "Currency": NumberFormatFor = "#,##0"
        Case "Integer": NumberFormatFor = "0"
        Case "Decimal": NumberFormatFor = "0.00"
        Case "Percent": NumberFormatFor = "0%"
        Case "Text", "LongText", "URL", "Email", "EntityRef", "List": NumberFormatFor = "@"
    End Select
End Function

Public Sub ApplyValidation(rng As Range, ByVal attrCode As String)
    rng.Validation.Delete
    If Not modMeta.gAttrs.Exists(attrCode) Then Exit Sub
    Dim a As Object
    Set a = modMeta.gAttrs(attrCode)
    On Error Resume Next
    Select Case a("DataType")
        Case "List"
            If a("ListName") <> "" Then ApplyListValidation rng, a("ListName")
        Case "Boolean"
            ApplyListValidation rng, "YesNo"
        Case "Date"
            ' date serials, so the rule works under any regional date format
            rng.Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                Formula1:=CStr(CLng(DateSerial(1900, 1, 1))), Formula2:=CStr(CLng(DateSerial(2200, 12, 31)))
            rng.Validation.ErrorMessage = "Enter a date (yyyy-mm-dd)."
        Case "Integer"
            rng.Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                Formula1:="-1000000000", Formula2:="1000000000"
            rng.Validation.ErrorMessage = "Enter a whole number."
        Case "Decimal", "Currency", "Percent"
            rng.Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                Formula1:="-1E+15", Formula2:="1E+15"
            rng.Validation.ErrorMessage = "Enter a number."
    End Select
    If a("Desc") <> "" Then
        Select Case a("DataType")
            Case "Text", "LongText", "URL", "Email", "EntityRef"
                rng.Validation.Add Type:=xlValidateInputOnly
        End Select
        rng.Validation.InputTitle = Left$(a("Name"), 32)
        rng.Validation.InputMessage = Left$(a("Desc"), 250)
        rng.Validation.ShowInput = True
    End If
End Sub

Public Sub ApplyListValidation(rng As Range, ByVal listName As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:="=lst_" & listName
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    rng.Validation.ErrorMessage = "Pick a value from the " & listName & " list."
End Sub

' ---------------------------------------------------------------- EA_Entities caches
Public Sub RefreshEntitiesCaches(wbE As Workbook)
    WriteListCache wbE, "SYS_Lists"
    WriteRuleCache wbE
    SetSyncState wbE, "SchemaVersion", modMeta.SchemaVersion()
    SetSyncState wbE, "CoreFile", ThisWorkbook.Name
    SetSyncState wbE, "ModelID", modCore.Setting("ModelID")
End Sub

Private Sub WriteRuleCache(wbE As Workbook)
    Dim ws As Worksheet, r As Long, t As Variant, a As Variant, x As Variant
    If Not modCore.SheetExists(wbE, "SYS_RuleCache") Then
        Set ws = wbE.Worksheets.Add(After:=wbE.Worksheets(wbE.Worksheets.Count))
        ws.Name = "SYS_RuleCache"
    Else
        Set ws = wbE.Worksheets("SYS_RuleCache")
    End If
    modCore.UnprotectSheet ws
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    ws.Cells.Clear
    ws.Range("A1:H1").Value = Array("TypeCode", "AttrCode", "Name", "DataType", "ListName", "Mandatory", _
                                    "DisplayOrder", "Section")
    r = 1
    For Each t In modMeta.gTypeAttrs.Keys
        For Each a In modMeta.gTypeAttrs(t)
            r = r + 1
            ws.Cells(r, 1).Resize(1, 8).Value = Array(t, a, modMeta.gAttrs(a)("Name"), modMeta.gAttrs(a)("DataType"), _
                modMeta.gAttrs(a)("ListName"), IIf(modMeta.IsMandatory(CStr(t), CStr(a)), "Y", "N"), _
                modMeta.gMapInfo(t & "|" & a)(1), modMeta.gMapInfo(t & "|" & a)(2))
        Next a
    Next t
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(IIf(r < 2, 2, r), 8), , xlYes).Name = "tblCacheAttrs"
    Dim start As Long
    start = r + 3
    ws.Cells(start, 1).Resize(1, 7).Value = Array("AttrCode", "TypeCode", "RuleType", "Param1", "Param2", _
                                                  "Severity", "Message")
    r = start
    For Each x In modMeta.gRules
        r = r + 1
        ws.Cells(r, 1).Resize(1, 7).Value = Array(x(0), x(1), x(2), "'" & x(3), "'" & x(4), x(5), x(6))
    Next x
    ws.ListObjects.Add(xlSrcRange, ws.Cells(start, 1).Resize(IIf(r = start, 2, r - start + 1), 7), , xlYes).Name = "tblCacheRules"
    ws.Visible = xlSheetHidden
End Sub

Public Function GetSyncState(wbE As Workbook, ByVal key As String) As Variant
    Dim lo As ListObject, r As Variant
    On Error GoTo Missing
    Set lo = wbE.Worksheets("SYS_SyncState").ListObjects("tblSyncState")
    r = Application.Match(key, lo.ListColumns("Key").DataBodyRange, 0)
    If IsError(r) Then GetSyncState = "": Exit Function
    GetSyncState = lo.ListColumns("Value").DataBodyRange.Cells(r, 1).Value
    If IsEmpty(GetSyncState) Then GetSyncState = ""
    Exit Function
Missing:
    GetSyncState = ""
End Function

Public Sub SetSyncState(wbE As Workbook, ByVal key As String, ByVal value As Variant)
    Dim lo As ListObject, r As Variant
    Set lo = wbE.Worksheets("SYS_SyncState").ListObjects("tblSyncState")
    r = Application.Match(key, lo.ListColumns("Key").DataBodyRange, 0)
    If IsError(r) Then
        modCore.AppendRow lo, Array(key, value)
    Else
        lo.ListColumns("Value").DataBodyRange.Cells(r, 1).Value = value
    End If
End Sub

Public Sub UpdateEntitiesHome(wbE As Workbook)
    Dim ws As Worksheet, i As Long, t As Variant, order As Variant, d As Variant, types As Collection
    Set ws = wbE.Worksheets("Home")
    modCore.UnprotectSheet ws
    ws.Range("B6:C60").ClearContents
    ws.Range("E8:E80").Hyperlinks.Delete
    ws.Range("E8:E80").Clear
    ws.Range("E7").Value = "Type sheets (click to open):"
    ws.Range("E7").Font.Bold = True
    Set types = modPaint.TypesInSheetOrder()
    i = 0
    For Each t In types
        With ws.Cells(8 + i, 5)
            .Value = modMeta.gTypes(t)("Name") & "  (" & modMeta.gTypes(t)("Domain") & ")"
            ws.Hyperlinks.Add Anchor:=ws.Cells(8 + i, 5), Address:="", _
                SubAddress:="'" & modMeta.gTypes(t)("Sheet") & "'!E3", TextToDisplay:=CStr(.Value)
        End With
        i = i + 1
    Next t
    i = 0
    For Each t In types
        If modPaint.EntityCountOf(CStr(t)) > 0 Then
            ws.Cells(6 + i, 2).Value = modMeta.gTypes(t)("Name")
            ws.Cells(6 + i, 3).Formula = "=COUNTIF('" & modMeta.gTypes(t)("Sheet") & "'!B:B,""E"")"
            ws.Cells(6 + i, 3).Font.Bold = True
            i = i + 1
            If i > 50 Then Exit For
        End If
    Next t
End Sub
