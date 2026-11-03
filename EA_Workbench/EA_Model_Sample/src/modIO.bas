Attribute VB_Name = "modIO"
' =====================================================================
' Phase 3: data exchange.
'   Templates  - macro-free .xlsx per entity type (+ its allowed
'                relationships), with dropdowns and validation, saved
'                in Templates\. A hidden _meta sheet ties it to the model.
'   Import     - filled templates: model check, rows matched by Entity ID
'                then by name within the type, validated, previewed on
'                IMPORT_Preview, committed after an automatic snapshot,
'                logged under one batch, files moved to Imports\Done\.
'   Export     - all entities, one type, everything around one entity,
'                or the full model with metamodel; Excel or CSV.
' =====================================================================
Option Explicit

Private Const SH_FORM As String = "FORM_IO"
Private Const SH_PREVIEW As String = "IMPORT_Preview"
Private Const TEMPLATE_KIND As String = "EAWB-Template-1"
Private Const TPL_FIRST_ROW As Long = 3
Private Const TPL_MAX_ROWS As Long = 2000

' ================================================================ form
Public Sub ShowIO()
    Dim ws As Worksheet, core As String
    modCore.Heal
    modMeta.LoadMeta
    ThisWorkbook.Activate
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    core = "'" & ThisWorkbook.Name & "'!modIO."
    modCore.BeginBusy
    On Error GoTo Done
    ws.Activate
    modForms.ResetForm ws, "Templates, import and export", _
        "Templates and exports are saved in the model folder (Templates\ and Exports\). Put filled templates in Imports\."
    SectionTitle ws, 4, "1. Generate an import template"
    modForms.Field ws, 5, "Entity type *", "_ttype"
    modForms.ListDropdown ws, "_ttype", "TypeCodes"
    modForms.Field ws, 6, "Pre-fill existing entities (Y/N)", "_prefill", "N"
    modForms.ListDropdown ws, "_prefill", "YesNo"
    ws.Cells(6, 4).Value = "Y = bulk update: a blank cell clears that value on import"
    modForms.Field ws, 7, "Include relationships sheet (Y/N)", "_rels", "Y"
    modForms.ListDropdown ws, "_rels", "YesNo"
    modCore.AddButton ws, "Generate template", core & "GenerateTemplateFromForm", ws.Range("F5"), 170

    SectionTitle ws, 9, "2. Import filled templates"
    ws.Cells(10, 1).Value = "Choose one or more filled templates; nothing is written until you commit the preview."
    ws.Cells(10, 1).Font.Color = RGB(89, 89, 89)
    modCore.AddButton ws, "Choose files and preview", core & "ImportPreview", ws.Range("F9"), 170
    modCore.AddButton ws, "Commit import", core & "ImportCommit", ws.Range("F11"), 170

    SectionTitle ws, 13, "3. Export"
    modForms.Field ws, 14, "Scope *", "_scope", "All entities"
    modForms.ListDropdown ws, "_scope", "ExportScope"
    modForms.Field ws, 15, "Entity type (for one type)", "_etype"
    modForms.ListDropdown ws, "_etype", "TypeCodes"
    modForms.Field ws, 16, "Start entity Display ID (for around one entity)", "_start"
    modForms.Field ws, 17, "Relationship hops (1-5)", "_hops", 2
    modForms.Field ws, 18, "Include retired (Y/N)", "_retired", "N"
    modForms.ListDropdown ws, "_retired", "YesNo"
    modForms.Field ws, 19, "Format *", "_format", "Excel"
    modForms.ListDropdown ws, "_format", "ExportFormat"
    modCore.AddButton ws, "Export", core & "ExportFromForm", ws.Range("F14"), 170
    modCore.AddButton ws, "Back to Home", "'" & ThisWorkbook.Name & "'!modForms.GoHome", ws.Range("F21"), 170
    ws.Range("B5").Select
Done:
    Dim eMsg As String
    If Err.Number <> 0 Then eMsg = Err.Description
    On Error Resume Next
    If eMsg <> "" Then modForms.FormMessage ws, "Form not fully built: " & eMsg, True
    modCore.ProtectSheet ws
    modCore.EndBusy
End Sub

Private Sub SectionTitle(ws As Worksheet, ByVal r As Long, ByVal text As String)
    ws.Cells(r, 1).Value = text
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    ws.Cells(r, 1).Font.Color = RGB(31, 56, 100)
End Sub

Private Function ModelFolder() As String
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then Err.Raise vbObjectError + 401, , _
        "The model folder is on a web location; move it to a local folder."
    ModelFolder = ThisWorkbook.Path & Application.PathSeparator
End Function

Private Function SubFolder(ByVal nm As String) As String
    Dim p As String
    p = ModelFolder() & nm
    If Dir(p, vbDirectory) = "" Then MkDir p
    SubFolder = p & Application.PathSeparator
End Function

Private Function Stamp() As String
    Stamp = Format$(Now, "yyyymmdd-hhnnss")
End Function

' ================================================================ templates
Public Sub GenerateTemplateFromForm()
    Dim ws As Worksheet, t As String, p As String
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    t = CStr(modForms.FieldValue(ws, "_ttype"))
    If Not modMeta.TypeExists(t) Then
        modCore.UnprotectSheet ws
        modForms.FormMessage ws, "Pick an entity type first.", True
        modCore.ProtectSheet ws
        Exit Sub
    End If
    p = GenerateTemplate(t, UCase$(CStr(modForms.FieldValue(ws, "_prefill"))) = "Y", _
                         UCase$(CStr(modForms.FieldValue(ws, "_rels"))) <> "N")
    If p <> "" Then
        modCore.UnprotectSheet ws
        modForms.FormMessage ws, "Template saved: " & p
        modCore.ProtectSheet ws
        MsgBox "Template saved:" & vbLf & p & vbLf & vbLf & _
               "Fill it in (no macros needed), put it in the Imports folder, then use Choose files and preview.", _
               vbInformation, "EA Workbench"
    End If
End Sub

' Builds a template workbook and returns its path ("" if it failed)
Public Function GenerateTemplate(ByVal typeCode As String, ByVal prefill As Boolean, ByVal withRels As Boolean) As String
    Dim wb As Workbook, ws As Worksheet, path As String, keys As Collection, k As Variant, c As Long, r As Long
    Dim lists As Object, rels As Collection, lastRow As Long
    On Error GoTo Fail
    modMeta.LoadMeta
    SubFolder "Imports"           ' ready for the filled template to come back
    path = SubFolder("Templates") & typeCode & "-template-" & Stamp() & IIf(prefill, "-prefilled", "") & ".xlsx"
    modCore.BeginBusy
    Set wb = Workbooks.Add(xlWBATWorksheet)
    ' lists the validation needs, written to the template's own Lists sheet first
    Set lists = modCore.NewDict()
    lists("Status") = True: lists("RepositoryArea") = True: lists("YesNo") = True
    For Each k In modMeta.AttrsForType(typeCode)
        If modMeta.gAttrs(k)("DataType") = "List" And modMeta.gAttrs(k)("ListName") <> "" Then lists(modMeta.gAttrs(k)("ListName")) = True
    Next k
    BuildListsSheet wb, lists

    ' ---- README
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "README"
    ws.Range("A1").Value = modMeta.TypeName_(typeCode) & " import template"
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True
    WriteLines ws, 3, Array( _
        "Model: " & CStr(modCore.Setting("ModelName")), _
        "Entity type: " & modMeta.TypeName_(typeCode) & " (" & typeCode & ")", _
        "", _
        "How to fill it in", _
        "1. Entities sheet: one row per entity. Columns marked * are required. Use the dropdowns where offered.", _
        "2. To update an existing entity, keep its row (the hidden Entity ID links it) or use its exact name.", _
        IIf(prefill, "3. This template is pre-filled: a cell you clear will be cleared in the repository.", _
                     "3. Blank cells are ignored for existing entities; only filled cells change values."), _
        IIf(withRels, "4. Relationships sheet: one row per relationship from an entity of this type. The Allowed sheet lists valid combinations.", ""), _
        "5. Save the file and put it in the model's Imports folder, then in EA_Core use Templates, import and export.", _
        "", _
        "Do not rename sheets, delete the hidden first row or change the _meta sheet.")
    ws.Columns("A").ColumnWidth = 110

    ' ---- Entities
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Entities"
    Set keys = New Collection
    For Each k In Array("EntityID", "DisplayID", "Name", "Description", "Status", "RepositoryArea"): keys.Add CStr(k): Next k
    For Each k In modMeta.AttrsForType(typeCode): keys.Add CStr(k): Next k
    c = 0
    For Each k In keys
        c = c + 1
        ws.Cells(1, c).Value = k
        ws.Cells(2, c).Value = TemplateHeader(typeCode, CStr(k))
        ws.Range(ws.Cells(TPL_FIRST_ROW, c), ws.Cells(TPL_MAX_ROWS, c)).NumberFormat = TemplateFormat(CStr(k))
    Next k
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, keys.Count))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .WrapText = True
    End With
    ws.Rows(1).Hidden = True
    ws.Columns(1).Hidden = True
    ws.Columns(2).Interior.Color = RGB(242, 242, 242)
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 34
    ws.Columns(4).ColumnWidth = 40
    For c = 5 To keys.Count
        ws.Columns(c).ColumnWidth = 18
    Next c
    If prefill Then WritePrefill ws, typeCode, keys
    ' validation (lists point at the template's own Lists sheet)
    lastRow = TPL_MAX_ROWS
    c = 0
    For Each k In keys
        c = c + 1
        Select Case CStr(k)
            Case "Status": TplList ws.Range(ws.Cells(TPL_FIRST_ROW, c), ws.Cells(lastRow, c)), "Status", lists
            Case "RepositoryArea": TplList ws.Range(ws.Cells(TPL_FIRST_ROW, c), ws.Cells(lastRow, c)), "RepositoryArea", lists
            Case "EntityID", "DisplayID", "Name", "Description"
            Case Else: TplValidation ws.Range(ws.Cells(TPL_FIRST_ROW, c), ws.Cells(lastRow, c)), CStr(k), lists
        End Select
    Next k
    ws.Activate
    ws.Range("C3").Select
    ActiveWindow.FreezePanes = True

    ' ---- Relationships and Allowed
    If withRels Then
        Set rels = AllowedFrom(typeCode)
        BuildRelSheets wb, typeCode, rels, prefill, lists
    End If

    ' ---- _meta
    BuildMeta wb, typeCode, prefill
    wb.Worksheets("Lists").Visible = xlSheetHidden
    wb.Worksheets("Entities").Activate
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=51
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    modCore.EndBusy
    GenerateTemplate = path
    Exit Function
Fail:
    Dim msg As String
    msg = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    modCore.ResetBusy
    MsgBox "Template not created: " & msg, vbExclamation, "EA Workbench"
End Function

Private Sub WriteLines(ws As Worksheet, ByVal r As Long, lines As Variant)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        ws.Cells(r + i - LBound(lines), 1).Value = modCore.SafeValue(CStr(lines(i)))
    Next i
End Sub

Private Function TemplateHeader(ByVal typeCode As String, ByVal key As String) As String
    Select Case key
        Case "EntityID": TemplateHeader = "Entity ID"
        Case "DisplayID": TemplateHeader = "Display ID (info)"
        Case "Name": TemplateHeader = "Name *"
        Case "Description": TemplateHeader = "Description"
        Case "Status": TemplateHeader = "Status"
        Case "RepositoryArea": TemplateHeader = "Repository Area"
        Case Else: TemplateHeader = modMeta.AttrName(key) & IIf(modMeta.IsMandatory(typeCode, key), " *", "")
    End Select
End Function

Private Function TemplateFormat(ByVal key As String) As String
    Select Case key
        Case "EntityID", "DisplayID", "Name", "Description", "Status", "RepositoryArea": TemplateFormat = "@"
        Case Else: TemplateFormat = modGenerator.NumberFormatFor(key)
    End Select
End Function

Private Sub WritePrefill(ws As Worksheet, ByVal typeCode As String, keys As Collection)
    Dim lo As ListObject, cols As Object, d As Variant, i As Long, base As Long, r As Long, c As Long, k As Variant
    Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    d = modCore.TableData(lo)
    If Not IsArray(d) Then Exit Sub
    base = lo.Range.Column - 1
    r = TPL_FIRST_ROW - 1
    For i = 1 To UBound(d, 1)
        If CStr(d(i, cols("TypeCode") - base)) = typeCode And CStr(d(i, cols("EntityID") - base)) <> "" Then
            r = r + 1
            If r > TPL_MAX_ROWS Then Exit For
            c = 0
            For Each k In keys
                c = c + 1
                If cols.Exists(k) Then ws.Cells(r, c).Value = modCore.SafeFor(d(i, cols(k) - base), CStr(ws.Cells(r, c).NumberFormat))
            Next k
        End If
    Next i
End Sub

Private Sub TplList(rng As Range, ByVal listName As String, lists As Object)
    lists(listName) = True
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:="=lst_" & listName
    rng.Validation.IgnoreBlank = True
End Sub

Private Sub TplValidation(rng As Range, ByVal attrCode As String, lists As Object)
    Dim a As Object
    If Not modMeta.gAttrs.Exists(attrCode) Then Exit Sub
    Set a = modMeta.gAttrs(attrCode)
    If a("DataType") = "List" And a("ListName") <> "" Then
        TplList rng, a("ListName"), lists
    ElseIf a("DataType") = "Boolean" Then
        TplList rng, "YesNo", lists
    Else
        modGenerator.ApplyValidation rng, attrCode          ' dates, numbers, input messages
    End If
End Sub

' Allowed (RelCode, TargetType) pairs from a source type, expanded over active types
Private Function AllowedFrom(ByVal typeCode As String) As Collection
    Dim c As New Collection, rc As Variant, t As Variant
    For Each rc In modMeta.RelsFromType(typeCode)
        For Each t In modMeta.TargetTypesFor(typeCode, CStr(rc))
            c.Add Array(CStr(rc), CStr(t))
        Next t
    Next rc
    Set AllowedFrom = c
End Function

Private Sub BuildRelSheets(wb As Workbook, ByVal typeCode As String, rels As Collection, ByVal prefill As Boolean, lists As Object)
    Dim ws As Worksheet, i As Long, x As Variant, codes As Object, tl As Object, r As Long
    ' Allowed
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Allowed"
    ws.Range("A1:D1").Value = Array("Relationship", "Verb", "Target type", "Target type name")
    ws.Range("A1:D1").Font.Bold = True
    Set codes = modCore.NewDict(): Set tl = modCore.NewDict()
    i = 1
    For Each x In rels
        i = i + 1
        ws.Cells(i, 1).Value = x(0)
        ws.Cells(i, 2).Value = modMeta.RelVerb(CStr(x(0)))
        ws.Cells(i, 3).Value = x(1)
        ws.Cells(i, 4).Value = modMeta.TypeName_(CStr(x(1)))
        codes(x(0)) = True: tl(x(1)) = True
    Next x
    Dim kk As Variant, n As Long
    ws.Range("F1").Value = "Relationship codes"
    ws.Range("F1").Font.Bold = True
    n = 1
    For Each kk In codes.Keys
        n = n + 1
        ws.Cells(n, 6).Value = kk
    Next kk
    If n = 1 Then ws.Cells(2, 6).Value = "(none)": n = 2
    Dim nt As Long
    ws.Range("G1").Value = "Target types"
    ws.Range("G1").Font.Bold = True
    nt = 1
    For Each kk In tl.Keys
        nt = nt + 1
        ws.Cells(nt, 7).Value = kk
    Next kk
    If nt = 1 Then ws.Cells(2, 7).Value = "(none)": nt = 2
    ws.Columns("A:G").AutoFit
    ' Relationships
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets("Allowed"))
    ws.Name = "Relationships"
    ws.Range("A1:F1").Value = Array("SourceName", "RelCode", "TargetType", "TargetName", "Label", "Description")
    ws.Range("A2:F2").Value = Array("Source name (" & modMeta.TypeName_(typeCode) & ") *", "Relationship *", _
                                    "Target type *", "Target name *", "Label", "Description")
    With ws.Range("A2:F2")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    ws.Rows(1).Hidden = True
    ws.Range("A3:F" & TPL_MAX_ROWS).NumberFormat = "@"
    ws.Columns("A").ColumnWidth = 34: ws.Columns("B").ColumnWidth = 18: ws.Columns("C").ColumnWidth = 24
    ws.Columns("D").ColumnWidth = 34: ws.Columns("E").ColumnWidth = 20: ws.Columns("F").ColumnWidth = 36
    ' dropdowns: source names from the Entities sheet, codes and target types from Allowed
    On Error Resume Next
    With ws.Range("A3:A" & TPL_MAX_ROWS).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=Entities!$C$3:$C$" & TPL_MAX_ROWS
    End With
    With ws.Range("B3:B" & TPL_MAX_ROWS).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Allowed!$F$2:$F$" & n
    End With
    With ws.Range("C3:C" & TPL_MAX_ROWS).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Allowed!$G$2:$G$" & nt
    End With
    On Error GoTo 0
    If prefill Then
        Dim d As Variant, srcType As String
        d = modCore.TableData(modData.RelLO())
        r = TPL_FIRST_ROW - 1
        If IsArray(d) Then
            For i = 1 To UBound(d, 1)
                If CStr(d(i, 11)) <> "Retired" And CStr(d(i, 1)) <> "" Then
                    srcType = modData.EntityTypeOf(CStr(d(i, 3)))
                    If srcType = typeCode Then
                        r = r + 1
                        ws.Cells(r, 1).Value = modData.EntityNameOf(CStr(d(i, 3)))
                        ws.Cells(r, 2).Value = d(i, 2)
                        ws.Cells(r, 3).Value = modData.EntityTypeOf(CStr(d(i, 4)))
                        ws.Cells(r, 4).Value = modData.EntityNameOf(CStr(d(i, 4)))
                        ws.Cells(r, 5).Value = d(i, 9)
                        ws.Cells(r, 6).Value = d(i, 10)
                    End If
                End If
            Next i
        End If
    End If
End Sub


Private Sub BuildListsSheet(wb As Workbook, lists As Object)
    Dim ws As Worksheet, k As Variant, c As Long, r As Long, v As Variant
    lists("YesNo") = True
    Set ws = wb.Worksheets(1)            ' the new workbook's only sheet
    ws.Name = "Lists"
    For Each k In lists.Keys
        If modMeta.gLists.Exists(k) Then
            c = c + 1
            ws.Cells(1, c).Value = k
            ws.Cells(1, c).Font.Bold = True
            r = 1
            For Each v In modMeta.gLists(k)
                r = r + 1
                ws.Cells(r, c).Value = "'" & v
            Next v
            wb.Names.Add Name:="lst_" & k, RefersTo:="=Lists!" & ws.Range(ws.Cells(2, c), ws.Cells(IIf(r < 2, 2, r), c)).Address
        End If
    Next k
    ' hidden later: a workbook's only sheet cannot be hidden
End Sub

Private Sub BuildMeta(wb As Workbook, ByVal typeCode As String, ByVal prefill As Boolean)
    Dim ws As Worksheet, info As Variant, i As Long
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "_meta"
    info = Array("Kind", TEMPLATE_KIND, "ModelID", CStr(modCore.Setting("ModelID")), _
                 "ModelName", CStr(modCore.Setting("ModelName")), "TypeCode", typeCode, _
                 "SchemaVersion", modMeta.SchemaVersion(), "Prefilled", IIf(prefill, "Y", "N"), _
                 "TemplateID", modCore.NewGUID(), "Generated", modCore.UtcNow())
    For i = 0 To UBound(info) Step 2
        ws.Cells(i \ 2 + 1, 1).Value = info(i)
        ws.Cells(i \ 2 + 1, 2).Value = modCore.SafeValue(info(i + 1))
    Next i
    ws.Visible = xlSheetVeryHidden
End Sub

' ================================================================ import
Private Function PreviewLO() As ListObject
    Set PreviewLO = ThisWorkbook.Worksheets(SH_PREVIEW).ListObjects("tblImportPreview")
End Function

Public Sub ImportPreview()
    Dim files As Variant, list As String, i As Long, startDir As String
    modCore.Heal
    On Error Resume Next
    startDir = SubFolder("Imports")
    ChDrive Left$(startDir, 1)
    ChDir startDir
    On Error GoTo 0
    files = Application.GetOpenFilename("Excel workbooks (*.xlsx),*.xlsx", , "Choose filled templates", , True)
    If Not IsArray(files) Then Exit Sub
    For i = LBound(files) To UBound(files)
        list = list & IIf(list = "", "", "|") & files(i)
    Next i
    RunImport list, False
End Sub

Public Sub ImportCommit()
    Dim list As String, ws As Worksheet
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_PREVIEW)
    list = CStr(ws.Range("B4").Value)
    If list = "" Then
        MsgBox "Preview an import first (Choose files and preview).", vbInformation, "EA Workbench"
        Exit Sub
    End If
    If MsgBox("Commit this import? Rows marked Error are skipped. A snapshot is taken first.", _
              vbOKCancel + vbQuestion, "EA Workbench") = vbCancel Then Exit Sub
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    If Not modHistory.AutoSnapshot("Before import") Then Exit Sub
    RunImport list, True
End Sub

' Reads the templates, writes the preview, and (commit = True) applies the valid rows.
Private Sub RunImport(ByVal fileList As String, ByVal commit As Boolean)
    Dim files() As String, f As Variant, rows As New Collection, batch As String, touched As Object
    Dim nNew As Long, nUpd As Long, nRel As Long, nErr As Long, nSame As Long, moved As String, skipMove As Object
    On Error GoTo Fail
    modCore.BeginBusy
    modMeta.LoadMeta True
    files = Split(fileList, "|")
    batch = modCore.NewGUID()
    Set touched = modCore.NewDict()
    Set skipMove = modCore.NewDict()
    For Each f In files
        ImportOneFile CStr(f), commit, batch, rows, touched, skipMove, nNew, nUpd, nRel, nErr, nSame
    Next f
    WritePreview fileList, rows, commit
    If commit Then
        For Each f In files
            If skipMove.Exists(LCase$(CStr(f))) Then
                moved = moved & vbLf & Mid$(CStr(f), InStrRev(CStr(f), Application.PathSeparator) + 1) & " was not imported, so it was left in place."
            Else
                moved = moved & MoveToDone(CStr(f))
            End If
        Next f
        If touched.Count > 0 Then modPaint.PaintTypes touched
        modSync.SaveWithStamp modCore.EntitiesWB(False)
        modSync.SaveWithStamp ThisWorkbook
        modCore.EndBusy
        MsgBox "Import committed." & vbLf & vbLf & nNew & " new entit" & IIf(nNew = 1, "y", "ies") & ", " & _
               nUpd & " updated, " & nSame & " unchanged, " & nRel & " new relationship(s)." & vbLf & _
               nErr & " row(s) skipped with errors - see IMPORT_Preview." & moved, _
               IIf(nErr > 0, vbExclamation, vbInformation), "EA Workbench"
    Else
        modCore.EndBusy
        MsgBox "Preview ready: " & nNew & " new, " & nUpd & " to update, " & nSame & " unchanged, " & _
               nRel & " new relationship(s), " & nErr & " error(s)." & vbLf & _
               "Check IMPORT_Preview, then click Commit import.", _
               IIf(nErr > 0, vbExclamation, vbInformation), "EA Workbench"
    End If
    Exit Sub
Fail:
    Dim msg As String
    msg = Err.Description
    modCore.ResetBusy
    MsgBox "Import stopped: " & msg, vbExclamation, "EA Workbench"
End Sub

Private Sub ImportOneFile(ByVal path As String, ByVal commit As Boolean, ByVal batch As String, rows As Collection, _
                          touched As Object, skipMove As Object, ByRef nNew As Long, ByRef nUpd As Long, ByRef nRel As Long, _
                          ByRef nErr As Long, ByRef nSame As Long)
    Dim wb As Workbook, meta As Object, fname As String, typeCode As String, prefill As Boolean, ob As Workbook
    Dim ws As Worksheet, keys As Variant, lastRow As Long, lastCol As Long, v As Variant, i As Long, j As Long
    Dim nameIdx As Object, newNames As Object, src As String
    fname = Mid$(path, InStrRev(path, Application.PathSeparator) + 1)
    src = "Import " & fname
    For Each ob In Application.Workbooks
        If LCase$(ob.FullName) = LCase$(path) Or LCase$(ob.Name) = LCase$(fname) Then
            rows.Add Array(fname, "", "", "File", "Error", "", "", "", "", "", "", "", "Error", _
                           "File is open in Excel - save and close it, then preview again")
            nErr = nErr + 1
            skipMove(LCase$(path)) = True
            Exit Sub
        End If
    Next ob
    On Error GoTo Fail
    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)
    ThisWorkbook.Activate
    Set meta = ReadMeta(wb)
    If CStr(meta("Kind")) <> TEMPLATE_KIND Then
        rows.Add Array(fname, "", "", "File", "Error", "", "", "", "", "", "", "", "Error", "Not a Workbench import template")
        nErr = nErr + 1
        skipMove(LCase$(path)) = True
        GoTo CloseFile
    End If
    If CStr(meta("ModelID")) <> CStr(modCore.Setting("ModelID")) Then
        rows.Add Array(fname, "", "", "File", "Error", "", "", "", "", "", "", "", "Error", _
                       "Template belongs to another model (" & meta("ModelName") & ")")
        nErr = nErr + 1
        skipMove(LCase$(path)) = True
        GoTo CloseFile
    End If
    typeCode = CStr(meta("TypeCode"))
    If Not modMeta.TypeExists(typeCode) Then
        rows.Add Array(fname, "", "", "File", "Error", "", "", "", typeCode, "", "", "", "Error", "Entity type no longer exists")
        nErr = nErr + 1
        skipMove(LCase$(path)) = True
        GoTo CloseFile
    End If
    If Val(CStr(meta("SchemaVersion"))) <> modMeta.SchemaVersion() Then
        rows.Add Array(fname, "", "", "File", "Info", "", "", "", typeCode, "", "", "", "Warning", _
                       "Template made with metamodel version " & meta("SchemaVersion") & "; columns are matched by attribute code")
    End If
    prefill = (CStr(meta("Prefilled")) = "Y")

    ' ---- entities
    Set nameIdx = modData.BuildNameIndex()
    Set newNames = modCore.NewDict()
    If modCore.SheetExists(wb, "Entities") Then
        Set ws = wb.Worksheets("Entities")
        lastRow = modCore.LastRowOf(ws): lastCol = modCore.LastColOf(ws)
        If lastRow >= TPL_FIRST_ROW And lastCol > 0 Then
            v = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
            ReDim keys(1 To lastCol)
            For j = 1 To lastCol
                keys(j) = CStr(modCore.Nz(v(1, j), ""))
            Next j
            For i = TPL_FIRST_ROW To lastRow
                ImportEntityRow fname, i, v, keys, typeCode, prefill, commit, batch, src, nameIdx, newNames, rows, _
                                touched, nNew, nUpd, nErr, nSame
            Next i
        End If
    End If

    ' ---- relationships
    If modCore.SheetExists(wb, "Relationships") Then
        Set ws = wb.Worksheets("Relationships")
        lastRow = modCore.LastRowOf(ws)
        If lastRow >= TPL_FIRST_ROW Then
            v = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 6)).Value
            For i = TPL_FIRST_ROW To lastRow
                ImportRelRow fname, i, v, typeCode, commit, batch, src, nameIdx, newNames, rows, touched, nRel, nErr, nSame
            Next i
        End If
    End If
CloseFile:
    wb.Close SaveChanges:=False
    Exit Sub
Fail:
    Dim eMsg As String
    eMsg = fname & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 501, "modIO", eMsg
End Sub

' Keeps the name index current while a commit adds or renames entities
Private Sub IndexSet(nameIdx As Object, ByVal typeCode As String, ByVal nm As String, ByVal id As String, ByVal add As Boolean)
    Dim k As Variant
    For Each k In Array(typeCode & "|" & LCase$(Trim$(nm)), "*|" & LCase$(Trim$(nm)))
        If add Then
            If Not nameIdx.Exists(k) Then nameIdx.Add k, modCore.NewDict()
            nameIdx(k)(id) = True
        ElseIf nameIdx.Exists(k) Then
            If nameIdx(k).Exists(id) Then nameIdx(k).Remove id
            If nameIdx(k).Count = 0 Then nameIdx.Remove k
        End If
    Next k
End Sub

Private Function ReadMeta(wb As Workbook) As Object
    Dim d As Object, ws As Worksheet, r As Long, k As Variant
    Set d = modCore.NewDict()
    For Each k In Array("Kind", "ModelID", "ModelName", "TypeCode", "SchemaVersion", "Prefilled")
        d(k) = ""
    Next k
    Set ReadMeta = d
    If Not modCore.SheetExists(wb, "_meta") Then Exit Function
    Set ws = wb.Worksheets("_meta")
    For r = 1 To 30
        If CStr(ws.Cells(r, 1).Value) <> "" Then d(CStr(ws.Cells(r, 1).Value)) = ws.Cells(r, 2).Value
    Next r
End Function

Private Sub ImportEntityRow(ByVal fname As String, ByVal i As Long, v As Variant, keys As Variant, ByVal typeCode As String, _
                            ByVal prefill As Boolean, ByVal commit As Boolean, ByVal batch As String, ByVal src As String, _
                            nameIdx As Object, newNames As Object, rows As Collection, touched As Object, _
                            ByRef nNew As Long, ByRef nUpd As Long, ByRef nErr As Long, ByRef nSame As Long)
    Dim vals As Object, j As Long, k As String, id As String, nm As String, any_ As Boolean, r As Long
    Dim cur As Object, changed As Object, issues As Collection, it As Variant, p() As String, merged As Object
    Dim action As String, blocked As Boolean, disp As String, newID As String, matches As Object
    Set vals = modCore.NewDict()
    For j = 1 To UBound(keys)
        k = keys(j)
        If k <> "" And k <> "DisplayID" And k <> "EntityID" Then
            vals(k) = modData.ToCellValue(k, v(i, j))
            If Not modCore.IsBlankValue(v(i, j)) And k <> "Status" And k <> "RepositoryArea" Then any_ = True
        End If
        If k = "EntityID" Then id = Trim$(CStr(modCore.Nz(v(i, j), "")))
    Next j
    If Not any_ Then Exit Sub
    nm = Trim$(CStr(modCore.Nz(RowVal(vals, "Name"), "")))

    ' ---- match: Entity ID, then name within the type
    If id <> "" Then
        r = modData.FindEntityRow(id)
        If r = 0 Then
            AddRow rows, fname, "Entities", i, "Entity", "Error", id, "", nm, typeCode, "", "", "", "Error", _
                   "Entity ID is not in this model"
            nErr = nErr + 1
            Exit Sub
        End If
        If modData.EntityTypeOf(id) <> typeCode Then
            AddRow rows, fname, "Entities", i, "Entity", "Error", id, "", nm, typeCode, "", "", "", "Error", _
                   "That entity is a " & modMeta.TypeName_(modData.EntityTypeOf(id)) & ", not a " & modMeta.TypeName_(typeCode)
            nErr = nErr + 1
            Exit Sub
        End If
    ElseIf nm <> "" Then
        k = typeCode & "|" & LCase$(nm)
        If nameIdx.Exists(k) Then
            Set matches = nameIdx(k)
            If matches.Count > 1 Then
                AddRow rows, fname, "Entities", i, "Entity", "Error", "", "", nm, typeCode, "", "", "", "Error", _
                       matches.Count & " entities of this type have this name; add the Entity ID to choose one"
                nErr = nErr + 1
                Exit Sub
            End If
            Dim mk As Variant
            mk = matches.Keys
            id = CStr(mk(0))
            r = modData.FindEntityRow(id)
        End If
    End If

    If id = "" Then
        ' ---- new entity
        If nm = "" Then
            AddRow rows, fname, "Entities", i, "Entity", "Error", "", "", "", typeCode, "Name", "", "", "Error", "Name is required"
            nErr = nErr + 1
            Exit Sub
        End If
        If newNames.Exists(LCase$(nm)) Then
            AddRow rows, fname, "Entities", i, "Entity", "Error", "", "", nm, typeCode, "Name", "", nm, "Error", _
                   "The same name appears twice in this file"
            nErr = nErr + 1
            Exit Sub
        End If
        Set merged = modCore.NewDict()
        For Each it In vals.Keys
            If Not modCore.IsBlankValue(vals(it)) Then merged(it) = vals(it)
        Next it
        If Not merged.Exists("Status") Then merged("Status") = "Active"
        If Not merged.Exists("RepositoryArea") Then merged("RepositoryArea") = "Architecture Landscape"
        Set issues = modRules.ValidateRow(typeCode, merged, nameIdx, "")
        If modRules.HasErrors(issues) Then
            AddIssueRows rows, fname, i, "", "", nm, typeCode, issues, merged
            nErr = nErr + 1
            Exit Sub
        End If
        newNames(LCase$(nm)) = ""
        If commit Then
            Set issues = New Collection
            newID = modData.CreateEntity(typeCode, merged, src, batch, issues)
            If newID = "" Then
                newNames.Remove LCase$(nm)
                AddIssueRows rows, fname, i, "", "", nm, typeCode, issues, merged
                nErr = nErr + 1
                Exit Sub
            End If
            newNames(LCase$(nm)) = newID
            IndexSet nameIdx, typeCode, nm, newID, True
            touched(typeCode) = True
            disp = CStr(modData.MasterSheet().Cells(modData.FindEntityRow(newID), modData.MasterCols()("DisplayID")).Value)
            AddRow rows, fname, "Entities", i, "Entity", "Added", newID, disp, nm, typeCode, "(new entity)", "", nm, "", ""
        Else
            AddRow rows, fname, "Entities", i, "Entity", "New", "", "(on commit)", nm, typeCode, "(new entity)", "", nm, "", ""
        End If
        AddWarnings rows, fname, i, "", nm, typeCode, issues
        nNew = nNew + 1
        Exit Sub
    End If

    ' ---- existing entity: changed values only
    Set cur = modData.ReadEntityRow(r)
    disp = CStr(cur("DisplayID"))
    Set changed = modCore.NewDict()
    For Each it In vals.Keys
        If modMeta.IsSystemAttr(CStr(it)) Or modMeta.IsMapped(typeCode, CStr(it)) Then
            If modCore.IsBlankValue(vals(it)) And Not prefill Then
                ' blank in a fresh template: leave the current value alone
            ElseIf modCore.CellText(vals(it)) <> modCore.CellText(cur(it)) Then
                changed(it) = vals(it)
            End If
        ElseIf Not modCore.IsBlankValue(vals(it)) Then
            AddRow rows, fname, "Entities", i, "Entity", "Info", CStr(cur("EntityID")), disp, CStr(cur("Name")), typeCode, _
                   CStr(it), "", vals(it), "Warning", modMeta.AttrName(CStr(it)) & " no longer applies to this type - ignored"
        End If
    Next it
    If changed.Count = 0 Then
        AddRow rows, fname, "Entities", i, "Entity", "Unchanged", CStr(cur("EntityID")), disp, CStr(cur("Name")), typeCode, "", "", "", "", ""
        nSame = nSame + 1
        Exit Sub
    End If
    Set merged = modCore.NewDict()
    For Each it In cur.Keys: merged(it) = cur(it): Next it
    For Each it In changed.Keys: merged(it) = changed(it): Next it
    Set issues = modRules.ValidateRow(typeCode, merged, nameIdx, CStr(cur("EntityID")))
    If BlockingIssue(issues, changed) Then
        AddIssueRows rows, fname, i, CStr(cur("EntityID")), disp, CStr(cur("Name")), typeCode, issues, changed
        nErr = nErr + 1
        Exit Sub
    End If
    If commit Then
        Set issues = New Collection
        blocked = False
        modData.UpdateEntity CStr(cur("EntityID")), changed, src, batch, issues, blocked
        If blocked Then
            AddIssueRows rows, fname, i, CStr(cur("EntityID")), disp, CStr(cur("Name")), typeCode, issues, changed
            nErr = nErr + 1
            Exit Sub
        End If
        touched(typeCode) = True
        If changed.Exists("Name") Then
            IndexSet nameIdx, typeCode, CStr(cur("Name")), CStr(cur("EntityID")), False
            IndexSet nameIdx, typeCode, CStr(changed("Name")), CStr(cur("EntityID")), True
        End If
    End If
    If changed.Exists("Name") Then newNames(LCase$(Trim$(CStr(changed("Name"))))) = CStr(cur("EntityID"))
    action = IIf(commit, "Updated", "Update")
    For Each it In changed.Keys
        AddRow rows, fname, "Entities", i, "Entity", action, CStr(cur("EntityID")), disp, CStr(cur("Name")), typeCode, _
               modMeta.AttrName(CStr(it)), cur(it), changed(it), "", ""
    Next it
    AddWarnings rows, fname, i, CStr(cur("EntityID")), CStr(cur("Name")), typeCode, issues
    nUpd = nUpd + 1
End Sub

Private Function RowVal(d As Object, ByVal k As String) As Variant
    If d.Exists(k) Then RowVal = d(k)
End Function

' An error blocks an update only when it concerns a field the import changes
Private Function BlockingIssue(issues As Collection, changed As Object) As Boolean
    Dim it As Variant, p() As String
    For Each it In issues
        p = Split(it, "|", 4)
        If p(0) = "Error" And changed.Exists(p(2)) Then BlockingIssue = True: Exit Function
    Next it
End Function

Private Sub AddIssueRows(rows As Collection, ByVal fname As String, ByVal i As Long, ByVal id As String, ByVal disp As String, _
                         ByVal nm As String, ByVal typeCode As String, issues As Collection, vals As Object)
    Dim it As Variant, p() As String
    For Each it In issues
        p = Split(it, "|", 4)
        AddRow rows, fname, "Entities", i, "Entity", IIf(p(0) = "Error", "Error", "Info"), id, disp, nm, typeCode, _
               modMeta.AttrName(p(2)), "", RowVal(vals, p(2)), p(0), p(3)
    Next it
End Sub

Private Sub AddWarnings(rows As Collection, ByVal fname As String, ByVal i As Long, ByVal id As String, ByVal nm As String, _
                        ByVal typeCode As String, issues As Collection)
    Dim it As Variant, p() As String
    If issues Is Nothing Then Exit Sub
    For Each it In issues
        p = Split(it, "|", 4)
        If p(0) = "Warning" Then
            AddRow rows, fname, "Entities", i, "Entity", "Info", id, "", nm, typeCode, modMeta.AttrName(p(2)), "", "", "Warning", p(3)
        End If
    Next it
End Sub

Private Sub ImportRelRow(ByVal fname As String, ByVal i As Long, v As Variant, ByVal typeCode As String, ByVal commit As Boolean, _
                         ByVal batch As String, ByVal src As String, nameIdx As Object, newNames As Object, rows As Collection, touched As Object, _
                         ByRef nRel As Long, ByRef nErr As Long, ByRef nSame As Long)
    Dim sn As String, rc As String, tt As String, tn As String, lbl As String, ds As String
    Dim sid As String, tid As String, msg As String, issues As Collection, id As String, desc As String
    sn = Trim$(CStr(modCore.Nz(v(i, 1), ""))): rc = Trim$(CStr(modCore.Nz(v(i, 2), "")))
    tt = Trim$(CStr(modCore.Nz(v(i, 3), ""))): tn = Trim$(CStr(modCore.Nz(v(i, 4), "")))
    lbl = CStr(modCore.Nz(v(i, 5), "")): ds = CStr(modCore.Nz(v(i, 6), ""))
    If sn & rc & tt & tn & lbl & ds = "" Then Exit Sub
    desc = sn & " " & modMeta.RelVerb(rc) & " " & tn
    If sn = "" Or rc = "" Or tt = "" Or tn = "" Then
        msg = "Source name, relationship, target type and target name are all required"
    ElseIf Not modMeta.gRelTypes.Exists(rc) Then
        msg = "Unknown relationship " & rc
    ElseIf Not modMeta.TypeExists(tt) Then
        msg = "Unknown target type " & tt
    ElseIf Not modMeta.IsRelAllowed(typeCode, rc, tt) Then
        msg = modMeta.TypeName_(typeCode) & " " & modMeta.RelVerb(rc) & " " & modMeta.TypeName_(tt) & " is not allowed"
    Else
        Dim targetNew As Object
        If tt = typeCode Then Set targetNew = newNames         ' only this file's own type can be new
        sid = ResolveName(typeCode, sn, nameIdx, newNames, msg)
        If msg = "" Then tid = ResolveName(tt, tn, nameIdx, targetNew, msg)
    End If
    If msg <> "" Then
        AddRow rows, fname, "Relationships", i, "Relationship", "Error", "", "", desc, rc, "", "", "", "Error", msg
        nErr = nErr + 1
        Exit Sub
    End If
    If sid <> "(new)" And tid <> "(new)" Then
        If modData.RelationshipExists(rc, sid, tid) Then
            AddRow rows, fname, "Relationships", i, "Relationship", "Unchanged", "", "", desc, rc, "", "", "", "", ""
            nSame = nSame + 1
            Exit Sub
        End If
    End If
    If commit Then
        Set issues = New Collection
        id = modData.CreateRelationship(rc, sid, tid, lbl, ds, src, batch, issues)
        If id = "" Then
            AddRow rows, fname, "Relationships", i, "Relationship", "Error", "", "", desc, rc, "", "", "", "Error", _
                   Replace(modRules.IssueText(issues, 2), vbLf, "; ")
            nErr = nErr + 1
            Exit Sub
        End If
        touched(typeCode) = True: touched(tt) = True
        AddRow rows, fname, "Relationships", i, "Relationship", "Added", id, "", desc, rc, "", "", "", "", ""
    Else
        AddRow rows, fname, "Relationships", i, "Relationship", "New", "", "", desc, rc, "", "", "", "", ""
    End If
    nRel = nRel + 1
End Sub

' EntityID for a name within a type; "(new)" for a name created earlier in this file (preview only);
' msg set when nothing or more than one entity matches.
Private Function ResolveName(ByVal typeCode As String, ByVal nm As String, idx As Object, newNames As Object, _
                             ByRef msg As String) As String
    Dim k As String
    k = typeCode & "|" & LCase$(nm)
    If idx.Exists(k) Then
        If idx(k).Count > 1 Then
            msg = "More than one " & modMeta.TypeName_(typeCode) & " is called " & nm
        Else
            Dim ik As Variant
            ik = idx(k).Keys
            ResolveName = CStr(ik(0))
        End If
        Exit Function
    End If
    If Not newNames Is Nothing Then
        If newNames.Exists(LCase$(nm)) Then
            ResolveName = IIf(CStr(newNames(LCase$(nm))) = "", "(new)", CStr(newNames(LCase$(nm))))
            Exit Function
        End If
    End If
    msg = "No " & modMeta.TypeName_(typeCode) & " called " & nm
End Function

Private Sub AddRow(rows As Collection, ByVal fname As String, ByVal sheetName As String, ByVal r As Long, ByVal kind As String, _
                   ByVal action As String, ByVal id As String, ByVal disp As String, ByVal nm As String, ByVal typeCode As String, _
                   ByVal field As String, ByVal curV As Variant, ByVal newV As Variant, ByVal sev As String, ByVal msg As String)
    rows.Add Array(fname, sheetName, r, kind, action, id, disp, nm, typeCode, field, curV, newV, sev, msg)
End Sub

Private Sub WritePreview(ByVal fileList As String, rows As Collection, ByVal committed As Boolean)
    Dim ws As Worksheet, lo As ListObject, x As Variant, core As String
    Set ws = ThisWorkbook.Worksheets(SH_PREVIEW)
    Set lo = PreviewLO()
    core = "'" & ThisWorkbook.Name & "'!modIO."
    modCore.BeginBusy
    modCore.UnprotectSheet ws
    modCore.DeleteButtons ws
    modCore.ClearTable lo
    For Each x In rows
        modCore.AppendRow lo, x
    Next x
    ws.Range("A3").Value = IIf(committed, "Committed: ", "Preview of: ") & rows.Count & " row(s)"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value = "Files"
    ws.Range("B4").Value = IIf(committed, "", fileList)     ' cleared after a commit so it cannot run twice
    ws.Range("B4").Font.Color = RGB(89, 89, 89)
    If Not committed Then modCore.AddButton ws, "Commit import", core & "ImportCommit", ws.Range("H1"), 150
    modCore.AddButton ws, "Templates, import and export", core & "ShowIO", ws.Range("J1"), 190
    modCore.AddButton ws, "Back to Home", "'" & ThisWorkbook.Name & "'!modForms.GoHome", ws.Range("M1"), 120
    modCore.ProtectSheet ws
    modCore.EndBusy
    ThisWorkbook.Activate
    ws.Activate
End Sub

' Moves an imported file into Imports\Done\ (only files that were in Imports\)
Private Function MoveToDone(ByVal path As String) As String
    Dim imports As String, done As String, fname As String
    On Error GoTo Fail
    imports = LCase$(SubFolder("Imports"))
    fname = Mid$(path, InStrRev(path, Application.PathSeparator) + 1)
    If LCase$(Left$(path, Len(imports))) <> imports Then
        MoveToDone = vbLf & fname & " was not in the Imports folder, so it was left where it is."
        Exit Function
    End If
    done = SubFolder("Imports" & Application.PathSeparator & "Done")
    Name path As done & Stamp() & "_" & fname
    Exit Function
Fail:
    MoveToDone = vbLf & fname & " could not be moved to Imports\Done (" & Err.Description & ")."
End Function

' ================================================================ export
Public Sub ExportFromForm()
    Dim ws As Worksheet, scope As String, ids As Object, msg As String, p As String
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    scope = CStr(modForms.FieldValue(ws, "_scope"))
    Set ids = SelectEntities(scope, CStr(modForms.FieldValue(ws, "_etype")), CStr(modForms.FieldValue(ws, "_start")), _
                             Val(CStr(modForms.FieldValue(ws, "_hops"))), _
                             UCase$(CStr(modForms.FieldValue(ws, "_retired"))) = "Y", msg)
    modCore.UnprotectSheet ws
    If msg <> "" Then
        modForms.FormMessage ws, msg, True
        modCore.ProtectSheet ws
        Exit Sub
    End If
    modCore.ProtectSheet ws
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    p = ExportEntities(ids, scope, CStr(modForms.FieldValue(ws, "_format")) = "CSV", scope = "Full model with metamodel")
    If p <> "" Then
        modCore.UnprotectSheet ws
        modForms.FormMessage ws, "Exported " & ids.Count & " entit" & IIf(ids.Count = 1, "y", "ies") & " to " & p
        modCore.ProtectSheet ws
        MsgBox "Exported " & ids.Count & " entit" & IIf(ids.Count = 1, "y", "ies") & " to:" & vbLf & p, vbInformation, "EA Workbench"
    End If
End Sub

' Entity IDs for an export scope (Dictionary id -> True)
Public Function SelectEntities(ByVal scope As String, ByVal typeCode As String, ByVal startDisp As String, _
                               ByVal hops As Long, ByVal withRetired As Boolean, ByRef msg As String) As Object
    Dim ids As Object, lo As ListObject, cols As Object, d As Variant, i As Long, base As Long, id As String
    Dim st As String, tp As String, startID As String
    Set ids = modCore.NewDict()
    Set SelectEntities = ids
    Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    d = modCore.TableData(lo)
    If Not IsArray(d) Then msg = "The repository has no entities.": Exit Function
    base = lo.Range.Column - 1
    Select Case scope
        Case "One entity type"
            If Not modMeta.TypeExists(typeCode) Then msg = "Pick the entity type to export.": Exit Function
        Case "Around one entity"
            For i = 1 To UBound(d, 1)
                If StrComp(CStr(d(i, cols("DisplayID") - base)), Trim$(startDisp), vbTextCompare) = 0 Then
                    startID = CStr(d(i, cols("EntityID") - base))
                End If
            Next i
            If startID = "" Then msg = "Enter the Display ID of the start entity (e.g. APP-0002).": Exit Function
            If hops < 1 Then hops = 1
            If hops > 5 Then hops = 5
            Set SelectEntities = Neighbourhood(startID, hops)
            Exit Function
        Case "All entities", "Full model with metamodel"
        Case Else
            msg = "Pick a scope.": Exit Function
    End Select
    For i = 1 To UBound(d, 1)
        id = CStr(d(i, cols("EntityID") - base))
        st = CStr(d(i, cols("Status") - base))
        tp = CStr(d(i, cols("TypeCode") - base))
        If id <> "" And (withRetired Or st <> "Retired" Or scope = "Full model with metamodel") Then
            If scope <> "One entity type" Or tp = typeCode Then ids(id) = True
        End If
    Next i
End Function

' Entities within n relationship hops of a start entity, in both directions (Dictionary id -> hop level)
Public Function Neighbourhood(ByVal startID As String, ByVal hops As Long) As Object
    Dim seen As Object, frontier As Collection, nextF As Collection, d As Variant, i As Long, lvl As Long, x As Variant
    Dim adj As Object, a As String, b As String
    Set seen = modCore.NewDict()
    seen(startID) = 0
    Set adj = modCore.NewDict()
    d = modCore.TableData(modData.RelLO())
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, 1)) <> "" And CStr(d(i, 11)) <> "Retired" Then
                a = CStr(d(i, 3)): b = CStr(d(i, 4))
                If Not adj.Exists(a) Then adj.Add a, New Collection
                If Not adj.Exists(b) Then adj.Add b, New Collection
                adj(a).Add b: adj(b).Add a
            End If
        Next i
    End If
    Set frontier = New Collection
    frontier.Add startID
    For lvl = 1 To hops
        Set nextF = New Collection
        For Each x In frontier
            If adj.Exists(x) Then
                Dim y As Variant
                For Each y In adj(x)
                    If Not seen.Exists(y) Then seen(y) = lvl: nextF.Add y
                Next y
            End If
        Next x
        Set frontier = nextF
        If frontier.Count = 0 Then Exit For
    Next lvl
    Set Neighbourhood = seen
End Function

' Writes the export; returns the path (file or folder), "" on failure
Public Function ExportEntities(ids As Object, ByVal scope As String, ByVal asCsv As Boolean, ByVal withMeta As Boolean) As String
    Dim wb As Workbook, ws As Worksheet, folder As String, base As String, path As String, t As Variant
    Dim lo As ListObject, cols As Object, d As Variant, i As Long, b As Long, r As Long, c As Long, k As Variant
    Dim keys As Collection, byType As Object, nRel As Long, i_ As Variant, order As Collection, seen As Object
    On Error GoTo Fail
    modMeta.LoadMeta
    folder = SubFolder("Exports")
    base = SafeFile(CStr(modCore.Setting("ModelName"))) & "-export-" & SafeFile(scope) & "-" & Stamp()
    modCore.BeginBusy
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "README"
    WriteLines ws, 1, Array("Export of " & CStr(modCore.Setting("ModelName")), "Scope: " & scope, _
                            "Exported (UTC): " & modCore.UtcNow(), "Entities: " & ids.Count, "", _
                            "One sheet per entity type, then Relationships (both ends in the export).")
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Columns("A").ColumnWidth = 80

    ' ---- entities, one sheet per type
    Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    d = modCore.TableData(lo)
    b = lo.Range.Column - 1
    Set byType = modCore.NewDict()
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If ids.Exists(CStr(d(i, cols("EntityID") - b))) Then
                t = CStr(d(i, cols("TypeCode") - b))
                If Not byType.Exists(t) Then byType.Add t, New Collection
                byType(t).Add i
            End If
        Next i
    End If
    ' sheet order first, then any type without a sheet (inactive types)
    Set order = New Collection: Set seen = modCore.NewDict()
    For Each t In modPaint.TypesInSheetOrder()
        If byType.Exists(t) And Not seen.Exists(t) Then order.Add CStr(t): seen(t) = True
    Next t
    For Each t In byType.Keys
        If Not seen.Exists(t) Then order.Add CStr(t): seen(t) = True
    Next t
    For Each t In order
        If byType.Exists(t) Then
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = UniqueSheetName(wb, modMeta.TypeName_(CStr(t)))
            Set keys = New Collection
            For Each k In Array("DisplayID", "Name", "Description", "Status", "RepositoryArea"): keys.Add CStr(k): Next k
            For Each k In modMeta.AttrsForType(CStr(t)): keys.Add CStr(k): Next k
            keys.Add "Modified": keys.Add "EntityID"
            c = 0
            For Each k In keys
                c = c + 1
                ws.Cells(1, c).Value = HeaderOf(CStr(k))
                ws.Range(ws.Cells(2, c), ws.Cells(byType(t).Count + 1, c)).NumberFormat = ExportFormat(CStr(k))
            Next k
            ws.Range(ws.Cells(1, 1), ws.Cells(1, keys.Count)).Font.Bold = True
            r = 1
            For Each i_ In byType(t)
                r = r + 1
                c = 0
                For Each k In keys
                    c = c + 1
                    If cols.Exists(k) Then ws.Cells(r, c).Value = modCore.SafeFor(d(i_, cols(k) - b), CStr(ws.Cells(r, c).NumberFormat))
                Next k
            Next i_
            ws.Columns.AutoFit
        End If
    Next t

    ' ---- relationships with both ends in the export
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Relationships"
    ws.Range("A1:L1").Value = Array("Relationship", "Verb", "Source Display ID", "Source name", "Source type", _
                                    "Target Display ID", "Target name", "Target type", "Label", "Description", "Status", "RelID")
    ws.Range("A1:L1").Font.Bold = True
    Dim relRows As New Collection, rr As Variant
    d = modCore.TableData(modData.RelLO())
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If ids.Exists(CStr(d(i, 3))) And ids.Exists(CStr(d(i, 4))) Then
                relRows.Add Array(CStr(d(i, 2)), modMeta.RelVerb(CStr(d(i, 2))), _
                    DispOf(CStr(d(i, 3))), CStr(d(i, 5)), CStr(d(i, 6)), DispOf(CStr(d(i, 4))), CStr(d(i, 7)), CStr(d(i, 8)), _
                    CStr(modCore.Nz(d(i, 9), "")), CStr(modCore.Nz(d(i, 10), "")), CStr(modCore.Nz(d(i, 11), "")), CStr(d(i, 1)))
            End If
        Next i
    End If
    nRel = relRows.Count
    If nRel > 0 Then ws.Range("A2:L" & (nRel + 1)).NumberFormat = "@"
    r = 1
    For Each rr In relRows
        r = r + 1
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Value = rr
    Next rr
    ws.Columns.AutoFit

    ' ---- metamodel (full model only)
    If withMeta Then
        Dim mws As Worksheet
        For Each mws In ThisWorkbook.Worksheets
            If mws.Name Like "MM_*" Then CopySheetValues mws, wb
        Next mws
    End If

    Application.DisplayAlerts = False
    If asCsv Then
        path = folder & base & Application.PathSeparator
        MkDir path
        For Each ws In wb.Worksheets
            If ws.Name <> "README" Then
                ws.Copy
                ActiveWorkbook.SaveAs Filename:=path & SafeFile(ws.Name) & ".csv", FileFormat:=62   ' xlCSVUTF8
                ActiveWorkbook.Close SaveChanges:=False
            End If
        Next ws
        wb.Close SaveChanges:=False
    Else
        path = folder & base & ".xlsx"
        wb.Worksheets(1).Activate
        wb.SaveAs Filename:=path, FileFormat:=51
    End If
    Application.DisplayAlerts = True
    modCore.EndBusy
    ExportEntities = path
    Exit Function
Fail:
    Dim msg As String
    msg = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    modCore.ResetBusy
    MsgBox "Export failed: " & msg, vbExclamation, "EA Workbench"
End Function

Private Function HeaderOf(ByVal k As String) As String
    Select Case k
        Case "DisplayID": HeaderOf = "Display ID"
        Case "RepositoryArea": HeaderOf = "Repository Area"
        Case "EntityID": HeaderOf = "Entity ID"
        Case "Name", "Description", "Status", "Modified": HeaderOf = k
        Case Else: HeaderOf = modMeta.AttrName(k)
    End Select
End Function

Private Function ExportFormat(ByVal k As String) As String
    Select Case k
        Case "DisplayID", "Name", "Description", "Status", "RepositoryArea", "Modified", "EntityID": ExportFormat = "@"
        Case Else: ExportFormat = modGenerator.NumberFormatFor(k)
    End Select
End Function

Private Function DispOf(ByVal id As String) As String
    Dim r As Long
    r = modData.FindEntityRow(id)
    If r > 0 Then DispOf = CStr(modData.MasterSheet().Cells(r, modData.MasterCols()("DisplayID")).Value)
End Function

Private Function UniqueSheetName(wb As Workbook, ByVal wanted As String) As String
    Dim nm As String, n As Long
    nm = Left$(SafeFile(wanted), 31)
    Do While modCore.SheetExists(wb, nm)
        n = n + 1
        nm = Left$(SafeFile(wanted), 27) & " (" & n & ")"
    Loop
    UniqueSheetName = nm
End Function

Private Function SafeFile(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9 _&().-]" Then out = out & ch Else out = out & "_"
    Next i
    SafeFile = Trim$(Left$(out, 60))
End Function

Private Sub CopySheetValues(src As Worksheet, wb As Workbook)
    Dim dst As Worksheet, lastR As Long, lastC As Long, v As Variant, i As Long, j As Long
    Set dst = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    dst.Name = src.Name
    lastR = modCore.LastRowOf(src): lastC = modCore.LastColOf(src)
    If lastR = 0 Or lastC = 0 Then Exit Sub
    v = src.Range(src.Cells(1, 1), src.Cells(lastR, lastC)).Value
    If Not IsArray(v) Then Exit Sub
    For i = 1 To UBound(v, 1)
        For j = 1 To UBound(v, 2)
            If IsError(v(i, j)) Then
                v(i, j) = Empty
            ElseIf VarType(v(i, j)) = vbString Then
                v(i, j) = "'" & v(i, j)
            End If
        Next j
    Next i
    dst.Range("A1").Resize(lastR, lastC).Value = v
End Sub
