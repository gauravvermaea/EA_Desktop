Attribute VB_Name = "modForms"
' =====================================================================
' Wizards, painted on FORM_ sheets from the metamodel at run time:
'   FORM_Entity        add / edit / retire an entity
'   FORM_Relationship  add a relationship (only allowed types offered)
'   FORM_Type          add / edit entity, attribute and relationship types
'   FORM_Rule          define and test validation rules
' Layout: column A label, B input, C hidden field code, D message.
' Helper lists for dropdowns live in hidden columns X:AF.
' =====================================================================
Option Explicit

Private Const FIRST_FIELD_ROW As Long = 7
Private Const MSG_CELL As String = "A3"

' ---------------------------------------------------------------- shared form helpers
Private Function FormWS(ByVal nm As String) As Worksheet
    Set FormWS = ThisWorkbook.Worksheets(nm)
End Function

Public Sub ResetForm(ws As Worksheet, ByVal title As String, ByVal hint As String, Optional ByVal fromRow As Long = 4)
    modCore.UnprotectSheet ws
    modCore.DeleteButtons ws
    ws.Range(ws.Cells(fromRow, 1), ws.Cells(600, 20)).Clear
    ws.Range(ws.Cells(fromRow, 1), ws.Cells(600, 20)).Validation.Delete
    ws.Range("X:AF").Clear
    ws.Range("A1").Value = title
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(31, 56, 100)
    ws.Range("A2").Value = hint
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Range(MSG_CELL).ClearContents
    ws.Columns("A").ColumnWidth = 30: ws.Columns("B").ColumnWidth = 46
    ws.Columns("C").Hidden = True
    ws.Columns("D").ColumnWidth = 60: ws.Columns("E").ColumnWidth = 16
    ws.Columns("X:AF").Hidden = True
    ActiveWindow.DisplayGridlines = False
End Sub

Public Sub Field(ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal code As String, _
                  Optional ByVal value As Variant, Optional ByVal readOnly As Boolean = False)
    If IsMissing(value) Then value = Empty
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 1).Font.Bold = (Right$(label, 1) = "*")
    ws.Cells(r, 3).Value = code
    ws.Cells(r, 2).Value = modCore.SafeFor(value, CStr(ws.Cells(r, 2).NumberFormat))
    With ws.Cells(r, 2)
        .Locked = readOnly
        .Interior.Color = IIf(readOnly, RGB(242, 242, 242), RGB(255, 255, 255))
        .Borders.Color = RGB(191, 191, 191)
    End With
End Sub

Public Function FieldRow(ws As Worksheet, ByVal code As String) As Long
    Dim r As Variant
    r = Application.Match(code, ws.Range("C1:C600"), 0)
    If Not IsError(r) Then FieldRow = CLng(r)
End Function

Public Function FieldValue(ws As Worksheet, ByVal code As String) As Variant
    Dim r As Long
    r = FieldRow(ws, code)
    If r > 0 Then FieldValue = ws.Cells(r, 2).Value
End Function

Public Sub SetFieldValue(ws As Worksheet, ByVal code As String, ByVal v As Variant)
    Dim r As Long
    r = FieldRow(ws, code)
    If r > 0 Then ws.Cells(r, 2).Value = v
End Sub

Public Sub FieldMsg(ws As Worksheet, ByVal code As String, ByVal text As String, ByVal sev As String)
    Dim r As Long
    r = FieldRow(ws, code)
    If r = 0 Then Exit Sub
    ws.Cells(r, 4).Value = text
    ws.Cells(r, 4).Font.Color = IIf(sev = "Error", RGB(192, 0, 0), RGB(156, 101, 0))
End Sub

Public Sub FormMessage(ws As Worksheet, ByVal text As String, Optional ByVal isError As Boolean = False)
    With ws.Range(MSG_CELL)
        .Value = text
        .Font.Bold = True
        .Font.Color = IIf(isError, RGB(192, 0, 0), RGB(0, 97, 0))
    End With
End Sub

' Clears field messages in column D, stopping at a grid (whose column D holds data)
Public Sub ClearFieldMsgs(ws As Worksheet)
    Dim r As Long
    For r = 4 To 600
        If CStr(ws.Cells(r, 3).Value) = "_grid" Then Exit For
        ws.Cells(r, 4).ClearContents
    Next r
    ws.Range(MSG_CELL).ClearContents
End Sub

' Writes a helper list in a hidden column and points the field's dropdown at it.
Private Sub HelperList(ws As Worksheet, ByVal code As String, ByVal labelCol As String, items As Collection, _
                       Optional ids As Collection = Nothing, Optional ByVal idCol As String = "")
    Dim r As Long, i As Long, fr As Long
    ws.Range(labelCol & ":" & labelCol).ClearContents
    If idCol <> "" Then ws.Range(idCol & ":" & idCol).ClearContents
    ws.Range(labelCol & "1").Value = code
    For i = 1 To items.Count
        ws.Range(labelCol & (i + 1)).Value = "'" & items(i)
        If Not ids Is Nothing Then ws.Range(idCol & (i + 1)).Value = ids(i)
    Next i
    r = IIf(items.Count = 0, 2, items.Count + 1)
    If items.Count = 0 Then ws.Range(labelCol & "2").Value = "(none)"
    fr = FieldRow(ws, code)
    If fr = 0 Then Exit Sub
    With ws.Cells(fr, 2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$" & labelCol & "$2:$" & labelCol & "$" & r
        .IgnoreBlank = True
    End With
End Sub

Private Function LookupId(ws As Worksheet, ByVal labelCol As String, ByVal idCol As String, ByVal label As String) As String
    Dim r As Variant
    If label = "" Then Exit Function
    r = Application.Match(label, ws.Range(labelCol & ":" & labelCol), 0)
    If Not IsError(r) Then LookupId = CStr(ws.Range(idCol & r).Value)
End Function

Public Sub ListDropdown(ws As Worksheet, ByVal code As String, ByVal listName As String)
    Dim r As Long
    r = FieldRow(ws, code)
    If r > 0 Then modGenerator.ApplyListValidation ws.Cells(r, 2), listName
End Sub

Public Sub FinishForm(ws As Worksheet)
    ws.Activate
    modCore.ProtectSheet ws
End Sub

Private Function MacroRef(ByVal procName As String) As String
    MacroRef = "'" & ThisWorkbook.Name & "'!modForms." & procName
End Function

Private Function EntityLabel(ByVal displayID As String, ByVal nm As String) As String
    EntityLabel = displayID & " " & ChrW$(183) & " " & nm
End Function

' Entities of a type: labels and IDs, sorted by name
Private Sub EntitiesOfType(ByVal typeCode As String, labels As Collection, ids As Collection)
    Dim d As Variant, i As Long, cols As Object, base As Long, lo As ListObject, j As Long, inserted As Boolean
    Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    base = lo.Range.Column - 1
    d = modCore.TableData(lo)
    If Not IsArray(d) Then Exit Sub
    For i = 1 To UBound(d, 1)
        If CStr(d(i, cols("TypeCode") - base)) = typeCode And CStr(d(i, cols("EntityID") - base)) <> "" Then
            Dim lbl As String
            lbl = EntityLabel(CStr(d(i, cols("DisplayID") - base)), CStr(d(i, cols("Name") - base)))
            inserted = False
            For j = 1 To labels.Count
                If LCase$(Mid$(lbl, InStr(lbl, " ") + 3)) < LCase$(Mid$(labels(j), InStr(labels(j), " ") + 3)) Then
                    labels.Add lbl, Before:=j: ids.Add CStr(d(i, cols("EntityID") - base)), Before:=j
                    inserted = True: Exit For
                End If
            Next j
            If Not inserted Then labels.Add lbl: ids.Add CStr(d(i, cols("EntityID") - base))
        End If
    Next i
End Sub

Public Sub GoHome()
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(modCore.SH_HOME).Activate
End Sub

' ---------------------------------------------------------------- routing of form edits
Public Sub OnFormChange(ByVal ws As Worksheet, ByVal Target As Range)
    If Target.Column <> 2 Or Target.CountLarge > 1 Then Exit Sub
    Dim code As String
    code = CStr(ws.Cells(Target.Row, 3).Value)
    Select Case ws.Name
        Case "FORM_Entity"
            If code = "_type" Then ShowEntityForm CStr(Target.Value), ""
            If code = "_entity" Then ShowEntityForm CStr(FieldValue(ws, "_type")), _
                                                    LookupId(ws, "X", "Y", CStr(Target.Value))
        Case "FORM_Relationship"
            Select Case code
                Case "_stype": RefreshRelForm ws, 1
                Case "_rel": RefreshRelForm ws, 2
                Case "_ttype": RefreshRelForm ws, 3
            End Select
        Case "FORM_Type"
            If code = "_kind" Then ShowTypeForm CStr(Target.Value), ""
            If code = "_existing" Then ShowTypeForm CStr(FieldValue(ws, "_kind")), CStr(Target.Value)
            If code = "_copyfrom" Then CopyMappingsInto ws, CStr(Target.Value)
        Case "FORM_Rule"
            If code = "_attr" Then ShowRuleForm CStr(Target.Value), ""
            If code = "_existing" Then ShowRuleForm CStr(FieldValue(ws, "_attr")), LookupId(ws, "X", "Y", CStr(Target.Value))
            If code = "RuleType" Then RuleHint ws
    End Select
End Sub

' =====================================================================
' FORM_Entity
' =====================================================================
Public Sub ShowEntityFormBlank()
    ShowEntityForm "", ""
End Sub

Public Sub ShowEntityForm(Optional ByVal typeCode As String = "", Optional ByVal entityID As String = "")
    modCore.Heal
    Dim ws As Worksheet, labels As New Collection, ids As New Collection, r As Long, cur As Object
    Dim a As Variant, row As Long
    modMeta.LoadMeta
    ThisWorkbook.Activate
    Set ws = FormWS("FORM_Entity")
    modCore.BeginBusy
    On Error GoTo Fail
    ws.Activate
    ResetForm ws, "Add / Edit Entity", "Pick a type, then pick an existing entity to edit or leave it blank to add a new one. * = required."
    Field ws, 4, "Entity type *", "_type", typeCode
    ListDropdown ws, "_type", "TypeCodes"
    If typeCode <> "" And modMeta.TypeExists(typeCode) Then
        EntitiesOfType typeCode, labels, ids
        Dim curLabel As String
        If entityID <> "" Then
            row = modData.FindEntityRow(entityID)
            If row > 0 Then
                Set cur = modData.ReadEntityRow(row)
                curLabel = EntityLabel(CStr(cur("DisplayID")), CStr(cur("Name")))
            Else
                entityID = ""
            End If
        End If
        Field ws, 5, "Existing entity (blank = new)", "_entity", curLabel
        HelperList ws, "_entity", "X", labels, ids, "Y"
        Field ws, 6, "Display ID", "_display", IIf(cur Is Nothing, "(assigned on save)", Nz2(cur, "DisplayID")), True
        ws.Range("Z1").Value = entityID
        r = FIRST_FIELD_ROW
        Field ws, r, "Name *", "Name", Nz2(cur, "Name"): r = r + 1
        Field ws, r, "Description", "Description", Nz2(cur, "Description"): r = r + 1
        Field ws, r, "Status", "Status", IIf(cur Is Nothing, "Active", Nz2(cur, "Status")): ListDropdown ws, "Status", "Status": r = r + 1
        Field ws, r, "Repository area", "RepositoryArea", IIf(cur Is Nothing, "Architecture Landscape", Nz2(cur, "RepositoryArea"))
        ListDropdown ws, "RepositoryArea", "RepositoryArea": r = r + 1
        Dim lastSection As String, sec As String
        For Each a In modMeta.AttrsForType(typeCode)
            sec = modMeta.MapSection(typeCode, CStr(a))
            If sec <> lastSection And sec <> "" Then
                r = r + 1
                ws.Cells(r, 1).Value = sec
                ws.Cells(r, 1).Font.Bold = True: ws.Cells(r, 1).Font.Color = RGB(31, 56, 100)
                r = r + 1
                lastSection = sec
            End If
            Field ws, r, modMeta.AttrName(CStr(a)) & IIf(modMeta.IsMandatory(typeCode, CStr(a)), " *", ""), CStr(a), _
                  Nz2(cur, CStr(a))
            ws.Cells(r, 2).NumberFormat = modGenerator.NumberFormatFor(CStr(a))
            If Not cur Is Nothing Then ws.Cells(r, 2).Value = modCore.SafeFor(Nz2(cur, CStr(a)), CStr(ws.Cells(r, 2).NumberFormat))
            modGenerator.ApplyValidation ws.Cells(r, 2), CStr(a)
            If modMeta.gAttrs(a)("DataType") = "EntityRef" Then ws.Cells(r, 4).Value = "Enter the Display ID, Entity ID or exact name"
            r = r + 1
        Next a
        AddButton ws, "Save entity", MacroRef("SaveEntityForm"), ws.Range("F4"), 150
        AddButton ws, "New " & modMeta.TypeName_(typeCode), MacroRef("NewEntityOfType"), ws.Range("F6"), 150
        If entityID <> "" Then
            AddButton ws, "Retire entity", MacroRef("RetireEntityForm"), ws.Range("F8"), 150
            AddButton ws, "Add relationship", MacroRef("RelFromEntityForm"), ws.Range("F10"), 150
        End If
    End If
    AddButton ws, "Back to Home", MacroRef("GoHome"), ws.Range("F12"), 150
    FinishForm ws
    modCore.EndBusy
    If typeCode <> "" Then ws.Cells(FIRST_FIELD_ROW, 2).Select Else ws.Range("B4").Select
    Exit Sub
Fail:
    modCore.ResetBusy
    MsgBox "Could not open the entity form: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

Private Function Nz2(d As Object, ByVal k As String) As Variant
    If d Is Nothing Then Nz2 = Empty: Exit Function
    If d.Exists(k) Then Nz2 = d(k)
End Function

Public Sub NewEntityOfType()
    ShowEntityForm CStr(FieldValue(FormWS("FORM_Entity"), "_type")), ""
End Sub

Public Sub SaveEntityForm()
    Dim ws As Worksheet, typeCode As String, entityID As String, values As Object, r As Long, code As String
    Dim issues As New Collection, id As String, it As Variant, p() As String, ok As Boolean
    Set ws = FormWS("FORM_Entity")
    typeCode = CStr(FieldValue(ws, "_type"))
    entityID = CStr(ws.Range("Z1").Value)
    If Not modMeta.TypeExists(typeCode) Then FormMessage ws, "Pick an entity type first.", True: Exit Sub
    Set values = modCore.NewDict()
    For r = FIRST_FIELD_ROW To 600
        code = CStr(ws.Cells(r, 3).Value)
        If code <> "" And Left$(code, 1) <> "_" Then values(code) = modData.ToCellValue(code, ws.Cells(r, 2).Value)
    Next r
    modCore.UnprotectSheet ws
    ClearFieldMsgs ws
    If entityID = "" Then
        id = modData.CreateEntity(typeCode, values, "Form", "", issues)
        ok = (id <> "")
    Else
        Dim blocked As Boolean
        modData.UpdateEntity entityID, values, "Form", "", issues, blocked
        ok = Not blocked
        id = entityID
    End If
    For Each it In issues
        p = Split(it, "|", 4)
        FieldMsg ws, p(2), p(3), p(0)
    Next it
    modCore.ProtectSheet ws
    If Not ok Then
        FormMessage ws, "Not saved - fix the fields marked in red.", True
        Exit Sub
    End If
    Dim touched As Object
    Set touched = modCore.NewDict()
    touched(typeCode) = True
    modPaint.PaintTypes touched
    ShowEntityForm typeCode, id
    FormMessage ws, "Saved " & modData.MasterSheet().Cells(modData.FindEntityRow(id), modData.MasterCols()("DisplayID")).Value & _
                    IIf(issues.Count > 0, " with warnings - see column D.", "."), False
    For Each it In issues
        p = Split(it, "|", 4)
        If p(0) = "Warning" Then FieldMsg ws, p(2), p(3), p(0)
    Next it
End Sub

Public Sub RetireEntityForm()
    Dim ws As Worksheet, id As String, typeCode As String
    Set ws = FormWS("FORM_Entity")
    id = CStr(ws.Range("Z1").Value)
    If id = "" Then Exit Sub
    If MsgBox("Retire this entity? It stays in the repository with Status = Retired.", vbOKCancel + vbQuestion, _
              "EA Workbench") = vbCancel Then Exit Sub
    typeCode = CStr(FieldValue(ws, "_type"))
    modData.RetireEntity id, "Form"
    Dim touched As Object
    Set touched = modCore.NewDict(): touched(typeCode) = True
    modPaint.PaintTypes touched
    ShowEntityForm typeCode, id
    FormMessage ws, "Entity retired."
End Sub

Public Sub RelFromEntityForm()
    ShowRelForm CStr(FormWS("FORM_Entity").Range("Z1").Value)
End Sub

' Called from EA_Entities: edit the entity on the active row
Public Sub EditSelectedEntity()
    Dim ws As Worksheet, r As Long, id As String, t As String
    Set ws = ActiveSheet
    If Not ws.Name Like "T_*" Then MsgBox "Select a row on a type sheet first.", vbInformation, "EA Workbench": Exit Sub
    r = ActiveCell.Row
    t = modEvents.TypeOfSheet(ws.Name)
    id = CStr(ws.Cells(r, 1).Value)
    If CStr(ws.Cells(r, 2).Value) = "R" Then
        MsgBox "That row is a relationship. Select the entity row above it.", vbInformation, "EA Workbench"
        Exit Sub
    End If
    ShowEntityForm t, id
End Sub

' Right-click on a relationship row of a type sheet: change its label or retire it
Public Sub EditSelectedRelationship()
    Dim ws As Worksheet, relID As String, d As Variant, m As Variant, lo As ListObject, cur As String, ans As VbMsgBoxResult
    Dim upd As Object, s As String, types As Object
    modCore.Heal
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    If ws.Name Like "T_*" Then
        If CStr(ws.Cells(ActiveCell.Row, 2).Value) = "R" Then relID = CStr(ws.Cells(ActiveCell.Row, 1).Value)
    End If
    If relID = "" Then
        MsgBox "Click + beside an entity, then right-click one of its relationship rows.", vbInformation, "EA Workbench"
        Exit Sub
    End If
    Set lo = modData.RelLO()
    m = Application.Match(relID, lo.ListColumns(1).DataBodyRange, 0)
    If IsError(m) Then Exit Sub
    d = lo.ListRows(CLng(m)).Range.Value
    s = d(1, 5) & " " & modMeta.RelVerb(CStr(d(1, 2))) & " " & d(1, 7)
    cur = CStr(modCore.Nz(d(1, 9), ""))
    ans = MsgBox(s & IIf(cur <> "", " (" & cur & ")", "") & vbLf & vbLf & _
                 "Yes = change the label" & vbLf & "No = retire this relationship" & vbLf & "Cancel = leave it", _
                 vbYesNoCancel + vbQuestion, "EA Workbench")
    If ans = vbCancel Then Exit Sub
    Set upd = modCore.NewDict()
    If ans = vbYes Then
        s = InputBox("New label (blank removes it):", "EA Workbench", cur)
        If StrPtr(s) = 0 Then Exit Sub
        upd("Label") = s
    Else
        upd("Status") = "Retired"
    End If
    If Not modData.UpdateRelationship(relID, upd, "Form", modCore.NewGUID()) Then
        MsgBox "The relationship could not be changed.", vbExclamation, "EA Workbench"
        Exit Sub
    End If
    Set types = modCore.NewDict()
    types(modData.EntityTypeOf(CStr(d(1, 3)))) = True
    types(modData.EntityTypeOf(CStr(d(1, 4)))) = True
    modPaint.PaintTypes types
    modSync.SaveWithStamp ThisWorkbook
End Sub

Public Sub AddRelationshipForSelected()
    Dim ws As Worksheet, id As String
    Set ws = ActiveSheet
    If ws.Name Like "T_*" Then
        If CStr(ws.Cells(ActiveCell.Row, 2).Value) = "E" Then id = CStr(ws.Cells(ActiveCell.Row, 1).Value)
    End If
    ShowRelForm id
End Sub

' =====================================================================
' FORM_Relationship
' =====================================================================
Public Sub ShowRelFormBlank()
    ShowRelForm ""
End Sub

Public Sub ShowRelForm(Optional ByVal sourceID As String = "")
    modCore.Heal
    Dim ws As Worksheet, st As String, srcLabel As String, row As Long, cur As Object
    modMeta.LoadMeta
    ThisWorkbook.Activate
    Set ws = FormWS("FORM_Relationship")
    modCore.BeginBusy
    On Error GoTo Fail
    ws.Activate
    ResetForm ws, "Add Relationship", "Only relationship types and targets allowed by the metamodel are offered."
    If sourceID <> "" Then
        row = modData.FindEntityRow(sourceID)
        If row > 0 Then
            Set cur = modData.ReadEntityRow(row)
            st = CStr(cur("TypeCode"))
            srcLabel = EntityLabel(CStr(cur("DisplayID")), CStr(cur("Name")))
        End If
    End If
    Field ws, 4, "Source type *", "_stype", st
    ListDropdown ws, "_stype", "TypeCodes"
    Field ws, 5, "Source entity *", "_src", srcLabel
    Field ws, 6, "Relationship *", "_rel"
    Field ws, 7, "Target type *", "_ttype"
    Field ws, 8, "Target entity *", "_tgt"
    Field ws, 9, "Label", "_label"
    Field ws, 10, "Description", "_desc"
    RefreshRelForm ws, 1, True
    If srcLabel <> "" Then SetFieldValue ws, "_src", srcLabel
    AddButton ws, "Save relationship", MacroRef("SaveRelForm"), ws.Range("F4"), 150
    AddButton ws, "Clear", MacroRef("ShowRelFormBlank"), ws.Range("F6"), 150
    AddButton ws, "Back to Home", MacroRef("GoHome"), ws.Range("F8"), 150
    FinishForm ws
    modCore.EndBusy
    Exit Sub
Fail:
    modCore.ResetBusy
    MsgBox "Could not open the relationship form: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

' level 1 = source type changed, 2 = relationship changed, 3 = target type changed
Private Sub RefreshRelForm(ws As Worksheet, ByVal level As Long, Optional ByVal painting As Boolean = False)
    Dim st As String, rc As String, tt As String, labels As Collection, ids As Collection, c As Collection
    If Not painting Then modCore.BeginBusy
    On Error GoTo Done
    modCore.UnprotectSheet ws
    st = CStr(FieldValue(ws, "_stype"))
    If level <= 1 Then
        Set labels = New Collection: Set ids = New Collection
        If modMeta.TypeExists(st) Then EntitiesOfType st, labels, ids
        HelperList ws, "_src", "X", labels, ids, "Y"
        If Not painting Then SetFieldValue ws, "_src", ""
        If modMeta.TypeExists(st) Then Set c = modMeta.RelsFromType(st) Else Set c = New Collection
        HelperList ws, "_rel", "Z", c
        SetFieldValue ws, "_rel", "": SetFieldValue ws, "_ttype", "": SetFieldValue ws, "_tgt", ""
        HelperList ws, "_ttype", "AA", New Collection
        HelperList ws, "_tgt", "AB", New Collection, New Collection, "AC"
    End If
    rc = CStr(FieldValue(ws, "_rel"))
    If level = 2 Then
        If modMeta.TypeExists(st) And rc <> "" Then Set c = modMeta.TargetTypesFor(st, rc) Else Set c = New Collection
        HelperList ws, "_ttype", "AA", c
        SetFieldValue ws, "_ttype", "": SetFieldValue ws, "_tgt", ""
        If c.Count = 1 Then SetFieldValue ws, "_ttype", c(1): level = 3
    End If
    tt = CStr(FieldValue(ws, "_ttype"))
    If level = 3 Then
        Set labels = New Collection: Set ids = New Collection
        If modMeta.TypeExists(tt) Then EntitiesOfType tt, labels, ids
        HelperList ws, "_tgt", "AB", labels, ids, "AC"
        SetFieldValue ws, "_tgt", ""
    End If
Done:
    modCore.ProtectSheet ws
    If Not painting Then modCore.EndBusy
End Sub

Public Sub SaveRelForm()
    Dim ws As Worksheet, src As String, tgt As String, rc As String, issues As New Collection, id As String
    Set ws = FormWS("FORM_Relationship")
    modCore.UnprotectSheet ws
    ClearFieldMsgs ws
    src = LookupId(ws, "X", "Y", CStr(FieldValue(ws, "_src")))
    tgt = LookupId(ws, "AB", "AC", CStr(FieldValue(ws, "_tgt")))
    rc = CStr(FieldValue(ws, "_rel"))
    If src = "" Then FieldMsg ws, "_src", "Pick a source entity", "Error"
    If rc = "" Then FieldMsg ws, "_rel", "Pick a relationship", "Error"
    If tgt = "" Then FieldMsg ws, "_tgt", "Pick a target entity", "Error"
    If src = "" Or tgt = "" Or rc = "" Then
        FormMessage ws, "Not saved - complete the fields marked in red.", True
        modCore.ProtectSheet ws
        Exit Sub
    End If
    id = modData.CreateRelationship(rc, src, tgt, CStr(FieldValue(ws, "_label")), CStr(FieldValue(ws, "_desc")), _
                                    "Form", "", issues)
    If id = "" Then
        FormMessage ws, "Not saved: " & Replace(modRules.IssueText(issues), vbLf, "; "), True
        modCore.ProtectSheet ws
        Exit Sub
    End If
    Dim touched As Object
    Set touched = modCore.NewDict()
    touched(modData.EntityTypeOf(src)) = True: touched(modData.EntityTypeOf(tgt)) = True
    modPaint.PaintTypes touched
    ThisWorkbook.Activate
    ws.Activate
    FormMessage ws, "Saved: " & modData.EntityNameOf(src) & " " & modMeta.RelVerb(rc) & " " & modData.EntityNameOf(tgt) & "."
    SetFieldValue ws, "_tgt", "": SetFieldValue ws, "_label", "": SetFieldValue ws, "_desc", ""
    modCore.ProtectSheet ws
End Sub

' =====================================================================
' FORM_Type  (entity, attribute and relationship types)
' =====================================================================
Public Sub ShowTypeFormBlank()
    ShowTypeForm "Entity type", ""
End Sub

Public Sub ShowTypeForm(Optional ByVal kind As String = "Entity type", Optional ByVal code As String = "")
    modCore.Heal
    Dim ws As Worksheet, c As New Collection, k As Variant
    modMeta.LoadMeta
    ThisWorkbook.Activate
    Set ws = FormWS("FORM_Type")
    modCore.BeginBusy
    On Error GoTo Fail
    ws.Activate
    ResetForm ws, "Add / Edit Type", "Changing the metamodel regenerates DATA_Master columns and the EA_Entities type sheets. Codes cannot be changed once saved."
    Field ws, 4, "Kind *", "_kind", kind
    ListDropdown ws, "_kind", "TypeKinds"
    Select Case kind
        Case "Entity type": For Each k In modMeta.gTypes.Keys: c.Add CStr(k): Next k
        Case "Attribute type"
            For Each k In modMeta.gAttrs.Keys
                If Not modMeta.gAttrs(k)("IsSystem") Then c.Add CStr(k)
            Next k
        Case "Relationship type": For Each k In modMeta.gRelTypes.Keys: c.Add CStr(k): Next k
    End Select
    Field ws, 5, "Existing (blank = new)", "_existing", code
    HelperList ws, "_existing", "X", c
    Select Case kind
        Case "Entity type": PaintEntityTypeFields ws, code
        Case "Attribute type": PaintAttrTypeFields ws, code
        Case "Relationship type": PaintRelTypeFields ws, code
    End Select
    AddButton ws, "Save " & LCase$(kind), MacroRef("SaveTypeForm"), ws.Range("F4"), 170
    AddButton ws, "New " & LCase$(kind), MacroRef("NewTypeOfKind"), ws.Range("F6"), 170
    AddButton ws, "Back to Home", MacroRef("GoHome"), ws.Range("F8"), 170
    FinishForm ws
    modCore.EndBusy
    Exit Sub
Fail:
    modCore.ResetBusy
    MsgBox "Could not open the type form: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

Public Sub NewTypeOfKind()
    ShowTypeForm CStr(FieldValue(FormWS("FORM_Type"), "_kind")), ""
End Sub

Private Function GridStart() As Long
    GridStart = 22
End Function

Private Sub GridHeader(ws As Worksheet, ByVal title As String, hdrs As Variant)
    Dim i As Long
    ws.Cells(GridStart() - 2, 1).Value = title
    ws.Cells(GridStart() - 2, 1).Font.Bold = True
    ws.Cells(GridStart() - 2, 1).Font.Color = RGB(31, 56, 100)
    For i = 0 To UBound(hdrs)
        ws.Cells(GridStart() - 1, i + 1).Value = hdrs(i)
    Next i
    With ws.Range(ws.Cells(GridStart() - 1, 1), ws.Cells(GridStart() - 1, UBound(hdrs) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    ws.Cells(GridStart() - 1, 3).Value = "_grid"
End Sub

Private Sub GridInput(ws As Worksheet, ByVal r As Long, ByVal fromCol As Long, ByVal toCol As Long)
    With ws.Range(ws.Cells(r, fromCol), ws.Cells(r, toCol))
        .Locked = False
        .Borders.Color = RGB(191, 191, 191)
    End With
End Sub

Private Sub PaintEntityTypeFields(ws As Worksheet, ByVal code As String)
    Dim t As Object, r As Long, a As Variant, isNew As Boolean
    isNew = (code = "")
    If Not isNew Then Set t = modMeta.gTypes(code)
    r = FIRST_FIELD_ROW
    Field ws, r, "Type code *", "TypeCode", code, Not isNew: r = r + 1
    ws.Cells(r - 1, 4).Value = "Letters and digits, no spaces (e.g. BusinessCapabilityMap)"
    Field ws, r, "Name *", "Name", IIf(isNew, "", Nz3(t, "Name")): r = r + 1
    Field ws, r, "Domain *", "Domain", IIf(isNew, "", Nz3(t, "Domain")): ListDropdown ws, "Domain", "Domain": r = r + 1
    Field ws, r, "Aspect", "Aspect", IIf(isNew, "", Nz3(t, "Aspect")): ListDropdown ws, "Aspect", "Aspects": r = r + 1
    Field ws, r, "Display ID prefix *", "IDPrefix", IIf(isNew, "", Nz3(t, "Prefix")), Not isNew: r = r + 1
    ws.Cells(r - 1, 4).Value = "2 to 5 capital letters, unique (e.g. BCM)"
    Field ws, r, "Description", "Description", IIf(isNew, "", TypeDescription(code)): r = r + 1
    Field ws, r, "Status", "Status", IIf(isNew, "Active", Nz3(t, "Status")): ListDropdown ws, "Status", "Status": r = r + 1
    If isNew Then
        Field ws, r, "Copy attribute mappings from", "_copyfrom": ListDropdown ws, "_copyfrom", "TypeCodes": r = r + 1
    End If

    GridHeader ws, "Attributes for this type", Array("Attribute", "Map? (Y/N)", "", "Mandatory? (Y/N)", "Section")
    r = GridStart()
    For Each a In modMeta.gAttrs.Keys
        If Not modMeta.gAttrs(a)("IsSystem") And modMeta.gAttrs(a)("Status") = "Active" Then
            ws.Cells(r, 1).Value = modMeta.gAttrs(a)("Name") & "  [" & a & "]"
            ws.Cells(r, 3).Value = "attr:" & a
            If Not isNew And modMeta.IsMapped(code, CStr(a)) Then
                ws.Cells(r, 2).Value = "Y"
                ws.Cells(r, 4).Value = IIf(modMeta.IsMandatory(code, CStr(a)), "Y", "N")
                ws.Cells(r, 5).Value = modMeta.MapSection(code, CStr(a))
            Else
                ws.Cells(r, 2).Value = "N": ws.Cells(r, 4).Value = "N"
            End If
            GridInput ws, r, 2, 2: GridInput ws, r, 4, 5
            modGenerator.ApplyListValidation ws.Cells(r, 2), "YesNo"
            modGenerator.ApplyListValidation ws.Cells(r, 4), "YesNo"
            r = r + 1
        End If
    Next a
End Sub

Private Function TypeDescription(ByVal code As String) As String
    Dim lo As ListObject, r As Long
    Set lo = modCore.Tbl("MM_EntityTypes", "tblEntityTypes")
    r = FindTableRow(lo, "TypeCode", code)
    If r > 0 Then TypeDescription = CStr(lo.ListColumns("Description").DataBodyRange.Cells(r, 1).Value)
End Function

Private Function Nz3(d As Object, ByVal k As String) As Variant
    If d Is Nothing Then Exit Function
    If d.Exists(k) Then Nz3 = d(k)
End Function

Private Sub CopyMappingsInto(ws As Worksheet, ByVal fromType As String)
    Dim r As Long, a As String
    If Not modMeta.TypeExists(fromType) Then Exit Sub
    modCore.BeginBusy
    modCore.UnprotectSheet ws
    For r = GridStart() To 600
        If Left$(CStr(ws.Cells(r, 3).Value), 5) = "attr:" Then
            a = Mid$(CStr(ws.Cells(r, 3).Value), 6)
            If modMeta.IsMapped(fromType, a) Then
                ws.Cells(r, 2).Value = "Y"
                ws.Cells(r, 4).Value = IIf(modMeta.IsMandatory(fromType, a), "Y", "N")
                ws.Cells(r, 5).Value = modMeta.MapSection(fromType, a)
            End If
        End If
    Next r
    modCore.ProtectSheet ws
    modCore.EndBusy
End Sub

Private Sub PaintAttrTypeFields(ws As Worksheet, ByVal code As String)
    Dim a As Object, r As Long, t As Variant, isNew As Boolean
    isNew = (code = "")
    If Not isNew Then Set a = modMeta.gAttrs(code)
    r = FIRST_FIELD_ROW
    Field ws, r, "Attribute code *", "AttrCode", code, Not isNew: r = r + 1
    ws.Cells(r - 1, 4).Value = "Letters and digits, no spaces (e.g. RecoveryTimeObjective)"
    Field ws, r, "Name *", "Name", Nz3(a, "Name"): r = r + 1
    Field ws, r, "Data type *", "DataType", Nz3(a, "DataType"), Not isNew: ListDropdown ws, "DataType", "DataType": r = r + 1
    Field ws, r, "List name (for List)", "ListName", Nz3(a, "ListName"): ListDropdown ws, "ListName", "ListNames": r = r + 1
    Field ws, r, "New list: name and values", "_newlist": r = r + 1
    ws.Cells(r - 1, 4).Value = "Optional. Format  ListName: value 1; value 2; value 3"
    Field ws, r, "Unit", "Unit", Nz3(a, "Unit"): r = r + 1
    Field ws, r, "Default value", "DefaultValue", Nz3(a, "Default"): r = r + 1
    Field ws, r, "Description", "Description", Nz3(a, "Desc"): r = r + 1
    Field ws, r, "Status", "Status", IIf(isNew, "Active", Nz3(a, "Status")): ListDropdown ws, "Status", "Status": r = r + 1

    GridHeader ws, "Entity types that carry this attribute", Array("Entity type", "Map? (Y/N)", "", "Mandatory? (Y/N)", "Section")
    r = GridStart()
    For Each t In modPaint.TypesInSheetOrder()
        ws.Cells(r, 1).Value = modMeta.gTypes(t)("Name") & "  [" & t & "]"
        ws.Cells(r, 3).Value = "type:" & t
        If Not isNew And modMeta.IsMapped(CStr(t), code) Then
            ws.Cells(r, 2).Value = "Y"
            ws.Cells(r, 4).Value = IIf(modMeta.IsMandatory(CStr(t), code), "Y", "N")
            ws.Cells(r, 5).Value = modMeta.MapSection(CStr(t), code)
        Else
            ws.Cells(r, 2).Value = "N": ws.Cells(r, 4).Value = "N"
        End If
        GridInput ws, r, 2, 2: GridInput ws, r, 4, 5
        modGenerator.ApplyListValidation ws.Cells(r, 2), "YesNo"
        modGenerator.ApplyListValidation ws.Cells(r, 4), "YesNo"
        r = r + 1
    Next t
End Sub

Private Sub PaintRelTypeFields(ws As Worksheet, ByVal code As String)
    Dim rt As Object, r As Long, isNew As Boolean, lo As ListObject, d As Variant, i As Long, n As Long
    isNew = (code = "")
    If Not isNew Then Set rt = modMeta.gRelTypes(code)
    r = FIRST_FIELD_ROW
    Field ws, r, "Relationship code *", "RelCode", code, Not isNew: r = r + 1
    Field ws, r, "Verb (shown on sheets) *", "Name", Nz3(rt, "Name"): r = r + 1
    ws.Cells(r - 1, 4).Value = "e.g. ""serves"", ""realizes"""
    Field ws, r, "Category *", "Category", Nz3(rt, "Category"): ListDropdown ws, "Category", "RelCategory": r = r + 1
    Field ws, r, "Directed (Y/N)", "Directed", IIf(isNew, "Y", Nz3(rt, "Directed")): ListDropdown ws, "Directed", "YesNo": r = r + 1
    Field ws, r, "Description", "Description", "": r = r + 1
    Field ws, r, "Status", "Status", IIf(isNew, "Active", Nz3(rt, "Status")): ListDropdown ws, "Status", "Status": r = r + 1

    GridHeader ws, "Allowed source and target types  (TypeCode, * = any, SAME = same as source, G:<group>)", _
               Array("Source type", "Target type", "", "Min", "Max", "Status")
    r = GridStart()
    Set lo = modCore.Tbl("MM_EntityRelationships", "tblEntityRels")
    d = modCore.TableData(lo)
    If Not isNew And IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, lo.ListColumns("RelCode").Index)) = code Then
                ws.Cells(r, 1).Value = "'" & d(i, lo.ListColumns("SourceType").Index)
                ws.Cells(r, 2).Value = "'" & d(i, lo.ListColumns("TargetType").Index)
                ws.Cells(r, 3).Value = "allow:" & d(i, lo.ListColumns("MapID").Index)
                ws.Cells(r, 4).Value = d(i, lo.ListColumns("MinCard").Index)
                ws.Cells(r, 5).Value = d(i, lo.ListColumns("MaxCard").Index)
                ws.Cells(r, 6).Value = d(i, lo.ListColumns("Status").Index)
                GridInput ws, r, 1, 2: GridInput ws, r, 4, 6
                modGenerator.ApplyListValidation ws.Cells(r, 6), "Status"
                r = r + 1: n = n + 1
            End If
        Next i
    End If
    For i = 1 To 10
        ws.Cells(r, 3).Value = "allow:new"
        ws.Cells(r, 6).Value = "Active"
        GridInput ws, r, 1, 2: GridInput ws, r, 4, 6
        modGenerator.ApplyListValidation ws.Cells(r, 6), "Status"
        r = r + 1
    Next i
End Sub

Public Sub SaveTypeForm()
    Dim ws As Worksheet, kind As String, ok As Boolean
    Set ws = FormWS("FORM_Type")
    kind = CStr(FieldValue(ws, "_kind"))
    modCore.UnprotectSheet ws
    ClearFieldMsgs ws
    Select Case kind
        Case "Entity type": ok = SaveEntityType(ws)
        Case "Attribute type": ok = SaveAttrType(ws)
        Case "Relationship type": ok = SaveRelType(ws)
    End Select
    modCore.ProtectSheet ws
End Sub

Private Function ValidCode(ByVal s As String, ByVal maxLen As Long) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Or Len(s) > maxLen Then Exit Function
    If Not Mid$(s, 1, 1) Like "[A-Za-z]" Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not ch Like "[A-Za-z0-9]" Then Exit Function
    Next i
    ValidCode = True
End Function

Private Function SaveEntityType(ws As Worksheet) As Boolean
    Dim code As String, nm As String, dom As String, pfx As String, isNew As Boolean, lo As ListObject, r As Long
    code = Trim$(CStr(FieldValue(ws, "TypeCode")))
    nm = Trim$(CStr(FieldValue(ws, "Name")))
    dom = Trim$(CStr(FieldValue(ws, "Domain")))
    pfx = UCase$(Trim$(CStr(FieldValue(ws, "IDPrefix"))))
    isNew = (CStr(FieldValue(ws, "_existing")) = "")
    Dim bad As Boolean
    If Not ValidCode(code, 28) Then FieldMsg ws, "TypeCode", "Letters and digits only, up to 28 characters", "Error": bad = True
    If isNew And modMeta.TypeExists(code) Then FieldMsg ws, "TypeCode", "That code already exists", "Error": bad = True
    If nm = "" Then FieldMsg ws, "Name", "Name is required", "Error": bad = True
    If dom = "" Then FieldMsg ws, "Domain", "Pick a domain", "Error": bad = True
    If Not pfx Like "[A-Z][A-Z]*" Or Len(pfx) > 5 Then FieldMsg ws, "IDPrefix", "2 to 5 capital letters", "Error": bad = True
    If isNew And PrefixInUse(pfx) Then FieldMsg ws, "IDPrefix", "Prefix already used by another type", "Error": bad = True
    If bad Then FormMessage ws, "Not saved - fix the fields marked in red.", True: Exit Function

    Dim batch As String, ts As String, u As String
    If Not modHistory.AutoSnapshot("Before metamodel change") Then FormMessage ws, "Not saved.", True: Exit Function
    batch = modCore.NewGUID(): ts = modCore.UtcNow(): u = modCore.CurrentUser()
    modCore.BeginBusy
    On Error GoTo Fail
    Set lo = modCore.Tbl("MM_EntityTypes", "tblEntityTypes")
    modCore.UnprotectSheet lo.Parent
    If isNew Then
        modCore.AppendRow lo, Array(modCore.NewGUID(), code, nm, dom, CStr(FieldValue(ws, "Aspect")), _
            CStr(FieldValue(ws, "Description")), Left$("T_" & code, 31), pfx, "Entity:" & code, _
            CStr(FieldValue(ws, "Status")), ts, u)
        EnsureStyle code, dom
        modData.WriteHistory "", "Create", code, "TypeCode", "", nm, "Type form", batch, modCore.SH_HMETA
    Else
        r = FindTableRow(lo, "TypeCode", code)
        SetTableValue lo, r, "Name", nm, "TypeCode", batch
        SetTableValue lo, r, "Domain", dom, "TypeCode", batch
        SetTableValue lo, r, "Aspect", CStr(FieldValue(ws, "Aspect")), "TypeCode", batch
        SetTableValue lo, r, "Description", CStr(FieldValue(ws, "Description")), "TypeCode", batch
        SetTableValue lo, r, "Status", CStr(FieldValue(ws, "Status")), "TypeCode", batch
        SetTableValue lo, r, "Modified", ts, "", ""
        SetTableValue lo, r, "ModifiedBy", u, "", ""
    End If
    modCore.ProtectSheet lo.Parent
    SaveGridMappings ws, code, "attr", batch
    modCore.EndBusy
    ApplyAfterMetaChange
    ShowTypeForm "Entity type", code
    FormMessage FormWS("FORM_Type"), "Saved entity type " & code & ". DATA_Master and EA_Entities have been regenerated."
    SaveEntityType = True
    Exit Function
Fail:
    modCore.ResetBusy
    FormMessage ws, "Not saved: " & Err.Description, True
End Function

Private Function PrefixInUse(ByVal pfx As String) As Boolean
    Dim t As Variant
    For Each t In modMeta.gTypes.Keys
        If UCase$(modMeta.gTypes(t)("Prefix")) = pfx Then PrefixInUse = True: Exit Function
    Next t
End Function

Private Sub EnsureStyle(ByVal code As String, ByVal dom As String)
    Dim lo As ListObject, fillHex As String
    Set lo = modCore.Tbl("MM_Styles", "tblStyles")
    If FindTableRow(lo, "StyleKey", "Entity:" & code) > 0 Then Exit Sub
    Select Case dom
        Case "Business": fillHex = "FFFFB5"
        Case "Application": fillHex = "B5FFFF"
        Case "Technology": fillHex = "C9E7B7"
        Case "Strategy": fillHex = "F5DEAA"
        Case "Motivation": fillHex = "CCCCFF"
        Case "Implementation & Migration": fillHex = "FFE0E0"
        Case Else: fillHex = "EEEEEE"
    End Select
    modCore.UnprotectSheet lo.Parent
    modCore.AppendRow lo, Array("Entity:" & code, "Entity", fillHex, modDiagram.DefaultEntityStyle())
    modCore.ProtectSheet lo.Parent
End Sub

' Grid rows carry "attr:<code>" (entity type form) or "type:<code>" (attribute form) in column C
Private Sub SaveGridMappings(ws As Worksheet, ByVal fixedCode As String, ByVal prefix As String, ByVal batch As String)
    Dim lo As ListObject, r As Long, other As String, tc As String, ac As String, mapOn As Boolean
    Dim row As Long, order As Long, ts As String, u As String
    Set lo = modCore.Tbl("MM_EntityAttributes", "tblEntityAttributes")
    modCore.UnprotectSheet lo.Parent
    ts = modCore.UtcNow(): u = modCore.CurrentUser()
    For r = GridStart() To 600
        If Left$(CStr(ws.Cells(r, 3).Value), Len(prefix) + 1) = prefix & ":" Then
            other = Mid$(CStr(ws.Cells(r, 3).Value), Len(prefix) + 2)
            If prefix = "attr" Then tc = fixedCode: ac = other Else tc = other: ac = fixedCode
            mapOn = (UCase$(CStr(ws.Cells(r, 2).Value)) = "Y")
            row = FindMapRow(lo, tc, ac)
            If mapOn Then
                If row = 0 Then
                    order = NextOrder(lo, tc)
                    modCore.AppendRow lo, Array(modCore.NewGUID(), tc, ac, IIf(UCase$(CStr(ws.Cells(r, 4).Value)) = "Y", "Y", "N"), _
                                                order, CStr(ws.Cells(r, 5).Value), "", "Active", ts, u)
                    modData.WriteHistory "", "Create", tc & "|" & ac, "Mapping", "", "mapped", "Type form", batch, modCore.SH_HMETA
                Else
                    SetTableValue lo, row, "Mandatory", IIf(UCase$(CStr(ws.Cells(r, 4).Value)) = "Y", "Y", "N"), "MapID", batch
                    SetTableValue lo, row, "Section", CStr(ws.Cells(r, 5).Value), "MapID", batch
                    SetTableValue lo, row, "Status", "Active", "MapID", batch
                End If
            ElseIf row > 0 Then
                If CStr(lo.ListColumns("Status").DataBodyRange.Cells(row, 1).Value) = "Active" Then
                    SetTableValue lo, row, "Status", "Retired", "MapID", batch
                    ArchiveUnmappedValues tc, ac, batch
                End If
            End If
        End If
    Next r
    modCore.ProtectSheet lo.Parent
End Sub

Private Function FindMapRow(lo As ListObject, ByVal tc As String, ByVal ac As String) As Long
    Dim d As Variant, i As Long, cT As Long, cA As Long
    d = modCore.TableData(lo)
    If Not IsArray(d) Then Exit Function
    cT = lo.ListColumns("TypeCode").Index: cA = lo.ListColumns("AttrCode").Index
    For i = 1 To UBound(d, 1)
        If CStr(d(i, cT)) = tc And CStr(d(i, cA)) = ac Then FindMapRow = i: Exit Function
    Next i
End Function

Private Function NextOrder(lo As ListObject, ByVal tc As String) As Long
    Dim d As Variant, i As Long, m As Long
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, lo.ListColumns("TypeCode").Index)) = tc Then
                If Val(CStr(d(i, lo.ListColumns("DisplayOrder").Index))) > m Then m = Val(CStr(d(i, lo.ListColumns("DisplayOrder").Index)))
            End If
        Next i
    End If
    NextOrder = m + 1
End Function

' Removing a mapping: values of that attribute on entities of the type go to history, then are cleared.
Private Sub ArchiveUnmappedValues(ByVal tc As String, ByVal ac As String, ByVal batch As String)
    Dim ws As Worksheet, cols As Object, lo As ListObject, r As Long, last As Long
    Set ws = modData.MasterSheet(): Set cols = modData.MasterCols(): Set lo = modData.MasterLO()
    If Not cols.Exists(ac) Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    last = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1
    modCore.UnprotectSheet ws
    For r = lo.DataBodyRange.Row To last
        If CStr(ws.Cells(r, cols("TypeCode")).Value) = tc Then
            If Not modCore.IsBlankValue(ws.Cells(r, cols(ac)).Value) Then
                modData.WriteHistory tc, "Archive", CStr(ws.Cells(r, cols("EntityID")).Value), ac, _
                                     ws.Cells(r, cols(ac)).Value, "", "Mapping removed", batch
                ws.Cells(r, cols(ac)).ClearContents
                modData.StampRow r, False, cols
            End If
        End If
    Next r
    modCore.ProtectSheet ws
End Sub

Private Function SaveAttrType(ws As Worksheet) As Boolean
    Dim code As String, nm As String, dt As String, ln As String, isNew As Boolean, lo As ListObject, r As Long
    code = Trim$(CStr(FieldValue(ws, "AttrCode")))
    nm = Trim$(CStr(FieldValue(ws, "Name")))
    dt = Trim$(CStr(FieldValue(ws, "DataType")))
    ln = Trim$(CStr(FieldValue(ws, "ListName")))
    isNew = (CStr(FieldValue(ws, "_existing")) = "")
    Dim bad As Boolean, newList As String
    newList = Trim$(CStr(FieldValue(ws, "_newlist")))
    If Not ValidCode(code, 40) Then FieldMsg ws, "AttrCode", "Letters and digits only, up to 40 characters", "Error": bad = True
    If isNew And modMeta.gAttrs.Exists(code) Then FieldMsg ws, "AttrCode", "That code already exists", "Error": bad = True
    If nm = "" Then FieldMsg ws, "Name", "Name is required", "Error": bad = True
    If dt = "" Then FieldMsg ws, "DataType", "Pick a data type", "Error": bad = True
    If newList <> "" Then
        If InStr(newList, ":") = 0 Then
            FieldMsg ws, "_newlist", "Use  ListName: value 1; value 2", "Error": bad = True
        Else
            ln = Trim$(Split(newList, ":")(0))
            If Not ValidCode(ln, 40) Then FieldMsg ws, "_newlist", "List name: letters and digits only", "Error": bad = True
        End If
    End If
    If dt = "List" And ln = "" Then FieldMsg ws, "ListName", "A List attribute needs a list", "Error": bad = True
    If bad Then FormMessage ws, "Not saved - fix the fields marked in red.", True: Exit Function

    Dim batch As String, ts As String, u As String, v As Variant, n As Long
    If Not modHistory.AutoSnapshot("Before metamodel change") Then FormMessage ws, "Not saved.", True: Exit Function
    batch = modCore.NewGUID(): ts = modCore.UtcNow(): u = modCore.CurrentUser()
    modCore.BeginBusy
    On Error GoTo Fail
    If newList <> "" Then
        Set lo = modCore.Tbl("MM_Lists", "tblLists")
        modCore.UnprotectSheet lo.Parent
        For Each v In Split(Mid$(newList, InStr(newList, ":") + 1), ";")
            If Trim$(v) <> "" Then
                n = n + 1
                If Not modMeta.InList(ln, Trim$(v)) Then modCore.AppendRow lo, Array(ln, Trim$(v), n, "", "Active")
            End If
        Next v
        modCore.ProtectSheet lo.Parent
        modData.WriteHistory "", "Create", ln, "List", "", Mid$(newList, InStr(newList, ":") + 1), "Type form", batch, modCore.SH_HMETA
    End If
    Set lo = modCore.Tbl("MM_AttributeTypes", "tblAttributeTypes")
    modCore.UnprotectSheet lo.Parent
    If isNew Then
        modCore.AppendRow lo, Array(modCore.NewGUID(), code, nm, dt, ln, CStr(FieldValue(ws, "Unit")), _
            FieldValue(ws, "DefaultValue"), CStr(FieldValue(ws, "Description")), "N", CStr(FieldValue(ws, "Status")), ts, u)
        modData.WriteHistory "", "Create", code, "AttrCode", "", nm, "Type form", batch, modCore.SH_HMETA
    Else
        r = FindTableRow(lo, "AttrCode", code)
        SetTableValue lo, r, "Name", nm, "AttrCode", batch
        SetTableValue lo, r, "ListName", ln, "AttrCode", batch
        SetTableValue lo, r, "Unit", CStr(FieldValue(ws, "Unit")), "AttrCode", batch
        SetTableValue lo, r, "DefaultValue", FieldValue(ws, "DefaultValue"), "AttrCode", batch
        SetTableValue lo, r, "Description", CStr(FieldValue(ws, "Description")), "AttrCode", batch
        SetTableValue lo, r, "Status", CStr(FieldValue(ws, "Status")), "AttrCode", batch
        SetTableValue lo, r, "Modified", ts, "", ""
        SetTableValue lo, r, "ModifiedBy", u, "", ""
    End If
    modCore.ProtectSheet lo.Parent
    modMeta.LoadMeta True
    SaveGridMappings ws, code, "type", batch
    modCore.EndBusy
    ApplyAfterMetaChange
    ShowTypeForm "Attribute type", code
    FormMessage FormWS("FORM_Type"), "Saved attribute " & code & ". DATA_Master and EA_Entities have been regenerated."
    SaveAttrType = True
    Exit Function
Fail:
    modCore.ResetBusy
    FormMessage ws, "Not saved: " & Err.Description, True
End Function

Private Function SaveRelType(ws As Worksheet) As Boolean
    Dim code As String, nm As String, cat As String, isNew As Boolean, lo As ListObject, r As Long
    code = Trim$(CStr(FieldValue(ws, "RelCode")))
    nm = Trim$(CStr(FieldValue(ws, "Name")))
    cat = Trim$(CStr(FieldValue(ws, "Category")))
    isNew = (CStr(FieldValue(ws, "_existing")) = "")
    Dim bad As Boolean
    If Not ValidCode(code, 40) Then FieldMsg ws, "RelCode", "Letters and digits only", "Error": bad = True
    If isNew And modMeta.gRelTypes.Exists(code) Then FieldMsg ws, "RelCode", "That code already exists", "Error": bad = True
    If nm = "" Then FieldMsg ws, "Name", "A verb is required", "Error": bad = True
    If cat = "" Then FieldMsg ws, "Category", "Pick a category", "Error": bad = True
    ' check tokens in the grid
    Dim gr As Long, tok As Variant
    ws.Range(ws.Cells(GridStart(), 7), ws.Cells(600, 7)).ClearContents
    For gr = GridStart() To 600
        If Left$(CStr(ws.Cells(gr, 3).Value), 6) = "allow:" Then
            For Each tok In Array(CStr(ws.Cells(gr, 1).Value), CStr(ws.Cells(gr, 2).Value))
                If tok <> "" And Not TokenValid(CStr(tok)) Then
                    ws.Cells(gr, 7).Value = "Unknown type or group: " & tok
                    ws.Cells(gr, 7).Font.Color = RGB(192, 0, 0)
                    bad = True
                End If
            Next tok
        End If
    Next gr
    If bad Then FormMessage ws, "Not saved - fix the fields marked in red.", True: Exit Function

    Dim batch As String, ts As String, u As String
    If Not modHistory.AutoSnapshot("Before metamodel change") Then FormMessage ws, "Not saved.", True: Exit Function
    batch = modCore.NewGUID(): ts = modCore.UtcNow(): u = modCore.CurrentUser()
    modCore.BeginBusy
    On Error GoTo Fail
    Set lo = modCore.Tbl("MM_RelationshipTypes", "tblRelTypes")
    modCore.UnprotectSheet lo.Parent
    If isNew Then
        modCore.AppendRow lo, Array(modCore.NewGUID(), code, nm, cat, CStr(FieldValue(ws, "Directed")), "Rel:" & code, _
                                    CStr(FieldValue(ws, "Description")), CStr(FieldValue(ws, "Status")), ts, u)
        modData.WriteHistory "", "Create", code, "RelCode", "", nm, "Type form", batch, modCore.SH_HMETA
    Else
        r = FindTableRow(lo, "RelCode", code)
        SetTableValue lo, r, "Name", nm, "RelCode", batch
        SetTableValue lo, r, "Category", cat, "RelCode", batch
        SetTableValue lo, r, "Directed", CStr(FieldValue(ws, "Directed")), "RelCode", batch
        SetTableValue lo, r, "Status", CStr(FieldValue(ws, "Status")), "RelCode", batch
        SetTableValue lo, r, "Modified", ts, "", ""
        SetTableValue lo, r, "ModifiedBy", u, "", ""
        If CStr(FieldValue(ws, "Description")) <> "" Then SetTableValue lo, r, "Description", CStr(FieldValue(ws, "Description")), "RelCode", batch
    End If
    modCore.ProtectSheet lo.Parent

    Set lo = modCore.Tbl("MM_EntityRelationships", "tblEntityRels")
    modCore.UnprotectSheet lo.Parent
    Dim mapID As String, row As Long
    For gr = GridStart() To 600
        If Left$(CStr(ws.Cells(gr, 3).Value), 6) = "allow:" Then
            mapID = Mid$(CStr(ws.Cells(gr, 3).Value), 7)
            If mapID = "new" Then
                If Trim$(CStr(ws.Cells(gr, 1).Value)) <> "" And Trim$(CStr(ws.Cells(gr, 2).Value)) <> "" Then
                    modCore.AppendRow lo, Array(modCore.NewGUID(), Trim$(CStr(ws.Cells(gr, 1).Value)), code, _
                        Trim$(CStr(ws.Cells(gr, 2).Value)), ws.Cells(gr, 4).Value, ws.Cells(gr, 5).Value, "Custom", "", _
                        IIf(CStr(ws.Cells(gr, 6).Value) = "", "Active", CStr(ws.Cells(gr, 6).Value)), ts, u)
                    modData.WriteHistory "", "Create", code, "Allowed", "", ws.Cells(gr, 1).Value & " -> " & ws.Cells(gr, 2).Value, _
                                         "Type form", batch, modCore.SH_HMETA
                End If
            Else
                row = FindTableRow(lo, "MapID", mapID)
                If row > 0 Then
                    SetTableValue lo, row, "SourceType", Trim$(CStr(ws.Cells(gr, 1).Value)), "MapID", batch
                    SetTableValue lo, row, "TargetType", Trim$(CStr(ws.Cells(gr, 2).Value)), "MapID", batch
                    SetTableValue lo, row, "MinCard", ws.Cells(gr, 4).Value, "MapID", batch
                    SetTableValue lo, row, "MaxCard", ws.Cells(gr, 5).Value, "MapID", batch
                    SetTableValue lo, row, "Status", CStr(ws.Cells(gr, 6).Value), "MapID", batch
                End If
            End If
        End If
    Next gr
    modCore.ProtectSheet lo.Parent
    modCore.EndBusy
    ApplyAfterMetaChange
    ShowTypeForm "Relationship type", code
    FormMessage FormWS("FORM_Type"), "Saved relationship type " & code & "."
    SaveRelType = True
    Exit Function
Fail:
    modCore.ResetBusy
    FormMessage ws, "Not saved: " & Err.Description, True
End Function

Private Function TokenValid(ByVal tok As String) As Boolean
    tok = Trim$(tok)
    If tok = "*" Or tok = "SAME" Then TokenValid = True: Exit Function
    If Left$(tok, 2) = "G:" Then TokenValid = modMeta.gGroups.Exists(Mid$(tok, 3)): Exit Function
    TokenValid = modMeta.TypeExists(tok)
End Function

Private Sub ApplyAfterMetaChange()
    modMeta.BumpSchemaVersion
    modGenerator.RegenerateAll
End Sub

' ---------------------------------------------------------------- metamodel table helpers
Public Function FindTableRow(lo As ListObject, ByVal header As String, ByVal value As String) As Long
    Dim r As Variant
    If lo.DataBodyRange Is Nothing Then Exit Function
    r = Application.Match(value, lo.ListColumns(header).DataBodyRange, 0)
    If Not IsError(r) Then FindTableRow = CLng(r)
End Function

' Sets a cell in a metamodel table and logs the change to H_Metamodel when keyHeader is given.
Private Sub SetTableValue(lo As ListObject, ByVal rowIdx As Long, ByVal header As String, ByVal v As Variant, _
                          ByVal keyHeader As String, ByVal batch As String)
    Dim c As Range, oldV As Variant
    If rowIdx = 0 Then Exit Sub
    Set c = lo.ListColumns(header).DataBodyRange.Cells(rowIdx, 1)
    oldV = c.Value
    If modCore.CellText(oldV) = modCore.CellText(v) Then Exit Sub
    c.Value = modCore.SafeFor(v, CStr(c.NumberFormat))
    If keyHeader <> "" Then
        modData.WriteHistory "", "Update", CStr(lo.ListColumns(keyHeader).DataBodyRange.Cells(rowIdx, 1).Value), _
                             lo.Name & "." & header, oldV, v, "Type form", batch, modCore.SH_HMETA
    End If
End Sub

' =====================================================================
' FORM_Rule
' =====================================================================
Public Sub ShowRuleFormBlank()
    ShowRuleForm "", ""
End Sub

Public Sub ShowRuleForm(Optional ByVal attrCode As String = "", Optional ByVal ruleID As String = "")
    modCore.Heal
    Dim ws As Worksheet, labels As New Collection, ids As New Collection, x As Variant, cur As Variant, r As Long
    modMeta.LoadMeta
    ThisWorkbook.Activate
    Set ws = FormWS("FORM_Rule")
    modCore.BeginBusy
    On Error GoTo Fail
    ws.Activate
    ResetForm ws, "Define Rule", "Rules are checked on every edit, on save from EA_Entities and by Validate All."
    Field ws, 4, "Attribute *", "_attr", attrCode
    ListDropdown ws, "_attr", "RuleAttrCodes"
    Dim lo As ListObject, d As Variant, i As Long
    Set lo = modCore.Tbl("MM_Rules", "tblRules")
    d = modCore.TableData(lo)
    If attrCode <> "" And IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, lo.ListColumns("AttrCode").Index)) = attrCode Then
                labels.Add d(i, lo.ListColumns("RuleType").Index) & " " & d(i, lo.ListColumns("Param1").Index) & _
                           IIf(CStr(d(i, lo.ListColumns("TypeCode").Index)) <> "", " (" & d(i, lo.ListColumns("TypeCode").Index) & ")", "") & _
                           " - " & d(i, lo.ListColumns("Status").Index)
                ids.Add CStr(d(i, lo.ListColumns("RuleID").Index))
                If CStr(d(i, lo.ListColumns("RuleID").Index)) = ruleID Then cur = Application.Index(lo.DataBodyRange.Value, i, 0)
            End If
        Next i
    End If
    Dim curLabel As String
    If ruleID <> "" Then
        For i = 1 To ids.Count
            If ids(i) = ruleID Then curLabel = labels(i)
        Next i
    End If
    Field ws, 5, "Existing rule (blank = new)", "_existing", curLabel
    HelperList ws, "_existing", "X", labels, ids, "Y"
    ws.Range("Z1").Value = ruleID
    r = FIRST_FIELD_ROW
    Field ws, r, "Only for type (blank = all)", "TypeCode", RuleVal(cur, lo, "TypeCode"): ListDropdown ws, "TypeCode", "TypeCodes": r = r + 1
    Field ws, r, "Rule type *", "RuleType", RuleVal(cur, lo, "RuleType"): ListDropdown ws, "RuleType", "RuleType": r = r + 1
    ws.Range(ws.Cells(r, 2), ws.Cells(r + 1, 2)).NumberFormat = "@"     ' keep dates and patterns as text
    Field ws, r, "Parameter 1", "Param1", CStr(Nz(RuleVal(cur, lo, "Param1"), "")): r = r + 1
    Field ws, r, "Parameter 2", "Param2", CStr(Nz(RuleVal(cur, lo, "Param2"), "")): r = r + 1
    Field ws, r, "Severity *", "Severity", IIf(IsEmpty(cur), "Error", RuleVal(cur, lo, "Severity")): ListDropdown ws, "Severity", "Severity": r = r + 1
    Field ws, r, "Message *", "Message", RuleVal(cur, lo, "Message"): r = r + 1
    Field ws, r, "Status", "Status", IIf(IsEmpty(cur), "Active", RuleVal(cur, lo, "Status")): ListDropdown ws, "Status", "Status": r = r + 2
    Field ws, r, "Test value", "_test": r = r + 1
    RuleHint ws
    AddButton ws, "Save rule", MacroRef("SaveRuleForm"), ws.Range("F4"), 150
    AddButton ws, "Test rule", MacroRef("TestRuleForm"), ws.Range("F6"), 150
    AddButton ws, "New rule", MacroRef("NewRuleForAttr"), ws.Range("F8"), 150
    AddButton ws, "Back to Home", MacroRef("GoHome"), ws.Range("F10"), 150
    FinishForm ws
    modCore.EndBusy
    Exit Sub
Fail:
    modCore.ResetBusy
    MsgBox "Could not open the rule form: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

Private Function RuleVal(cur As Variant, lo As ListObject, ByVal header As String) As Variant
    If IsEmpty(cur) Then Exit Function
    RuleVal = cur(lo.ListColumns(header).Index)
End Function

Public Sub NewRuleForAttr()
    ShowRuleForm CStr(FieldValue(FormWS("FORM_Rule"), "_attr")), ""
End Sub

Private Sub RuleHint(ws As Worksheet)
    Dim h As String
    Select Case CStr(FieldValue(ws, "RuleType"))
        Case "Mandatory": h = "No parameters."
        Case "List": h = "Param 1 = list name (MM_Lists)."
        Case "MinValue", "MaxValue": h = "Param 1 = the limit."
        Case "Range": h = "Param 1 = minimum, Param 2 = maximum."
        Case "MinLength", "MaxLength": h = "Param 1 = number of characters."
        Case "Regex": h = "Param 1 = regular expression, e.g. ^\d+(\.\d+)*$"
        Case "DateRange": h = "Param 1 = earliest date, Param 2 = latest date (yyyy-mm-dd)."
        Case "DateOrder": h = "Param 1 = the attribute this date must come after."
        Case "Unique": h = "Param 1 = Type (unique within type) or Model."
        Case "EntityRef": h = "Param 1 = allowed type codes, comma separated (blank = any)."
        Case "RequiredIf": h = "Param 1 = other attribute, Param 2 = value that makes this required."
    End Select
    Dim r As Long
    r = FieldRow(ws, "Param1")
    If r > 0 Then
        Dim wasProt As Boolean
        wasProt = ws.ProtectContents
        If wasProt Then modCore.UnprotectSheet ws
        ws.Cells(r, 4).Value = h
        ws.Cells(r, 4).Font.Color = RGB(89, 89, 89)
        If wasProt Then modCore.ProtectSheet ws
    End If
End Sub

Public Sub TestRuleForm()
    Dim ws As Worksheet, rt As String, tv As Variant, ok As Boolean, r As Long, attrCode As String
    Set ws = FormWS("FORM_Rule")
    rt = CStr(FieldValue(ws, "RuleType"))
    tv = FieldValue(ws, "_test")
    attrCode = CStr(FieldValue(ws, "_attr"))
    modCore.UnprotectSheet ws
    r = FieldRow(ws, "_test")
    Select Case rt
        Case "DateOrder", "RequiredIf", "Unique"
            ws.Cells(r, 4).Value = "This rule needs a whole entity - check it with Validate All."
            ws.Cells(r, 4).Font.Color = RGB(89, 89, 89)
        Case "Mandatory"
            ok = Not modCore.IsBlankValue(tv)
            ShowTest ws, r, ok
        Case Else
            If modCore.IsBlankValue(tv) Then
                ws.Cells(r, 4).Value = "Enter a test value."
            Else
                If rt = "Regex" And Not modRules.RegexAvailable() Then
                    ws.Cells(r, 4).Value = "The regular-expression engine is not available on this PC; Regex rules are skipped."
                Else
                    ok = modRules.RulePasses(rt, CStr(FieldValue(ws, "Param1")), CStr(FieldValue(ws, "Param2")), _
                                             modData.ToCellValue(attrCode, tv), Nothing, "", Nothing, "")
                    ShowTest ws, r, ok
                End If
            End If
    End Select
    modCore.ProtectSheet ws
End Sub

Private Sub ShowTest(ws As Worksheet, ByVal r As Long, ByVal ok As Boolean)
    ws.Cells(r, 4).Value = IIf(ok, "Passes", "Fails: " & CStr(FieldValue(ws, "Message")))
    ws.Cells(r, 4).Font.Color = IIf(ok, RGB(0, 97, 0), RGB(192, 0, 0))
End Sub

Public Sub SaveRuleForm()
    Dim ws As Worksheet, attrCode As String, rt As String, sev As String, msg As String, ruleID As String
    Dim lo As ListObject, r As Long, bad As Boolean, batch As String, ts As String, u As String
    Set ws = FormWS("FORM_Rule")
    modCore.UnprotectSheet ws
    ClearFieldMsgs ws
    attrCode = CStr(FieldValue(ws, "_attr"))
    rt = CStr(FieldValue(ws, "RuleType"))
    sev = CStr(FieldValue(ws, "Severity"))
    msg = Trim$(CStr(FieldValue(ws, "Message")))
    ruleID = CStr(ws.Range("Z1").Value)
    If Not modMeta.gAttrs.Exists(attrCode) Then FieldMsg ws, "_attr", "Pick an attribute", "Error": bad = True
    If rt = "" Then FieldMsg ws, "RuleType", "Pick a rule type", "Error": bad = True
    If sev = "" Then FieldMsg ws, "Severity", "Pick a severity", "Error": bad = True
    If msg = "" Then FieldMsg ws, "Message", "Write the message users will see", "Error": bad = True
    If rt = "Regex" Then
        If modRules.RegexAvailable() Then
            On Error Resume Next
            Dim rx As Object
            Set rx = CreateObject("VBScript.RegExp")
            rx.Pattern = CStr(FieldValue(ws, "Param1"))
            rx.Test "x"
            If Err.Number <> 0 Then FieldMsg ws, "Param1", "Invalid regular expression", "Error": bad = True
            On Error GoTo 0
        End If
    End If
    If bad Then
        FormMessage ws, "Not saved - fix the fields marked in red.", True
        modCore.ProtectSheet ws
        Exit Sub
    End If
    If Not modHistory.AutoSnapshot("Before rule change") Then
        FormMessage ws, "Not saved.", True
        modCore.ProtectSheet ws
        Exit Sub
    End If
    batch = modCore.NewGUID(): ts = modCore.UtcNow(): u = modCore.CurrentUser()
    Set lo = modCore.Tbl("MM_Rules", "tblRules")
    modCore.UnprotectSheet lo.Parent
    lo.ListColumns("Param1").DataBodyRange.NumberFormat = "@"
    lo.ListColumns("Param2").DataBodyRange.NumberFormat = "@"
    If ruleID = "" Then
        ruleID = modCore.NewGUID()
        modCore.AppendRow lo, Array(ruleID, attrCode, CStr(FieldValue(ws, "TypeCode")), rt, "'" & CStr(FieldValue(ws, "Param1")), _
                                    "'" & CStr(FieldValue(ws, "Param2")), sev, msg, CStr(FieldValue(ws, "Status")), ts, u)
        modData.WriteHistory "", "Create", ruleID, "Rule", "", attrCode & " " & rt, "Rule form", batch, modCore.SH_HMETA
    Else
        r = FindTableRow(lo, "RuleID", ruleID)
        SetTableValue lo, r, "TypeCode", CStr(FieldValue(ws, "TypeCode")), "RuleID", batch
        SetTableValue lo, r, "RuleType", rt, "RuleID", batch
        SetTableValue lo, r, "Param1", CStr(FieldValue(ws, "Param1")), "RuleID", batch
        SetTableValue lo, r, "Param2", CStr(FieldValue(ws, "Param2")), "RuleID", batch
        SetTableValue lo, r, "Severity", sev, "RuleID", batch
        SetTableValue lo, r, "Message", msg, "RuleID", batch
        SetTableValue lo, r, "Status", CStr(FieldValue(ws, "Status")), "RuleID", batch
        SetTableValue lo, r, "Modified", ts, "", ""
        SetTableValue lo, r, "ModifiedBy", u, "", ""
    End If
    modCore.ProtectSheet lo.Parent
    modMeta.LoadMeta True
    Dim wbE As Workbook
    Set wbE = modCore.EntitiesWB(False)
    If Not wbE Is Nothing Then modGenerator.RefreshEntitiesCaches wbE
    ShowRuleForm attrCode, ruleID
    FormMessage FormWS("FORM_Rule"), "Rule saved. Run Validate All to check existing data against it."
End Sub
