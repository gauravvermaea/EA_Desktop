Attribute VB_Name = "modMerge"
' =====================================================================
' Parent and child models (phase 5).
'   Create child  (parent)  copies the model folder with a scope of entities;
'                           the child finishes its own setup when first opened
'   Merge package (child)   values-only copy of the child's metamodel and data
'                           in <child>\Merge\ChildPackage.xlsx
'   Merge preview (parent)  three-way comparison of base, parent and child by
'                           GUID, field by field, listed on MERGE_Review
'   Apply merge   (parent)  metamodel first, then entities, relationships and
'                           diagrams; the package becomes the child's new base
' Excel cannot open two workbooks with the same name, so the parent and the
' child are never open together: they meet only through the package and base
' files, which have their own names.
' =====================================================================
Option Explicit

Private Const SH_FORM As String = "FORM_Merge"
Private Const SH_REVIEW As String = "MERGE_Review"
Private Const SH_CHILDREN As String = "SYS_Children"
Private Const SH_SCOPE As String = "SYS_ChildScope"
Private Const SETUP_FILE As String = "ChildSetup.txt"
Private Const PACKAGE_KIND As String = "EAWB-MergePackage-1"
Private Const PACKAGE_FILE As String = "ChildPackage.xlsx"
Private Const BASE_FILE As String = "MergeBase.xlsx"
Private Const KEEP_PARENT As String = "Parent"
Private Const KEEP_CHILD As String = "Child"
Private Const KEEP_SKIP As String = "Skip"
' sheet|table|key headers (comma)|headers never compared (comma)
Private Const MM_TABLES As String = _
    "MM_EntityTypes|tblEntityTypes|TypeCode|TypeID,Modified,ModifiedBy;" & _
    "MM_AttributeTypes|tblAttributeTypes|AttrCode|AttrID,Modified,ModifiedBy;" & _
    "MM_RelationshipTypes|tblRelTypes|RelCode|RelTypeID,Modified,ModifiedBy;" & _
    "MM_EntityAttributes|tblEntityAttributes|TypeCode,AttrCode|MapID,Modified,ModifiedBy;" & _
    "MM_EntityRelationships|tblEntityRels|SourceType,RelCode,TargetType|MapID,Modified,ModifiedBy;" & _
    "MM_TypeGroups|tblTypeGroups|GroupName,TypeCode|;" & _
    "MM_Rules|tblRules|RuleID|Modified,ModifiedBy;" & _
    "MM_Lists|tblLists|ListName,Value|;" & _
    "MM_Styles|tblStyles|StyleKey|"
Private Const ENTITY_SYS As String = "Name|Description|Status|RepositoryArea"
Private Const REL_FIELDS As String = "Label|Description|Status"

Private mFso As Object

' ================================================================ small helpers
Private Function Fso() As Object
    If mFso Is Nothing Then Set mFso = CreateObject("Scripting.FileSystemObject")
    Set Fso = mFso
End Function

Private Function Sep() As String
    Sep = Application.PathSeparator
End Function

Private Function SafeName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ")
        s = Replace(s, CStr(ch), "_")
    Next ch
    SafeName = s
End Function

Private Function FileStamp(ByVal path As String) As String
    On Error Resume Next
    FileStamp = Format$(FileDateTime(path), "yyyymmddhhnnss")
End Function

Private Function T(ByVal v As Variant) As String
    T = modCore.CellText(v)
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub

Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
    If Left$(ReadUtf8, 1) = ChrW$(&HFEFF) Then ReadUtf8 = Mid$(ReadUtf8, 2)
End Function

Private Sub EnsureFolder(ByVal p As String)
    If Not Fso().FolderExists(p) Then Fso().CreateFolder p
End Sub

' Reads a sheet whose first row (or keyRow) holds column keys: returns key -> Dictionary(column key -> value)
Private Function ReadKeyed(ws As Worksheet, ByVal keyRow As Long, ByVal firstRow As Long, ByVal idKey As String) As Object
    Dim d As Object, v As Variant, lastR As Long, lastC As Long, i As Long, j As Long, keys() As String, idCol As Long
    Dim rec As Object
    Set d = modCore.NewDict()
    Set ReadKeyed = d
    lastR = modCore.LastRowOf(ws): lastC = modCore.LastColOf(ws)
    If lastR < firstRow Or lastC = 0 Then Exit Function
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    ReDim keys(1 To lastC)
    For j = 1 To lastC
        keys(j) = CStr(modCore.Nz(v(keyRow, j), ""))
        If keys(j) = idKey Then idCol = j
    Next j
    If idCol = 0 Then Exit Function
    For i = firstRow To lastR
        If CStr(modCore.Nz(v(i, idCol), "")) <> "" Then
            Set rec = modCore.NewDict()
            For j = 1 To lastC
                If keys(j) <> "" Then rec(keys(j)) = v(i, j)
            Next j
            Set d(CStr(v(i, idCol))) = rec
        End If
    Next i
End Function

' Rows of a metamodel table keyed by its key headers joined with "|"
Private Function ReadTableKeyed(ws As Worksheet, ByVal keyHeaders As String) As Object
    Dim d As Object, v As Variant, lastR As Long, lastC As Long, i As Long, j As Long, hdr() As String
    Dim rec As Object, k As String, kh As Variant
    Set d = modCore.NewDict()
    Set ReadTableKeyed = d
    If ws Is Nothing Then Exit Function
    lastR = modCore.LastRowOf(ws): lastC = modCore.LastColOf(ws)
    If lastR < 2 Or lastC = 0 Then Exit Function
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    ReDim hdr(1 To lastC)
    For j = 1 To lastC
        hdr(j) = CStr(modCore.Nz(v(1, j), ""))
    Next j
    For i = 2 To lastR
        Set rec = modCore.NewDict()
        For j = 1 To lastC
            If hdr(j) <> "" Then rec(hdr(j)) = v(i, j)
        Next j
        k = ""
        For Each kh In Split(keyHeaders, ",")
            If rec.Exists(kh) Then k = k & IIf(k = "", "", "|") & CStr(modCore.Nz(rec(kh), ""))
        Next kh
        If Replace(k, "|", "") <> "" Then Set d(k) = rec
    Next i
End Function

Private Function SheetOf(wb As Workbook, ByVal nm As String) As Worksheet
    If modCore.SheetExists(wb, nm) Then Set SheetOf = wb.Worksheets(nm)
End Function

Private Sub CopyValuesTo(src As Worksheet, wb As Workbook, Optional ByVal newName As String = "")
    Dim dst As Worksheet, lastR As Long, lastC As Long, v As Variant
    Set dst = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    dst.Name = IIf(newName = "", src.Name, newName)
    lastR = modCore.LastRowOf(src): lastC = modCore.LastColOf(src)
    If lastR = 0 Or lastC = 0 Then Exit Sub
    dst.Range(dst.Cells(1, 1), dst.Cells(lastR, lastC)).NumberFormat = "@"
    v = src.Range(src.Cells(1, 1), src.Cells(lastR, lastC)).Value
    If IsArray(v) Then
        Dim i As Long, j As Long
        For i = 1 To UBound(v, 1)
            For j = 1 To UBound(v, 2)
                If Not IsEmpty(v(i, j)) Then v(i, j) = modCore.CellText(v(i, j))
            Next j
        Next i
    End If
    dst.Range(dst.Cells(1, 1), dst.Cells(lastR, lastC)).Value = v
End Sub

' ================================================================ form
Public Sub ShowMerge()
    Dim ws As Worksheet, core As String, isChild As Boolean
    modCore.Heal
    modMeta.LoadMeta
    ThisWorkbook.Activate
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    core = "'" & ThisWorkbook.Name & "'!modMerge."
    isChild = (CStr(modCore.Setting("ParentModelID")) <> "")
    modCore.BeginBusy
    On Error GoTo Done
    ws.Activate
    modForms.ResetForm ws, "Parent and child models", _
        "A child is a copy of part of this model for a project or team. Its changes are merged back here."
    Section ws, 4, "1. Create a child model from this model"
    modForms.Field ws, 5, "Child name *", "_child", "Project"
    modForms.Field ws, 6, "Scope *", "_scope", "Around entities"
    modForms.ListDropdown ws, "_scope", "MergeScope"
    modForms.Field ws, 7, "Entity type (for one type)", "_etype"
    modForms.ListDropdown ws, "_etype", "TypeCodes"
    modForms.Field ws, 8, "Repository area (for an area)", "_area"
    modForms.ListDropdown ws, "_area", "RepositoryArea"
    modForms.Field ws, 9, "Start Display IDs, comma separated (for around entities)", "_start"
    modForms.Field ws, 10, "Relationship hops (0-5)", "_hops", 1
    ws.Cells(5, 4).Value = "The child goes in a new folder next to this one"
    modCore.AddButton ws, "Create child", core & "CreateChildFromForm", ws.Range("F5"), 190

    Section ws, 12, "2. Merge a child into this model"
    ws.Cells(13, 1).Value = "In the child first click Prepare merge package; close the child; then preview here."
    ws.Cells(13, 1).Font.Color = RGB(89, 89, 89)
    modCore.AddButton ws, "Choose child and preview", core & "MergePreview", ws.Range("F12"), 190
    modCore.AddButton ws, "Apply merge", core & "MergeApply", ws.Range("F14"), 190
    modCore.AddButton ws, "Children of this model", core & "ShowChildren", ws.Range("F16"), 190

    Section ws, 18, "3. This model as a child"
    If isChild Then
        ws.Cells(19, 1).Value = "Child of " & modCore.Setting("ParentModelName") & " (" & modCore.Setting("ParentFolder") & _
                                "), created " & modCore.Setting("ChildCreated")
        modCore.AddButton ws, "Prepare merge package", core & "PrepareMergePackage", ws.Range("F18"), 190
    Else
        ws.Cells(19, 1).Value = "This model has no parent."
    End If
    ws.Cells(19, 1).Font.Color = RGB(89, 89, 89)
    modCore.AddButton ws, "Back to Home", "'" & ThisWorkbook.Name & "'!modForms.GoHome", ws.Range("F21"), 190
    ws.Range("B5").Select
Done:
    Dim eMsg As String
    If Err.Number <> 0 Then eMsg = Err.Description
    On Error Resume Next
    If eMsg <> "" Then modForms.FormMessage ws, "Form not fully built: " & eMsg, True
    modCore.ProtectSheet ws
    modCore.EndBusy
End Sub

Private Sub Section(ws As Worksheet, ByVal r As Long, ByVal text As String)
    ws.Cells(r, 1).Value = text
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    ws.Cells(r, 1).Font.Color = RGB(31, 56, 100)
End Sub

Private Sub FormMsg(ByVal text As String, ByVal isError As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    modCore.UnprotectSheet ws
    modForms.FormMessage ws, text, isError
    modCore.ProtectSheet ws
End Sub

' ================================================================ create child (parent side)
Public Sub CreateChildFromForm()
    Dim ws As Worksheet, scope As String, childName As String, msg As String, ids As Object, refs As Object
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    FormMsg "", False
    childName = Trim$(CStr(modForms.FieldValue(ws, "_child")))
    scope = CStr(modForms.FieldValue(ws, "_scope"))
    If childName = "" Then FormMsg "Enter a child name.", True: Exit Sub
    Set ids = ScopeIDs(scope, CStr(modForms.FieldValue(ws, "_etype")), CStr(modForms.FieldValue(ws, "_area")), _
                       CStr(modForms.FieldValue(ws, "_start")), Val(CStr(modForms.FieldValue(ws, "_hops"))), msg)
    If msg <> "" Then FormMsg msg, True: Exit Sub
    CreateChild childName, scope, ids
End Sub

' Entities in scope (EntityID -> True); retired entities are left behind except for the whole model
Public Function ScopeIDs(ByVal scope As String, ByVal typeCode As String, ByVal area As String, _
                         ByVal starts As String, ByVal hops As Long, ByRef msg As String) As Object
    Dim ids As Object, lo As ListObject, cols As Object, d As Variant, i As Long, b As Long, part As Variant
    Dim id As String, nb As Object, k As Variant, found As Boolean
    Set ids = modCore.NewDict()
    Set ScopeIDs = ids
    Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    d = modCore.TableData(lo)
    If Not IsArray(d) Then msg = "The model has no entities.": Exit Function
    b = lo.Range.Column - 1
    Select Case scope
        Case "Whole model"
            For i = 1 To UBound(d, 1)
                If CStr(d(i, cols("EntityID") - b)) <> "" Then ids(CStr(d(i, cols("EntityID") - b))) = True
            Next i
        Case "One entity type"
            If Not modMeta.TypeExists(typeCode) Then msg = "Pick the entity type.": Exit Function
            For i = 1 To UBound(d, 1)
                If CStr(d(i, cols("TypeCode") - b)) = typeCode And CStr(d(i, cols("Status") - b)) <> "Retired" Then
                    ids(CStr(d(i, cols("EntityID") - b))) = True
                End If
            Next i
        Case "Repository area"
            If area = "" Then msg = "Pick the repository area.": Exit Function
            For i = 1 To UBound(d, 1)
                If CStr(d(i, cols("RepositoryArea") - b)) = area And CStr(d(i, cols("Status") - b)) <> "Retired" Then
                    ids(CStr(d(i, cols("EntityID") - b))) = True
                End If
            Next i
        Case "Around entities"
            If hops < 0 Then hops = 0
            If hops > 5 Then hops = 5
            For Each part In Split(starts, ",")
                found = False
                For i = 1 To UBound(d, 1)
                    If StrComp(CStr(d(i, cols("DisplayID") - b)), Trim$(CStr(part)), vbTextCompare) = 0 Then
                        id = CStr(d(i, cols("EntityID") - b))
                        found = True
                        ids(id) = True
                        If hops > 0 Then
                            Set nb = modIO.Neighbourhood(id, hops)
                            For Each k In nb.Keys: ids(k) = True: Next k
                        End If
                    End If
                Next i
                If Not found And Trim$(CStr(part)) <> "" Then msg = "Display ID not found: " & Trim$(CStr(part)): Exit Function
            Next part
        Case Else
            msg = "Pick a scope."
            Exit Function
    End Select
    If ids.Count = 0 Then msg = "Nothing is in that scope."
End Function

Private Sub CreateChild(ByVal childName As String, ByVal scope As String, ids As Object)
    Dim parentFolder As String, childFolder As String, refs As Object, d As Variant, i As Long, a As String, b As String
    Dim wbE As Workbook, f As String, ff As Integer, k As Variant, newID As String, copied As Long, skippedD As Long
    Dim dIDs As Object, ok As Boolean
    On Error GoTo Fail
    modMeta.LoadMeta
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    parentFolder = ThisWorkbook.Path
    childFolder = Fso().GetParentFolderName(parentFolder) & Sep() & "EA_Model_" & _
                  SafeName(CStr(modCore.Setting("ModelName"))) & "_" & SafeName(childName)
    If Fso().FolderExists(childFolder) Then
        FormMsg "A folder named " & childFolder & " already exists. Pick another child name.", True
        Exit Sub
    End If
    ' entities outside the scope that in-scope relationships point at come along as references
    Set refs = modCore.NewDict()
    d = modCore.TableData(modData.RelLO())
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, 1)) <> "" And CStr(d(i, 11)) <> "Retired" Then
                a = CStr(d(i, 3)): b = CStr(d(i, 4))
                If ids.Exists(a) And Not ids.Exists(b) Then refs(b) = True
                If ids.Exists(b) And Not ids.Exists(a) Then refs(a) = True
            End If
        Next i
    End If
    If MsgBox("Create child model """ & childName & """?" & vbLf & vbLf & ids.Count & " entities in scope, " & refs.Count & _
              " more as read-only references." & vbLf & "Folder: " & childFolder, vbOKCancel + vbQuestion, _
              "EA Workbench") = vbCancel Then Exit Sub

    modCore.BeginBusy
    modCore.Status "creating child folder"
    Set wbE = modCore.EntitiesWB(False)
    If Not wbE Is Nothing Then modSync.SaveWithStamp wbE
    modSync.SaveWithStamp ThisWorkbook
    EnsureFolder childFolder
    For Each k In Array("History", "Imports", "Imports\Done", "Exports", "Templates", "Diagrams", "Diagrams\Export", "Merge")
        EnsureFolder childFolder & Sep() & k
    Next k
    ThisWorkbook.SaveCopyAs childFolder & Sep() & ThisWorkbook.Name
    If Not wbE Is Nothing Then
        wbE.SaveCopyAs childFolder & Sep() & wbE.Name
    ElseIf Fso().FileExists(parentFolder & Sep() & modCore.EntitiesFileName()) Then
        Fso().CopyFile parentFolder & Sep() & modCore.EntitiesFileName(), childFolder & Sep() & modCore.EntitiesFileName()
    End If
    For Each k In Array("model.json", "README.md")
        If Fso().FileExists(parentFolder & Sep() & k) Then Fso().CopyFile parentFolder & Sep() & k, childFolder & Sep() & k
    Next k
    For Each k In Array("src", "Library", "Reports")
        If Fso().FolderExists(parentFolder & Sep() & k) Then Fso().CopyFolder parentFolder & Sep() & k, childFolder & Sep() & k
    Next k
    ' diagrams whose repository shapes are all in the child
    f = Dir(parentFolder & Sep() & "Diagrams" & Sep() & "*.drawio")
    Dim dFiles As New Collection, fv As Variant
    Do While f <> ""
        dFiles.Add f
        f = Dir()
    Loop
    For Each fv In dFiles
        Set dIDs = modDiagram.DiagramEntityIDs(parentFolder & Sep() & "Diagrams" & Sep() & fv)
        ok = Not dIDs Is Nothing
        If ok Then
            For Each k In dIDs.Keys
                If Not ids.Exists(k) And Not refs.Exists(k) Then ok = False: Exit For
            Next k
        End If
        If ok Then
            Fso().CopyFile parentFolder & Sep() & "Diagrams" & Sep() & fv, childFolder & Sep() & "Diagrams" & Sep() & fv
            copied = copied + 1
        Else
            skippedD = skippedD + 1
        End If
    Next fv
    ' instructions the child carries out when it is first opened
    newID = LCase$(Replace(Replace(modCore.NewGUID(), "{", ""), "}", ""))
    Dim txt As String
    txt = "Kind=EAWB-ChildSetup-1" & vbLf & "ModelID=" & newID & vbLf & "ModelName=" & childName & vbLf & _
          "ParentModelID=" & CStr(modCore.Setting("ModelID")) & vbLf & "ParentModelName=" & CStr(modCore.Setting("ModelName")) & _
          vbLf & "ParentFolder=" & parentFolder & vbLf & "Scope=" & scope & vbLf
    For Each k In ids.Keys
        txt = txt & "E|" & k & "|In scope" & vbLf
    Next k
    For Each k In refs.Keys
        txt = txt & "E|" & k & "|Reference" & vbLf
    Next k
    WriteUtf8 childFolder & Sep() & SETUP_FILE, txt
    ' register the child
    modCore.UnprotectSheet ThisWorkbook.Worksheets(SH_CHILDREN)
    modCore.AppendRow ChildrenLO(), Array(newID, childName, childFolder, modCore.UtcNow(), scope & " (" & ids.Count & ")", _
                                          "", "Created; finishes setup when first opened")
    modCore.ProtectSheet ThisWorkbook.Worksheets(SH_CHILDREN)
    modSync.SaveWithStamp ThisWorkbook
    modCore.EndBusy
    MsgBox "Child model created in:" & vbLf & childFolder & vbLf & vbLf & copied & " diagram(s) copied" & _
           IIf(skippedD > 0, ", " & skippedD & " left out because they show entities outside the scope", "") & "." & vbLf & vbLf & _
           "To work on it: close this model (both workbooks), then open EA_Entities.xlsm in the child folder. " & _
           "The first open removes everything outside the scope and records the merge base.", vbInformation, "EA Workbench"
    Exit Sub
Fail:
    Dim m As String
    m = Err.Description
    On Error Resume Next
    modCore.ResetBusy
    MsgBox "Child not created: " & m, vbExclamation, "EA Workbench"
End Sub

Private Function ChildrenLO() As ListObject
    Set ChildrenLO = modCore.Tbl(SH_CHILDREN, "tblChildren")
End Function

Public Sub ShowChildren()
    modCore.Heal
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(SH_CHILDREN).Activate
End Sub

' ================================================================ child setup (runs in the child, on first open)
Public Function HasPendingSetup() As Boolean
    HasPendingSetup = Fso().FileExists(ThisWorkbook.Path & Sep() & SETUP_FILE)
End Function

Public Sub CompleteChildSetup()
    Dim path As String, ff As Integer, ln As String, p() As String, info As Object, scope As Object
    Dim lo As ListObject, d As Variant, i As Long, j As Long, keep() As Boolean, n As Long, out() As Variant
    Dim cols As Object, b As Long, idCol As Long, ws As Worksheet, wbE As Workbook, nE As Long, nR As Long
    On Error GoTo Fail
    path = ThisWorkbook.Path & Sep() & SETUP_FILE
    Set info = modCore.NewDict(): Set scope = modCore.NewDict()
    Dim lnv As Variant
    For Each lnv In Split(Replace(ReadUtf8(path), vbCr, ""), vbLf)
        ln = CStr(lnv)
        If Left$(ln, 2) = "E|" Then
            p = Split(ln, "|")
            If UBound(p) >= 2 Then scope(p(1)) = p(2)
        ElseIf InStr(ln, "=") > 0 Then
            info(Left$(ln, InStr(ln, "=") - 1)) = Mid$(ln, InStr(ln, "=") + 1)
        End If
    Next lnv
    If CStr(info("Kind")) <> "EAWB-ChildSetup-1" Then Exit Sub
    modCore.BeginBusy
    modCore.Status "setting up the child model"

    ' ---- entities: keep the scope and its references
    Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    b = lo.Range.Column - 1
    idCol = cols("EntityID") - b
    nE = FilterTable(lo, idCol, scope)
    ' ---- relationships: both ends kept, at least one end in scope
    nR = FilterRelationships(scope)

    ' ---- settings, lineage, scope register
    modCore.SetSetting "ModelID", info("ModelID")
    modCore.SetSetting "ModelName", info("ModelName")
    modCore.SetSetting "ParentModelID", info("ParentModelID")
    modCore.SetSetting "ParentModelName", info("ParentModelName")
    modCore.SetSetting "ParentFolder", info("ParentFolder")
    modCore.SetSetting "ChildCreated", modCore.UtcNow()
    modSync.JsonSet "modelId", """" & info("ModelID") & """"
    modSync.JsonSet "name", """" & Replace(info("ModelName"), """", "'") & """"
    modSync.JsonSet "parentModelId", """" & info("ParentModelID") & """"
    WriteScopeSheet scope
    ClearTableIfAny "SYNC_Review", "tblSyncReview"
    ClearTableIfAny "SYS_Children", "tblChildren"
    ClearTableIfAny "IMPORT_Preview", "tblImportPreview"
    ClearTableIfAny "HARVEST_Review", "tblHarvest"
    ClearTableIfAny "SYS_Versions", "tblVersions"
    ClearTableIfAny "SYS_Views", "tblViews"
    modCore.EndBusy

    ' ---- type sheets, libraries, base
    modGenerator.RegenerateAll
    If modGenerator.gRegenFailed Then
        MsgBox "Child setup is not finished because the type sheets could not be regenerated. Close without saving " & _
               "and open the child again to retry.", vbExclamation, "EA Workbench"
        Exit Sub
    End If
    modDiagram.AfterFullSync
    WritePackage ThisWorkbook.Path & Sep() & "History" & Sep() & BASE_FILE
    Kill path
    modSync.SaveWithStamp ThisWorkbook
    Set wbE = modCore.EntitiesWB(False)
    If Not wbE Is Nothing Then modSync.SaveWithStamp wbE
    MsgBox "Child model """ & info("ModelName") & """ is ready: " & nE & " entities and " & nR & " relationships." & vbLf & _
           "Entities marked as references (outside the scope) can be read, but changes to them are not merged back." & vbLf & vbLf & _
           "When you are ready to merge: Parent and child models > Prepare merge package.", vbInformation, "EA Workbench"
    Exit Sub
Fail:
    Dim m As String
    m = Err.Description
    On Error Resume Next
    modCore.ResetBusy
    MsgBox "Child setup stopped: " & m & vbLf & "Close without saving, then open the child again to retry.", _
           vbExclamation, "EA Workbench"
End Sub

' Deletes the rows whose key column is not in keepIDs, in place and in blocks from the bottom,
' so the kept rows keep their values, formats and formulas. Returns rows kept.
Private Function FilterTable(lo As ListObject, ByVal keyCol As Long, keepIDs As Object) As Long
    Dim d As Variant, i As Long, j As Long, n As Long, ws As Worksheet
    d = modCore.TableData(lo)
    If Not IsArray(d) Then Exit Function
    Set ws = lo.Parent
    modCore.UnprotectSheet ws
    If Not lo.AutoFilter Is Nothing Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    i = UBound(d, 1)
    Do While i >= 1
        If keepIDs.Exists(CStr(d(i, keyCol))) Then
            n = n + 1
            i = i - 1
        Else
            j = i
            Do While j > 1
                If keepIDs.Exists(CStr(d(j - 1, keyCol))) Then Exit Do
                j = j - 1
            Loop
            lo.ListRows(j).Range.Resize(i - j + 1).Delete xlShiftUp
            i = j - 1
        End If
    Loop
    modCore.ProtectSheet ws
    FilterTable = n
End Function

Private Function FilterRelationships(scope As Object) As Long
    Dim lo As ListObject, d As Variant, i As Long, keep As Object, a As String, b As String
    Set lo = modData.RelLO()
    Set keep = modCore.NewDict()
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            a = CStr(d(i, 3)): b = CStr(d(i, 4))
            If scope.Exists(a) And scope.Exists(b) Then
                If scope(a) = "In scope" Or scope(b) = "In scope" Then keep(CStr(d(i, 1))) = True
            End If
        Next i
    End If
    FilterRelationships = FilterTable(lo, 1, keep)
End Function

Private Sub ClearTableIfAny(ByVal sheetName As String, ByVal tableName As String)
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If lo Is Nothing Then Exit Sub
    modCore.UnprotectSheet lo.Parent
    modCore.ClearTable lo
    modCore.ProtectSheet lo.Parent
End Sub

Private Sub WriteScopeSheet(scope As Object)
    Dim ws As Worksheet, k As Variant, r As Long, v() As Variant
    If modCore.SheetExists(ThisWorkbook, SH_SCOPE) Then
        Set ws = ThisWorkbook.Worksheets(SH_SCOPE)
        modCore.UnprotectSheet ws
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_SCOPE
    End If
    ReDim v(1 To scope.Count + 1, 1 To 2)
    v(1, 1) = "EntityID": v(1, 2) = "Role"
    r = 1
    For Each k In scope.Keys
        r = r + 1
        v(r, 1) = k: v(r, 2) = scope(k)
    Next k
    ws.Range("A1").Resize(scope.Count + 1, 2).Value = v
    ws.Visible = xlSheetHidden
    modCore.ProtectSheet ws
End Sub

' Role of each entity in this child (In scope / Reference); empty when not a child
Private Function ScopeRoles() As Object
    Dim d As Object, ws As Worksheet, v As Variant, i As Long
    Set d = modCore.NewDict()
    Set ScopeRoles = d
    If Not modCore.SheetExists(ThisWorkbook, SH_SCOPE) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_SCOPE)
    If modCore.LastRowOf(ws) < 2 Then Exit Function
    v = ws.Range("A2:B" & modCore.LastRowOf(ws)).Value
    For i = 1 To UBound(v, 1)
        d(CStr(v(i, 1))) = CStr(v(i, 2))
    Next i
End Function

' ================================================================ merge package (child side)
Public Sub PrepareMergePackage()
    Dim p As String
    modCore.Heal
    If CStr(modCore.Setting("ParentModelID")) = "" Then
        MsgBox "This model has no parent, so there is nothing to merge back.", vbInformation, "EA Workbench"
        Exit Sub
    End If
    On Error GoTo Fail
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    p = ThisWorkbook.Path & Sep() & "Merge"
    EnsureFolder p
    p = p & Sep() & PACKAGE_FILE
    WritePackage p
    modSync.SaveWithStamp ThisWorkbook
    MsgBox "Merge package written:" & vbLf & p & vbLf & vbLf & "Now close this model (both workbooks), open the parent (" & _
           modCore.Setting("ParentFolder") & ") and use Parent and child models > Choose child and preview.", _
           vbInformation, "EA Workbench"
    Exit Sub
Fail:
    MsgBox "Package not written: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

' Values-only copy of this model: Info, Scope, DATA_Master, DATA_Relationships and the MM_ tables
Private Sub WritePackage(ByVal path As String)
    Dim wb As Workbook, ws As Worksheet, info As Variant, i As Long
    modCore.BeginBusy
    On Error GoTo Fail
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Info"
    info = Array("Kind", PACKAGE_KIND, "ModelID", CStr(modCore.Setting("ModelID")), "ModelName", CStr(modCore.Setting("ModelName")), _
                 "ParentModelID", CStr(modCore.Setting("ParentModelID")), "Created", modCore.UtcNow(), _
                 "CreatedBy", modCore.CurrentUser(), "SchemaVersion", modMeta.SchemaVersion(), "Folder", ThisWorkbook.Path)
    ws.Range("A1:B" & (UBound(info) + 1) \ 2).NumberFormat = "@"
    For i = 0 To UBound(info) Step 2
        ws.Cells(i \ 2 + 1, 1).Value = info(i)
        ws.Cells(i \ 2 + 1, 2).Value = CStr(info(i + 1))
    Next i
    If modCore.SheetExists(ThisWorkbook, SH_SCOPE) Then CopyValuesTo ThisWorkbook.Worksheets(SH_SCOPE), wb, "Scope"
    WriteDiagramStamps wb
    CopyValuesTo modData.MasterSheet(), wb
    CopyValuesTo modData.RelLO().Parent, wb
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "MM_*" Then CopyValuesTo ws, wb
    Next ws
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=51
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    modCore.EndBusy
    Exit Sub
Fail:
    Dim m As String
    m = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    modCore.ResetBusy
    Err.Raise vbObjectError + 601, "modMerge", m
End Sub

' ================================================================ merge (parent side)
Private Sub WriteDiagramStamps(wb As Workbook)
    Dim ws As Worksheet, f As String, r As Long, folder As String
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Diagrams"
    ws.Range("A:C").NumberFormat = "@"
    ws.Range("A1:C1").Value = Array("File", "Size", "Modified")
    folder = ThisWorkbook.Path & Sep() & "Diagrams" & Sep()
    If Not Fso().FolderExists(folder) Then Exit Sub
    r = 1
    f = Dir(folder & "*.drawio")
    Do While f <> ""
        r = r + 1
        ws.Cells(r, 1).Value = f
        ws.Cells(r, 2).Value = CStr(FileLen(folder & f))
        ws.Cells(r, 3).Value = Format$(FileDateTime(folder & f), "yyyy-mm-dd hh:nn:ss")
        f = Dir()
    Loop
End Sub

Private Function ReviewLO() As ListObject
    Set ReviewLO = modCore.Tbl(SH_REVIEW, "tblMergeReview")
End Function

Public Sub MergePreview()
    Dim f As Variant
    modCore.Heal
    f = Application.GetOpenFilename("Merge package (ChildPackage.xlsx),*.xlsx", , _
                                    "Choose the child's Merge\ChildPackage.xlsx (close the child first)")
    If VarType(f) = vbBoolean Then Exit Sub
    RunMerge CStr(f), False
End Sub

Public Sub MergeApply()
    Dim ws As Worksheet, p() As String
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_REVIEW)
    p = Split(CStr(ws.Range("B4").Value) & "*", "*")
    If p(0) = "" Then
        MsgBox "Preview a merge first (Choose child and preview).", vbInformation, "EA Workbench"
        Exit Sub
    End If
    If FileStamp(p(0)) <> p(1) Then
        MsgBox "The merge package changed after the preview. Preview again.", vbExclamation, "EA Workbench"
        Exit Sub
    End If
    If MsgBox("Apply this merge? Rows set to Keep = Child are written to this model; Parent and Skip leave it as it is. " & _
              "A snapshot is taken first.", vbOKCancel + vbQuestion, "EA Workbench") = vbCancel Then Exit Sub
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    If Not modHistory.AutoSnapshot("Before merge") Then Exit Sub
    RunMerge p(0), True
End Sub

Private Function OpenReadOnly(ByVal path As String) As Workbook
    Dim wb As Workbook, nm As String, ev As Boolean
    nm = Mid$(path, InStrRev(path, Sep()) + 1)
    For Each wb In Application.Workbooks
        If LCase$(wb.Name) = LCase$(nm) Then
            Err.Raise vbObjectError + 602, "modMerge", nm & " is open in Excel; close it first"
        End If
    Next wb
    ev = Application.EnableEvents
    Application.EnableEvents = False
    Set OpenReadOnly = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)
    Application.EnableEvents = ev
    ThisWorkbook.Activate
End Function

Private Function InfoOf(wb As Workbook) As Object
    Dim d As Object, ws As Worksheet, v As Variant, i As Long
    Set d = modCore.NewDict()
    Set InfoOf = d
    If Not modCore.SheetExists(wb, "Info") Then Exit Function
    Set ws = wb.Worksheets("Info")
    v = ws.Range("A1:B20").Value
    For i = 1 To 20
        If CStr(modCore.Nz(v(i, 1), "")) <> "" Then d(CStr(v(i, 1))) = CStr(modCore.Nz(v(i, 2), ""))
    Next i
End Function

Private Sub RunMerge(ByVal pkgPath As String, ByVal commit As Boolean)
    Dim wbC As Workbook, wbB As Workbook, basePath As String, childFolder As String, info As Object, rows As New Collection
    Dim decisions As Object, batch As String, src As String, mmChanged As Boolean, touched As Object, cnt(1 To 6) As Long
    Dim roles As Object, results As Object
    On Error GoTo Fail
    modMeta.LoadMeta True
    childFolder = Fso().GetParentFolderName(Fso().GetParentFolderName(pkgPath))
    basePath = childFolder & Sep() & "History" & Sep() & BASE_FILE
    If Not Fso().FileExists(basePath) Then Err.Raise vbObjectError + 603, , "No merge base in " & basePath & _
        " - was the child set up by opening it once?"
    modCore.BeginBusy
    modCore.Status "reading the child"
    Set wbC = OpenReadOnly(pkgPath)
    Set wbB = OpenReadOnly(basePath)
    Set info = InfoOf(wbC)
    If info("Kind") <> PACKAGE_KIND Then Err.Raise vbObjectError + 604, , "That file is not a Workbench merge package"
    If info("ParentModelID") <> CStr(modCore.Setting("ModelID")) Then Err.Raise vbObjectError + 605, , _
        "That child was not created from this model (its parent is another model)"
    src = "Merge " & info("ModelName")
    batch = modCore.NewGUID()
    Set decisions = ReadDecisions(commit)
    Set touched = modCore.NewDict()
    Set results = modCore.NewDict()
    Set roles = ReadRoles(wbC)

    ' ---- 1. metamodel
    MergeMetamodel wbB, wbC, commit, decisions, rows, src, batch, mmChanged, cnt
    If commit And mmChanged Then
        modCore.EndBusy
        modMeta.BumpSchemaVersion
        modGenerator.RegenerateAll
        modMeta.LoadMeta True
        modCore.BeginBusy
    End If
    ' ---- 2. entities, 3. relationships
    MergeEntities wbB, wbC, roles, commit, decisions, rows, src, batch, touched, cnt
    MergeRelationships wbB, wbC, roles, commit, decisions, rows, src, batch, touched, cnt
    ' ---- 4. diagrams
    MergeDiagrams childFolder, wbB, wbC, commit, decisions, rows, cnt

    wbC.Close SaveChanges:=False: Set wbC = Nothing
    wbB.Close SaveChanges:=False: Set wbB = Nothing
    WriteReview pkgPath, rows, commit
    If commit Then
        ' the package becomes the child's new base, so the next merge only sees later changes - unless something
        ' new was skipped or failed: then the base stays, and those items show again next time
        Dim baseMoved As Boolean
        baseMoved = (cnt(5) + cnt(6) = 0)
        If baseMoved Then Fso().CopyFile pkgPath, basePath, True
        EnsureFolder ThisWorkbook.Path & Sep() & "History"
        Fso().CopyFile pkgPath, ThisWorkbook.Path & Sep() & "History" & Sep() & "Merge_" & SafeName(info("ModelName")) & "_" & _
                                Format$(Now, "yyyymmdd-hhnnss") & ".xlsx", True
        MarkChildMerged info("ModelID"), info("ModelName"), childFolder
        If touched.Exists("") Then touched.Remove ""
        If touched.Count > 0 Then modPaint.PaintTypes touched
        modDiagram.AfterFullSync
        modSync.SaveWithStamp modCore.EntitiesWB(False)
        modSync.SaveWithStamp ThisWorkbook
        modCore.EndBusy
        MsgBox "Merge applied from " & info("ModelName") & "." & vbLf & vbLf & cnt(1) & " metamodel change(s), " & cnt(2) & _
               " entity change(s), " & cnt(3) & " relationship change(s), " & cnt(4) & " diagram(s) copied." & vbLf & _
               cnt(5) & " item(s) could not be applied - see the Note column on MERGE_Review." & vbLf & vbLf & _
               IIf(baseMoved, "The child can carry on; its next merge brings only the changes made after this one.", _
                   "The merge base was kept, because new items were skipped or failed: they, and conflicts you kept as " & _
                   "Parent, show again at the next merge. Records already merged compare as equal."), _
               IIf(cnt(5) > 0, vbExclamation, vbInformation), "EA Workbench"
    Else
        modCore.EndBusy
        MsgBox "Merge preview ready: " & rows.Count & " row(s) on MERGE_Review." & vbLf & _
               "Conflicts are set to Keep = Parent. Change Keep where you want the child's value, then click Apply merge.", _
               vbInformation, "EA Workbench"
    End If
    Exit Sub
Fail:
    Dim m As String
    m = Err.Description
    On Error Resume Next
    If Not wbC Is Nothing Then wbC.Close SaveChanges:=False
    If Not wbB Is Nothing Then wbB.Close SaveChanges:=False
    modCore.ResetBusy
    MsgBox "Merge stopped: " & m, vbExclamation, "EA Workbench"
End Sub

Private Function ReadRoles(wbC As Workbook) As Object
    Dim d As Object, ws As Worksheet, v As Variant, i As Long
    Set d = modCore.NewDict()
    Set ReadRoles = d
    If Not modCore.SheetExists(wbC, "Scope") Then Exit Function
    Set ws = wbC.Worksheets("Scope")
    If modCore.LastRowOf(ws) < 2 Then Exit Function
    v = ws.Range("A2:B" & modCore.LastRowOf(ws)).Value
    For i = 1 To UBound(v, 1)
        d(CStr(v(i, 1))) = CStr(v(i, 2))
    Next i
End Function

' Keep choices from MERGE_Review (apply only): "Kind|RecordID|Field" -> Keep
Private Function ReadDecisions(ByVal commit As Boolean) As Object
    Dim d As Object, data As Variant, i As Long
    Set d = modCore.NewDict()
    Set ReadDecisions = d
    If Not commit Then Exit Function
    data = modCore.TableData(ReviewLO())
    If Not IsArray(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        d(CStr(data(i, 1)) & "|" & CStr(data(i, 3)) & "|" & CStr(data(i, 7))) = CStr(data(i, 11))
    Next i
End Function

Private Function Decide(decisions As Object, ByVal kind As String, ByVal id As String, ByVal field As String, _
                        ByVal dflt As String) As String
    Dim k As String
    k = kind & "|" & id & "|" & field
    If decisions.Exists(k) Then
        Decide = decisions(k)
        If Decide <> KEEP_PARENT And Decide <> KEEP_CHILD And Decide <> KEEP_SKIP Then Decide = dflt
    Else
        Decide = dflt
    End If
End Function

' Columns: Kind, Action, RecordID, DisplayID, Name, Table/Type, Field, Base, Parent, Child, Keep, Note
Private Sub AddRow(rows As Collection, ByVal kind As String, ByVal action As String, ByVal id As String, ByVal disp As String, _
                   ByVal nm As String, ByVal tbl As String, ByVal field As String, ByVal bV As Variant, ByVal pV As Variant, _
                   ByVal cV As Variant, ByVal keep As String, ByVal note As String)
    rows.Add Array(kind, action, id, disp, nm, tbl, field, T(bV), T(pV), T(cV), keep, note)
End Sub

' Three-way outcome for one value: "" (nothing to do), "child" (take child), "conflict"
Private Function ThreeWay(ByVal bV As String, ByVal pV As String, ByVal cV As String) As String
    If cV = bV Or cV = pV Then
        ThreeWay = ""
    ElseIf pV = bV Then
        ThreeWay = "child"
    Else
        ThreeWay = "conflict"
    End If
End Function

' ---------------------------------------------------------------- metamodel
Private Sub MergeMetamodel(wbB As Workbook, wbC As Workbook, ByVal commit As Boolean, decisions As Object, rows As Collection, _
                           ByVal src As String, ByVal batch As String, ByRef changed As Boolean, cnt() As Long)
    Dim spec As Variant, p() As String, B As Object, P_ As Object, C As Object, k As Variant, hdr As Variant
    Dim sig As String, sb As String, sp As String, sc As String, outcome As String, keep As String, lo As ListObject
    Dim excl As String, action As String
    For Each spec In Split(MM_TABLES, ";")
        p = Split(spec, "|")
        excl = "," & p(3) & ","
        Set lo = modCore.Tbl(p(0), p(1))
        Set B = ReadTableKeyed(SheetOf(wbB, p(0)), p(2))
        Set P_ = ReadTableKeyed(lo.Parent, p(2))
        Set C = ReadTableKeyed(SheetOf(wbC, p(0)), p(2))
        For Each k In C.Keys
            sc = RowSig(C(k), excl)
            If Not P_.Exists(k) Then
                ' new in the child (or removed from the parent's table, which the Workbench never does)
                keep = Decide(decisions, "Metamodel", p(0) & ":" & k, "(row)", KEEP_CHILD)
                AddRow rows, "Metamodel", "Add", p(0) & ":" & k, "", CStr(k), p(0), "(row)", "", "", sc, keep, "New in the child"
                If commit And keep = KEEP_CHILD Then
                    WriteMetaRow lo, 0, C(k)
                    modData.WriteHistory "", "Merge add", CStr(k), p(0), "", sc, src, batch, modCore.SH_HMETA
                    changed = True: cnt(1) = cnt(1) + 1
                End If
            Else
                sp = RowSig(P_(k), excl)
                If B.Exists(k) Then sb = RowSig(B(k), excl) Else sb = sp
                outcome = ThreeWay(sb, sp, sc)
                If outcome <> "" Then
                    keep = Decide(decisions, "Metamodel", p(0) & ":" & k, "(row)", IIf(outcome = "child", KEEP_CHILD, KEEP_PARENT))
                    action = IIf(outcome = "child", "Update", "Conflict")
                    AddRow rows, "Metamodel", action, p(0) & ":" & k, "", CStr(k), p(0), "(row)", sb, sp, sc, keep, _
                           IIf(outcome = "child", "Changed in the child", "Changed in both")
                    If commit And keep = KEEP_CHILD Then
                        WriteMetaRow lo, MetaRowIndex(lo, p(2), CStr(k)), C(k)
                        modData.WriteHistory "", "Merge update", CStr(k), p(0), sp, sc, src, batch, modCore.SH_HMETA
                        changed = True: cnt(1) = cnt(1) + 1
                    End If
                End If
            End If
        Next k
    Next spec
End Sub

Private Function RowSig(rec As Object, ByVal excl As String) As String
    Dim k As Variant, s As String
    For Each k In rec.Keys
        If InStr(excl, "," & k & ",") = 0 Then s = s & k & "=" & T(rec(k)) & "; "
    Next k
    RowSig = s
End Function

Private Function MetaRowIndex(lo As ListObject, ByVal keyHeaders As String, ByVal key As String) As Long
    Dim d As Variant, i As Long, kh As Variant, k As String, c As Long
    d = modCore.TableData(lo)
    If Not IsArray(d) Then Exit Function
    For i = 1 To UBound(d, 1)
        k = ""
        For Each kh In Split(keyHeaders, ",")
            c = modCore.ColIndex(lo, CStr(kh))
            If c > 0 Then k = k & IIf(k = "", "", "|") & CStr(modCore.Nz(d(i, c), ""))
        Next kh
        If k = key Then MetaRowIndex = i: Exit Function
    Next i
End Function

Private Function IsPlainNumber(ByVal s As String) As Boolean
    s = Trim$(s)
    If s = "" Then Exit Function
    IsPlainNumber = Not (s Like "*[!0-9.-]*")
End Function

' Writes a metamodel row by header names (index 0 appends)
Private Sub WriteMetaRow(lo As ListObject, ByVal idx As Long, rec As Object)
    Dim rng As Range, k As Variant, c As Long, ws As Worksheet
    Set ws = lo.Parent
    modCore.UnprotectSheet ws
    If idx = 0 Then
        Set rng = modCore.AppendRow(lo, Array(Empty))
    Else
        Set rng = lo.ListRows(idx).Range
    End If
    For Each k In rec.Keys
        c = modCore.ColIndex(lo, CStr(k))
        If c > 0 Then
            If InStr(",DisplayOrder,SortOrder,MinCard,MaxCard,NextNo,", "," & k & ",") > 0 And IsPlainNumber(CStr(modCore.Nz(rec(k), ""))) Then
                rng.Cells(1, c).Value = Val(CStr(rec(k)))
            Else
                rng.Cells(1, c).Value = modCore.SafeFor(rec(k), CStr(rng.Cells(1, c).NumberFormat))
            End If
        End If
    Next k
    c = modCore.ColIndex(lo, "Modified")
    If c > 0 Then rng.Cells(1, c).Value = modCore.UtcNow()
    c = modCore.ColIndex(lo, "ModifiedBy")
    If c > 0 Then rng.Cells(1, c).Value = modCore.CurrentUser()
    modCore.ProtectSheet ws
End Sub

' ---------------------------------------------------------------- entities
Private Sub MergeEntities(wbB As Workbook, wbC As Workbook, roles As Object, ByVal commit As Boolean, decisions As Object, _
                          rows As Collection, ByVal src As String, ByVal batch As String, touched As Object, cnt() As Long)
    Dim B As Object, C As Object, P_ As Object, k As Variant, rb As Object, rc As Object, rp As Object, f As Variant
    Dim fields As Collection, outcome As String, keep As String, upd As Object, issues As Collection, blocked As Boolean
    Dim tc As String, nm As String, disp As String, newID As String, vals As Object, pRetired As Boolean, cRetired As Boolean
    Dim bRetired As Boolean, childChanged As Boolean, parentChanged As Boolean, r As Long, note As String
    Set B = ReadKeyed(wbB.Worksheets("DATA_Master"), 1, 3, "EntityID")
    Set C = ReadKeyed(wbC.Worksheets("DATA_Master"), 1, 3, "EntityID")
    Set P_ = ReadKeyed(modData.MasterSheet(), 1, 3, "EntityID")
    For Each k In C.Keys
        Set rc = C(k)
        tc = CStr(rc("TypeCode")): nm = CStr(rc("Name")): disp = CStr(rc("DisplayID"))
        If roles.Exists(k) Then
            If roles(k) = "Reference" Then
                If B.Exists(k) Then
                    If RecordChanged(B(k), rc) Then AddRow rows, "Entity", "Info", CStr(k), disp, nm, tc, "", "", "", "", _
                        KEEP_SKIP, "Reference entity (outside the child's scope) was changed in the child - not merged"
                End If
                GoTo NextEntity
            End If
        End If
        If Not P_.Exists(k) Then
            If B.Exists(k) Then
                AddRow rows, "Entity", "Info", CStr(k), disp, nm, tc, "", "", "", "", KEEP_SKIP, "No longer in the parent - not merged"
                GoTo NextEntity
            End If
            ' ---- new in the child
            keep = Decide(decisions, "Entity", CStr(k), "(new)", KEEP_CHILD)
            note = "New in the child"
            If Not modMeta.TypeExists(tc) Then note = "New in the child; its type " & tc & " comes with the metamodel merge"
            If commit And keep = KEEP_CHILD Then
                Set vals = modCore.NewDict()
                For Each f In EntityFields(rc)
                    If Not modCore.IsBlankValue(rc(f)) Then vals(f) = FromText(CStr(f), rc(f))
                Next f
                Set issues = New Collection
                newID = modData.CreateEntity(tc, vals, src, batch, issues, CStr(k), disp)
                If newID = "" Then
                    note = "Not added: " & Replace(modRules.IssueText(issues, 3), vbLf, "; ")
                    cnt(5) = cnt(5) + 1
                Else
                    r = modData.FindEntityRow(newID)
                    note = "Added as " & modData.MasterSheet().Cells(r, modData.MasterCols()("DisplayID")).Value
                    touched(tc) = True: cnt(2) = cnt(2) + 1
                End If
            End If
            If commit And keep <> KEEP_CHILD Then cnt(6) = cnt(6) + 1
            AddRow rows, "Entity", "Add", CStr(k), disp, nm, tc, "(new)", "", "", nm, keep, note
            GoTo NextEntity
        End If
        ' ---- in both: field by field
        Set rp = P_(k)
        If B.Exists(k) Then Set rb = B(k) Else Set rb = rp
        Set fields = EntityFields(rc)
        pRetired = (T(rp("Status")) = "Retired"): cRetired = (T(rc("Status")) = "Retired"): bRetired = (T(rb("Status")) = "Retired")
        childChanged = False: parentChanged = False
        For Each f In fields
            If T(Val_(rc, f)) <> T(Val_(rb, f)) Then childChanged = True
            If T(Val_(rp, f)) <> T(Val_(rb, f)) Then parentChanged = True
        Next f
        Set upd = modCore.NewDict()
        For Each f In fields
            outcome = ThreeWay(T(Val_(rb, f)), T(Val_(rp, f)), T(Val_(rc, f)))
            ' retired on one side and changed on the other: always a decision
            If outcome = "child" And pRetired And Not bRetired Then outcome = "conflict"
            If outcome = "child" And f = "Status" And cRetired And Not bRetired And parentChanged Then outcome = "conflict"
            If outcome <> "" Then
                keep = Decide(decisions, "Entity", CStr(k), CStr(f), IIf(outcome = "child", KEEP_CHILD, KEEP_PARENT))
                note = IIf(outcome = "child", "Changed in the child", "Changed in both")
                If pRetired And Not bRetired And outcome = "conflict" Then note = "Retired in the parent, changed in the child"
                If cRetired And f = "Status" And outcome = "conflict" Then note = "Retired in the child, changed in the parent"
                AddRow rows, "Entity", IIf(outcome = "child", "Update", "Conflict"), CStr(k), T(rp("DisplayID")), T(rp("Name")), _
                       tc, CStr(f), Val_(rb, f), Val_(rp, f), Val_(rc, f), keep, note
                If keep = KEEP_CHILD Then upd(f) = FromText(CStr(f), Val_(rc, f))
            End If
        Next f
        If commit And upd.Count > 0 Then
            Set issues = New Collection
            blocked = False
            modData.UpdateEntity CStr(k), upd, src, batch, issues, blocked
            If blocked Then
                cnt(5) = cnt(5) + 1
                AddRow rows, "Entity", "Error", CStr(k), T(rp("DisplayID")), T(rp("Name")), tc, "", "", "", "", KEEP_SKIP, _
                       "Not applied: " & Replace(modRules.IssueText(issues, 3), vbLf, "; ")
            Else
                touched(tc) = True: cnt(2) = cnt(2) + 1
            End If
        End If
NextEntity:
    Next k
End Sub

' Package values are text written with a "." decimal point; convert numbers with Val, not the regional settings
Private Function FromText(ByVal attrCode As String, ByVal v As Variant) As Variant
    Dim dt As String
    If modMeta.gAttrs.Exists(attrCode) Then dt = modMeta.gAttrs(attrCode)("DataType")
    Select Case dt
        Case "Integer", "Decimal", "Currency", "Percent", "Number"
            If Trim$(CStr(modCore.Nz(v, ""))) = "" Then FromText = Empty Else FromText = Val(CStr(v))
        Case Else
            FromText = modData.ToCellValue(attrCode, v)
    End Select
End Function

Private Function Val_(rec As Object, ByVal f As Variant) As Variant
    If rec.Exists(CStr(f)) Then Val_ = rec(CStr(f))
End Function

' The fields that merge: Name, Description, Status, RepositoryArea and every attribute column of the child row
Private Function EntityFields(rc As Object) As Collection
    Dim c As New Collection, k As Variant, skip As String
    skip = ",EntityID,TypeCode,DisplayID,Created,CreatedBy,Modified,ModifiedBy,RowHash,"
    For Each k In Split(ENTITY_SYS, "|"): c.Add CStr(k): Next k
    For Each k In rc.Keys
        If InStr(skip, "," & k & ",") = 0 And InStr("|" & ENTITY_SYS & "|", "|" & k & "|") = 0 Then
            If modMeta.gAttrs.Exists(CStr(k)) Then c.Add CStr(k)
        End If
    Next k
    Set EntityFields = c
End Function

Private Function RecordChanged(a As Object, b As Object) As Boolean
    Dim k As Variant
    For Each k In b.Keys
        If InStr(",Created,CreatedBy,Modified,ModifiedBy,RowHash,", "," & k & ",") = 0 Then
            If T(Val_(a, k)) <> T(b(k)) Then RecordChanged = True: Exit Function
        End If
    Next k
End Function

' ---------------------------------------------------------------- relationships
Private Sub MergeRelationships(wbB As Workbook, wbC As Workbook, roles As Object, ByVal commit As Boolean, decisions As Object, _
                               rows As Collection, ByVal src As String, ByVal batch As String, touched As Object, cnt() As Long)
    Dim B As Object, C As Object, P_ As Object, k As Variant, rb As Object, rc As Object, rp As Object, f As Variant
    Dim outcome As String, keep As String, desc As String, issues As Collection, newID As String, note As String
    Dim upd As Object, s As String, tg As String
    Set B = ReadKeyed(wbB.Worksheets("DATA_Relationships"), 1, 2, "RelID")
    Set C = ReadKeyed(wbC.Worksheets("DATA_Relationships"), 1, 2, "RelID")
    Set P_ = ReadKeyed(modData.RelLO().Parent, 1, 2, "RelID")
    For Each k In C.Keys
        Set rc = C(k)
        s = T(rc("SourceID")): tg = T(rc("TargetID"))
        desc = T(rc("SourceName")) & " " & modMeta.RelVerb(T(rc("RelCode"))) & " " & T(rc("TargetName"))
        If Not P_.Exists(k) Then
            If B.Exists(k) Then
                AddRow rows, "Relationship", "Info", CStr(k), "", desc, T(rc("RelCode")), "", "", "", "", KEEP_SKIP, _
                       "No longer in the parent - not merged"
            Else
                keep = Decide(decisions, "Relationship", CStr(k), "(new)", KEEP_CHILD)
                note = "New in the child"
                If commit And keep = KEEP_CHILD Then
                    Set issues = New Collection
                    newID = modData.CreateRelationship(T(rc("RelCode")), s, tg, T(rc("Label")), T(rc("Description")), src, batch, _
                                                       issues, CStr(k))
                    If newID = "" Then
                        note = "Not added: " & Replace(modRules.IssueText(issues, 3), vbLf, "; ")
                        cnt(5) = cnt(5) + 1
                    Else
                        If T(rc("Status")) <> "Active" And T(rc("Status")) <> "" Then
                            Dim st As Object
                            Set st = modCore.NewDict()
                            st("Status") = T(rc("Status"))
                            modData.UpdateRelationship newID, st, src, batch
                        End If
                        touched(modData.EntityTypeOf(s)) = True: touched(modData.EntityTypeOf(tg)) = True
                        cnt(3) = cnt(3) + 1
                    End If
                End If
                If commit And keep <> KEEP_CHILD Then cnt(6) = cnt(6) + 1
                AddRow rows, "Relationship", "Add", CStr(k), "", desc, T(rc("RelCode")), "(new)", "", "", desc, keep, note
            End If
        Else
            Set rp = P_(k)
            If B.Exists(k) Then Set rb = B(k) Else Set rb = rp
            Set upd = modCore.NewDict()
            For Each f In Split(REL_FIELDS, "|")
                outcome = ThreeWay(T(Val_(rb, f)), T(Val_(rp, f)), T(Val_(rc, f)))
                If outcome <> "" Then
                    keep = Decide(decisions, "Relationship", CStr(k), CStr(f), IIf(outcome = "child", KEEP_CHILD, KEEP_PARENT))
                    AddRow rows, "Relationship", IIf(outcome = "child", "Update", "Conflict"), CStr(k), "", desc, T(rc("RelCode")), _
                           CStr(f), Val_(rb, f), Val_(rp, f), Val_(rc, f), keep, _
                           IIf(outcome = "child", "Changed in the child", "Changed in both")
                    If keep = KEEP_CHILD Then upd(CStr(f)) = T(Val_(rc, f))
                End If
            Next f
            If commit And upd.Count > 0 Then
                If modData.UpdateRelationship(CStr(k), upd, src, batch) Then
                    cnt(3) = cnt(3) + 1
                    touched(modData.EntityTypeOf(s)) = True: touched(modData.EntityTypeOf(tg)) = True
                Else
                    cnt(5) = cnt(5) + 1
                End If
            End If
        End If
    Next k
End Sub



' ---------------------------------------------------------------- diagrams
' Diagram stamps (name -> size|modified) recorded in a package
Private Function DiagramStamps(wb As Workbook) As Object
    Dim d As Object, ws As Worksheet, v As Variant, i As Long
    Set d = modCore.NewDict()
    Set DiagramStamps = d
    If Not modCore.SheetExists(wb, "Diagrams") Then Exit Function
    Set ws = wb.Worksheets("Diagrams")
    If modCore.LastRowOf(ws) < 2 Then Exit Function
    v = ws.Range("A2:C" & modCore.LastRowOf(ws)).Value
    For i = 1 To UBound(v, 1)
        If CStr(v(i, 1)) <> "" Then d(LCase$(CStr(v(i, 1)))) = CStr(v(i, 2)) & "|" & CStr(v(i, 3))
    Next i
End Function

' A child diagram is merged when it is new or changed since the base (by its recorded size and date)
Private Sub MergeDiagrams(ByVal childFolder As String, wbB As Workbook, wbC As Workbook, ByVal commit As Boolean, _
                          decisions As Object, rows As Collection, cnt() As Long)
    Dim cDir As String, pDir As String, f As String, target As String, keep As String, action As String, note As String
    Dim bStamps As Object, cStamps As Object, k As Variant, nm As String, ws As Worksheet, v As Variant, i As Long
    cDir = childFolder & Sep() & "Diagrams" & Sep()
    pDir = ThisWorkbook.Path & Sep() & "Diagrams" & Sep()
    If Not Fso().FolderExists(cDir) Then Exit Sub
    EnsureFolder Left$(pDir, Len(pDir) - 1)
    Set bStamps = DiagramStamps(wbB)
    If Not modCore.SheetExists(wbC, "Diagrams") Then Exit Sub
    Set ws = wbC.Worksheets("Diagrams")
    If modCore.LastRowOf(ws) < 2 Then Exit Sub
    v = ws.Range("A2:C" & modCore.LastRowOf(ws)).Value
    For i = 1 To UBound(v, 1)
        f = CStr(v(i, 1))
        target = ""
        If f <> "" And Fso().FileExists(cDir & f) Then
            If Not bStamps.Exists(LCase$(f)) Then
                note = "New in the child"
            ElseIf bStamps(LCase$(f)) <> CStr(v(i, 2)) & "|" & CStr(v(i, 3)) Then
                note = "Changed in the child"
            Else
                note = ""
            End If
            ' already copied by an earlier merge: same size and date as the child's file
            If note <> "" And Fso().FileExists(pDir & f) Then
                If Fso().GetFile(pDir & f).Size = Fso().GetFile(cDir & f).Size And _
                   Fso().GetFile(pDir & f).DateLastModified = Fso().GetFile(cDir & f).DateLastModified Then note = ""
            End If
            If note <> "" Then
                If Fso().FileExists(pDir & f) Then
                    action = "Copy": target = pDir & Left$(f, Len(f) - 7) & "_child.drawio"
                    note = note & " - copied next to the parent's diagram, not merged"
                Else
                    action = "Add": target = pDir & f
                End If
            End If
        End If
        If target <> "" Then
            keep = Decide(decisions, "Diagram", f, "(file)", KEEP_CHILD)
            If commit And keep = KEEP_CHILD Then
                Fso().CopyFile cDir & f, target, True
                cnt(4) = cnt(4) + 1
            End If
            AddRow rows, "Diagram", action, f, "", f, "Diagrams", "(file)", "", "", Mid$(target, InStrRev(target, Sep()) + 1), keep, note
        End If
    Next i
End Sub

Private Sub MarkChildMerged(ByVal childID As String, ByVal childName As String, ByVal folder As String)
    Dim lo As ListObject, d As Variant, i As Long, found As Boolean
    Set lo = ChildrenLO()
    modCore.UnprotectSheet lo.Parent
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, 1)) = childID Then
                lo.ListRows(i).Range.Cells(1, 6).Value = modCore.UtcNow()
                lo.ListRows(i).Range.Cells(1, 7).Value = "Merged"
                found = True
            End If
        Next i
    End If
    If Not found Then modCore.AppendRow lo, Array(childID, childName, folder, "", "", modCore.UtcNow(), "Merged")
    modCore.ProtectSheet lo.Parent
End Sub

Private Sub WriteReview(ByVal pkgPath As String, rows As Collection, ByVal committed As Boolean)
    Dim ws As Worksheet, lo As ListObject, x As Variant, core As String
    Set ws = ThisWorkbook.Worksheets(SH_REVIEW)
    Set lo = ReviewLO()
    core = "'" & ThisWorkbook.Name & "'!modMerge."
    modCore.BeginBusy
    modCore.UnprotectSheet ws
    modCore.DeleteButtons ws
    modCore.ClearTable lo
    For Each x In rows
        modCore.AppendRow lo, x
    Next x
    ws.Range("A3").Value = IIf(committed, "Applied: ", "Preview: ") & rows.Count & " row(s)"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value = "Package"
    ws.Range("B4").Value = IIf(committed, "", pkgPath & "*" & FileStamp(pkgPath))   ' cleared after applying
    ws.Range("B4").Font.Color = RGB(89, 89, 89)
    If Not lo.DataBodyRange Is Nothing Then
        lo.ListColumns(11).DataBodyRange.Locked = committed
        modGenerator.ApplyListValidation lo.ListColumns(11).DataBodyRange, "MergeKeep"
    End If
    If Not committed Then modCore.AddButton ws, "Apply merge", core & "MergeApply", ws.Range("H1"), 150
    modCore.AddButton ws, "Parent and child models", core & "ShowMerge", ws.Range("J1"), 170
    modCore.AddButton ws, "Back to Home", "'" & ThisWorkbook.Name & "'!modForms.GoHome", ws.Range("M1"), 120
    modCore.ProtectSheet ws
    modCore.EndBusy
    ThisWorkbook.Activate
    ws.Activate
End Sub
