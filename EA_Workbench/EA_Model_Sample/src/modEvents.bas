Attribute VB_Name = "modEvents"
' =====================================================================
' Event routing.
'   DATA_Master edits      -> validate, stamp, history, repaint type sheet
'   EA_Entities T_ edits   -> validate locally, highlight, queue in SYS_Pending
'   EA_Entities save       -> PushPending: queued rows written to EA_Core
'   FORM_ sheets           -> modForms
' =====================================================================
Option Explicit

Private mOld As Object          ' "book|sheet|address" -> value before the edit
Private Const MAX_SNAPSHOT As Long = 2000

' ---------------------------------------------------------------- selection snapshot (old values for history)
Public Sub OnSelection(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Quiet
    If modCore.IsBusy() Then Exit Sub
    If Not IsTrackedSheet(Sh) Then Exit Sub
    Set mOld = modCore.NewDict()
    If Target.CountLarge > MAX_SNAPSHOT Then Exit Sub
    Dim c As Range
    For Each c In Target.Cells
        mOld(Key(Sh, c)) = c.Value
    Next c
Quiet:
End Sub

Private Function Key(ByVal Sh As Object, ByVal c As Range) As String
    Key = Sh.Parent.Name & "|" & Sh.Name & "|" & c.Address(False, False)
End Function

Private Function OldValue(ByVal Sh As Object, ByVal c As Range, ByRef known As Boolean) As Variant
    known = False
    If mOld Is Nothing Then Exit Function
    If mOld.Exists(Key(Sh, c)) Then OldValue = mOld(Key(Sh, c)): known = True
End Function

Private Function IsTrackedSheet(ByVal Sh As Object) As Boolean
    If Sh.Parent Is ThisWorkbook Then
        IsTrackedSheet = (Sh.Name = modCore.SH_MASTER)
    ElseIf LCase$(Sh.Parent.Name) = LCase$(modCore.EntitiesFileName()) Then
        IsTrackedSheet = (Sh.Name Like "T_*")
    End If
End Function

' ---------------------------------------------------------------- change router
Public Sub OnChange(ByVal Sh As Object, ByVal Target As Range)
    If modCore.IsBusy() Then Exit Sub
    On Error GoTo Fail
    If Sh.Parent Is ThisWorkbook Then
        Select Case Sh.Name
            Case modCore.SH_MASTER: OnMasterChange Sh, Target
            Case "FORM_Entity", "FORM_Relationship", "FORM_Type", "FORM_Rule": modForms.OnFormChange Sh, Target
        End Select
    ElseIf LCase$(Sh.Parent.Name) = LCase$(modCore.EntitiesFileName()) Then
        If Sh.Name Like "T_*" Then OnEntitiesChange Sh, Target
    End If
    Exit Sub
Fail:
    modCore.ResetBusy
    MsgBox "Workbench could not process that edit: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

' ---------------------------------------------------------------- DATA_Master edits
Private Sub OnMasterChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim lo As ListObject, area As Range, c As Range, cols As Object, keyOf As String
    Set lo = modData.MasterLO()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set area = Application.Intersect(Target, lo.DataBodyRange)
    If area Is Nothing Then Exit Sub
    modMeta.LoadMeta
    Set cols = modData.MasterCols()
    If area.CountLarge > 5000 Then
        MsgBox "Large paste detected. Run Validate All to check it; history was not recorded for this paste.", _
               vbInformation, "EA Workbench"
        Exit Sub
    End If

    Dim rowsChanged As Object, msgs As String, known As Boolean, oldV As Variant, r As Long, typeCode As String
    Set rowsChanged = modCore.NewDict()
    modCore.BeginBusy
    For Each c In area.Cells
        r = c.Row
        keyOf = CStr(ws.Cells(modCore.MASTER_KEY_ROW, c.Column).Value)
        typeCode = CStr(ws.Cells(r, cols("TypeCode")).Value)
        oldV = OldValue(ws, c, known)
        Select Case keyOf
            Case "EntityID", "TypeCode", "DisplayID", "Created", "CreatedBy", "Modified", "ModifiedBy", "RowHash"
                If known Then c.Value = oldV
                msgs = msgs & vbLf & "System column " & keyOf & " cannot be edited here."
            Case Else
                If Not modMeta.IsSystemAttr(keyOf) And Not modMeta.IsMapped(typeCode, keyOf) _
                   And Not modCore.IsBlankValue(c.Value) Then
                    If known Then c.Value = oldV Else c.ClearContents
                    msgs = msgs & vbLf & modMeta.AttrName(keyOf) & " does not apply to " & modMeta.TypeName_(typeCode) & "."
                Else
                    If Not rowsChanged.Exists(r) Then rowsChanged.Add r, modCore.NewDict()
                    rowsChanged(r)(keyOf) = Array(IIf(known, oldV, Empty), known)
                End If
        End Select
    Next c

    Dim rk As Variant, row As Object, issues As Collection, ak As Variant, worst As String, entityID As String
    Dim batch As String, touched As Object, nameIdx As Object
    batch = modCore.NewGUID()
    Set touched = modCore.NewDict()
    Set nameIdx = modData.BuildNameIndex()
    For Each rk In rowsChanged.Keys
        r = CLng(rk)
        Set row = modData.ReadEntityRow(r, cols)
        typeCode = CStr(row("TypeCode")): entityID = CStr(row("EntityID"))
        Set issues = modRules.ValidateRow(typeCode, row, nameIdx, entityID)
        For Each ak In rowsChanged(rk).Keys
            worst = modRules.WorstFor(issues, CStr(ak))
            Set c = ws.Cells(r, cols(ak))
            If worst = "Error" Then
                If rowsChanged(rk)(ak)(1) Then
                    c.Value = rowsChanged(rk)(ak)(0)
                    msgs = msgs & vbLf & modMeta.AttrName(CStr(ak)) & ": " & FirstMessage(issues, CStr(ak)) & " (change undone)"
                Else
                    c.Interior.Color = modCore.COLOR_ERROR
                    msgs = msgs & vbLf & modMeta.AttrName(CStr(ak)) & ": " & FirstMessage(issues, CStr(ak)) & _
                           " (please correct the highlighted cell)"
                End If
                rowsChanged(rk).Remove ak
            Else
                If worst = "Warning" Then
                    c.Interior.Color = modCore.COLOR_WARNING
                Else
                    c.Interior.ColorIndex = xlColorIndexNone
                End If
                modData.WriteHistory typeCode, "Update", entityID, CStr(ak), rowsChanged(rk)(ak)(0), c.Value, _
                                     "DATA_Master", batch
            End If
        Next ak
        If rowsChanged(rk).Count > 0 Then
            modData.StampRow r, False, cols
            touched(typeCode) = True
            ' related entities show this name on their sheets too
            If rowsChanged(rk).Exists("Name") Then AddRelatedTypes entityID, touched
        End If
    Next rk
    modCore.EndBusy
    If touched.Count > 0 Then modPaint.PaintTypes touched
    If msgs <> "" Then MsgBox Mid$(msgs, 2), vbExclamation, "EA Workbench"
End Sub

Private Function FirstMessage(issues As Collection, ByVal attrCode As String) As String
    Dim it As Variant, p() As String
    For Each it In issues
        p = Split(it, "|", 4)
        If p(2) = attrCode And p(0) = "Error" Then FirstMessage = p(3): Exit Function
    Next it
End Function

Private Sub AddRelatedTypes(ByVal entityID As String, touched As Object)
    Dim d As Variant, i As Long, other As String
    d = modCore.TableData(modData.RelLO())
    If Not IsArray(d) Then Exit Sub
    For i = 1 To UBound(d, 1)
        other = ""
        If d(i, 3) = entityID Then other = CStr(d(i, 4))
        If d(i, 4) = entityID Then other = CStr(d(i, 3))
        If other <> "" Then touched(modData.EntityTypeOf(other)) = True
    Next i
End Sub

' ---------------------------------------------------------------- EA_Entities edits
Public Function TypeOfSheet(ByVal sheetName As String) As String
    Dim t As Variant
    modMeta.LoadMeta
    For Each t In modMeta.gTypes.Keys
        If StrComp(modMeta.gTypes(t)("Sheet"), sheetName, vbTextCompare) = 0 Then TypeOfSheet = CStr(t): Exit Function
    Next t
End Function

Private Sub OnEntitiesChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim area As Range, c As Range, typeCode As String, lastCol As Long, keyOf As String
    Dim rowsSeen As Object, row As Object, issues As Collection, worst As String, r As Long
    typeCode = TypeOfSheet(ws.Name)
    If typeCode = "" Then Exit Sub
    lastCol = ws.Cells(modCore.TS_KEY_ROW, ws.Columns.Count).End(xlToLeft).Column
    Set area = Application.Intersect(Target, ws.Range(ws.Cells(modCore.TS_FIRST_ROW, 5), ws.Cells(ws.Rows.Count, lastCol)))
    If area Is Nothing Then Exit Sub
    If area.CountLarge > 5000 Then
        MsgBox "Large paste: those rows are queued and will be checked when you save.", vbInformation, "EA Workbench"
    End If
    Set rowsSeen = modCore.NewDict()
    Dim nameIdx As Object
    Set nameIdx = modData.BuildNameIndex()
    modCore.BeginBusy
    On Error GoTo Done
    For Each c In area.Cells
        r = c.Row
        If CStr(ws.Cells(r, 2).Value) <> "R" Then
            keyOf = CStr(ws.Cells(modCore.TS_KEY_ROW, c.Column).Value)
            If Not rowsSeen.Exists(r) Then Set rowsSeen(r) = SheetRow(ws, r, lastCol)
            Set row = rowsSeen(r)
            If area.CountLarge <= 5000 Then
                Set issues = modRules.EvaluateValue(typeCode, keyOf, c.Value, row, nameIdx, CStr(ws.Cells(r, 1).Value))
                worst = ""
                If issues.Count > 0 Then worst = Split(issues(1), "|")(0)
                If modRules.HasErrors(issues) Then worst = "Error"
                MarkCell c, worst, modRules.IssueText(issues)
            End If
            RecordPending ws.Name, r, CStr(ws.Cells(r, 1).Value), keyOf, c.Value
        End If
    Next c
Done:
    modCore.EndBusy
    modCore.Status PendingCount() & " edit(s) waiting - they go to EA_Core when you save EA_Entities"
End Sub

Private Function SheetRow(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long) As Object
    Dim d As Object, c As Long, k As String
    Set d = modCore.NewDict()
    For c = 4 To lastCol
        k = CStr(ws.Cells(modCore.TS_KEY_ROW, c).Value)
        If k <> "" Then d(k) = ws.Cells(r, c).Value
    Next c
    Set SheetRow = d
End Function

Private Sub MarkCell(c As Range, ByVal worst As String, ByVal msg As String)
    On Error Resume Next
    If Not c.Comment Is Nothing Then c.Comment.Delete
    Select Case worst
        Case "Error"
            c.Interior.Color = modCore.COLOR_ERROR
            c.AddComment msg
        Case "Warning"
            c.Interior.Color = modCore.COLOR_WARNING
            c.AddComment msg
        Case Else
            c.Interior.ColorIndex = xlColorIndexNone
    End Select
End Sub

Private Function PendingLO() As ListObject
    Dim wbE As Workbook
    Set wbE = modCore.EntitiesWB(False)
    If wbE Is Nothing Then Exit Function
    Set PendingLO = wbE.Worksheets("SYS_Pending").ListObjects("tblPending")
End Function

Private Sub RecordPending(ByVal sheetName As String, ByVal r As Long, ByVal entityID As String, _
                          ByVal attrCode As String, ByVal v As Variant)
    Dim lo As ListObject
    Set lo = PendingLO()
    If lo Is Nothing Then Exit Sub
    modCore.AppendRow lo, Array(sheetName, r, entityID, attrCode, "'" & modCore.CellText(v), modCore.UtcNow())
End Sub

' Sheet names with edits queued in SYS_Pending (these must not be repainted)
Public Function PendingSheets() As Object
    Dim d As Object, lo As ListObject, v As Variant, i As Long
    Set d = modCore.NewDict()
    Set PendingSheets = d
    Set lo = PendingLO()
    If lo Is Nothing Then Exit Function
    v = modCore.TableData(lo)
    If Not IsArray(v) Then Exit Function
    For i = 1 To UBound(v, 1)
        If CStr(v(i, 1)) <> "" Then d(CStr(v(i, 1))) = True
    Next i
End Function

' Keys of queued edits: "sheet|row" and "id:<EntityID>"
Public Function PendingKeys() As Object
    Dim d As Object, lo As ListObject, v As Variant, i As Long
    Set d = modCore.NewDict()
    Set PendingKeys = d
    Set lo = PendingLO()
    If lo Is Nothing Then Exit Function
    v = modCore.TableData(lo)
    If Not IsArray(v) Then Exit Function
    For i = 1 To UBound(v, 1)
        d(CStr(v(i, 1)) & "|" & CStr(v(i, 2))) = True
        If CStr(v(i, 3)) <> "" Then d("id:" & CStr(v(i, 3))) = True
    Next i
End Function

Public Function PendingCount() As Long
    Dim lo As ListObject, d As Variant
    Set lo = PendingLO()
    If lo Is Nothing Then Exit Function
    d = modCore.TableData(lo)
    If IsArray(d) Then PendingCount = UBound(d, 1)
End Function

' ---------------------------------------------------------------- push EA_Entities -> EA_Core
' Writes every queued row through modData (validation, stamps, history).
' Rows with errors stay queued and highlighted; their sheet is not repainted.
Public Sub PushPending(Optional ByVal fromSave As Boolean = False, Optional ByVal noPaint As Boolean = False)
    Dim wbE As Workbook, lo As ListObject, d As Variant, i As Long, k As String
    modCore.Heal
    On Error GoTo Crash
    Set wbE = modCore.EntitiesWB(False)
    If wbE Is Nothing Then Exit Sub
    Set lo = wbE.Worksheets("SYS_Pending").ListObjects("tblPending")
    d = modCore.TableData(lo)
    If Not IsArray(d) Then
        If Not fromSave Then MsgBox "Nothing waiting to push.", vbInformation, "EA Workbench"
        Exit Sub
    End If

    ' rows to push, and which attributes were edited on each (only those are sent)
    Dim rowsTodo As Object, attrsOf As Object
    Set rowsTodo = modCore.NewDict(): Set attrsOf = modCore.NewDict()
    For i = 1 To UBound(d, 1)
        k = CStr(d(i, 1)) & "|" & CStr(d(i, 2))
        If Not rowsTodo.Exists(k) Then rowsTodo(k) = True
        If Not attrsOf.Exists(k) Then attrsOf.Add k, modCore.NewDict()
        attrsOf(k)(CStr(d(i, 4))) = True
    Next i

    Dim rk As Variant, parts() As String, ws As Worksheet, r As Long, typeCode As String, lastCol As Long
    Dim values As Object, issues As Collection, entityID As String, batch As String, newID As String
    Dim okRows As Object, failTypes As Object, okTypes As Object, msgs As String, nOK As Long, nFail As Long
    Set okRows = modCore.NewDict(): Set failTypes = modCore.NewDict(): Set okTypes = modCore.NewDict()
    batch = modCore.NewGUID()
    modMeta.LoadMeta
    For Each rk In rowsTodo.Keys
        parts = Split(rk, "|")
        If modCore.SheetExists(wbE, parts(0)) Then
            Set ws = wbE.Worksheets(parts(0))
            r = CLng(parts(1))
            typeCode = TypeOfSheet(ws.Name)
            lastCol = ws.Cells(modCore.TS_KEY_ROW, ws.Columns.Count).End(xlToLeft).Column
            Set values = SheetRow(ws, r, lastCol)
            values.Remove "DisplayID"
            entityID = CStr(ws.Cells(r, 1).Value)
            Set issues = New Collection
            If CStr(ws.Cells(r, 2).Value) = "R" Then
                okRows(rk) = True                                ' relationship rows are read-only
            ElseIf entityID = "" Then
                If RowIsEmpty(values) Then
                    okRows(rk) = True
                Else
                    newID = modData.CreateEntity(typeCode, values, "EA_Entities", batch, issues)
                    If newID <> "" Then
                        modCore.UnprotectSheet ws
                        ws.Cells(r, 1).Value = newID: ws.Cells(r, 2).Value = "E"
                        modCore.ProtectSheet ws
                        okRows(rk) = True: okTypes(typeCode) = True: nOK = nOK + 1
                    End If
                End If
            Else
                Dim blocked As Boolean, edited As Object, ak As Variant
                Set edited = modCore.NewDict()
                For Each ak In attrsOf(rk).Keys
                    If values.Exists(ak) Then edited(ak) = values(ak)
                Next ak
                blocked = False
                modData.UpdateEntity entityID, edited, "EA_Entities", batch, issues, blocked
                If Not blocked Then
                    okRows(rk) = True: okTypes(typeCode) = True: nOK = nOK + 1
                    If edited.Exists("Name") Then AddRelatedTypes entityID, okTypes
                End If
            End If
            If Not okRows.Exists(rk) Then
                nFail = nFail + 1
                failTypes(typeCode) = True
                msgs = msgs & vbLf & ws.Name & " row " & r & ": " & Replace(modRules.IssueText(issues, 3), vbLf, "; ")
            End If
        Else
            okRows(rk) = True
        End If
    Next rk

    ' keep only the failed rows in the queue (if anything fails here, Crash restores the full queue)
    Dim queueCleared As Boolean
    modCore.BeginBusy
    Dim keep As New Collection
    For i = 1 To UBound(d, 1)
        k = CStr(d(i, 1)) & "|" & CStr(d(i, 2))
        If Not okRows.Exists(k) Then keep.Add Array(d(i, 1), d(i, 2), d(i, 3), d(i, 4), "'" & d(i, 5), d(i, 6))
    Next i
    modCore.ClearTable lo
    queueCleared = True
    Dim x As Variant
    For Each x In keep
        modCore.AppendRow lo, x
    Next x
    queueCleared = False
    ' repaint types that pushed cleanly (types with failures keep the user's edits on screen)
    Dim t As Variant
    For Each t In failTypes.Keys
        If okTypes.Exists(t) Then okTypes.Remove t
    Next t
    modGenerator.SetSyncState wbE, "LastPush", modCore.UtcNow()
    modCore.EndBusy
    If okTypes.Count > 0 And Not noPaint Then modPaint.PaintTypes okTypes
    If nOK > 0 Then
        modSync.SaveWithStamp ThisWorkbook
        ' a push from the button also saves EA_Entities, so new rows keep their IDs
        If Not fromSave Then modSync.SaveWithStamp wbE
    End If
    If nFail > 0 Then
        MsgBox nOK & " row(s) saved to EA_Core. " & nFail & " row(s) have errors and are still waiting:" & _
               Left$(msgs, 1500), vbExclamation, "EA Workbench"
    ElseIf Not fromSave Then
        MsgBox nOK & " row(s) saved to EA_Core.", vbInformation, "EA Workbench"
    End If
    modCore.Status IIf(nFail > 0, nFail & " row(s) still waiting", "EA_Core is up to date")
    Exit Sub
Crash:
    Dim crashMsg As String
    crashMsg = Err.Description
    On Error Resume Next
    If queueCleared Then                      ' put the whole original queue back
        modCore.ClearTable lo
        For i = 1 To UBound(d, 1)
            modCore.AppendRow lo, Array(d(i, 1), d(i, 2), d(i, 3), d(i, 4), "'" & d(i, 5), d(i, 6))
        Next i
    End If
    modCore.ResetBusy
    MsgBox "Saving edits to EA_Core stopped: " & crashMsg & vbLf & _
           "Your edits are still queued; save EA_Entities again to retry.", vbExclamation, "EA Workbench"
End Sub

Private Function RowIsEmpty(values As Object) As Boolean
    Dim k As Variant
    RowIsEmpty = True
    For Each k In values.Keys
        If k <> "Status" And k <> "RepositoryArea" Then
            If Not modCore.IsBlankValue(values(k)) Then RowIsEmpty = False: Exit Function
        End If
    Next k
End Function

