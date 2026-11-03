Attribute VB_Name = "modValidate"
' =====================================================================
' Validate All: checks every entity row and relationship against the
' metamodel and the rules, highlights failing cells in DATA_Master and
' lists every issue on SYS_Issues with a link to the cell.
' =====================================================================
Option Explicit

Public Sub ValidateAll()
    modCore.Heal
    RunValidation False
End Sub

' Runs the checks. quiet = no message boxes (used by Sync Now). Returns a one-line summary ("" if clean).
Public Function RunValidation(ByVal quiet As Boolean) As String
    Dim ws As Worksheet, lo As ListObject, cols As Object, d As Variant, i As Long, base As Long
    Dim out As New Collection, nErr As Long, nWarn As Long, nameIdx As Object
    On Error GoTo Fail
    modCore.BeginBusy
    modCore.Status "validating entities"
    modMeta.LoadMeta True
    Set ws = modData.MasterSheet(): Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    base = lo.Range.Column - 1
    d = modCore.TableData(lo)
    Set nameIdx = modData.BuildNameIndex()
    modCore.UnprotectSheet ws

    ' clear previous highlights on business columns
    Dim k As Variant, firstRow As Long, lastRow As Long
    If Not lo.DataBodyRange Is Nothing Then
        firstRow = lo.DataBodyRange.Row: lastRow = firstRow + lo.DataBodyRange.Rows.Count - 1
        For Each k In cols.Keys
            Select Case k
                Case "EntityID", "TypeCode", "DisplayID", "Created", "CreatedBy", "Modified", "ModifiedBy", "RowHash"
                Case Else
                    ws.Range(ws.Cells(firstRow, cols(k)), ws.Cells(lastRow, cols(k))).Interior.ColorIndex = xlColorIndexNone
            End Select
        Next k
    End If

    If IsArray(d) Then
        Dim row As Object, issues As Collection, it As Variant, p() As String, typeCode As String, id As String, r As Long
        For i = 1 To UBound(d, 1)
            id = CStr(d(i, cols("EntityID") - base))
            If id <> "" Then
                r = lo.DataBodyRange.Row + i - 1
                typeCode = CStr(d(i, cols("TypeCode") - base))
                Set row = modCore.NewDict()
                For Each k In cols.Keys
                    row(k) = d(i, cols(k) - base)
                Next k
                If Not modMeta.TypeExists(typeCode) Then
                    Set issues = New Collection
                    issues.Add "Error|Type|TypeCode|Unknown entity type " & typeCode
                Else
                    Set issues = modRules.ValidateRow(typeCode, row, nameIdx, id)
                End If
                For Each it In issues
                    p = Split(it, "|", 4)
                    If p(0) = "Error" Then nErr = nErr + 1 Else nWarn = nWarn + 1
                    out.Add Array(modCore.UtcNow(), modCore.SH_MASTER, id, row("DisplayID"), row("Name"), typeCode, _
                                  p(2), p(1), p(0), p(3), CellAddress(r, p(2), cols))
                    If cols.Exists(p(2)) Then
                        ws.Cells(r, cols(p(2))).Interior.Color = IIf(p(0) = "Error", modCore.COLOR_ERROR, modCore.COLOR_WARNING)
                    End If
                Next it
                If CStr(row("RowHash")) = "" Then ws.Cells(r, cols("RowHash")).Value = modData.MasterRowHash(r, cols)
            End If
        Next i
    End If
    modCore.ProtectSheet ws

    ' relationships
    modCore.Status "validating relationships"
    Dim rl As Variant, st As String, tt As String, rid As String, rr As Long, rlo As ListObject
    Set rlo = modData.RelLO()
    rl = modCore.TableData(rlo)
    If IsArray(rl) Then
        For i = 1 To UBound(rl, 1)
            rid = CStr(rl(i, 1))
            If rid <> "" And CStr(rl(i, 11)) <> "Retired" Then
                rr = rlo.DataBodyRange.Row + i - 1
                st = modData.EntityTypeOf(CStr(rl(i, 3)))
                tt = modData.EntityTypeOf(CStr(rl(i, 4)))
                If st = "" Then
                    out.Add RelIssue(rid, rl, i, "Missing", "Error", "Source entity does not exist", rr, 3): nErr = nErr + 1
                ElseIf tt = "" Then
                    out.Add RelIssue(rid, rl, i, "Missing", "Error", "Target entity does not exist", rr, 4): nErr = nErr + 1
                ElseIf Not modMeta.IsRelAllowed(st, CStr(rl(i, 2)), tt) Then
                    out.Add RelIssue(rid, rl, i, "NotAllowed", "Error", modMeta.TypeName_(st) & " " & _
                        modMeta.RelVerb(CStr(rl(i, 2))) & " " & modMeta.TypeName_(tt) & " is not allowed by the metamodel", rr, 2)
                    nErr = nErr + 1
                End If
            End If
        Next i
    End If

    ' write SYS_Issues
    Dim ilo As ListObject, iws As Worksheet, x As Variant, rng As Range
    Set ilo = modCore.Tbl(modCore.SH_ISSUES, "tblIssues")
    Set iws = ilo.Parent
    modCore.UnprotectSheet iws
    modCore.ClearTable ilo
    For Each x In out
        Set rng = modCore.AppendRow(ilo, Array(x(0), x(1), x(2), x(3), x(4), x(5), x(6), x(7), x(8), x(9), Empty))
        If x(10) <> "" Then
            rng.Cells(1, 11).Formula = "=HYPERLINK(""#'" & x(1) & "'!" & x(10) & """,""Go to"")"
        End If
    Next x
    modCore.ProtectSheet iws
    modCore.SetSetting "LastValidateAll", modCore.UtcNow()
    modCore.EndBusy
    If nErr + nWarn > 0 Then RunValidation = nErr & " error(s) and " & nWarn & " warning(s) - see SYS_Issues"
    modCore.Banner RunValidation
    If quiet Then Exit Function
    If Not modRules.RegexAvailable() Then
        MsgBox "Note: the regular-expression engine (VBScript) is not available on this PC, so Regex rules were skipped.", _
               vbInformation, "EA Workbench"
    End If
    If nErr + nWarn = 0 Then
        MsgBox "Validate All: no issues found.", vbInformation, "EA Workbench"
    Else
        If MsgBox("Validate All found " & nErr & " error(s) and " & nWarn & " warning(s). Open the issue list?", _
                  vbYesNo + vbQuestion, "EA Workbench") = vbYes Then
            ThisWorkbook.Activate
            iws.Activate
        End If
    End If
    Exit Function
Fail:
    modCore.ResetBusy
    RunValidation = "Validate All stopped: " & Err.Description
    If Not quiet Then MsgBox RunValidation, vbExclamation, "EA Workbench"
End Function

Private Function CellAddress(ByVal r As Long, ByVal attrCode As String, cols As Object) As String
    If cols.Exists(attrCode) Then CellAddress = modData.MasterSheet().Cells(r, cols(attrCode)).Address(False, False)
End Function

Private Function RelIssue(ByVal rid As String, rl As Variant, ByVal i As Long, ByVal ruleType As String, _
                          ByVal sev As String, ByVal msg As String, ByVal sheetRow As Long, ByVal col As Long) As Variant
    RelIssue = Array(modCore.UtcNow(), modCore.SH_REL, rid, "", CStr(rl(i, 5)) & " -> " & CStr(rl(i, 7)), _
                     CStr(rl(i, 2)), "", ruleType, sev, msg, _
                     ThisWorkbook.Worksheets(modCore.SH_REL).Cells(sheetRow, col).Address(False, False))
End Function
