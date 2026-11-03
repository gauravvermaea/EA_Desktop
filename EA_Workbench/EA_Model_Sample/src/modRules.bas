Attribute VB_Name = "modRules"
' =====================================================================
' Rules engine. One generic routine checks any value against its
' attribute's data type, its mapping (applicable / mandatory) and every
' active rule in MM_Rules. Mirrors build/model.py evaluate().
'
' Issues are returned as strings "Severity|RuleType|AttrCode|Message".
' =====================================================================
Option Explicit

Private mRegex As Object
Private mRegexChecked As Boolean
Private mRegexAvailable As Boolean

' Checks one value. row: Dictionary AttrCode -> value for the whole entity.
' nameIndex (optional): Dictionary "Type|lower name" -> Dictionary of EntityIDs, for Unique.
Public Function EvaluateValue(ByVal typeCode As String, ByVal attrCode As String, ByVal value As Variant, _
                              row As Object, Optional nameIndex As Object = Nothing, _
                              Optional ByVal entityID As String = "") As Collection
    Dim issues As New Collection
    Set EvaluateValue = issues
    modMeta.LoadMeta
    If Not modMeta.gAttrs.Exists(attrCode) Then Exit Function

    Dim a As Object, blank As Boolean, isSys As Boolean, nm As String, dt As String
    Set a = modMeta.gAttrs(attrCode)
    nm = a("Name"): dt = a("DataType"): isSys = a("IsSystem")
    blank = modCore.IsBlankValue(value)

    If Not isSys Then
        If Not modMeta.IsMapped(typeCode, attrCode) Then
            If Not blank Then issues.Add "Error|NotApplicable|" & attrCode & "|" & nm & " does not apply to " & _
                                         modMeta.TypeName_(typeCode)
            Exit Function
        End If
        If blank And modMeta.IsMandatory(typeCode, attrCode) Then
            issues.Add "Error|Mandatory|" & attrCode & "|" & nm & " is required for " & modMeta.TypeName_(typeCode)
        End If
    End If

    ' built-in data type checks
    If Not blank Then
        Select Case dt
            Case "Integer"
                If Not IsWhole(value) Then issues.Add "Error|DataType|" & attrCode & "|" & nm & " must be a whole number"
            Case "Decimal", "Currency", "Percent"
                If Not IsNumeric(value) Then issues.Add "Error|DataType|" & attrCode & "|" & nm & " must be a number"
            Case "Date"
                If Not IsDateValue(value) Then issues.Add "Error|DataType|" & attrCode & "|" & nm & " must be a date"
            Case "Boolean"
                Select Case UCase$(modCore.CellText(value))
                    Case "Y", "N", "TRUE", "FALSE"
                    Case Else: issues.Add "Error|DataType|" & attrCode & "|" & nm & " must be Y or N"
                End Select
            Case "List"
                If a("ListName") <> "" Then
                    If Not modMeta.InList(a("ListName"), modCore.CellText(value)) Then
                        issues.Add "Error|List|" & attrCode & "|" & nm & " must be one of the " & a("ListName") & " values"
                    End If
                End If
            Case "URL"
                If Not RegexTest("^https?://\S+$", modCore.CellText(value)) Then
                    issues.Add "Warning|DataType|" & attrCode & "|" & nm & " should be a web address"
                End If
            Case "Email"
                If Not RegexTest("^[^@\s]+@[^@\s]+\.[^@\s]+$", modCore.CellText(value)) Then
                    issues.Add "Warning|DataType|" & attrCode & "|" & nm & " should be an email address"
                End If
        End Select
    End If

    ' rules from MM_Rules
    Dim r As Variant, rt As String, p1 As String, p2 As String, fail As Boolean
    For Each r In modMeta.gRules
        If r(0) = attrCode And (r(1) = "" Or r(1) = typeCode) Then
            rt = r(2): p1 = r(3): p2 = r(4): fail = False
            Select Case rt
                Case "Mandatory"
                    fail = blank
                Case "RequiredIf"
                    If blank Then fail = (modCore.CellText(RowValue(row, p1)) = p2)
                Case Else
                    If Not blank Then fail = Not RulePasses(rt, p1, p2, value, row, typeCode, nameIndex, entityID)
            End Select
            If fail Then issues.Add r(5) & "|" & rt & "|" & attrCode & "|" & r(6)
        End If
    Next r
End Function

' One non-blank rule check (also used by the rule wizard's Test button)
Public Function RulePasses(ByVal rt As String, ByVal p1 As String, ByVal p2 As String, ByVal value As Variant, _
                           row As Object, ByVal typeCode As String, nameIndex As Object, _
                           ByVal entityID As String) As Boolean
    Dim s As String, n As Double
    s = modCore.CellText(value)
    RulePasses = True
    Select Case rt
        Case "List"
            RulePasses = modMeta.InList(p1, s)
        Case "MinValue"
            If Not IsNumeric(value) Then RulePasses = False Else RulePasses = (CDbl(value) >= Val(p1))
        Case "MaxValue"
            If Not IsNumeric(value) Then RulePasses = False Else RulePasses = (CDbl(value) <= Val(p1))
        Case "Range"
            If Not IsNumeric(value) Then
                RulePasses = False
            Else
                n = CDbl(value)
                RulePasses = (n >= Val(p1) And n <= Val(p2))
            End If
        Case "MinLength"
            RulePasses = (Len(s) >= Val(p1))
        Case "MaxLength"
            RulePasses = (Len(s) <= Val(p1))
        Case "Regex"
            RulePasses = RegexTest(p1, s)
        Case "DateRange"
            If Not IsDateValue(value) Then
                RulePasses = False
            Else
                RulePasses = (ToDate(value) >= ToDate(p1) And ToDate(value) <= ToDate(p2))
            End If
        Case "DateOrder"
            Dim other As Variant
            other = RowValue(row, p1)
            If IsDateValue(value) And IsDateValue(other) Then RulePasses = (ToDate(value) > ToDate(other))
        Case "Unique"
            If Not nameIndex Is Nothing Then
                Dim k As String
                k = IIf(p1 = "Type", typeCode, "*") & "|" & LCase$(Trim$(s))
                If nameIndex.Exists(k) Then
                    Dim ids As Object, id As Variant
                    Set ids = nameIndex(k)
                    For Each id In ids.Keys
                        If CStr(id) <> entityID Then RulePasses = False: Exit For
                    Next id
                End If
            End If
        Case "EntityRef"
            ' p1 = comma separated allowed TypeCodes; value = EntityID or exact Name
            RulePasses = modData.EntityRefOK(s, p1)
    End Select
End Function

' Checks a whole entity row. Returns all issues.
Public Function ValidateRow(ByVal typeCode As String, row As Object, Optional nameIndex As Object = Nothing, _
                            Optional ByVal entityID As String = "") As Collection
    Dim all As New Collection, c As Collection, it As Variant, k As Variant
    modMeta.LoadMeta
    For Each k In Array("Name", "Description", "Status", "RepositoryArea")
        Set c = EvaluateValue(typeCode, CStr(k), RowValue(row, CStr(k)), row, nameIndex, entityID)
        For Each it In c: all.Add it: Next it
    Next k
    ' every non-system attribute present in the row or mapped to the type
    Dim seen As Object
    Set seen = modCore.NewDict()
    Dim ac As Variant
    For Each ac In modMeta.AttrsForType(typeCode)
        seen(ac) = True
        Set c = EvaluateValue(typeCode, CStr(ac), RowValue(row, CStr(ac)), row, nameIndex, entityID)
        For Each it In c: all.Add it: Next it
    Next ac
    For Each k In row.Keys
        If Not seen.Exists(k) And Not modMeta.IsSystemAttr(CStr(k)) Then
            If modMeta.gAttrs.Exists(k) Then
                Set c = EvaluateValue(typeCode, CStr(k), row(k), row, nameIndex, entityID)
                For Each it In c: all.Add it: Next it
            End If
        End If
    Next k
    Set ValidateRow = all
End Function

Public Function HasErrors(issues As Collection) As Boolean
    Dim it As Variant
    For Each it In issues
        If Left$(it, 6) = "Error|" Then HasErrors = True: Exit Function
    Next it
End Function

Public Function IssueText(issues As Collection, Optional ByVal maxItems As Long = 12) As String
    Dim it As Variant, parts() As String, s As String, n As Long
    For Each it In issues
        n = n + 1
        If n > maxItems Then s = s & vbLf & "...and " & (issues.Count - maxItems) & " more": Exit For
        parts = Split(it, "|", 4)
        s = s & IIf(s = "", "", vbLf) & parts(0) & ": " & parts(3)
    Next it
    IssueText = s
End Function

' Worst severity for one attribute in an issue collection: "", "Warning" or "Error"
Public Function WorstFor(issues As Collection, ByVal attrCode As String) As String
    Dim it As Variant, parts() As String
    For Each it In issues
        parts = Split(it, "|", 4)
        If parts(2) = attrCode Then
            If parts(0) = "Error" Then WorstFor = "Error": Exit Function
            WorstFor = "Warning"
        End If
    Next it
End Function

' ---------------------------------------------------------------- helpers
Private Function RowValue(row As Object, ByVal key As String) As Variant
    If row Is Nothing Then Exit Function
    If row.Exists(key) Then RowValue = row(key)
End Function

Public Function IsWhole(ByVal v As Variant) As Boolean
    If Not IsNumeric(v) Then Exit Function
    If VarType(v) = vbString Then
        If InStr(v, ".") > 0 Or InStr(v, ",") > 0 Then Exit Function
    End If
    IsWhole = (CDbl(v) = Fix(CDbl(v)))
End Function

Public Function IsDateValue(ByVal v As Variant) As Boolean
    If VarType(v) = vbDate Then IsDateValue = True: Exit Function
    If IsEmpty(v) Then Exit Function
    Dim s As String
    s = CStr(v)
    If s Like "####-##-##*" Then
        On Error Resume Next
        Dim d As Date
        d = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
        IsDateValue = (Err.Number = 0) And Month(d) = CInt(Mid$(s, 6, 2))
        On Error GoTo 0
    ElseIf IsNumeric(v) Then
        IsDateValue = False
    Else
        IsDateValue = IsDate(v)
    End If
End Function

Public Function ToDate(ByVal v As Variant) As Date
    If VarType(v) = vbDate Then ToDate = v: Exit Function
    Dim s As String
    s = CStr(v)
    If s Like "####-##-##*" Then
        ToDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    Else
        ToDate = CDate(v)
    End If
End Function

' Regex via VBScript.RegExp. If the engine is unavailable the check passes
' (and Validate All reports once that regex rules were skipped).
Public Function RegexTest(ByVal pattern As String, ByVal text As String) As Boolean
    If Not mRegexChecked Then
        mRegexChecked = True
        On Error Resume Next
        Set mRegex = CreateObject("VBScript.RegExp")
        mRegexAvailable = (Err.Number = 0) And Not mRegex Is Nothing
        On Error GoTo 0
    End If
    If Not mRegexAvailable Then RegexTest = True: Exit Function
    On Error GoTo BadPattern
    mRegex.Pattern = pattern
    mRegex.IgnoreCase = False
    mRegex.Global = False
    RegexTest = mRegex.Test(text)
    Exit Function
BadPattern:
    RegexTest = True
End Function

Public Function RegexAvailable() As Boolean
    Call RegexTest("a", "a")
    RegexAvailable = mRegexAvailable
End Function
