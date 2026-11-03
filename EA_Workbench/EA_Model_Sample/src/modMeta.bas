Attribute VB_Name = "modMeta"
' =====================================================================
' Metamodel cache. Loads the MM_ sheets into dictionaries so rule checks,
' painting and forms never read metamodel cells one at a time.
' Mirrors build/model.py (Meta class).
' =====================================================================
Option Explicit

Public gTypes As Object        ' TypeCode -> Dictionary(Name, Domain, Aspect, Prefix, Sheet, Status, Fill)
Public gAttrs As Object        ' AttrCode -> Dictionary(Name, DataType, ListName, Unit, Default, IsSystem, Status, Desc)
Public gRelTypes As Object     ' RelCode  -> Dictionary(Name, Category, Directed, Status)
Public gTypeAttrs As Object    ' TypeCode -> Collection of AttrCodes in DisplayOrder (active mappings)
Public gMapInfo As Object      ' "Type|Attr" -> Array(Mandatory, DisplayOrder, Section)
Public gGroups As Object       ' GroupName -> Dictionary of TypeCodes
Public gAllowed As Collection  ' Array(SourceToken, RelCode, TargetToken, MinCard, MaxCard)
Public gRules As Collection    ' Array(AttrCode, TypeCode, RuleType, P1, P2, Severity, Message, RuleID)
Public gLists As Object        ' ListName -> Collection of values (ordered)
Public gListSets As Object     ' "List|Value" -> True
Private mLoaded As Boolean

Public Sub LoadMeta(Optional ByVal force As Boolean = False)
    If mLoaded And Not force Then Exit Sub
    Dim d As Variant, i As Long, lo As ListObject, x As Object, key As String

    ' ---- entity types
    Set gTypes = modCore.NewDict()
    Set lo = modCore.Tbl("MM_EntityTypes", "tblEntityTypes")
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, Col(lo, "TypeCode"))) <> "" Then
                Set x = modCore.NewDict()
                x("Name") = CStr(d(i, Col(lo, "Name")))
                x("Domain") = CStr(d(i, Col(lo, "Domain")))
                x("Aspect") = CStr(d(i, Col(lo, "Aspect")))
                x("Prefix") = CStr(d(i, Col(lo, "IDPrefix")))
                x("Sheet") = CStr(d(i, Col(lo, "SheetName")))
                x("Status") = CStr(d(i, Col(lo, "Status")))
                x("StyleKey") = CStr(d(i, Col(lo, "StyleKey")))
                x("Fill") = ""
                Set gTypes(CStr(d(i, Col(lo, "TypeCode")))) = x
            End If
        Next i
    End If

    ' ---- domain colours from styles
    Set lo = modCore.Tbl("MM_Styles", "tblStyles")
    d = modCore.TableData(lo)
    If IsArray(d) Then
        Dim tc As Variant
        For Each tc In gTypes.Keys
            For i = 1 To UBound(d, 1)
                If CStr(d(i, 1)) = gTypes(tc)("StyleKey") Then gTypes(tc)("Fill") = CStr(d(i, 3)): Exit For
            Next i
        Next tc
    End If

    ' ---- attribute types
    Set gAttrs = modCore.NewDict()
    Set lo = modCore.Tbl("MM_AttributeTypes", "tblAttributeTypes")
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, Col(lo, "AttrCode"))) <> "" Then
                Set x = modCore.NewDict()
                x("Name") = CStr(d(i, Col(lo, "Name")))
                x("DataType") = CStr(d(i, Col(lo, "DataType")))
                x("ListName") = CStr(d(i, Col(lo, "ListName")))
                x("Unit") = CStr(d(i, Col(lo, "Unit")))
                x("Default") = d(i, Col(lo, "DefaultValue"))
                x("Desc") = CStr(d(i, Col(lo, "Description")))
                x("IsSystem") = (UCase$(CStr(d(i, Col(lo, "IsSystem")))) = "Y")
                x("Status") = CStr(d(i, Col(lo, "Status")))
                Set gAttrs(CStr(d(i, Col(lo, "AttrCode")))) = x
            End If
        Next i
    End If

    ' ---- relationship types
    Set gRelTypes = modCore.NewDict()
    Set lo = modCore.Tbl("MM_RelationshipTypes", "tblRelTypes")
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, Col(lo, "RelCode"))) <> "" Then
                Set x = modCore.NewDict()
                x("Name") = CStr(d(i, Col(lo, "Name")))
                x("Category") = CStr(d(i, Col(lo, "Category")))
                x("Directed") = CStr(d(i, Col(lo, "Directed")))
                x("Status") = CStr(d(i, Col(lo, "Status")))
                Set gRelTypes(CStr(d(i, Col(lo, "RelCode")))) = x
            End If
        Next i
    End If

    ' ---- entity-attribute mappings (ordered by DisplayOrder)
    Set gTypeAttrs = modCore.NewDict()
    Set gMapInfo = modCore.NewDict()
    Set lo = modCore.Tbl("MM_EntityAttributes", "tblEntityAttributes")
    d = modCore.TableData(lo)
    If IsArray(d) Then
        Dim cT As Long, cA As Long, cM As Long, cO As Long, cS As Long, cSt As Long
        cT = Col(lo, "TypeCode"): cA = Col(lo, "AttrCode"): cM = Col(lo, "Mandatory")
        cO = Col(lo, "DisplayOrder"): cS = Col(lo, "Section"): cSt = Col(lo, "Status")
        ' insertion sort into per-type collections by DisplayOrder
        For i = 1 To UBound(d, 1)
            If CStr(d(i, cSt)) = "Active" And CStr(d(i, cT)) <> "" And gAttrs.Exists(CStr(d(i, cA))) Then
                key = CStr(d(i, cT)) & "|" & CStr(d(i, cA))
                gMapInfo(key) = Array(UCase$(CStr(d(i, cM))) = "Y", Val(CStr(d(i, cO))), CStr(d(i, cS)))
                If Not gTypeAttrs.Exists(CStr(d(i, cT))) Then gTypeAttrs.Add CStr(d(i, cT)), New Collection
                InsertOrdered gTypeAttrs(CStr(d(i, cT))), CStr(d(i, cA)), Val(CStr(d(i, cO))), CStr(d(i, cT))
            End If
        Next i
    End If

    ' ---- type groups
    Set gGroups = modCore.NewDict()
    Set lo = modCore.Tbl("MM_TypeGroups", "tblTypeGroups")
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, 1)) <> "" Then
                If Not gGroups.Exists(CStr(d(i, 1))) Then gGroups.Add CStr(d(i, 1)), modCore.NewDict()
                gGroups(CStr(d(i, 1)))(CStr(d(i, 2))) = True
            End If
        Next i
    End If

    ' ---- allowed relationships
    Set gAllowed = New Collection
    Set lo = modCore.Tbl("MM_EntityRelationships", "tblEntityRels")
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, Col(lo, "Status"))) = "Active" Then
                gAllowed.Add Array(CStr(d(i, Col(lo, "SourceType"))), CStr(d(i, Col(lo, "RelCode"))), _
                                   CStr(d(i, Col(lo, "TargetType"))), d(i, Col(lo, "MinCard")), d(i, Col(lo, "MaxCard")))
            End If
        Next i
    End If

    ' ---- rules
    Set gRules = New Collection
    Set lo = modCore.Tbl("MM_Rules", "tblRules")
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, Col(lo, "Status"))) = "Active" And CStr(d(i, Col(lo, "AttrCode"))) <> "" Then
                gRules.Add Array(CStr(d(i, Col(lo, "AttrCode"))), CStr(d(i, Col(lo, "TypeCode"))), _
                                 CStr(d(i, Col(lo, "RuleType"))), CStr(d(i, Col(lo, "Param1"))), _
                                 CStr(d(i, Col(lo, "Param2"))), CStr(d(i, Col(lo, "Severity"))), _
                                 CStr(d(i, Col(lo, "Message"))), CStr(d(i, Col(lo, "RuleID"))))
            End If
        Next i
    End If

    ' ---- lists
    Set gLists = modCore.NewDict()
    Set gListSets = modCore.NewDict()
    Set lo = modCore.Tbl("MM_Lists", "tblLists")
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, Col(lo, "Status"))) <> "Retired" And CStr(d(i, 1)) <> "" Then
                If Not gLists.Exists(CStr(d(i, 1))) Then gLists.Add CStr(d(i, 1)), New Collection
                gLists(CStr(d(i, 1))).Add CStr(d(i, 2))
                gListSets(CStr(d(i, 1)) & "|" & CStr(d(i, 2))) = True
            End If
        Next i
    End If
    mLoaded = True
End Sub

Private Function Col(lo As ListObject, ByVal header As String) As Long
    Col = modCore.ColIndex(lo, header)
    If Col = 0 Then Err.Raise vbObjectError + 101, "modMeta", "Column '" & header & "' missing in " & lo.Name
End Function

Private Sub InsertOrdered(c As Collection, ByVal attrCode As String, ByVal order As Double, ByVal typeCode As String)
    Dim i As Long, o As Double
    For i = 1 To c.Count
        o = gMapInfo(typeCode & "|" & c(i))(1)
        If order < o Then
            c.Add attrCode, Before:=i
            Exit Sub
        End If
    Next i
    c.Add attrCode
End Sub

' ---------------------------------------------------------------- queries
Public Function TypeExists(ByVal typeCode As String) As Boolean
    LoadMeta
    TypeExists = gTypes.Exists(typeCode)
End Function

Public Function TypeName_(ByVal typeCode As String) As String
    LoadMeta
    If gTypes.Exists(typeCode) Then TypeName_ = gTypes(typeCode)("Name") Else TypeName_ = typeCode
End Function

Public Function AttrName(ByVal attrCode As String) As String
    LoadMeta
    If gAttrs.Exists(attrCode) Then AttrName = gAttrs(attrCode)("Name") Else AttrName = attrCode
End Function

Public Function IsSystemAttr(ByVal attrCode As String) As Boolean
    LoadMeta
    If gAttrs.Exists(attrCode) Then IsSystemAttr = gAttrs(attrCode)("IsSystem")
End Function

' Attribute codes mapped to a type, in display order (empty collection if none)
Public Function AttrsForType(ByVal typeCode As String) As Collection
    LoadMeta
    If gTypeAttrs.Exists(typeCode) Then
        Set AttrsForType = gTypeAttrs(typeCode)
    Else
        Set AttrsForType = New Collection
    End If
End Function

Public Function IsMapped(ByVal typeCode As String, ByVal attrCode As String) As Boolean
    LoadMeta
    IsMapped = gMapInfo.Exists(typeCode & "|" & attrCode)
End Function

Public Function IsMandatory(ByVal typeCode As String, ByVal attrCode As String) As Boolean
    LoadMeta
    If gMapInfo.Exists(typeCode & "|" & attrCode) Then IsMandatory = gMapInfo(typeCode & "|" & attrCode)(0)
End Function

Public Function MapSection(ByVal typeCode As String, ByVal attrCode As String) As String
    LoadMeta
    If gMapInfo.Exists(typeCode & "|" & attrCode) Then MapSection = gMapInfo(typeCode & "|" & attrCode)(2)
End Function

Public Function InList(ByVal listName As String, ByVal value As String) As Boolean
    LoadMeta
    InList = gListSets.Exists(listName & "|" & value)
End Function

' Token: TypeCode | "*" | "SAME" | "G:<group>"   (mirrors model.token_matches)
Public Function TokenMatches(ByVal token As String, ByVal typeCode As String, _
                             Optional ByVal sourceType As String = "") As Boolean
    token = Trim$(token)
    If token = "*" Then
        TokenMatches = True
    ElseIf token = "SAME" Then
        TokenMatches = (typeCode = sourceType)
    ElseIf Left$(token, 2) = "G:" Then
        If gGroups.Exists(Mid$(token, 3)) Then TokenMatches = gGroups(Mid$(token, 3)).Exists(typeCode)
    Else
        TokenMatches = (StrComp(token, typeCode, vbTextCompare) = 0)
    End If
End Function

Public Function IsRelAllowed(ByVal srcType As String, ByVal relCode As String, ByVal tgtType As String) As Boolean
    LoadMeta
    Dim a As Variant
    For Each a In gAllowed
        If a(1) = relCode Then
            If TokenMatches(a(0), srcType) Then
                If TokenMatches(a(2), tgtType, srcType) Then IsRelAllowed = True: Exit Function
            End If
        End If
    Next a
End Function

' Relationship codes allowed from a source type to any target type
Public Function RelsFromType(ByVal srcType As String) As Collection
    LoadMeta
    Dim c As New Collection, seen As Object, rc As Variant, t As Variant
    Set seen = modCore.NewDict()
    For Each rc In gRelTypes.Keys
        If gRelTypes(rc)("Status") = "Active" Then
            For Each t In gTypes.Keys
                If gTypes(t)("Status") = "Active" Then
                    If IsRelAllowed(srcType, CStr(rc), CStr(t)) Then
                        If Not seen.Exists(rc) Then seen(rc) = True: c.Add CStr(rc)
                        Exit For
                    End If
                End If
            Next t
        End If
    Next rc
    Set RelsFromType = c
End Function

' Target types allowed for a source type and relationship
Public Function TargetTypesFor(ByVal srcType As String, ByVal relCode As String) As Collection
    LoadMeta
    Dim c As New Collection, t As Variant
    For Each t In gTypes.Keys
        If gTypes(t)("Status") = "Active" Then
            If IsRelAllowed(srcType, relCode, CStr(t)) Then c.Add CStr(t)
        End If
    Next t
    Set TargetTypesFor = c
End Function

Public Function InverseVerb(ByVal verb As String) As String
    Select Case LCase$(verb)
        Case "composed of": InverseVerb = "part of"
        Case "aggregates": InverseVerb = "aggregated by"
        Case "assigned to": InverseVerb = "assigned from"
        Case "realizes": InverseVerb = "realized by"
        Case "serves": InverseVerb = "served by"
        Case "accesses": InverseVerb = "accessed by"
        Case "influences": InverseVerb = "influenced by"
        Case "triggers": InverseVerb = "triggered by"
        Case "flows to": InverseVerb = "flows from"
        Case "specializes": InverseVerb = "specialized by"
        Case Else: InverseVerb = verb
    End Select
End Function

Public Function RelVerb(ByVal relCode As String) As String
    LoadMeta
    If gRelTypes.Exists(relCode) Then RelVerb = gRelTypes(relCode)("Name") Else RelVerb = relCode
End Function

Public Function ActiveTypeCodes() As Collection
    LoadMeta
    Dim c As New Collection, t As Variant
    For Each t In gTypes.Keys
        If gTypes(t)("Status") = "Active" Then c.Add CStr(t)
    Next t
    Set ActiveTypeCodes = c
End Function

Public Function SchemaVersion() As Long
    SchemaVersion = Val(CStr(modCore.Setting("SchemaVersion")))
End Function

Public Sub BumpSchemaVersion()
    modCore.SetSetting "SchemaVersion", SchemaVersion() + 1
End Sub
