Attribute VB_Name = "modReports"
' =====================================================================
' Phase 3: reports.
'   EA_Reports.xlsx  - macro-free workbook next to EA_Core. Power Query
'                      reads the saved EA_Core.xlsm from the workbook's
'                      own folder, so the folder can move. Refreshes on
'                      open or with Data > Refresh All. Built (first time)
'                      and refreshed from EA_Core with Refresh reports.
'   Entity detail    - one or more entities, painted on RPT_EntityDetail
'   Impact           - everything within n hops, painted on RPT_Impact
'                      (both in EA_Entities; macros needed to paint).
' =====================================================================
Option Explicit

Private Const REPORTS_FILE As String = "EA_Reports.xlsx"
Private mProblems As String

' ================================================================ EA_Reports.xlsx
Public Sub RefreshReports()
    Dim path As String, wb As Workbook, isNew As Boolean
    modCore.Heal
    On Error GoTo Fail
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Move the model folder to a local folder to use the reports workbook.", vbExclamation, "EA Workbench"
        Exit Sub
    End If
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    modCore.Status "reports: validating"
    modValidate.RunValidation True                 ' fresh SYS_Issues for the data-quality pages
    modSync.SaveWithStamp ThisWorkbook             ' Power Query reads the saved file
    path = ThisWorkbook.Path & Application.PathSeparator & REPORTS_FILE
    mProblems = ""
    Set wb = OpenIfOpen(REPORTS_FILE)
    If wb Is Nothing Then
        If Dir(path) = "" Then
            isNew = True
            Set wb = BuildReportsWorkbook(path)
        Else
            Set wb = Workbooks.Open(path, UpdateLinks:=0)
        End If
    End If
    If Not isNew And Not wb Is Nothing Then
        ' keep an existing workbook's queries current (new columns, renamed EA_Core)
        Dim q As Object, k As Variant
        Set q = QueryFormulas()
        On Error Resume Next
        For Each k In q.Keys
            wb.Queries(CStr(k)).Formula = CStr(q(k))
            If Err.Number <> 0 Then Problem "update query " & k
        Next k
        On Error GoTo Fail
    End If
    If wb Is Nothing Then Exit Sub
    modCore.Status "reports: refreshing"
    RefreshWorkbook wb
    On Error Resume Next
    wb.Worksheets("Settings").Range("B4").Value = modCore.UtcNow()
    On Error GoTo Fail
    Application.DisplayAlerts = False
    wb.Save
    Application.DisplayAlerts = True
    wb.Activate
    wb.Worksheets("Dashboard").Activate
    modCore.Status "reports refreshed"
    If mProblems <> "" Then
        MsgBox "EA_Reports was " & IIf(isNew, "built", "refreshed") & ", but some parts failed:" & mProblems, _
               vbExclamation, "EA Workbench"
    End If
    Exit Sub
Fail:
    Dim msg As String
    msg = Err.Description
    Application.DisplayAlerts = True
    modCore.ResetBusy
    MsgBox "Reports stopped: " & msg, vbExclamation, "EA Workbench"
End Sub

Private Function OpenIfOpen(ByVal fname As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.Name) = LCase$(fname) Then Set OpenIfOpen = wb: Exit Function
    Next wb
End Function

Private Sub Problem(ByVal what As String)
    mProblems = mProblems & vbLf & "- " & what & ": " & Err.Description
    Err.Clear
End Sub

Private Sub RefreshWorkbook(wb As Workbook)
    Dim ws As Worksheet, lo As ListObject, pc As PivotCache
    On Error Resume Next
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.SourceType = xlSrcQuery Then
                lo.QueryTable.Refresh BackgroundQuery:=False
                If Err.Number <> 0 Then Problem "refresh " & lo.Name
            End If
        Next lo
    Next ws
    Dim nQuery As Long
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.SourceType = xlSrcQuery Then nQuery = nQuery + 1
        Next lo
    Next ws
    If nQuery = 0 Then wb.RefreshAll                ' fall back if the tables are not seen as query tables
    For Each pc In wb.PivotCaches
        pc.Refresh
        If Err.Number <> 0 Then Problem "refresh pivot"
    Next pc
    wb.Worksheets("Dashboard").Calculate
End Sub

' Power Query source of every query: the EA_Core.xlsm saved next to this workbook.
' ` stands for " in these templates.
Private Function MHead() As String
    MHead = "let" & vbCrLf & _
        "    Folder = Excel.CurrentWorkbook(){[Name=`ModelFolder`]}[Content]{0}[Column1]," & vbCrLf & _
        "    Src = Excel.Workbook(File.Contents(Folder & `" & ThisWorkbook.Name & "`), null, true)," & vbCrLf
End Function

Private Function M(ByVal body As String) As String
    M = Replace(MHead() & body, "`", """")
End Function

Private Function QueryFormulas() As Object
    Dim q As Object
    Set q = modCore.NewDict()
    q("qEntities") = M( _
        "    Master = Src{[Item=`tblMaster`,Kind=`Table`]}[Data]," & vbCrLf & _
        "    Types = Table.SelectColumns(Src{[Item=`tblEntityTypes`,Kind=`Table`]}[Data], {`TypeCode`, `Name`, `Domain`})," & vbCrLf & _
        "    Joined = Table.NestedJoin(Master, {`Type`}, Types, {`TypeCode`}, `T`, JoinKind.LeftOuter)," & vbCrLf & _
        "    Expanded = Table.ExpandTableColumn(Joined, `T`, {`Name`, `Domain`}, {`Type Name`, `Type Domain`})," & vbCrLf & _
        "    Trimmed = Table.RemoveColumns(Expanded, {`Created`, `Created By`, `Modified By`, `Row Hash`}, MissingField.Ignore)," & vbCrLf & _
        "    Rows = Table.SelectRows(Trimmed, each [Entity ID] <> null and [Entity ID] <> ``)," & vbCrLf & _
        "    Front = {`Display ID`, `Name`, `Type Name`, `Type Domain`, `Status`}," & vbCrLf & _
        "    Result = Table.ReorderColumns(Rows, Front & List.Difference(Table.ColumnNames(Rows), Front))" & vbCrLf & _
        "in" & vbCrLf & "    Result")
    q("qRelationships") = M( _
        "    Rel = Table.SelectRows(Src{[Item=`tblRel`,Kind=`Table`]}[Data], each [RelID] <> null and [RelID] <> ``)," & vbCrLf & _
        "    Verbs = Table.SelectColumns(Src{[Item=`tblRelTypes`,Kind=`Table`]}[Data], {`RelCode`, `Name`})," & vbCrLf & _
        "    Joined = Table.NestedJoin(Rel, {`RelCode`}, Verbs, {`RelCode`}, `V`, JoinKind.LeftOuter)," & vbCrLf & _
        "    Expanded = Table.ExpandTableColumn(Joined, `V`, {`Name`}, {`Relationship`})," & vbCrLf & _
        "    Result = Table.SelectColumns(Expanded, {`SourceName`, `Relationship`, `TargetName`, `SourceType`, `TargetType`, `RelCode`, `Label`, `Status`, `SourceID`, `TargetID`, `RelID`})" & vbCrLf & _
        "in" & vbCrLf & "    Result")
    q("qIssues") = M( _
        "    Issues = Src{[Item=`tblIssues`,Kind=`Table`]}[Data]," & vbCrLf & _
        "    Rows = Table.SelectRows(Issues, each [Severity] <> null and [Severity] <> ``)," & vbCrLf & _
        "    Result = Table.RemoveColumns(Rows, {`Go to`}, MissingField.Ignore)" & vbCrLf & _
        "in" & vbCrLf & "    Result")
    q("qOrphans") = M( _
        "    Master = Table.SelectRows(Src{[Item=`tblMaster`,Kind=`Table`]}[Data], each [Entity ID] <> null and [Entity ID] <> `` and [Status] <> `Retired`)," & vbCrLf & _
        "    Rel = Table.SelectRows(Src{[Item=`tblRel`,Kind=`Table`]}[Data], each [RelID] <> null and [Status] <> `Retired`)," & vbCrLf & _
        "    Linked = List.Buffer(List.Distinct(Rel[SourceID] & Rel[TargetID]))," & vbCrLf & _
        "    Orphans = Table.SelectRows(Master, each not List.Contains(Linked, [Entity ID]))," & vbCrLf & _
        "    Result = Table.SelectColumns(Orphans, {`Display ID`, `Name`, `Type`, `Status`, `Entity ID`})" & vbCrLf & _
        "in" & vbCrLf & "    Result")
    Set QueryFormulas = q
End Function

Private Function BuildReportsWorkbook(ByVal path As String) As Workbook
    Dim wb As Workbook, ws As Worksheet, q As Object, k As Variant, nm As Variant
    On Error GoTo Fail
    modCore.Status "reports: creating EA_Reports.xlsx"
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "Dashboard"
    For Each nm In Array("Entity list", "Relationships", "Matrix", "Data quality", "Orphans", "Settings")
        wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = nm
    Next nm
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=51        ' saved first: the folder formula needs a file name
    Application.DisplayAlerts = True

    ' ---- settings: the workbook finds its own folder
    Set ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "EA Reports settings"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Model folder (this workbook's folder)"
    ' TEXTBEFORE (Microsoft 365) copes with "[" in folder names; older Excel falls back to FIND
    ws.Range("B2").Formula = "=IFERROR(TEXTBEFORE(CELL(""filename"",$A$1),""["",-1)," & _
                             "LEFT(CELL(""filename"",$A$1),FIND(""["",CELL(""filename"",$A$1))-1))"
    ws.Range("A3").Value = "Repository file"
    ws.Range("B3").Value = ThisWorkbook.Name
    ws.Range("A4").Value = "Last refreshed from EA_Core (UTC)"
    ws.Range("A6").Value = "Reports read the last saved EA_Core. Refresh with Data > Refresh All (twice if pivots lag), " & _
                           "or from EA_Core with Refresh reports."
    ws.Columns("A").ColumnWidth = 40: ws.Columns("B").ColumnWidth = 70
    wb.Names.Add Name:="ModelFolder", RefersTo:="=Settings!$B$2"

    ' ---- queries
    On Error Resume Next
    wb.Queries.FastCombine = True                   ' allows the folder cell and the file to be combined
    If Err.Number <> 0 Then Problem "privacy setting (set Data > Get Data > Query Options > Privacy to ignore)"
    On Error GoTo Fail
    Set q = QueryFormulas()
    For Each k In q.Keys
        AddQuery wb, CStr(k), CStr(q(k))
    Next k
    LoadQuery wb, "qEntities", wb.Worksheets("Entity list"), "A10", "tEntities"
    LoadQuery wb, "qRelationships", wb.Worksheets("Relationships"), "A4", "tRelationships"
    LoadQuery wb, "qIssues", wb.Worksheets("Data quality"), "A4", "tIssues"
    LoadQuery wb, "qOrphans", wb.Worksheets("Orphans"), "A4", "tOrphans"

    ' ---- pages
    BuildEntityList wb
    BuildRelationshipsPage wb
    BuildMatrix wb
    BuildDataQuality wb
    BuildDashboard wb
    Application.DisplayAlerts = False
    wb.Save
    Application.DisplayAlerts = True
    Set BuildReportsWorkbook = wb
    Exit Function
Fail:
    Problem "building EA_Reports"
    Application.DisplayAlerts = True
    Set BuildReportsWorkbook = wb
End Function

Private Sub AddQuery(wb As Workbook, ByVal qName As String, ByVal formula As String)
    On Error Resume Next
    wb.Queries(qName).Delete
    Err.Clear
    wb.Queries.Add Name:=qName, Formula:=formula
    If Err.Number <> 0 Then Problem "query " & qName
End Sub

Private Sub LoadQuery(wb As Workbook, ByVal qName As String, ws As Worksheet, ByVal at As String, ByVal tableName As String)
    Dim lo As ListObject
    On Error GoTo Fail
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qName & ";Extended Properties=""""", _
        Destination:=ws.Range(at))
    lo.Name = tableName                  ' named first, so later pages find it even if the first refresh fails
    lo.TableStyle = "TableStyleMedium2"
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & qName & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = True
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SaveData = True
        .AdjustColumnWidth = True
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    Exit Sub
Fail:
    Problem "loading " & qName
End Sub

Private Sub PageTitle(ws As Worksheet, ByVal title As String, ByVal subtitle As String)
    ws.Range("A1").Value = title
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = RGB(31, 56, 100)
    ws.Range("A2").Value = subtitle
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ActiveWindowGridlines ws
End Sub

Private Sub ActiveWindowGridlines(ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Private Function NewPivot(wb As Workbook, ByVal srcTable As String, ws As Worksheet, ByVal at As String, _
                          ByVal pvName As String) As PivotTable
    Dim pc As PivotCache
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcTable)
    pc.RefreshOnFileOpen = True
    Set NewPivot = pc.CreatePivotTable(TableDestination:=ws.Range(at), TableName:=pvName)
End Function

Private Sub BuildEntityList(wb As Workbook)
    Dim ws As Worksheet, lo As ListObject
    On Error GoTo Fail
    Set ws = wb.Worksheets("Entity list")
    PageTitle ws, "Entity list", "Every entity with all its attributes. Use the slicers or the column filters."
    Set lo = ws.ListObjects("tEntities")
    AddSlicer wb, lo, "Type Name", ws, "slType", "Type", 3, 1, 200
    AddSlicer wb, lo, "Type Domain", ws, "slDomain", "Domain", 3, 215, 160
    AddSlicer wb, lo, "Status", ws, "slStatus", "Status", 3, 390, 120
    Exit Sub
Fail:
    Problem "Entity list page"
End Sub

Private Sub AddSlicer(wb As Workbook, lo As ListObject, ByVal field As String, ws As Worksheet, ByVal sName As String, _
                      ByVal caption As String, ByVal topRow As Long, ByVal leftPts As Double, ByVal widthPts As Double)
    Dim sc As SlicerCache
    On Error GoTo Fail
    Set sc = wb.SlicerCaches.Add2(lo, field)
    sc.Slicers.Add ws, , sName, caption, ws.Rows(topRow).Top, leftPts + 5, widthPts, 100
    Exit Sub
Fail:
    Problem "slicer " & field
End Sub

Private Sub BuildRelationshipsPage(wb As Workbook)
    On Error GoTo Fail
    PageTitle wb.Worksheets("Relationships"), "Relationships", "Every relationship in the repository."
    Exit Sub
Fail:
    Problem "Relationships page"
End Sub

Private Sub BuildMatrix(wb As Workbook)
    Dim ws As Worksheet, pt As PivotTable
    On Error GoTo Fail
    Set ws = wb.Worksheets("Matrix")
    PageTitle ws, "Relationship matrix", _
        "Pick the source type, target type and relationship in the filters above the grid, e.g. ApplicationComponent x Capability."
    Set pt = NewPivot(wb, "tRelationships", ws, "A8", "pvMatrix")
    With pt
        .PivotFields("SourceType").Orientation = xlPageField
        .PivotFields("TargetType").Orientation = xlPageField
        .PivotFields("Relationship").Orientation = xlPageField
        .PivotFields("SourceName").Orientation = xlRowField
        .PivotFields("TargetName").Orientation = xlColumnField
        .AddDataField .PivotFields("RelID"), "Links", xlCount
        .RowAxisLayout xlTabularRow
        .NullString = ""
    End With
    On Error Resume Next
    pt.PivotFields("SourceType").CurrentPage = "ApplicationComponent"
    pt.PivotFields("TargetType").CurrentPage = "Capability"
    Exit Sub
Fail:
    Problem "Matrix page"
End Sub

Private Sub BuildDataQuality(wb As Workbook)
    On Error GoTo Fail
    PageTitle wb.Worksheets("Data quality"), "Data quality", _
        "Issues from the last Validate All in EA_Core (Refresh reports runs it first)."
    PageTitle wb.Worksheets("Orphans"), "Orphans", "Active entities with no relationships."
    Exit Sub
Fail:
    Problem "Data quality page"
End Sub

Private Sub BuildDashboard(wb As Workbook)
    Dim ws As Worksheet, pt As PivotTable, i As Long, labels As Variant, formulas As Variant
    On Error GoTo Fail
    Set ws = wb.Worksheets("Dashboard")
    PageTitle ws, "EA repository dashboard", "Model: " & CStr(modCore.Setting("ModelName")) & _
                  ". Refresh with Data > Refresh All, or from EA_Core with Refresh reports."
    labels = Array("Entities", "Relationships", "Errors", "Warnings", "Orphans", "Data quality")
    formulas = Array("=COUNTA(tEntities[Entity ID])", "=COUNTA(tRelationships[RelID])", _
                     "=COUNTIF(tIssues[Severity],""Error"")", "=COUNTIF(tIssues[Severity],""Warning"")", _
                     "=COUNTA(tOrphans[Entity ID])", _
                     "=IFERROR(1-ROWS(UNIQUE(FILTER(tIssues[RecordID],(tIssues[Severity]=""Error"")*(tIssues[Sheet]=""DATA_Master""))))/B5,1)")
    For i = 0 To 5
        ws.Cells(4, 2 + i * 2).Value = labels(i)
        ws.Cells(4, 2 + i * 2).Font.Color = RGB(89, 89, 89)
        ws.Cells(5, 2 + i * 2).Formula2 = formulas(i)
        ws.Cells(5, 2 + i * 2).Font.Size = 20
        ws.Cells(5, 2 + i * 2).Font.Bold = True
    Next i
    ws.Cells(5, 12).NumberFormat = "0%"
    ' entities by domain and type
    Set pt = NewPivot(wb, "tEntities", ws, "B8", "pvByType")
    With pt
        .PivotFields("Type Domain").Orientation = xlRowField
        .PivotFields("Type Name").Orientation = xlRowField
        .AddDataField .PivotFields("Entity ID"), "Entities", xlCount
        .RowAxisLayout xlOutlineRow
    End With
    ' by lifecycle and repository area
    On Error Resume Next
    Set pt = NewPivot(wb, "tEntities", ws, "F8", "pvLifecycle")
    pt.PivotFields("Lifecycle").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Entity ID"), "By lifecycle", xlCount
    If Err.Number <> 0 Then Problem "Dashboard lifecycle pivot"
    Set pt = NewPivot(wb, "tEntities", ws, "I8", "pvArea")
    pt.PivotFields("Repository Area").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Entity ID"), "By repository area", xlCount
    If Err.Number <> 0 Then Problem "Dashboard repository-area pivot"
    On Error GoTo Fail
    ws.Columns("A").ColumnWidth = 2
    ws.Range("B1").Select
    Exit Sub
Fail:
    Problem "Dashboard"
End Sub

' ================================================================ entity detail and impact (in EA_Entities)
Private Function ReportSheet(ByVal nm As String) As Worksheet
    Dim wbE As Workbook, ws As Worksheet
    Set wbE = modCore.EntitiesWB(True)
    If wbE Is Nothing Then Exit Function
    If modCore.SheetExists(wbE, nm) Then
        Set ws = wbE.Worksheets(nm)
        modCore.UnprotectSheet ws
        ws.Cells.Hyperlinks.Delete
        ws.Cells.Clear
    Else
        Set ws = wbE.Worksheets.Add(After:=wbE.Worksheets(wbE.Worksheets.Count))
        ws.Name = nm
    End If
    ws.Tab.Color = RGB(237, 125, 49)
    Set ReportSheet = ws
End Function

' Entity IDs selected on a type sheet, or asked for by Display ID
Public Function ChosenEntities(ByVal prompt As String) As Collection
    Dim c As New Collection, ws As Worksheet, cell As Range, s As String, part As Variant, id As String, seen As Object
    Set ChosenEntities = c
    Set seen = modCore.NewDict()
    If TypeName(ActiveSheet) = "Worksheet" Then
        Set ws = ActiveSheet
        If ws.Name Like "T_*" Then
            For Each cell In Selection.Rows
                If CStr(ws.Cells(cell.Row, 2).Value) = "E" Then
                    id = CStr(ws.Cells(cell.Row, 1).Value)
                    If id <> "" And Not seen.Exists(id) Then seen(id) = True: c.Add id
                End If
            Next cell
        End If
    End If
    If c.Count > 0 Then Exit Function
    s = InputBox(prompt, "EA Workbench")
    For Each part In Split(s, ",")
        id = IDByDisplay(Trim$(CStr(part)))
        If id <> "" And Not seen.Exists(id) Then seen(id) = True: c.Add id
    Next part
End Function

Private Function IDByDisplay(ByVal disp As String) As String
    Dim m As Variant, ws As Worksheet, cols As Object
    If disp = "" Then Exit Function
    Set ws = modData.MasterSheet(): Set cols = modData.MasterCols()
    m = Application.Match(disp, ws.Columns(cols("DisplayID")), 0)
    If Not IsError(m) Then IDByDisplay = CStr(ws.Cells(CLng(m), cols("EntityID")).Value)
End Function

Private Sub LinkToEntity(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal id As String, ByVal text As String)
    Dim t As String, sh As String
    t = modData.EntityTypeOf(id)
    ws.Cells(r, c).Value = modCore.SafeValue(text)
    If t = "" Then Exit Sub
    If Not modMeta.gTypes.Exists(t) Then Exit Sub
    sh = modMeta.gTypes(t)("Sheet")
    ws.Cells(r, c).Formula = "=HYPERLINK(""#'" & sh & "'!E""&MATCH(""" & id & """,'" & sh & "'!$A:$A,0),""" & _
                             Replace(text, """", """""") & """)"
End Sub

Public Sub EntityDetail()
    Dim ids As Collection, ws As Worksheet, r As Long, id As Variant, row As Object, typeCode As String
    Dim a As Variant, sec As String, lastSec As String, d As Variant, i As Long, other As String, verb As String
    Dim h As Collection, x As Variant, n As Long
    modCore.Heal
    Set ids = ChosenEntities("Display IDs of the entities to show, separated by commas (e.g. APP-0002, CAP-0001):")
    If ids.Count = 0 Then Exit Sub
    modMeta.LoadMeta
    modCore.BeginBusy
    On Error GoTo Fail
    Set ws = ReportSheet("RPT_EntityDetail")
    If ws Is Nothing Then modCore.EndBusy: MsgBox "EA_Entities was not found.", vbExclamation, "EA Workbench": Exit Sub
    ws.Range("A1").Value = "Entity detail"
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = RGB(31, 56, 100)
    ws.Range("A2").Value = "Generated " & modCore.UtcNow() & " UTC from EA_Core. Names in blue jump to the entity."
    ws.Range("A2").Font.Italic = True
    r = 4
    d = modCore.TableData(modData.RelLO())
    For Each id In ids
        If modData.FindEntityRow(CStr(id)) > 0 Then
            Set row = modData.ReadEntityRow(modData.FindEntityRow(CStr(id)))
            typeCode = CStr(row("TypeCode"))
            ws.Cells(r, 1).Value = row("DisplayID") & "  " & row("Name")
            ws.Cells(r, 1).Font.Size = 14: ws.Cells(r, 1).Font.Bold = True
            ws.Cells(r, 3).Value = modMeta.TypeName_(typeCode) & "  |  " & row("Status")
            ws.Cells(r, 3).Font.Color = RGB(89, 89, 89)
            r = r + 1
            ws.Cells(r, 1).Value = modCore.SafeValue(CStr(modCore.Nz(row("Description"), "")))
            ws.Cells(r, 1).Font.Italic = True
            r = r + 2
            ' attributes by section
            lastSec = "#"
            For Each a In modMeta.AttrsForType(typeCode)
                sec = modMeta.MapSection(typeCode, CStr(a))
                If sec <> lastSec Then
                    ws.Cells(r, 1).Value = IIf(sec = "", "Attributes", sec)
                    ws.Cells(r, 1).Font.Bold = True: ws.Cells(r, 1).Font.Color = RGB(31, 56, 100)
                    r = r + 1
                    lastSec = sec
                End If
                ws.Cells(r, 1).Value = modMeta.AttrName(CStr(a))
                ws.Cells(r, 2).NumberFormat = modGenerator.NumberFormatFor(CStr(a))
                ws.Cells(r, 2).Value = modCore.SafeFor(row(a), CStr(ws.Cells(r, 2).NumberFormat))
                ws.Cells(r, 2).HorizontalAlignment = xlLeft
                r = r + 1
            Next a
            ws.Cells(r, 1).Value = "Repository area": ws.Cells(r, 2).Value = row("RepositoryArea"): r = r + 1
            ws.Cells(r, 1).Value = "Last modified": ws.Cells(r, 2).Value = row("Modified") & "  by " & row("ModifiedBy"): r = r + 2
            ' relationships
            ws.Cells(r, 1).Value = "Relationships"
            ws.Cells(r, 1).Font.Bold = True: ws.Cells(r, 1).Font.Color = RGB(31, 56, 100)
            r = r + 1
            n = 0
            If IsArray(d) Then
                For i = 1 To UBound(d, 1)
                    If CStr(d(i, 11)) <> "Retired" Then
                        other = ""
                        If d(i, 3) = id Then other = CStr(d(i, 4)): verb = modMeta.RelVerb(CStr(d(i, 2)))
                        If d(i, 4) = id Then other = CStr(d(i, 3)): verb = modMeta.InverseVerb(modMeta.RelVerb(CStr(d(i, 2))))
                        If other <> "" Then
                            ws.Cells(r, 1).Value = verb
                            LinkToEntity ws, r, 2, other, modData.EntityNameOf(other)
                            ws.Cells(r, 3).Value = modMeta.TypeName_(modData.EntityTypeOf(other))
                            r = r + 1: n = n + 1
                        End If
                    End If
                Next i
            End If
            If n = 0 Then ws.Cells(r, 1).Value = "(none)": r = r + 1
            r = r + 1
            ' recent history
            Set h = RecentHistory(CStr(id), typeCode, 8)
            ws.Cells(r, 1).Value = "Recent changes"
            ws.Cells(r, 1).Font.Bold = True: ws.Cells(r, 1).Font.Color = RGB(31, 56, 100)
            r = r + 1
            For Each x In h
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Value = Array(x(0), x(1), modMeta.AttrName(CStr(x(2))), _
                    modCore.SafeValue(CStr(modCore.Nz(x(3), ""))), modCore.SafeValue(CStr(modCore.Nz(x(4), ""))))
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Font.Color = RGB(89, 89, 89)
                r = r + 1
            Next x
            If h.Count = 0 Then ws.Cells(r, 1).Value = "(no recorded changes)": r = r + 1
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Borders(xlEdgeBottom).Weight = xlMedium
            r = r + 2
        End If
    Next id
    ws.Columns("A").ColumnWidth = 30: ws.Columns("B").ColumnWidth = 40: ws.Columns("C").ColumnWidth = 26
    ws.Columns("D:E").ColumnWidth = 24
    modCore.EndBusy
    ws.Parent.Activate
    ws.Activate
    ws.Range("A1").Select
    Exit Sub
Fail:
    modCore.ResetBusy
    MsgBox "Entity detail stopped: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

' Last n history rows for an entity: Array(Timestamp, Action, AttrCode, Old, New), newest first
Private Function RecentHistory(ByVal id As String, ByVal typeCode As String, ByVal n As Long) As Collection
    Dim c As New Collection, sh As String, d As Variant, i As Long
    Set RecentHistory = c
    sh = Left$("H_" & typeCode, 31)
    If Not modCore.SheetExists(ThisWorkbook, sh) Then Exit Function
    d = modCore.TableData(ThisWorkbook.Worksheets(sh).ListObjects(1))
    If Not IsArray(d) Then Exit Function
    For i = UBound(d, 1) To 1 Step -1
        If CStr(d(i, 5)) = id Then
            c.Add Array(d(i, 2), d(i, 4), d(i, 6), d(i, 7), d(i, 8))
            If c.Count >= n Then Exit For
        End If
    Next i
End Function

Public Sub Impact()
    Dim ids As Collection, startID As String, s As String, hops As Long, levels As Object, ws As Worksheet
    Dim r As Long, lvl As Long, k As Variant, d As Variant, i As Long, via As Object, a As String, b As String
    modCore.Heal
    Set ids = ChosenEntities("Display ID of the entity to start from (e.g. APP-0002):")
    If ids.Count = 0 Then Exit Sub
    startID = ids(1)
    s = InputBox("How many relationship hops to follow (1-5)?", "Impact", "2")
    If s = "" Then Exit Sub
    hops = Val(s)
    If hops < 1 Then hops = 1
    If hops > 5 Then hops = 5
    modMeta.LoadMeta
    Set levels = modIO.Neighbourhood(startID, hops)
    ' how each entity was reached: first relationship found to an entity one level closer
    Set via = modCore.NewDict()
    d = modCore.TableData(modData.RelLO())
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, 1)) <> "" And CStr(d(i, 11)) <> "Retired" Then
                a = CStr(d(i, 3)): b = CStr(d(i, 4))
                If levels.Exists(a) And levels.Exists(b) Then
                    If levels(b) = levels(a) + 1 And Not via.Exists(b) Then via(b) = Array(a, modMeta.RelVerb(CStr(d(i, 2))))
                    If levels(a) = levels(b) + 1 And Not via.Exists(a) Then via(a) = Array(b, modMeta.InverseVerb(modMeta.RelVerb(CStr(d(i, 2)))))
                End If
            End If
        Next i
    End If
    modCore.BeginBusy
    On Error GoTo Fail
    Set ws = ReportSheet("RPT_Impact")
    If ws Is Nothing Then modCore.EndBusy: MsgBox "EA_Entities was not found.", vbExclamation, "EA Workbench": Exit Sub
    ws.Range("A1").Value = "Impact of " & modData.EntityNameOf(startID)
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = RGB(31, 56, 100)
    ws.Range("A2").Value = (levels.Count - 1) & " entities within " & hops & " hop(s), following relationships in both directions."
    ws.Range("A2").Font.Italic = True
    ws.Range("A4:F4").Value = Array("Hops", "Entity", "Type", "Display ID", "Reached from", "Via")
    ws.Range("A4:F4").Font.Bold = True
    ws.Range("A4:F4").Interior.Color = RGB(221, 235, 247)
    r = 4
    For lvl = 0 To hops
        For Each k In levels.Keys
            If levels(k) = lvl Then
                r = r + 1
                ws.Cells(r, 1).Value = lvl
                LinkToEntity ws, r, 2, CStr(k), modData.EntityNameOf(CStr(k))
                ws.Cells(r, 3).Value = modMeta.TypeName_(modData.EntityTypeOf(CStr(k)))
                ws.Cells(r, 4).Value = DisplayOf(CStr(k))
                If via.Exists(k) Then
                    ws.Cells(r, 5).Value = modCore.SafeValue(modData.EntityNameOf(CStr(via(k)(0))))
                    ws.Cells(r, 6).Value = via(k)(1)
                End If
            End If
        Next k
    Next lvl
    ws.Columns("A").ColumnWidth = 6: ws.Columns("B").ColumnWidth = 36: ws.Columns("C").ColumnWidth = 24
    ws.Columns("D").ColumnWidth = 12: ws.Columns("E").ColumnWidth = 34: ws.Columns("F").ColumnWidth = 18
    modCore.EndBusy
    ws.Parent.Activate
    ws.Activate
    ws.Range("A1").Select
    Exit Sub
Fail:
    modCore.ResetBusy
    MsgBox "Impact stopped: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

Private Function DisplayOf(ByVal id As String) As String
    Dim r As Long
    r = modData.FindEntityRow(id)
    If r > 0 Then DisplayOf = CStr(modData.MasterSheet().Cells(r, modData.MasterCols()("DisplayID")).Value)
End Function
