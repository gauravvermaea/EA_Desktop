Attribute VB_Name = "modCore"
' =====================================================================
' EA Repository Workbench - core utilities (EA_Core)
' Settings, paths, GUIDs, hashing, timestamps, busy flag, protection,
' table helpers. All other modules depend on this one.
' =====================================================================
Option Explicit

Public Const SH_HOME As String = "Home"
Public Const SH_MASTER As String = "DATA_Master"
Public Const SH_REL As String = "DATA_Relationships"
Public Const SH_SETTINGS As String = "SYS_Settings"
Public Const SH_ISSUES As String = "SYS_Issues"
Public Const SH_COUNTERS As String = "SYS_Counters"
Public Const SH_LISTCACHE As String = "SYS_ListCache"
Public Const SH_MAPCACHE As String = "SYS_MapCache"
Public Const SH_HMETA As String = "H_Metamodel"
Public Const SH_HREL As String = "H_Relationships"

Public Const MASTER_KEY_ROW As Long = 1
Public Const MASTER_HDR_ROW As Long = 2
Public Const MASTER_FIRST_ROW As Long = 3
Public Const TS_KEY_ROW As Long = 1
Public Const TS_HDR_ROW As Long = 2
Public Const TS_FIRST_ROW As Long = 3
Public Const TS_SPARE_ROWS As Long = 200

Public Const COLOR_ERROR As Long = 13421823     ' RGB(255,204,204)
Public Const COLOR_WARNING As Long = 10284031    ' RGB(255,235,156)

Private mBusy As Long
Public gSyncAt As Date              ' scheduled sync-on-open (cancelled if EA_Core closes first)
Private mSavedCalc As Long
Public gApp As clsAppEvents

' ---------------------------------------------------------------- startup
Public Sub StartupCore()
    On Error GoTo Fail
    Set gApp = New clsAppEvents
    Set gApp.App = Application
    modMeta.LoadMeta True
    ProtectCoreSheets
    AddContextMenu
    Banner ""
    ' a child model copied by Create child finishes its own setup on first open
    If modMerge.HasPendingSetup() Then modMerge.CompleteChildSetup
    ' quick sync once Excel has finished opening (skipped if EA_Entities already ran it)
    gSyncAt = Now + TimeSerial(0, 0, 1)
    Application.OnTime gSyncAt, "'" & ThisWorkbook.Name & "'!modSync.SyncOnOpenFromCore"
    Exit Sub
Fail:
    MsgBox "Workbench start-up problem: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

Public Sub EnsureEvents()
    If gApp Is Nothing Then
        Set gApp = New clsAppEvents
        Set gApp.App = Application
    End If
End Sub

' ---------------------------------------------------------------- busy / performance
Public Sub BeginBusy()
    mBusy = mBusy + 1
    If mBusy = 1 Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        mSavedCalc = Application.Calculation
        Application.Calculation = xlCalculationManual
    End If
End Sub

Public Sub EndBusy()
    If mBusy > 0 Then mBusy = mBusy - 1
    If mBusy = 0 Then
        Application.Calculation = IIf(mSavedCalc = 0, xlCalculationAutomatic, mSavedCalc)
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.StatusBar = False
    End If
End Sub

Public Function IsBusy() As Boolean
    IsBusy = (mBusy > 0)
End Function

' Call from error handlers so a failure never leaves events switched off.
Public Sub ResetBusy()
    mBusy = 1
    EndBusy
End Sub

Public Sub Status(ByVal text As String)
    Application.StatusBar = "EA Workbench: " & text
End Sub

' ---------------------------------------------------------------- workbooks
Public Function CoreWB() As Workbook
    Set CoreWB = ThisWorkbook
End Function

Public Function EntitiesFileName() As String
    EntitiesFileName = CStr(Setting("EntitiesFile"))
    If EntitiesFileName = "" Then EntitiesFileName = "EA_Entities.xlsm"
End Function

' Returns the open EA_Entities workbook; opens it from the model folder if asked.
Public Function EntitiesWB(Optional ByVal openIfClosed As Boolean = True) As Workbook
    Dim wb As Workbook, fname As String, path As String
    fname = EntitiesFileName()
    For Each wb In Application.Workbooks
        If LCase$(wb.Name) = LCase$(fname) Then
            If LCase$(wb.Path) = LCase$(ThisWorkbook.Path) Then
                Set EntitiesWB = wb
            Else
                ' another model's EA_Entities (same file name, other folder): never touch it
                Application.StatusBar = "EA Workbench: close the EA_Entities of " & wb.Path & " to work on this model"
            End If
            Exit Function
        End If
    Next wb
    If Not openIfClosed Then Exit Function
    path = ThisWorkbook.Path & Application.PathSeparator & fname
    If Dir(path) = "" Then
        ' fall back to the pre-install .xlsx name
        path = ThisWorkbook.Path & Application.PathSeparator & Replace(fname, ".xlsm", ".xlsx")
        If Dir(path) = "" Then Exit Function
    End If
    Dim ev As Boolean, s As Worksheet
    ev = Application.EnableEvents
    Application.EnableEvents = False
    Set EntitiesWB = Workbooks.Open(path, UpdateLinks:=0)
    Application.EnableEvents = ev
    ' its own Workbook_Open did not run: re-apply code-friendly protection with outlining
    For Each s In EntitiesWB.Worksheets
        If s.Name Like "T_*" Then ProtectSheet s
    Next s
    ThisWorkbook.Activate
End Function

Public Function SheetExists(wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' ---------------------------------------------------------------- settings
Public Function Setting(ByVal key As String) As Variant
    Dim lo As ListObject, r As Variant
    Set lo = ThisWorkbook.Worksheets(SH_SETTINGS).ListObjects("tblSettings")
    r = Application.Match(key, lo.ListColumns("Key").DataBodyRange, 0)
    If IsError(r) Then
        Setting = ""
    Else
        Setting = lo.ListColumns("Value").DataBodyRange.Cells(r, 1).Value
    End If
End Function

Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
    Dim lo As ListObject, r As Variant
    Set lo = ThisWorkbook.Worksheets(SH_SETTINGS).ListObjects("tblSettings")
    r = Application.Match(key, lo.ListColumns("Key").DataBodyRange, 0)
    If IsError(r) Then
        UnprotectSheet lo.Parent
        AppendRow lo, Array(key, value)
        ProtectSheet lo.Parent
    Else
        lo.ListColumns("Value").DataBodyRange.Cells(r, 1).Value = value
    End If
End Sub

Public Function Password() As String
    Password = CStr(Setting("ProtectPassword"))
End Function

' ---------------------------------------------------------------- identity, time, user
Public Function NewGUID() As String
    Static seeded As Boolean
    Dim i As Long, s As String, n As Long
    If Not seeded Then Randomize Timer: seeded = True
    For i = 1 To 32
        Select Case i
            Case 13: n = 4                          ' version 4
            Case 17: n = 8 + Int(Rnd() * 4)         ' variant 10xx
            Case Else: n = Int(Rnd() * 16)
        End Select
        s = s & LCase$(Hex$(n))
        If i = 8 Or i = 12 Or i = 16 Or i = 20 Then s = s & "-"
    Next i
    NewGUID = s
End Function

Public Function UtcNow() As String
    Dim dt As Date
    On Error GoTo LocalTime
    Dim swb As Object
    Set swb = CreateObject("WbemScripting.SWbemDateTime")
    swb.SetVarDate Now, True
    dt = swb.GetVarDate(False)
    UtcNow = Format$(dt, "yyyy-mm-dd") & "T" & Format$(dt, "hh:nn:ss") & "Z"
    Exit Function
LocalTime:
    UtcNow = Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss")
End Function

Public Function CurrentUser() As String
    CurrentUser = Environ$("USERNAME")
    If CurrentUser = "" Then CurrentUser = Application.UserName
End Function

' ---------------------------------------------------------------- hashing (mirrors model.fnv_hash)
Public Function HashText(ByVal s As String) As String
    Dim h As Double, i As Long, u As Long
    h = 2166136261#
    For i = 1 To Len(s)
        u = AscW(Mid$(s, i, 1))
        If u < 0 Then u = u + 65536
        h = FnvStep(h, u Mod 256)
        h = FnvStep(h, u \ 256)
    Next i
    HashText = LCase$(Right$("00000000" & HexOf32(h), 8))
End Function

Private Function FnvStep(ByVal h As Double, ByVal b As Long) As Double
    Dim lo As Double
    lo = h - Fix(h / 256#) * 256#                 ' low byte of h
    h = h - lo + (CLng(lo) Xor b)                 ' h XOR b (b < 256)
    ' h * 16777619 mod 2^32, split so every intermediate stays exact in a Double
    Dim t1 As Double, t2 As Double
    t1 = Mod32(h * 403#)
    t2 = Mod32((h - Fix(h / 256#) * 256#) * 16777216#)
    FnvStep = Mod32(t1 + t2)
End Function

Private Function Mod32(ByVal x As Double) As Double
    Mod32 = x - Fix(x / 4294967296#) * 4294967296#
End Function

Private Function HexOf32(ByVal h As Double) As String
    Dim hi As Long, lo As Long
    hi = CLng(Fix(h / 65536#))
    lo = CLng(h - CDbl(hi) * 65536#)
    HexOf32 = Right$("0000" & Hex$(hi), 4) & Right$("0000" & Hex$(lo), 4)
End Function

' Canonical text of a value (mirrors model.norm)
Public Function CellText(ByVal v As Variant) As String
    If IsError(v) Then CellText = "#ERR": Exit Function
    If IsEmpty(v) Or IsNull(v) Then CellText = "": Exit Function
    Select Case VarType(v)
        Case vbDate
            CellText = Format$(v, "yyyy-mm-dd")
        Case vbBoolean
            CellText = IIf(v, "TRUE", "FALSE")
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            If v = Fix(v) And Abs(v) < 1E+15 Then
                CellText = Trim$(Str$(CDec(v)))
            Else
                CellText = Trim$(Str$(v))
                If Left$(CellText, 1) = "." Then CellText = "0" & CellText
                If Left$(CellText, 2) = "-." Then CellText = "-0" & Mid$(CellText, 2)
            End If
        Case Else
            CellText = CStr(v)
    End Select
End Function

Public Function RowHashOf(values As Collection) As String
    Dim s As String, v As Variant, first As Boolean
    first = True
    For Each v In values
        If Not first Then s = s & ChrW$(30)
        s = s & CellText(v)
        first = False
    Next v
    RowHashOf = HashText(s)
End Function

' ---------------------------------------------------------------- protection
Public Sub ProtectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=Password(), UserInterfaceOnly:=True, AllowFormattingColumns:=True, _
               AllowFormattingRows:=True, AllowFiltering:=True, DrawingObjects:=False
    ws.EnableOutlining = True
    ws.EnableAutoFilter = True
End Sub

Public Sub UnprotectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=Password()
End Sub

Public Sub ProtectCoreSheets()
    Dim ws As Worksheet
    UnprotectSheet ThisWorkbook.Worksheets(SH_MASTER)
    UnlockMasterEditable                       ' set Locked flags before protecting
    ' metamodel description-style columns stay editable directly (set before protecting)
    UnlockColumns "MM_EntityTypes", Array("Name", "Description")
    UnlockColumns "MM_AttributeTypes", Array("Name", "Description", "Unit")
    UnlockColumns "MM_RelationshipTypes", Array("Name", "Description")
    UnlockColumns "MM_Rules", Array("Message")
    UnlockColumns "SYS_Settings", Array("Value")
    For Each ws In ThisWorkbook.Worksheets
        Select Case True
            Case ws.Name Like "MM_*", ws.Name Like "H_*", ws.Name Like "SYS_*", _
                 ws.Name = SH_MASTER, ws.Name = SH_REL
                ProtectSheet ws
        End Select
    Next ws
End Sub

' Call at the start of every button macro: recovers after an unhandled error or "End"
Public Sub Heal()
    EnsureEvents
    On Error Resume Next
    If Application.CommandBars("Cell").FindControl(Tag:="EAWB") Is Nothing Then AddContextMenu
    On Error GoTo 0
    If Not IsBusy() Then
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub

Private Sub UnlockColumns(ByVal sheetName As String, cols As Variant)
    Dim lo As ListObject, c As Variant
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(sheetName).ListObjects(1)
    UnprotectSheet lo.Parent
    For Each c In cols
        lo.ListColumns(CStr(c)).DataBodyRange.Locked = False
    Next c
End Sub

' Master: business columns open, system columns locked (enforced again by modEvents)
Public Sub UnlockMasterEditable()
    Dim ws As Worksheet, lo As ListObject, c As Long, key As String
    Set ws = ThisWorkbook.Worksheets(SH_MASTER)
    Set lo = ws.ListObjects("tblMaster")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For c = 1 To lo.ListColumns.Count
        key = CStr(ws.Cells(MASTER_KEY_ROW, lo.Range.Column + c - 1).Value)
        Select Case key
            Case "EntityID", "TypeCode", "DisplayID", "Created", "CreatedBy", "Modified", "ModifiedBy", "RowHash"
                lo.ListColumns(c).DataBodyRange.Locked = True
            Case Else
                lo.ListColumns(c).DataBodyRange.Locked = False
        End Select
    Next c
End Sub

' ---------------------------------------------------------------- table helpers
Public Function Tbl(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Set Tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
End Function

' Data of a table as a 2-D array (1-based). Returns Empty when the table has no real rows.
Public Function TableData(lo As ListObject) As Variant
    If lo.DataBodyRange Is Nothing Then Exit Function
    If lo.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function
    End If
    Dim v As Variant
    v = lo.DataBodyRange.Value
    If Not IsArray(v) Then
        Dim a(1 To 1, 1 To 1) As Variant
        a(1, 1) = v
        v = a
    End If
    TableData = v
End Function

Public Function ColIndex(lo As ListObject, ByVal header As String) As Long
    On Error Resume Next
    ColIndex = lo.ListColumns(header).Index
    On Error GoTo 0
End Function

' Appends a row; reuses the single blank row an empty table carries.
Public Function AppendRow(lo As ListObject, values As Variant) As Range
    Dim lr As ListRow, i As Long, n As Long
    If lo.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then
            Set lr = lo.ListRows(1)
        End If
    End If
    If lr Is Nothing Then Set lr = lo.ListRows.Add(AlwaysInsert:=True)
    n = UBound(values) - LBound(values) + 1
    For i = 1 To n
        If i > lo.ListColumns.Count Then Exit For
        lr.Range.Cells(1, i).Value = SafeFor(values(LBound(values) + i - 1), CStr(lr.Range.Cells(1, i).NumberFormat))
    Next i
    Set AppendRow = lr.Range
End Function

' Clears a table back to one blank row.
Public Sub ClearTable(lo As ListObject)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    On Error Resume Next
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData     ' a filter would hide rows from the delete
    End If
    On Error GoTo 0
    If lo.ListRows.Count > 1 Then
        lo.DataBodyRange.Offset(1, 0).Resize(lo.ListRows.Count - 1).Rows.Delete
    End If
    lo.DataBodyRange.ClearContents
End Sub

' Text that Excel would turn into a formula, number or date is stored as literal text
' (so "00123", "19.30" or "2026-01-01" typed into a Text attribute stay exactly as typed).
Public Function SafeValue(ByVal v As Variant) As Variant
    If VarType(v) = vbString Then
        If Len(v) > 0 Then
            Select Case Left$(v, 1)
                Case "=", "+", "-", "@"
                    SafeValue = "'" & v: Exit Function
            End Select
            If IsNumeric(v) Or IsDate(v) Then SafeValue = "'" & v: Exit Function
        End If
    End If
    SafeValue = v
End Function

' Like SafeValue, but for a cell already formatted as Text ("@") the raw string is written:
' Text cells keep text as typed, and an apostrophe there would be stored as a real character.
Public Function SafeFor(ByVal v As Variant, ByVal numberFormat As String) As Variant
    If numberFormat = "@" And VarType(v) = vbString Then
        If Left$(v, 1) = "'" Then SafeFor = Mid$(v, 2) Else SafeFor = v
    Else
        SafeFor = SafeValue(v)
    End If
End Function

Public Function Nz(ByVal v As Variant, Optional ByVal dflt As Variant = "") As Variant
    If IsError(v) Then
        Nz = dflt
    ElseIf IsEmpty(v) Or IsNull(v) Then
        Nz = dflt
    Else
        Nz = v
    End If
End Function

Public Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsError(v) Then IsBlankValue = False: Exit Function
    IsBlankValue = (Trim$(CellText(v)) = "")
End Function

' Last row / column holding anything (values or formulas), 0 when the sheet is empty
Public Function LastRowOf(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then LastRowOf = f.Row
End Function

Public Function LastColOf(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then LastColOf = f.Column
End Function

Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
    NewDict.CompareMode = vbTextCompare
End Function

' ---------------------------------------------------------------- UI helpers
Public Sub Banner(ByVal text As String)
    On Error Resume Next
    With ThisWorkbook.Worksheets(SH_HOME).Range("B4")
        .Value = text
        .Font.Bold = True
        .Font.Color = IIf(text = "", vbBlack, RGB(192, 0, 0))
    End With
End Sub

Public Sub AddButton(ws As Worksheet, ByVal caption As String, ByVal macro As String, _
                     ByVal topCell As Range, Optional ByVal widthPts As Double = 170, _
                     Optional ByVal nameTag As String = "")
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, topCell.Left, topCell.Top + 2, widthPts, 24)
    With shp
        If nameTag <> "" Then .Name = nameTag
        .OnAction = macro
        .Fill.ForeColor.RGB = RGB(31, 56, 100)
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = caption
            .Font.Size = 10
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Placement = xlFreeFloating
    End With
End Sub

Public Sub DeleteButtons(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Type = msoAutoShape Then ws.Shapes(i).Delete
    Next i
End Sub

' Right-click menu on cells: works on any EA_Entities type sheet row
Public Sub AddContextMenu()
    Dim cb As Object, ctl As Object
    On Error Resume Next
    RemoveContextMenu
    Set cb = Application.CommandBars("Cell")
    Set ctl = cb.Controls.Add(Type:=1, Before:=1, Temporary:=True)    ' msoControlButton
    ctl.Caption = "EA Workbench: edit this entity"
    ctl.Tag = "EAWB"
    ctl.OnAction = "'" & ThisWorkbook.Name & "'!modForms.EditSelectedEntity"
    Set ctl = cb.Controls.Add(Type:=1, Before:=2, Temporary:=True)
    ctl.Caption = "EA Workbench: add relationship from this entity"
    ctl.Tag = "EAWB"
    ctl.OnAction = "'" & ThisWorkbook.Name & "'!modForms.AddRelationshipForSelected"
    Set ctl = cb.Controls.Add(Type:=1, Before:=3, Temporary:=True)
    ctl.Caption = "EA Workbench: history of this entity"
    ctl.Tag = "EAWB"
    ctl.OnAction = "'" & ThisWorkbook.Name & "'!modHistory.ShowEntityHistory"
    Set ctl = cb.Controls.Add(Type:=1, Before:=4, Temporary:=True)
    ctl.Caption = "EA Workbench: detail of selected entities"
    ctl.Tag = "EAWB"
    ctl.OnAction = "'" & ThisWorkbook.Name & "'!modReports.EntityDetail"
    Set ctl = cb.Controls.Add(Type:=1, Before:=5, Temporary:=True)
    ctl.Caption = "EA Workbench: impact of this entity"
    ctl.Tag = "EAWB"
    ctl.OnAction = "'" & ThisWorkbook.Name & "'!modReports.Impact"
    Set ctl = cb.Controls.Add(Type:=1, Before:=6, Temporary:=True)
    ctl.Caption = "EA Workbench: diagram of selected entities"
    ctl.Tag = "EAWB"
    ctl.OnAction = "'" & ThisWorkbook.Name & "'!modDiagram.DiagramOfSelected"
    Set ctl = cb.Controls.Add(Type:=1, Before:=7, Temporary:=True)
    ctl.Caption = "EA Workbench: change or retire this relationship"
    ctl.Tag = "EAWB"
    ctl.OnAction = "'" & ThisWorkbook.Name & "'!modForms.EditSelectedRelationship"
    cb.Controls(8).BeginGroup = True
End Sub

Public Sub RemoveContextMenu()
    Dim ctl As Object
    On Error Resume Next
    Do
        Set ctl = Application.CommandBars("Cell").FindControl(Tag:="EAWB")
        If ctl Is Nothing Then Exit Do
        ctl.Delete
        Set ctl = Nothing
    Loop
End Sub

Public Sub OpenEntitiesWorkbook()
    Dim wb As Workbook
    Set wb = EntitiesWB(True)
    If wb Is Nothing Then
        MsgBox "EA_Entities was not found next to EA_Core.", vbExclamation, "EA Workbench"
    Else
        wb.Activate
    End If
End Sub
