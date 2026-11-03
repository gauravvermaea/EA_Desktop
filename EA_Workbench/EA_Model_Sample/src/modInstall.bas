Attribute VB_Name = "modInstall"
' =====================================================================
' One-time installer. Import ONLY this module into EA_Core.xlsx, then run
' InstallWorkbench. It imports the other modules from the src folder,
' adds the workbook event code, saves both workbooks as .xlsm, builds the
' Home buttons, regenerates the type sheets and runs Validate All.
'
' Needs: File > Options > Trust Center > Trust Center Settings >
'        Macro Settings > "Trust access to the VBA project object model".
' This module references the others only through Application.Run, so it
' compiles before they exist.
' =====================================================================
Option Explicit

Private Const CORE_MODULES As String = "modCore.bas|modMeta.bas|modRules.bas|modData.bas|modGenerator.bas|" & _
                                       "modPaint.bas|modEvents.bas|modForms.bas|modValidate.bas|modSync.bas|modHistory.bas|" & _
                                       "modIO.bas|modReports.bas|modDiagram.bas|modMerge.bas|" & _
                                       "clsAppEvents.cls"

Public Sub InstallWorkbench()
    Dim sep As String, folder As String, src As String, f As Variant, vbp As Object, n As Long
    sep = Application.PathSeparator
    folder = ThisWorkbook.Path & sep
    src = folder & "src" & sep
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "This workbook is open from OneDrive or SharePoint (a web address). Copy the model folder to a " & _
               "local folder such as C:\EA\ and run setup from there.", vbExclamation, "EA Workbench setup"
        Exit Sub
    End If

    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    n = vbp.VBComponents.Count
    If Err.Number <> 0 Or vbp Is Nothing Then
        On Error GoTo 0
        MsgBox "Excel is blocking access to the VBA project." & vbLf & vbLf & _
               "Turn on: File > Options > Trust Center > Trust Center Settings > Macro Settings > " & _
               """Trust access to the VBA project object model"", then run InstallWorkbench again." & vbLf & vbLf & _
               "You can switch it off again after installation.", vbExclamation, "EA Workbench setup"
        Exit Sub
    End If
    On Error GoTo Fail

    If Dir(src & "modCore.bas") = "" Then
        MsgBox "The src folder with the Workbench modules was not found next to this workbook:" & vbLf & src, _
               vbExclamation, "EA Workbench setup"
        Exit Sub
    End If
    If MsgBox("Install the EA Repository Workbench into this folder?" & vbLf & vbLf & folder & vbLf & vbLf & _
              "EA_Core and EA_Entities will be saved as macro-enabled .xlsm files. The original .xlsx files " & _
              "are moved to the History folder.", vbOKCancel + vbQuestion, "EA Workbench setup") = vbCancel Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- EA_Core: modules and workbook events
    For Each f In Split(CORE_MODULES, "|")
        ImportModule ThisWorkbook, src & f
    Next f
    SetWorkbookCode ThisWorkbook, src & "ThisWorkbook_Core.txt"
    SaveAsMacroBook ThisWorkbook, folder & "EA_Core.xlsm"

    ' ---- EA_Entities
    Dim wbE As Workbook, ePath As String
    ePath = folder & "EA_Entities.xlsx"
    If Dir(ePath) = "" Then ePath = folder & "EA_Entities.xlsm"
    If Dir(ePath) = "" Then
        MsgBox "EA_Entities was not found in " & folder, vbExclamation, "EA Workbench setup"
        GoTo Finish
    End If
    Set wbE = Workbooks.Open(ePath, UpdateLinks:=0)
    ImportModule wbE, src & "modEntitiesThin.bas"
    SetWorkbookCode wbE, src & "ThisWorkbook_Entities.txt"
    SaveAsMacroBook wbE, folder & "EA_Entities.xlsm"

    ' ---- start the engine, regenerate, buttons, validate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Run "'" & ThisWorkbook.Name & "'!modCore.StartupCore"
    Application.Run "'" & ThisWorkbook.Name & "'!modGenerator.RegenerateAll"
    Application.Run "'" & ThisWorkbook.Name & "'!modGenerator.SetupHomeButtons"
    Application.DisplayAlerts = False
    wbE.Save
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    ThisWorkbook.Worksheets("Home").Activate
    MsgBox "EA Repository Workbench installed." & vbLf & vbLf & _
           "Use the buttons on Home in EA_Core and EA_Entities. Validate All will run next.", _
           vbInformation, "EA Workbench setup"
    Application.Run "'" & ThisWorkbook.Name & "'!modValidate.ValidateAll"
Finish:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Setup stopped: " & Err.Description, vbExclamation, "EA Workbench setup"
End Sub

Private Sub ImportModule(wb As Workbook, ByVal path As String)
    Dim nm As String, comp As Object
    nm = Mid$(path, InStrRev(path, Application.PathSeparator) + 1)
    nm = Left$(nm, InStrRev(nm, ".") - 1)
    On Error Resume Next
    Set comp = wb.VBProject.VBComponents(nm)
    On Error GoTo 0
    ' Remove only takes effect when the macro ends, so rename first to avoid "modCore1" on a re-run
    If Not comp Is Nothing Then
        comp.Name = nm & "_old"
        wb.VBProject.VBComponents.Remove comp
    End If
    wb.VBProject.VBComponents.Import path
End Sub

Private Sub SetWorkbookCode(wb As Workbook, ByVal path As String)
    Dim cm As Object, txt As String, ff As Integer
    ff = FreeFile
    Open path For Input As #ff
    txt = Input$(LOF(ff), ff)
    Close #ff
    Set cm = wb.VBProject.VBComponents(wb.CodeName).CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString txt
End Sub

Private Sub SaveAsMacroBook(wb As Workbook, ByVal newPath As String)
    Dim oldPath As String, hist As String, sep As String
    sep = Application.PathSeparator
    oldPath = wb.FullName
    If LCase$(oldPath) = LCase$(newPath) Then
        wb.Save
        Exit Sub
    End If
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=newPath, FileFormat:=52      ' xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    hist = wb.Path & sep & "History" & sep
    If Dir(hist, vbDirectory) = "" Then MkDir hist
    If Dir(oldPath) <> "" Then
        Dim target As String
        target = hist & "preinstall_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Mid$(oldPath, InStrRev(oldPath, sep) + 1)
        Name oldPath As target
    End If
End Sub
