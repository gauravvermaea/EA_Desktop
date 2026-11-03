Attribute VB_Name = "modEntitiesThin"
' =====================================================================
' EA_Entities thin module. All logic lives in EA_Core; this only opens
' EA_Core when needed, starts the sync on open and routes the Home
' buttons to EA_Core. Edits on the type sheets are captured by EA_Core's
' application events.
' =====================================================================
Option Explicit

Public Function CoreName() As String
    Dim lo As ListObject, r As Variant
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets("SYS_SyncState").ListObjects("tblSyncState")
    r = Application.Match("CoreFile", lo.ListColumns("Key").DataBodyRange, 0)
    If Not IsError(r) Then CoreName = CStr(lo.ListColumns("Value").DataBodyRange.Cells(r, 1).Value)
    On Error GoTo 0
    If CoreName = "" Then CoreName = "EA_Core.xlsm"
End Function

' True only for the EA_Core in this workbook's own folder
Private Function CoreOpen() As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.Name) = LCase$(CoreName()) Then
            If LCase$(wb.Path) = LCase$(ThisWorkbook.Path) Then CoreOpen = True
            Exit Function
        End If
    Next wb
End Function

' Path of an EA_Core with the same name that belongs to another model folder, if one is open
Private Function OtherCorePath() As String
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.Name) = LCase$(CoreName()) And LCase$(wb.Path) <> LCase$(ThisWorkbook.Path) Then
            OtherCorePath = wb.Path: Exit Function
        End If
    Next wb
End Function

' Opens EA_Core (if needed) and asks it to run the quick sync
Public Sub ConnectCore()
    Dim path As String, openedCore As Boolean
    On Error GoTo Fail
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "This workbook is open from OneDrive or SharePoint (a web address). Copy the model folder to a " & _
               "local folder to use the Workbench.", vbExclamation, "EA Workbench"
        Exit Sub
    End If
    If OtherCorePath() <> "" Then
        MsgBox "Another model is open (" & OtherCorePath() & "). Excel cannot open two workbooks with the same name, " & _
               "so close that model (both workbooks), then open this one again.", vbExclamation, "EA Workbench"
        Exit Sub
    End If
    If Not CoreOpen() Then
        path = ThisWorkbook.Path & Application.PathSeparator & CoreName()
        If Dir(path) = "" Then
            MsgBox "EA_Core (" & CoreName() & ") was not found next to this workbook. You can edit, but nothing is " & _
                   "checked or saved to the repository until EA_Core is back. The next sync will pick up your edits.", _
                   vbExclamation, "EA Workbench"
            Exit Sub
        End If
        Workbooks.Open path, UpdateLinks:=0
        openedCore = True
    End If
    Application.Run "'" & CoreName() & "'!modSync.EntitiesOpened", openedCore
    ThisWorkbook.Activate
    Exit Sub
Fail:
    MsgBox "Could not connect to EA_Core: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

' Workbook_BeforeClose: EA_Core closes itself afterwards if this workbook opened it
Public Sub Closing()
    On Error Resume Next
    If CoreOpen() Then Application.Run "'" & CoreName() & "'!modSync.EntitiesClosing"
End Sub

Private Sub RunCore(ByVal proc As String)
    If Not CoreOpen() Then ConnectCore
    If Not CoreOpen() Then Exit Sub
    Application.Run "'" & CoreName() & "'!" & proc
End Sub

' ---- Home buttons
Public Sub Btn_AddEntity()
    RunCore "modForms.ShowEntityFormBlank"
End Sub

Public Sub Btn_AddRelationship()
    RunCore "modForms.ShowRelFormBlank"
End Sub

Public Sub Btn_Push()
    RunCore "modEvents.PushPending"
End Sub

Public Sub Btn_Merge()
    RunCore "modMerge.ShowMerge"
End Sub

Public Sub Btn_Diagrams()
    RunCore "modDiagram.ShowDiagrams"
End Sub

Public Sub Btn_SyncNow()
    RunCore "modSync.SyncNow"
End Sub

Public Sub Btn_Versions()
    RunCore "modHistory.ShowVersions"
End Sub

Public Sub Btn_Detail()
    RunCore "modReports.EntityDetail"
End Sub

Public Sub Btn_Impact()
    RunCore "modReports.Impact"
End Sub

Public Sub Btn_IO()
    RunCore "modIO.ShowIO"
End Sub

Public Sub Btn_Reports()
    RunCore "modReports.RefreshReports"
End Sub

Public Sub Btn_ValidateAll()
    RunCore "modValidate.ValidateAll"
End Sub

Public Sub Btn_Regenerate()
    RunCore "modGenerator.RegenerateAll"
End Sub

Public Sub Btn_OpenCore()
    RunCore "modForms.GoHome"
End Sub
