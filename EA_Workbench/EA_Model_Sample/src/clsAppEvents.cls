VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Application-level events, so EA_Core can see edits in both EA_Core and
' EA_Entities without code in every sheet.
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    modEvents.OnSelection Sh, Target
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    modEvents.OnChange Sh, Target
End Sub

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If modCore.IsBusy() Then Exit Sub
    If LCase$(Wb.Name) = LCase$(modCore.EntitiesFileName()) Then
        modEvents.PushPending True
        ' remember that the Workbench saw this save (see modSync.SavedOutside)
        On Error Resume Next
        modGenerator.SetSyncState Wb, "WBSaveStamp", CDbl(Now)
    End If
End Sub
