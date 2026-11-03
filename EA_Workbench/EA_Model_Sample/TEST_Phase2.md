# Phase 2 acceptance test (run after TEST_Phase1.md)

Tick each step. If one fails, note the step number and the exact message, and send them back.

## G. Sync on open and close

| # | Do this | Expect |
| --- | --- | --- |
| G1 | Close Excel. Open `EA_Entities.xlsm` only | EA_Core opens too; status bar ends with "quick sync done"; no messages |
| G2 | Close EA_Entities (save if asked) | EA_Core closes by itself a moment later |
| G3 | Open `EA_Core.xlsm` only | Quick sync runs; EA_Entities is opened, refreshed, saved and closed again |
| G4 | In EA_Core, change the Description of Claims Department in DATA_Master. Close EA_Core. Open EA_Entities | T_BusinessActor shows the new description |

## H. Edits made without the macros

| # | Do this | Expect |
| --- | --- | --- |
| H1 | Close Excel. Open EA_Entities with macros disabled (hold Shift while opening, or choose Disable Macros) | Opens with no Workbench activity |
| H2 | On T_Node, change the Vendor of Production Database Cluster to "Microsoft". Save and close | — |
| H3 | Open EA_Entities normally | Banner on Home: "1 difference(s) need a decision - see SYNC_Review" |
| H4 | EA_Core → SYNC_Review | One row: "Edited in EA_Entities outside the Workbench", Vendor, Oracle → Microsoft, Keep = EA_Entities |
| H5 | Click Apply decisions | "1 item(s) applied"; DATA_Master shows Microsoft; H_Node has the change with Source "Sync review" |
| H6 | Repeat H1–H2 changing Vendor to "IBM", and meanwhile in EA_Core (macros on) set the same Vendor to "SAP" | Next sync: one row "Changed in both", Keep = EA_Core |
| H7 | Leave Keep = EA_Core, Apply | Sheet shows SAP; review row says "Applied: kept EA_Core" |
| H8 | With macros off, type a new name on a blank row of T_Goal, save, reopen normally | Review row "New row in EA_Entities"; Apply adds it with a new GOL Display ID |
| H9 | With macros off, change a value in EA_Core DATA_Master, save, reopen | Sync logs it in the H_ sheet as Action "Sync", with the old value |

## I. Sync Now

| # | Do this | Expect |
| --- | --- | --- |
| I1 | On a type sheet, sort the rows by Name with macros off (or delete an entity row), save, reopen, click Sync Now | Summary message; the sheet is rebuilt in the right order with its relationship rows |
| I2 | Click Sync Now with nothing changed | "0 difference(s) need a decision"; Validate All summary shown |
| I3 | Set one review row to Skip, Apply | Row stays Open; that type sheet is not refreshed until it is settled |

## J. Versions

| # | Do this | Expect |
| --- | --- | --- |
| J1 | EA_Core Home → Version history → Take snapshot, label "baseline" | "Snapshot N saved"; two files in History\ (…-core.xlsx and …-entities.xlsx); listed on the sheet |
| J2 | Change Criticality of Claims Platform to 3 and add a relationship. Select the baseline row → Compare selected with current | VERSION_Compare lists the changed Criticality (5 → 3) and the added relationship |
| J3 | Version history → Entity history → APP-0002 | All changes to Claims Platform and its relationships, newest first |
| J4 | Select baseline → Restore entity or type → APP-0002 | Automatic snapshot first; Criticality back to 5; history shows Source "Restore from version N" |
| J5 | Add / edit type → change the Name of an attribute → Save | An automatic snapshot appears in the list with Reason "Before metamodel change" |
| J6 | Select baseline → Restore whole model → confirm | Automatic snapshot first; model back to baseline; type sheets regenerated; history keeps all rows |
| J7 | Compare two versions → enter the two latest version numbers | Differences between the two snapshots listed |
