# Phase 5 acceptance test (run after TEST_Phase4.md)

Tick each step. If one fails, note the step number and the exact message, and send them back.

## T. Create a child

| # | Do this | Expect |
| --- | --- | --- |
| T1 | EA_Core Home → Parent and child models. Child name "Claims", scope Around entities, start APP-0002, hops 1 → Create child | Confirmation with entity counts; a folder `EA_Model_Sample_Claims` next to the model folder; SYS_Children lists it |
| T2 | Close both workbooks. Open `EA_Model_Sample_Claims\EA_Entities.xlsm` | "Child model Claims is ready"; only Claims Platform, its neighbours and their references are on the type sheets |
| T3 | Parent and child models form | Section 3 shows "Child of Sample" with a Prepare merge package button |
| T4 | With the child open, open the parent's EA_Entities.xlsm | Message that another model is open; nothing is changed |

## U. Change both sides

| # | Do this | Expect |
| --- | --- | --- |
| U1 | In the child: Claims Platform Vendor → "Duck Creek"; add Application Component "Fraud Engine" (Owner, Lifecycle) with Serving to Claims Department; Claim Record Owner → "Claims Data". Save | — |
| U2 | Prepare merge package, then close the child | `Merge\ChildPackage.xlsx` written |
| U3 | Open the parent: Claims Platform Version → "9.1"; Claim Record Owner → "Data Governance". Save | — |

## V. Merge

| # | Do this | Expect |
| --- | --- | --- |
| V1 | Parent and child models → Choose child and preview → the package | MERGE_Review: Vendor Update (Keep Child); Fraud Engine Add; the Serving relationship Add; Claim Record Owner Conflict (Keep Parent); no row for Version |
| V2 | Set the Claim Record conflict to Keep = Child, Apply merge | Snapshot taken; summary counts; Claims Platform has Vendor Duck Creek and Version 9.1; Claim Record Owner Claims Data; Fraud Engine and its relationship exist with history Source "Merge Claims" |
| V3 | Apply merge again | "Preview a merge first" |
| V4 | Preview again straight away | No Update or Add rows: the child's base now matches |
| V5 | Edit the package or use a package from another model | "changed after the preview" / "not created from this model" |
