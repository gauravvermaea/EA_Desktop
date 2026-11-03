# Phase 1 acceptance test (run in Windows Excel after setup)

Tick each step. If one fails, note the step number and the exact message, and send them back.

## A. Setup and opening

| # | Do this | Expect |
| --- | --- | --- |
| A1 | Run InstallWorkbench (README steps 1–4) | "EA Repository Workbench installed", then a Validate All result |
| A2 | Close Excel. Open `EA_Entities.xlsm` only | EA_Core opens too; no error messages |
| A3 | EA_Core Home | Eight buttons; counts show 14 entities, 12 relationships, 42 entity types |
| A4 | EA_Entities Home | Six buttons; a clickable list of 42 type sheets |

## B. Type sheets and relationship rows

| # | Do this | Expect |
| --- | --- | --- |
| B1 | Open T_ApplicationComponent | 3 entities; header row coloured blue; Owner * and Lifecycle * marked mandatory |
| B2 | Click + beside Claims Platform | 5 relationship rows (realizes, accesses, serves, realizes, served by) in grey italics |
| B3 | Click "Claims Management" in a relationship row | Jumps to that row on T_Capability |
| B4 | Try typing in a relationship row | Excel refuses (cell is protected) |

## C. Editing in EA_Entities

| # | Do this | Expect |
| --- | --- | --- |
| C1 | On T_Node, set Criticality of Production Database Cluster to 7 | Cell turns red; hover shows "Criticality must be a whole number from 1 to 5" |
| C2 | Change it to 4 | Red clears |
| C3 | Set Version to "v19" | Cell turns amber (warning) |
| C4 | Save EA_Entities | Message about saved rows; EA_Core DATA_Master shows the new values, Modified stamp updated |
| C5 | EA_Core H_Node sheet | One history row per changed field, Source = EA_Entities |
| C6 | On T_Goal, type a new name in the first blank row below the data, then save | New entity with Display ID GOL-0002; row now shows its ID |
| C7 | On T_ApplicationComponent, blank the Owner of Customer Portal and save | "rows have errors and are still waiting"; cell stays red; EA_Core unchanged |
| C8 | Put the Owner back and save | Saved; queue empty |

## D. Editing in EA_Core

| # | Do this | Expect |
| --- | --- | --- |
| D1 | DATA_Master: on the Goal row, try typing in the grey Vendor cell | Refused: "Vendor does not apply to Goal" |
| D2 | DATA_Master: change the Description of Claims Platform | Accepted, stamped, logged in H_ApplicationComponent; T_ApplicationComponent repainted |
| D3 | DATA_Master: try editing a Display ID | Refused (system column) |

## E. Wizards

| # | Do this | Expect |
| --- | --- | --- |
| E1 | Add / edit entity → type ApplicationComponent → leave entity blank, Name "CRM", Owner "Sales IT", Lifecycle Production → Save entity | Saved APP-0004; appears on T_ApplicationComponent |
| E2 | Same form, Name "CRM" again → Save | Warning "Another entity of this type has the same name" |
| E3 | Add relationship → source type ApplicationComponent, source CRM, relationship Access | Target types offered are only Business Object, Data Object, Artifact and Material |
| E4 | Pick Serving → target type BusinessActor → Claims Department → Save | Saved; + rows on both type sheets show it |
| E5 | Add relationship: DataObject → try to find "Access" in the list | Not offered (passive structure cannot access) |
| E6 | Add / edit type → Entity type → new code "BusinessCapabilityMap", name, domain Strategy, prefix BCM, copy mappings from Capability → Save | DATA_Master unchanged in shape; new sheet T_BusinessCapabilityMap appears in EA_Entities |
| E7 | Add / edit type → Attribute type → new code "RecoveryTimeObjective", name "RTO (hours)", type Integer, map to ApplicationComponent → Save | New column in DATA_Master (grey for non-application rows); new column on T_ApplicationComponent |
| E8 | Define rule → attribute RecoveryTimeObjective → Range 0 / 72, Error → Test value 100 | "Fails: ..." ; Save rule |
| E9 | On T_ApplicationComponent enter RTO 100 for CRM | Cell turns red |

## F. Validate All

| # | Do this | Expect |
| --- | --- | --- |
| F1 | Validate All | Summary of errors and warnings; SYS_Issues lists each with a "Go to" link |
| F2 | Click a Go to link | Jumps to the offending cell, which is highlighted |
