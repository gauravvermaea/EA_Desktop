# Phase 4 acceptance test (run after TEST_Phase3.md)

Needs draw.io Desktop. Tick each step. If one fails, note the step number and the exact message, and send them back.

## P. Libraries

| # | Do this | Expect |
| --- | --- | --- |
| P1 | In draw.io: File → Open Library from → Device → `Library\ArchiMate4.xml` | A palette of 42 element shapes in domain colours, then 11 connectors |
| P2 | Same for `Library\ModelLibrary.xml` | 14 sample entities, titled "Type: Name (Display ID)" |
| P3 | EA_Core Home → Diagrams (draw.io) → Rebuild ArchiMate 4 stencil, then Rebuild Model Library | Both files rewritten; reopening them in draw.io shows the same palettes |

## Q. Generate

| # | Do this | Expect |
| --- | --- | --- |
| Q1 | Open `Diagrams\Claims_Overview.drawio` in draw.io | 14 shapes in bands (Motivation, Strategy, Business, Common, Application, Technology, Implementation & Migration) and 12 connectors |
| Q2 | Diagrams form: file "Claims", content Around one entity, APP-0002, hops 1 → Generate | Asks to open; draw.io shows Claims Platform and its direct neighbours with their connectors |
| Q3 | In draw.io move some shapes, save, close. Generate again with hops 2 | Your positions are kept; the new neighbours are added below |
| Q4 | Select two rows on T_ApplicationComponent → right-click → diagram of selected entities | A new page in the chosen file with both entities and any connector between them |

## R. Harvest

| # | Do this | Expect |
| --- | --- | --- |
| R1 | In `Claims.drawio` drag an Application Component from ArchiMate 4, name it "Fraud Engine". Ctrl+M: Owner "Claims IT", Lifecycle "Plan". Drag a Serving connector from it to Claims Department. Add another Application Component named "Broker Gateway" with no Owner. Save and close draw.io | — |
| R2 | Diagrams form → Choose diagrams and preview → Claims.drawio | HARVEST_Review: Fraud Engine "New"; Broker Gateway "Error: Owner is required - add it with Edit Data"; the connector "New" |
| R3 | Commit harvest | Snapshot taken; 1 new entity, 1 new relationship; T_ApplicationComponent shows Fraud Engine with its relationship |
| R4 | Reopen Claims.drawio in draw.io, select Fraud Engine, Ctrl+M | ea_id and ea_display are filled in |
| R5 | Rename Claims Platform on the diagram to "Claims Platform X", save, close, preview and commit | Preview: Rename; after commit the entity is renamed everywhere |
| R6 | Drag Claim Record from the Model Library and draw a Serving connector from it to Claims Department (Data Object serves Business Actor is not allowed), preview | Error: "Data Object serves Business Actor is not allowed" |
| R7 | Drag Application Component, name it "Customer Portal" (an existing name), preview | Link: the shape is linked to the existing APP entity, no duplicate |
| R8 | Edit the diagram after preview, then Commit | "These diagrams changed after the preview" - nothing written |

## S. Keep current, open and export

| # | Do this | Expect |
| --- | --- | --- |
| S1 | Set Customer Portal's Status to Retired in Excel, save. Diagrams form → Update diagrams from model | Customer Portal has a red outline on every diagram |
| S2 | Views register | Claims overview: "Stale: 1 entity retired or deleted"; Claims pages listed with entity and connector counts |
| S3 | Sync Now, then Views register | Status refreshed; ModelLibrary.xml no longer contains Customer Portal |
| S4 | Diagrams form → Export images (png) | One PNG per page in `Diagrams\Export\` |
| S5 | Views register → click a row → Open selected diagram | draw.io opens that file |
