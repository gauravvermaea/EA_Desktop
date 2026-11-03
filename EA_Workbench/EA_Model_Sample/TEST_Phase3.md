# Phase 3 acceptance test (run after TEST_Phase2.md)

Tick each step. If one fails, note the step number and the exact message, and send them back.

## K. Templates

| # | Do this | Expect |
| --- | --- | --- |
| K1 | EA_Core Home → Templates, import, export. Entity type ApplicationComponent, pre-fill N, relationships Y → Generate template | "Template saved"; a file in Templates\ |
| K2 | Open the template | Sheets README, Entities, Relationships, Allowed; Lifecycle and Hosting have dropdowns; Owner * and Lifecycle * are marked required |
| K3 | Relationships sheet: open the Relationship and Target type dropdowns | Only relationship codes and target types listed on the Allowed sheet |
| K4 | Generate again with pre-fill Y | The Entities sheet lists the existing application components with their values; Relationships lists their outgoing relationships |

## L. Import

| # | Do this | Expect |
| --- | --- | --- |
| L1 | In the K1 template, add rows: "Underwriting Workbench" (Owner "Underwriting IT", Lifecycle Plan) and "Broker Portal" (Owner blank, Lifecycle Production). Add a relationship row: Underwriting Workbench, Serving, BusinessActor, Claims Department. Save into Imports\ | — |
| L2 | Choose files and preview → pick the file | IMPORT_Preview: Underwriting Workbench "New"; Broker Portal "Error: Owner is required"; the relationship "New" |
| L3 | Commit import | Snapshot taken; "1 new entity … 1 new relationship(s), 1 row(s) skipped"; the file has moved to Imports\Done\; T_ApplicationComponent shows APP-0004 Underwriting Workbench with its relationship |
| L4 | Commit import again | "Preview an import first" (it cannot run twice) |
| L5 | In a fresh template, type the existing name "Claims Platform" with Vendor "Duck Creek" only, import | Preview: Update, Vendor Guidewire → Duck Creek; other values untouched after commit |
| L6 | In a pre-filled template, clear the Version of Policy Administration System, import | Preview: Update, Version 10.2 → blank |
| L7 | Import a template from another model (or edit _meta ModelID) | File-level error "belongs to another model" |

## M. Export

| # | Do this | Expect |
| --- | --- | --- |
| M1 | Export: scope All entities, Excel | File in Exports\: README, one sheet per type, Relationships |
| M2 | Export: Around one entity, APP-0002, hops 1 | Only Claims Platform and its direct neighbours, with the relationships between them |
| M3 | Export: One entity type, Capability, CSV | A folder in Exports\ with one CSV per sheet |
| M4 | Export: Full model with metamodel | Also contains the MM_ sheets |

## N. Reports

| # | Do this | Expect |
| --- | --- | --- |
| N1 | EA_Core Home → Refresh reports | EA_Reports.xlsx created in the model folder and opened on the Dashboard; counts match EA_Core |
| N2 | Entity list: click a Type slicer button | Table filters to that type |
| N3 | Matrix: filters set to ApplicationComponent and Capability | Grid of applications against capabilities with 1 where linked |
| N4 | Close EA_Reports. Change something in EA_Core and save. Open EA_Reports on its own, Data → Refresh All | The change appears; no macros were needed |
| N5 | On T_ApplicationComponent select two entity rows → right-click → detail of selected entities | RPT_EntityDetail in EA_Entities: both entities with attributes, relationships (links work) and recent changes |
| N6 | Right-click Claims Platform → impact of this entity → 2 | RPT_Impact lists entities by hop with how each was reached |

## O. Power BI (optional, needs Power BI Desktop)

| # | Do this | Expect |
| --- | --- | --- |
| O1 | Open Reports\EA_Reports.pbip, set the ModelFolder parameter, Refresh | Six tables load without errors; the Entities measure matches EA_Core |
| O2 | Add a card with Data quality % and a matrix of Entity[Name] × Target Entity[Target Name] | Values appear; filtering by EntityType works |
