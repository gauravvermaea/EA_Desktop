# EA Repository Workbench 1.1

One folder = one model. This folder holds a sample model seeded with the ArchiMate 4 metamodel
and 14 insurance-flavoured sample entities with 12 relationships.

- **Phase 1 (foundation):** metamodel, master data, type sheets with expandable relationships, validation,
  history, wizards, Validate All.
- **Phase 2 (sync and safety):** sync on open and on save, Sync Now, the SYNC_Review sheet for differences,
  version snapshots with compare and restore, entity history, history archiving.
- **Phase 3 (data exchange and reports):** import templates, import with preview, export, the EA_Reports
  workbook, entity detail and impact, and a Power BI project.
- **Phase 4 (draw.io):** an ArchiMate 4 stencil, a Model Library of real entities, generate diagrams from the
  repository, update them, harvest them back, a views register, open in draw.io and export images.
- **Phase 5 (parent and child models):** copy part of a model to a child folder for a project, and merge the child's
  changes back with a three-way, field-by-field review.

The `Guides` folder has the full documentation as PDFs: the **User Guide** (detailed, 36 pages), the **Setup Guide** and a **Quick Guide**. This README is the short version.

## What is in the folder

| Item | What it is |
| --- | --- |
| `EA_Core.xlsx` | Metamodel (MM_ sheets), master data (DATA_Master, DATA_Relationships), history (H_), SYNC_Review, version register, system and wizard sheets. Becomes `EA_Core.xlsm` at setup. |
| `EA_Entities.xlsx` | One type sheet per ArchiMate 4 element (T_ sheets) with expandable relationship rows. Becomes `EA_Entities.xlsm` at setup. |
| `src/` | The VBA engine as text files, imported at setup. |
| `model.json` | Model ID, name, file names, schema version, last sync. |
| `Guides/` | User Guide (detailed), Setup Guide and Quick Guide, as PDFs. |
| `EA_Reports.xlsx` | Created by **Refresh reports** the first time; macro-free reports that refresh themselves from EA_Core. |
| `Reports/` | `EA_Reports.pbip`, the Power BI project, with `README_PowerBI.md`. |
| `Library/` | `ArchiMate4.xml` (the ArchiMate 4 palette for draw.io) and `ModelLibrary.xml` (every active entity, ready to drag). |
| `Diagrams/` | draw.io diagrams (`Claims_Overview.drawio` is a sample). |

These working folders are created the first time they are needed, so they are not in the download:

| Folder | Created by | Holds |
| --- | --- | --- |
| `History/` | Setup, or the first snapshot | Version snapshots, pre-install copies, archived history, merge copies |
| `Templates/` | Generate template | Import templates |
| `Imports/` and `Imports\Done\` | Choose files and preview / Commit import | Filled templates waiting to import, then imported ones |
| `Exports/` | Export | Excel and CSV exports |
| `Diagrams\Export\` | Export images | PNG, SVG and PDF images of diagrams |
| `Merge/` | Prepare merge package (in a child model) | The merge package for the parent |

## Setup (once, about 2 minutes, Windows Excel)

**Easiest:** if Python is installed, double-click `Workbench_Tools\Setup.bat` and use step 3 in the window
(Unblock downloaded files, then Install into Excel). It does steps 1, 3 and 4 below for you; you still tick the
Trust Center option in step 2 yourself. The same window creates new models.

**By hand:**

1. **Copy the whole folder** to a local drive, such as `C:\EA\`, not OneDrive or a zip preview. If Windows
   marks the files as downloaded, right-click each `.xlsx` → Properties → tick **Unblock**.
2. **Allow setup to import code:** File → Options → Trust Center → Trust Center Settings → Macro Settings →
   tick **Trust access to the VBA project object model**. You can untick it again after setup.
3. **Open `EA_Core.xlsx`**, press **Alt+F11**, then File → Import File → `src\modInstall.bas`.
4. Press **F5** (or Run → Run Sub), choose **InstallWorkbench**, and confirm.

Setup imports the engine, saves both workbooks as `.xlsm`, moves the original `.xlsx` files to `History\`,
regenerates the type sheets, adds the Home buttons and runs Validate All.

**Recommended:** add the folder that holds your model folders as an Excel Trusted Location
(Trust Center → Trusted Locations), so the `.xlsm` files open with macros enabled.

## Daily use

- **Open EA_Entities.** It opens EA_Core and runs a quick sync. When you close EA_Entities, EA_Core
  closes too if EA_Entities opened it.
- **Edit on the type sheets.** Each cell is checked as you type: red = error, amber = warning; hover for the
  message. To add an entity, type on a blank row below the data.
- **Click the + in the left margin** to see an entity's relationships; click a related name to jump to it.
- **Save EA_Entities** to send your edits to EA_Core. Rows with errors stay highlighted and queued.
- **Right-click a row** for edit this entity, add relationship, or history of this entity.

## Sync

| When | What happens |
| --- | --- |
| Opening either workbook | Quick sync: queued edits are pushed, rows changed in EA_Core are repainted, structure and row order are fixed. |
| Saving EA_Entities | Queued edits are validated and written to EA_Core with history. |
| **Sync Now** (both Homes) | Full sync: every row in both workbooks is compared, then Validate All runs. |
| A workbook was saved without the macros | Detected automatically on the next open, and those rows are compared in full. |

EA_Core wins by default. Differences that need a decision appear on **SYNC_Review** in EA_Core:

| Situation | Default Keep |
| --- | --- |
| Edited in EA_Entities outside the Workbench | EA_Entities (apply the edit) |
| Changed in both EA_Core and EA_Entities | EA_Core |
| New row in EA_Entities | EA_Entities (add it) |
| Row in EA_Entities not found in EA_Core | EA_Core (remove it from the sheet) |

Set Keep to EA_Core, EA_Entities or Skip (decide later), then click **Apply decisions**. A type sheet with
open items is not refreshed until they are settled, so nothing on it is overwritten.

## Versions

**Version history** (both Homes) opens the version list in EA_Core:

- **Take snapshot:** saves a values-only copy of EA_Core and of the filled EA_Entities sheets to `History\`.
- **Compare selected with current / Compare two versions:** lists every added, removed and changed
  entity, relationship and metamodel row on VERSION_Compare.
- **Restore entity or type:** puts one entity (by Display ID) or every entity of a type back to a version's values.
- **Restore whole model:** replaces the metamodel, entities and relationships with a version.
- **Entity history:** every recorded change to one entity and its relationships, newest first.

A snapshot is taken automatically before every metamodel change and before every restore. History is never
overwritten; restores are logged in it. History sheets past 100,000 rows move their oldest rows to
`History\Archive_<year>.xlsx` during Sync Now.

## Templates, import and export

**Templates, import, export** on EA_Core's Home (or **Import / export** on EA_Entities' Home):

- **Generate template:** choose an entity type. You get a macro-free `.xlsx` in `Templates\` with an Entities sheet
  (dropdowns, required columns marked *) and a Relationships sheet that offers only the allowed relationship and target types.
  With **pre-fill**, existing entities are included for bulk updates, and a cleared cell clears the value on import. In a fresh
  template, blank cells are ignored for entities that already exist.
- **Import:** put filled templates in `Imports\`, then click **Choose files and preview**. Rows are matched by Entity ID, then by
  exact name within the type. IMPORT_Preview lists new, updated and unchanged rows and errors, field by field. **Commit import**
  takes a snapshot, writes the valid rows with history, and moves the files to `Imports\Done\`.
- **Export:** choose all entities, one type, everything within 1–5 relationships of one entity, or the full model with metamodel,
  in Excel or CSV. The file goes to `Exports\`.

## Reports

- **Refresh reports** (EA_Core Home): validates, saves EA_Core, then builds or refreshes `EA_Reports.xlsx`. It has a Dashboard (counts,
  errors, orphans, data-quality %, entities by domain and type, lifecycle, repository area), an Entity list with slicers,
  Relationships, a Matrix (pick source type, target type and relationship) and pages for Data quality and Orphans. EA_Reports has no macros
  and finds EA_Core in its own folder, so it can be opened on its own and refreshed with Data → Refresh All.
- **Entity detail** and **Impact** (both Homes, or right-click rows on a type sheet) paint RPT_EntityDetail and RPT_Impact in
  EA_Entities. Entity detail shows attributes, relationships and recent changes. Impact lists everything within n relationships, and how each one is reached.
- **Power BI:** see `Reports\README_PowerBI.md`.

## Diagrams (draw.io)

Needs draw.io Desktop (free, installs per user). **Diagrams (draw.io)** on either Home opens the form.

**First time in draw.io:** File → Open Library from → Device, pick `Library\ArchiMate4.xml`, then again for
`Library\ModelLibrary.xml`. Both palettes stay in draw.io's left panel.

- **Draw new things:** drag a shape from *ArchiMate 4*, type its name. For types with required attributes (such as
  Application Component's Owner and Lifecycle), select the shape, press Ctrl+M (Edit Data) and fill them in. Connect
  shapes with a connector from the bottom of the *ArchiMate 4* palette, which carries the relationship type.
- **Use existing things:** drag from *Model Library*. The shape is the real repository entity.
- **Generate a diagram:** pick around one entity (with hops), one type or all entities. Shapes are placed in bands by
  domain; tidy the layout in draw.io. Generating into an existing page adds only what is missing and keeps your layout.
  Right-click rows on a type sheet → **diagram of selected entities** does the same for a selection.
- **Harvest:** save and close the diagram in draw.io, then **Choose diagrams and preview**. HARVEST_Review lists new
  entities, renames, shapes linked by name to existing entities, and new relationships (checked against the allowed
  ArchiMate relationships). **Commit harvest** takes a snapshot, writes them with history, and stores the new IDs in
  the diagram file. Deleting a shape or connector in draw.io never deletes anything in the repository.
- **Update diagrams from model** refreshes names on every diagram and gives retired or deleted entities a red outline.
- **Views register** lists every page with its status: OK, stale (entities changed or retired), relationships not
  yet drawn, or file missing. Sync Now also refreshes it, together with both libraries.
- **Export images** writes PNG, SVG or PDF of every page to `Diagrams\Export\` using draw.io's command line.

draw.io must save diagrams uncompressed (the default in current versions). If the Workbench reports a compressed
page: in draw.io untick File → Properties → Compressed and save.

## Parent and child models

**Parent and child models** on either Home opens the form.

1. **Create child** (in the parent): name it and pick a scope: around entities (Display IDs plus hops), one entity type,
   a repository area, or the whole model. The child is a new folder next to this one, `EA_Model_<Parent>_<Child>`.
   Entities outside the scope that in-scope relationships point at come along as read-only references.
2. **Close the parent** (both workbooks). Excel cannot open two workbooks with the same name, so parent and child are
   never open together. Open the child's `EA_Entities.xlsm`; the first open removes everything outside the scope and
   saves the merge base.
3. Work in the child as normal.
4. **Prepare merge package** (in the child) writes `Merge\ChildPackage.xlsx`. Close the child.
5. **Choose child and preview** (in the parent) picks that package. MERGE_Review lists every difference with the base,
   parent and child values. Changes made only in the child are set to Keep = Child; changes made on both sides are
   conflicts set to Keep = Parent. Change Keep where you want, then **Apply merge**. A snapshot is taken first.
6. The child can keep working and merge again; each merge brings only what changed since the last one.

Diagrams new in the child are copied to the parent; diagrams changed in the child are copied next to the parent's as
`<name>_child.drawio`.

## Known limits

- The allowed-relationship rules are derived from the ArchiMate 4 generic metamodel and marked
  "ArchiMate4-derived (verify)" in MM_EntityRelationships. Check them against the relationship tables in your
  copy of the specification.
- Deleting a row on a type sheet does not delete the entity: the next sync puts the row back. To remove an
  entity, set its Status to Retired.
- Regex rules use the Windows VBScript regular-expression engine. If it is not installed, those rules are skipped and
  Validate All says so.
- The macros run in Windows Excel only. On Mac or the web, the workbooks open for reading only.
- Do not rename `EA_Core.xlsm` or `EA_Entities.xlsm`. The names are checked when either workbook opens.

- EA_Reports uses Power Query and dynamic-array formulas (Microsoft 365 or Excel 2021). If Excel asks about privacy levels
  when refreshing, choose to ignore them for this file.
- The Power BI project contains the model and measures; the report pages are laid out by you (about 15 minutes, see its README).

- Shape styles use draw.io's built-in ArchiMate 3 shapes (ArchiMate 4 keeps the same notation). The style of each
  type and relationship is in MM_Styles and can be changed there; then rebuild the stencil.
- Harvest reads names and new relationships from diagrams. Attribute values typed with Edit Data are used only when a
  shape creates a new entity; to change an existing entity's attributes, use Excel.

- Merge goes from child to parent. To bring later parent changes into a child, create a new child.
- Changes in the child to reference entities (outside its scope) are listed but not merged.

See `TEST_Phase1.md` to `TEST_Phase5.md` for the step-by-step acceptance tests.
