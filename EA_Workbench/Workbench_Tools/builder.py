"""Build EA_Core.xlsx and EA_Entities.xlsx for a model folder.

Usage: python3 builder.py <model_folder> [--no-sample]
The layout written here is the same layout the VBA generator maintains
(modGenerator / modPaint); change both together.
"""
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

import drawio
import model
import seed

NOW = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
BY = "seed"
SCHEMA_VERSION = 1
DATA_ROWS_SPARE = 200          # unlocked blank rows below type-sheet data

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
KEY_FONT = Font(color="808080", size=8)
REL_FONT = Font(italic=True, color="595959")
LINK_FONT = Font(italic=True, color="0563C1", underline="single")
NA_FILL = PatternFill("solid", fgColor="E7E6E6")
TITLE_FONT = Font(bold=True, size=16, color="1F3864")
THIN = Side(style="thin", color="BFBFBF")

meta = model.Meta()

SYNC_KEEP = ["EA_Core", "EA_Entities", "Skip"]
EXPORT_SCOPES = ["All entities", "One entity type", "Around one entity", "Full model with metamodel"]
DIAGRAM_SCOPES = ["Around one entity", "One entity type", "All entities"]
IMAGE_FORMATS = ["png", "svg", "pdf"]
VIEW_HEADERS = ["File", "Page", "Entities", "Relationships", "Generated", "Harvested", "Fingerprint", "Status", "Note"]
HARVEST_HEADERS = ["File", "Page", "CellID", "Kind", "Action", "EntityID", "DisplayID", "Name", "TypeCode", "Field",
                   "Current", "Proposed", "Severity", "Message"]
SAMPLE_DIAGRAM = ("Claims_Overview.drawio", "Claims overview")
MERGE_SCOPES = ["Around entities", "One entity type", "Repository area", "Whole model"]
MERGE_KEEP = ["Parent", "Child", "Skip"]
MERGE_HEADERS = ["Kind", "Action", "RecordID", "DisplayID", "Name", "TableOrType", "Field", "Base", "Parent", "Child",
                 "Keep", "Note"]
CHILDREN_HEADERS = ["ChildModelID", "ChildName", "Folder", "Created", "Scope", "LastMerged", "Status"]
SYNC_REVIEW_HEADERS = ["Found", "Situation", "Sheet", "SheetRow", "EntityID", "DisplayID", "Name", "TypeCode",
                       "AttrCode", "Attribute", "CoreValue", "SheetValue", "Keep", "Status", "Note"]
IMPORT_HEADERS = ["File", "Sheet", "Row", "Kind", "Action", "EntityID", "DisplayID", "Name", "TypeCode", "Field",
                  "Current", "Imported", "Severity", "Message"]
VERSION_HEADERS = ["VersionNo", "Label", "Taken", "TakenBy", "Auto", "Reason", "CoreFile", "EntitiesFile",
                   "Entities", "Relationships", "SchemaVersion"]


# ---------------------------------------------------------------- helpers
def add_table(ws, name, headers, rows, start_row=1, style="TableStyleMedium2", widths=None):
    for c, h in enumerate(headers, 1):
        ws.cell(start_row, c, h)
    for r, row in enumerate(rows, start_row + 1):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)
    last_row = start_row + max(len(rows), 1)
    ref = f"A{start_row}:{get_column_letter(len(headers))}{last_row}"
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = ws.cell(start_row + 1, 1)
    for c, h in enumerate(headers, 1):
        w = (widths or {}).get(h, max(12, min(40, len(h) + 4)))
        ws.column_dimensions[get_column_letter(c)].width = w
    return last_row


def define(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def list_validation(ws, sqref, named, allow_blank=True):
    dv = DataValidation(type="list", formula1=f"={named}", allow_blank=allow_blank, showErrorMessage=True,
                        errorTitle="Not in list", error="Pick a value from the list.")
    ws.add_data_validation(dv)
    dv.add(sqref)


def history_headers():
    return ["ChangeID", "Timestamp", "User", "Action", "RecordID", "AttrCode", "OldValue", "NewValue",
            "Source", "BatchID"]


def write_list_cache(wb, ws, lists, extra):
    """One column per list; named range lst_<Name> per column."""
    ws.sheet_state = "hidden"
    allv = dict(lists)
    allv.update(extra)
    for c, (name, vals) in enumerate(allv.items(), 1):
        ws.cell(1, c, name).font = Font(bold=True)
        for r, v in enumerate(vals, 2):
            ws.cell(r, c, v)
        col = get_column_letter(c)
        define(wb, f"lst_{name}", f"'{ws.title}'!${col}$2:${col}${max(2, len(vals) + 1)}")


def counters_from(entities):
    nxt = {}
    for t in seed.ENTITY_TYPES:
        nxt[t[4]] = 1
    for e in entities:
        p = meta.types[e["TypeCode"]]["prefix"]
        nxt[p] += 1
    return nxt


# ---------------------------------------------------------------- sample data
def build_sample(use_sample=True):
    entities, rels = [], []
    if not use_sample:
        return entities, rels
    seqs = {}
    for type_code, name, desc, attrs in seed.SAMPLE_ENTITIES:
        prefix = meta.types[type_code]["prefix"]
        seqs[prefix] = seqs.get(prefix, 0) + 1
        e = dict(EntityID=seed.guid("entity/" + name), TypeCode=type_code,
                 DisplayID=f"{prefix}-{seqs[prefix]:04d}", Name=name, Description=desc, Status="Active",
                 RepositoryArea="Architecture Landscape")
        for k, v in attrs.items():
            assert meta.is_mapped(type_code, k), f"{k} not mapped to {type_code}"
            if meta.attrs[k]["datatype"] == "Date":
                v = datetime.strptime(v, "%Y-%m-%d")
            e[k] = v
        entities.append(e)
    by_name = {e["Name"]: e for e in entities}
    for s, rel, t in seed.SAMPLE_RELS:
        se, te = by_name[s], by_name[t]
        assert meta.is_allowed(se["TypeCode"], rel, te["TypeCode"]), (s, rel, t)
        rels.append(dict(RelID=seed.guid(f"rel/{s}/{rel}/{t}"), RelCode=rel, SourceID=se["EntityID"],
                         TargetID=te["EntityID"], Label="", Description="", Status="Active"))
    return entities, rels


# ---------------------------------------------------------------- EA_Core
MASTER_SYS_FRONT = ["EntityID", "TypeCode", "DisplayID", "Name", "Description", "Status", "RepositoryArea"]
MASTER_SYS_BACK = ["Created", "CreatedBy", "Modified", "ModifiedBy", "RowHash"]


def master_columns():
    return MASTER_SYS_FRONT + [a[0] for a in seed.ATTRS] + MASTER_SYS_BACK


def entity_row_hash(e):
    hk = model.MASTER_SYS_HASHED + [a[0] for a in seed.ATTRS]
    return model.master_hash(hk, [e.get(c) for c in hk])


def sample_view_row(entities, rels):
    live = [e for e in entities if e.get("Status") != "Retired"]
    fp = drawio.fingerprint([(e["EntityID"], entity_row_hash(e)) for e in live])
    on = {e["EntityID"] for e in live}
    nrel = sum(1 for r in rels if r["SourceID"] in on and r["TargetID"] in on and r.get("Status") != "Retired")
    return [SAMPLE_DIAGRAM[0], SAMPLE_DIAGRAM[1], len(live), nrel, NOW, "", fp, "OK", "Generated by the builder"]


def write_drawio_files(folder: Path, entities, rels, model_id):
    """Library/ArchiMate4.xml, Library/ModelLibrary.xml and the sample diagram (VBA rebuilds them later)."""
    lib = folder / "Library"
    lib.mkdir(exist_ok=True)
    (folder / "Diagrams").mkdir(exist_ok=True)
    mandatory = {t: [m["attr"] for m in maps if m["mandatory"] == "Y"] for t, maps in meta.map.items()}
    (lib / "ArchiMate4.xml").write_text(drawio.stencil(model_id, mandatory), encoding="utf-8")
    (lib / "ModelLibrary.xml").write_text(drawio.model_library(entities, model_id), encoding="utf-8")
    if entities:
        (folder / "Diagrams" / SAMPLE_DIAGRAM[0]).write_text(
            drawio.generate_diagram(entities, rels, SAMPLE_DIAGRAM[1], model_id), encoding="utf-8")


def build_core(folder: Path, entities, rels, model_id, model_name):
    wb = Workbook()
    home = wb.active
    home.title = "Home"

    # ---- metamodel sheets
    ws = wb.create_sheet("MM_EntityTypes")
    rows = []
    for code, name, dom, asp, pfx in seed.ENTITY_TYPES:
        rows.append([seed.guid("type/" + code), code, name, dom, asp, "", f"T_{code}", pfx,
                     f"Entity:{code}", "Active", NOW, BY])
    add_table(ws, "tblEntityTypes", ["TypeID", "TypeCode", "Name", "Domain", "Aspect", "Description",
                                     "SheetName", "IDPrefix", "StyleKey", "Status", "Modified", "ModifiedBy"],
              rows, widths={"TypeID": 38, "TypeCode": 24, "Name": 24, "Aspect": 26, "SheetName": 26})

    ws = wb.create_sheet("MM_AttributeTypes")
    rows = []
    for code, name, dt, lst in seed.SYSTEM_ATTRS:
        rows.append([seed.guid("attr/" + code), code, name, dt, lst, "", "", "System attribute", "Y",
                     "Active", NOW, BY])
    for code, name, dt, lst, unit, desc in seed.ATTRS:
        rows.append([seed.guid("attr/" + code), code, name, dt, lst, unit, "", desc, "N", "Active", NOW, BY])
    add_table(ws, "tblAttributeTypes", ["AttrID", "AttrCode", "Name", "DataType", "ListName", "Unit",
                                        "DefaultValue", "Description", "IsSystem", "Status", "Modified",
                                        "ModifiedBy"],
              rows, widths={"AttrID": 38, "Description": 40})

    ws = wb.create_sheet("MM_RelationshipTypes")
    rows = [[seed.guid("rel/" + r[0]), r[0], r[1], r[2], r[3], f"Rel:{r[0]}", r[4], "Active", NOW, BY]
            for r in seed.REL_TYPES]
    add_table(ws, "tblRelTypes", ["RelTypeID", "RelCode", "Name", "Category", "Directed", "StyleKey",
                                  "Description", "Status", "Modified", "ModifiedBy"],
              rows, widths={"RelTypeID": 38, "Description": 60})

    ws = wb.create_sheet("MM_EntityAttributes")
    rows = []
    for t, maps in meta.map.items():
        for m in maps:
            rows.append([seed.guid(f"map/{t}/{m['attr']}"), t, m["attr"], m["mandatory"], m["order"],
                         m["section"], "", "Active", NOW, BY])
    add_table(ws, "tblEntityAttributes", ["MapID", "TypeCode", "AttrCode", "Mandatory", "DisplayOrder",
                                          "Section", "DefaultOverride", "Status", "Modified", "ModifiedBy"],
              rows, widths={"MapID": 38, "TypeCode": 24, "AttrCode": 22})

    ws = wb.create_sheet("MM_EntityRelationships")
    rows = [[seed.guid(f"allow/{s}/{r}/{t}"), s, r, t, "", "", "ArchiMate4-derived (verify)", n, "Active",
             NOW, BY] for s, r, t, n in seed.R]
    add_table(ws, "tblEntityRels", ["MapID", "SourceType", "RelCode", "TargetType", "MinCard", "MaxCard",
                                    "Origin", "Note", "Status", "Modified", "ModifiedBy"],
              rows, widths={"MapID": 38, "SourceType": 24, "TargetType": 24, "Origin": 28, "Note": 50})
    ws.cell(1, 13, "Tokens: TypeCode, * = any type, SAME = same type as source, G:<group> = MM_TypeGroups")
    ws.cell(1, 13).font = Font(italic=True, color="595959")

    ws = wb.create_sheet("MM_TypeGroups")
    rows = [[g, t] for g, members in seed.G.items() for t in members]
    add_table(ws, "tblTypeGroups", ["GroupName", "TypeCode"], rows, widths={"GroupName": 20, "TypeCode": 24})

    ws = wb.create_sheet("MM_Rules")
    rows = [[seed.guid(f"rule/{r[0]}/{r[2]}/{r[3]}"), *r, "Active", NOW, BY] for r in seed.RULES]
    add_table(ws, "tblRules", ["RuleID", "AttrCode", "TypeCode", "RuleType", "Param1", "Param2", "Severity",
                               "Message", "Status", "Modified", "ModifiedBy"],
              rows, widths={"RuleID": 38, "Param1": 24, "Message": 50})

    ws = wb.create_sheet("MM_Lists")
    rows = [[ln, v, i + 1, "", "Active"] for ln, vals in seed.LISTS.items() for i, v in enumerate(vals)]
    add_table(ws, "tblLists", ["ListName", "Value", "SortOrder", "Colour", "Status"], rows,
              widths={"ListName": 22, "Value": 36})

    ws = wb.create_sheet("MM_Styles")
    rows = [[f"Entity:{c}", "Entity", seed.DOMAIN_COLOURS[d], drawio.entity_style(c)]
            for c, _, d, _, _ in seed.ENTITY_TYPES]
    rows += [[f"Rel:{r[0]}", "Relationship", "", drawio.REL_STYLES[r[0]]] for r in seed.REL_TYPES]
    add_table(ws, "tblStyles", ["StyleKey", "AppliesTo", "FillColour", "DrawioStyle"], rows,
              widths={"StyleKey": 30, "DrawioStyle": 90})

    # ---- DATA_Master
    ws = wb.create_sheet("DATA_Master")
    cols = master_columns()
    hdr_names = {a[0]: a[1] for a in seed.SYSTEM_ATTRS}
    hdr_names.update({a[0]: a[1] for a in seed.ATTRS})
    for c, code in enumerate(cols, 1):
        ws.cell(1, c, code).font = KEY_FONT
    ws.row_dimensions[1].hidden = True
    rows = []
    for e in entities:
        row = []
        for code in cols:
            row.append(e.get(code))
        hk = model.MASTER_SYS_HASHED + [a[0] for a in seed.ATTRS]
        hv = [e.get(c) for c in hk]
        row[cols.index("Created")] = NOW
        row[cols.index("CreatedBy")] = BY
        row[cols.index("Modified")] = NOW
        row[cols.index("ModifiedBy")] = BY
        row[cols.index("RowHash")] = model.master_hash(hk, hv)
        rows.append(row)
    headers = [hdr_names[c] if c in hdr_names else c for c in cols]
    # Table headers must be unique; use display names
    last = add_table(ws, "tblMaster", headers, rows, start_row=2, style="TableStyleLight9",
                     widths={"Entity ID": 38, "Name": 32, "Description": 40, "Type": 24})
    ws.freeze_panes = "E3"
    ws.column_dimensions["A"].hidden = True
    attr_first = len(MASTER_SYS_FRONT) + 1
    attr_last = attr_first + len(seed.ATTRS) - 1
    fill_to = max(last, 2 + len(rows))
    # date / number formats
    for i, a in enumerate(seed.ATTRS):
        col = attr_first + i
        fmt = {"Date": "yyyy-mm-dd", "Currency": "#,##0", "Integer": "0", "Decimal": "0.00",
               "Percent": "0%"}.get(a[2])
        if fmt:
            for r in range(3, fill_to + 1):
                ws.cell(r, col).number_format = fmt
    # not-applicable shading: MapKeys holds "Type|Attr" for every active mapping
    rng = f"{get_column_letter(attr_first)}3:{get_column_letter(attr_last)}5000"
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=[f'AND($B3<>"",ISNA(MATCH($B3&"|"&{get_column_letter(attr_first)}$1,MapKeys,0)))'],
                         fill=NA_FILL, stopIfTrue=True))
    # list validations for system + list attributes
    list_validation(ws, "B3:B5000", "lst_TypeCodes", allow_blank=False)
    list_validation(ws, "F3:F5000", "lst_Status")
    list_validation(ws, "G3:G5000", "lst_RepositoryArea")
    for i, a in enumerate(seed.ATTRS):
        col = get_column_letter(attr_first + i)
        if a[2] == "List" and a[3]:
            list_validation(ws, f"{col}3:{col}5000", f"lst_{a[3]}")
        elif a[2] == "Date":
            dv = DataValidation(type="date", operator="between", formula1="DATE(1900,1,1)",
                                formula2="DATE(2200,12,31)", allow_blank=True, showErrorMessage=True,
                                error="Enter a date (yyyy-mm-dd).")
            ws.add_data_validation(dv)
            dv.add(f"{col}3:{col}5000")
        elif a[2] == "Integer":
            dv = DataValidation(type="whole", operator="between", formula1="-1000000000",
                                formula2="1000000000", allow_blank=True, showErrorMessage=True,
                                error="Enter a whole number.")
            ws.add_data_validation(dv)
            dv.add(f"{col}3:{col}5000")
        elif a[2] in ("Decimal", "Currency", "Percent"):
            dv = DataValidation(type="decimal", operator="between", formula1="-1E+15", formula2="1E+15",
                                allow_blank=True, showErrorMessage=True, error="Enter a number.")
            ws.add_data_validation(dv)
            dv.add(f"{col}3:{col}5000")

    # ---- DATA_Relationships
    ws = wb.create_sheet("DATA_Relationships")
    rheaders = ["RelID", "RelCode", "SourceID", "TargetID", "SourceName", "SourceType", "TargetName",
                "TargetType", "Label", "Description", "Status", "Created", "CreatedBy", "Modified",
                "ModifiedBy", "RowHash"]
    rrows = []
    for i, r in enumerate(rels, 2):
        rrows.append([r["RelID"], r["RelCode"], r["SourceID"], r["TargetID"], None, None, None, None,
                      r["Label"], r["Description"], r["Status"], NOW, BY, NOW, BY,
                      model.row_hash([r["RelCode"], r["SourceID"], r["TargetID"], r["Label"], r["Description"],
                                      r["Status"]])])
    last = add_table(ws, "tblRel", rheaders, rrows, widths={"RelID": 38, "SourceID": 38, "TargetID": 38,
                                                            "SourceName": 32, "TargetName": 32,
                                                            "SourceType": 22, "TargetType": 22})
    for r in range(2, max(last, 2 + len(rrows)) + 1):
        ws.cell(r, 5, f'=IFERROR(INDEX(DATA_Master!$D:$D,MATCH(C{r},DATA_Master!$A:$A,0)),"(missing)")')
        ws.cell(r, 6, f'=IFERROR(INDEX(DATA_Master!$B:$B,MATCH(C{r},DATA_Master!$A:$A,0)),"(missing)")')
        ws.cell(r, 7, f'=IFERROR(INDEX(DATA_Master!$D:$D,MATCH(D{r},DATA_Master!$A:$A,0)),"(missing)")')
        ws.cell(r, 8, f'=IFERROR(INDEX(DATA_Master!$B:$B,MATCH(D{r},DATA_Master!$A:$A,0)),"(missing)")')
    list_validation(ws, "B2:B5000", "lst_RelCodes", allow_blank=False)
    list_validation(ws, "K2:K5000", "lst_Status")

    # ---- history
    for name, tbl in (("H_Metamodel", "tblHMetamodel"), ("H_Relationships", "tblHRelationships")):
        ws = wb.create_sheet(name)
        hrows = []
        if name == "H_Metamodel":
            hrows.append([seed.guid("h/seed"), NOW, BY, "Create", "", "", "", "ArchiMate 4 seed loaded",
                          "Seed", seed.guid("batch/seed")])
        add_table(ws, tbl, history_headers(), hrows, style="TableStyleLight1",
                  widths={"ChangeID": 38, "RecordID": 38, "OldValue": 30, "NewValue": 30, "BatchID": 38})

    # ---- system sheets
    ws = wb.create_sheet("SYS_Issues")
    add_table(ws, "tblIssues", ["Found", "Sheet", "RecordID", "DisplayID", "Name", "TypeCode", "AttrCode",
                                "RuleType", "Severity", "Message", "Go to"], [],
              style="TableStyleLight10", widths={"RecordID": 38, "Message": 60, "Name": 30})

    ws = wb.create_sheet("SYS_Settings")
    settings = [
        ["ModelID", model_id], ["ModelName", model_name], ["SchemaVersion", SCHEMA_VERSION],
        ["CoreFile", "EA_Core.xlsm"], ["EntitiesFile", "EA_Entities.xlsm"], ["ReportsFile", "EA_Reports.xlsx"],
        ["ProtectPassword", "ea-workbench"], ["DrawioPath", r"C:\Program Files\draw.io\draw.io.exe"],
        ["LastValidateAll", ""], ["WorkbenchVersion", "1.1"],
        ["WBSaveStamp", ""], ["LastSync", ""], ["LastFullSync", ""],
        ["ParentModelID", ""], ["ParentModelName", ""], ["ParentFolder", ""], ["ChildCreated", ""],
    ]
    add_table(ws, "tblSettings", ["Key", "Value"], settings, widths={"Key": 22, "Value": 50})

    ws = wb.create_sheet("SYS_Counters")
    nxt = counters_from(entities)
    add_table(ws, "tblCounters", ["IDPrefix", "NextNo"], [[k, v] for k, v in nxt.items()])

    # ---- caches (hidden)
    ws = wb.create_sheet("SYS_ListCache")
    extra = {"TypeCodes": [t[0] for t in seed.ENTITY_TYPES], "RelCodes": [r[0] for r in seed.REL_TYPES],
             "AttrCodes": [a[0] for a in seed.ATTRS], "ListNames": list(seed.LISTS.keys()),
             "SyncKeep": SYNC_KEEP, "ExportScope": EXPORT_SCOPES, "ExportFormat": ["Excel", "CSV"],
             "DiagramScope": DIAGRAM_SCOPES, "ImageFormat": IMAGE_FORMATS,
             "MergeScope": MERGE_SCOPES, "MergeKeep": MERGE_KEEP}
    write_list_cache(wb, ws, {k: v for k, v in seed.LISTS.items()}, extra)

    ws = wb.create_sheet("SYS_MapCache")
    ws.sheet_state = "hidden"
    ws.cell(1, 1, "MapKey").font = Font(bold=True)
    keys = [f"{t}|{m['attr']}" for t, maps in meta.map.items() for m in maps]
    for r, k in enumerate(keys, 2):
        ws.cell(r, 1, k)
    define(wb, "MapKeys", f"'SYS_MapCache'!$A$2:$A${len(keys) + 1}")

    # ---- Phase 2: sync review, versions
    ws = wb.create_sheet("SYNC_Review")
    ws["A1"] = "Sync review"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Differences between EA_Core and EA_Entities that need a decision. Set Keep to EA_Core, "
                "EA_Entities or Skip (decide later), then click Apply decisions.")
    ws["A2"].font = Font(italic=True, color="595959")
    add_table(ws, "tblSyncReview", SYNC_REVIEW_HEADERS, [], start_row=5, style="TableStyleLight11",
              widths={"EntityID": 38, "Name": 30, "Situation": 34, "CoreValue": 26, "SheetValue": 26, "Note": 40})
    ws.freeze_panes = "A6"
    list_validation(ws, "M6:M5000", "lst_SyncKeep")

    ws = wb.create_sheet("SYS_Versions")
    add_table(ws, "tblVersions", VERSION_HEADERS, [], widths={"Label": 30, "Reason": 30, "CoreFile": 50,
                                                              "EntitiesFile": 50})

    ws = wb.create_sheet("IMPORT_Preview")
    ws["A1"] = "Import preview"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Nothing is written until you click Commit import. Rows with errors are skipped."
    ws["A2"].font = Font(italic=True, color="595959")
    add_table(ws, "tblImportPreview", IMPORT_HEADERS, [], start_row=6, style="TableStyleLight12",
              widths={"File": 30, "EntityID": 38, "Name": 30, "Current": 24, "Imported": 24, "Message": 50})
    ws.freeze_panes = "A7"

    # ---- Phase 4: draw.io views register and harvest review
    ws = wb.create_sheet("SYS_Views")
    ws["A1"] = "Views register"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Every diagram page generated, updated or harvested by the Workbench. Status is refreshed by Views register."
    ws["A2"].font = Font(italic=True, color="595959")
    view_rows = []
    if entities:
        view_rows.append(sample_view_row(entities, rels))
    add_table(ws, "tblViews", VIEW_HEADERS, view_rows, start_row=4, style="TableStyleLight13",
              widths={"File": 34, "Page": 24, "Generated": 22, "Harvested": 22, "Fingerprint": 12, "Status": 50,
                      "Note": 34})
    for r in range(5, 5 + max(1, len(view_rows))):
        ws.cell(r, 7).number_format = "@"

    ws = wb.create_sheet("HARVEST_Review")
    ws["A1"] = "Harvest review"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Nothing is written until you click Commit harvest. Rows with errors are skipped."
    ws["A2"].font = Font(italic=True, color="595959")
    add_table(ws, "tblHarvest", HARVEST_HEADERS, [], start_row=6, style="TableStyleLight12",
              widths={"File": 28, "Page": 18, "CellID": 20, "EntityID": 38, "Name": 34, "Current": 22,
                      "Proposed": 22, "Message": 60})
    ws.freeze_panes = "A7"

    # ---- Phase 5: parent and child models
    ws = wb.create_sheet("MERGE_Review")
    ws["A1"] = "Merge review"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Three-way comparison of the merge base, this model (Parent) and the child. Conflicts default to Keep = "
                "Parent. Set Keep, then click Apply merge.")
    ws["A2"].font = Font(italic=True, color="595959")
    add_table(ws, "tblMergeReview", MERGE_HEADERS, [], start_row=6, style="TableStyleLight14",
              widths={"RecordID": 38, "Name": 34, "TableOrType": 22, "Base": 24, "Parent": 24, "Child": 24, "Note": 50})
    ws.freeze_panes = "A7"
    list_validation(ws, "K7:K20000", "lst_MergeKeep")

    ws = wb.create_sheet("SYS_Children")
    ws["A1"] = "Child models"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Child models created from this model, and when each was last merged back."
    ws["A2"].font = Font(italic=True, color="595959")
    add_table(ws, "tblChildren", CHILDREN_HEADERS, [], start_row=4, style="TableStyleLight15",
              widths={"ChildModelID": 38, "ChildName": 24, "Folder": 60, "Created": 22, "Scope": 28, "LastMerged": 22,
                      "Status": 34})

    ws = wb.create_sheet("VERSION_Compare")
    ws["A1"] = "Version compare and entity history"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Filled by the Version history buttons."
    ws["A2"].font = Font(italic=True, color="595959")
    ws.sheet_view.showGridLines = False

    # ---- wizard sheets (painted by VBA at run time)
    for name, title in (("FORM_Entity", "Add / Edit Entity"), ("FORM_Relationship", "Add Relationship"),
                        ("FORM_Type", "Add / Edit Type"), ("FORM_Rule", "Define Rule"),
                        ("FORM_Versions", "Version history"), ("FORM_IO", "Templates, import and export"),
                        ("FORM_Diagram", "Diagrams (draw.io)"), ("FORM_Merge", "Parent and child models")):
        ws = wb.create_sheet(name)
        ws["A1"] = title
        ws["A1"].font = TITLE_FONT
        ws["A2"] = "This form is painted by the Workbench macros. Use the button on Home to open it."
        ws["A2"].font = Font(italic=True, color="595959")
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 44
        ws.column_dimensions["C"].width = 4
        ws.column_dimensions["D"].width = 60
        ws.sheet_view.showGridLines = False

    # ---- Home
    build_home(home, "EA_Core", model_name, [
        ("Entities", '=COUNTA(DATA_Master!A:A)-2'),
        ("Relationships", '=COUNTA(DATA_Relationships!A:A)-1'),
        ("Entity types", '=COUNTA(MM_EntityTypes!A:A)-1'),
        ("Attribute types", '=COUNTA(MM_AttributeTypes!A:A)-1'),
        ("Open issues", '=COUNTA(SYS_Issues!A:A)-1'),
        ("Schema version", '=INDEX(SYS_Settings!B:B,MATCH("SchemaVersion",SYS_Settings!A:A,0))'),
    ])
    for ws in wb.worksheets:
        if ws.title.startswith(("MM_", "DATA_")):
            ws.sheet_properties.tabColor = "1F3864" if ws.title.startswith("MM_") else "2E75B6"
        elif ws.title.startswith("H_"):
            ws.sheet_properties.tabColor = "7F7F7F"
        elif ws.title.startswith("FORM_"):
            ws.sheet_properties.tabColor = "70AD47"
        elif ws.title in ("SYNC_Review", "VERSION_Compare", "IMPORT_Preview"):
            ws.sheet_properties.tabColor = "ED7D31"
    out = folder / "EA_Core.xlsx"
    wb.save(out)
    return out


def build_home(ws, title, model_name, stats):
    ws.sheet_view.showGridLines = False
    ws["B2"] = f"EA Repository Workbench · {title}"
    ws["B2"].font = TITLE_FONT
    ws["B3"] = f"Model: {model_name}"
    ws["B3"].font = Font(size=12, color="595959")
    ws["B5"] = "Status"
    ws["B5"].font = Font(bold=True)
    for i, (label, f) in enumerate(stats, 6):
        ws.cell(i, 2, label)
        ws.cell(i, 3, f).font = Font(bold=True)
    ws["E5"] = "Actions"
    ws["E5"].font = Font(bold=True)
    ws["E6"] = "Buttons appear here after setup (see README.md, step 2)."
    ws["E6"].font = Font(italic=True, color="595959")
    ws["B4"] = ""  # banner row for sync / issues messages
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 4
    ws.column_dimensions["E"].width = 30
    ws.sheet_properties.tabColor = "C00000"


# ---------------------------------------------------------------- EA_Entities
TYPE_SHEET_FIXED = ["EntityID", "RowKind", "PaintHash", "DisplayID", "Name", "Description", "Status",
                    "RepositoryArea"]
FIXED_NAMES = {"EntityID": "Entity ID", "RowKind": "Row", "PaintHash": "Paint hash", "DisplayID": "Display ID",
               "Name": "Name", "Description": "Description", "Status": "Status", "RepositoryArea": "Repository Area"}


def type_sheet_columns(type_code):
    return TYPE_SHEET_FIXED + [m["attr"] for m in meta.attrs_for_type(type_code)]


def paint_type_sheet(wb, type_code, entities, rels, by_id):
    """Mirror of modPaint.PaintTypeSheet."""
    t = meta.types[type_code]
    ws = wb.create_sheet(f"T_{type_code}")
    cols = type_sheet_columns(type_code)
    fill = PatternFill("solid", fgColor=seed.DOMAIN_COLOURS[t["domain"]])
    for c, code in enumerate(cols, 1):
        ws.cell(1, c, code).font = KEY_FONT
        if code in FIXED_NAMES:
            name = FIXED_NAMES[code]
        else:
            name = meta.attrs[code]["name"] + (" *" if meta.mandatory(type_code, code) else "")
        cell = ws.cell(2, c, name)
        cell.font = Font(bold=True)
        cell.fill = fill
        cell.border = Border(bottom=Side(style="medium", color="404040"))
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        if code in meta.attrs and meta.attrs[code]["desc"]:
            cell.comment = Comment(meta.attrs[code]["desc"], "Workbench")
    ws.cell(1, 1).value = "EntityID"
    ws.row_dimensions[1].hidden = True
    ws.row_dimensions[2].height = 30
    for col in ("A", "B", "C"):
        ws.column_dimensions[col].hidden = True
    widths = {"DisplayID": 12, "Name": 34, "Description": 40, "Status": 10, "RepositoryArea": 24}
    for c, code in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(c)].width = widths.get(code, 16)
    ws.freeze_panes = "F3"
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.tabColor = seed.DOMAIN_COLOURS[t["domain"]]

    rel_names = {r[0]: r[1] for r in seed.REL_TYPES}
    inverse = {"composed of": "part of", "aggregates": "aggregated by", "assigned to": "assigned from",
               "realizes": "realized by", "serves": "served by", "accesses": "accessed by",
               "influences": "influenced by", "associated with": "associated with", "triggers": "triggered by",
               "flows to": "flows from", "specializes": "specialized by"}
    r = 3
    mine = sorted([e for e in entities if e["TypeCode"] == type_code], key=lambda e: e["Name"].lower())
    for e in mine:
        for c, code in enumerate(cols, 1):
            if code == "RowKind":
                v = "E"
            elif code == "PaintHash":
                v = ""      # computed by VBA on first open
            else:
                v = e.get(code)
            cell = ws.cell(r, c, v)
            if code in meta.attrs:
                fmt = {"Date": "yyyy-mm-dd", "Currency": "#,##0"}.get(meta.attrs[code]["datatype"])
                if fmt:
                    cell.number_format = fmt
        ws.cell(r, 5).font = Font(bold=True)
        r += 1
        out_rels = [x for x in rels if x["SourceID"] == e["EntityID"]]
        in_rels = [x for x in rels if x["TargetID"] == e["EntityID"]]
        for x, direction in [(x, "out") for x in out_rels] + [(x, "in") for x in in_rels]:
            other = by_id[x["TargetID"] if direction == "out" else x["SourceID"]]
            verb = rel_names[x["RelCode"]]
            if direction == "in":
                verb = inverse.get(verb, verb)
            arrow = "→" if direction == "out" else "←"
            ws.cell(r, 1, x["RelID"])
            ws.cell(r, 2, "R")
            ws.cell(r, 4, f"  {arrow} {verb}").font = REL_FONT
            other_sheet = f"T_{other['TypeCode']}"
            link = (f'=HYPERLINK("#\'{other_sheet}\'!E"&MATCH("{other["EntityID"]}",\'{other_sheet}\'!$A:$A,0),'
                    f'"{other["Name"]}")')
            ws.cell(r, 5, link).font = LINK_FONT
            ws.cell(r, 6, meta.types[other["TypeCode"]]["name"]).font = REL_FONT
            ws.cell(r, 7, x["Label"]).font = REL_FONT
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].hidden = True
            r += 1

    # validations on the editable area (data rows + spare rows)
    last = r + DATA_ROWS_SPARE
    list_validation(ws, f"G3:G{last}", "lst_Status")
    list_validation(ws, f"H3:H{last}", "lst_RepositoryArea")
    for c, code in enumerate(cols, 1):
        if code not in meta.attrs:
            continue
        a = meta.attrs[code]
        col = get_column_letter(c)
        if a["datatype"] == "List" and a["list"]:
            list_validation(ws, f"{col}3:{col}{last}", f"lst_{a['list']}")
        elif a["datatype"] == "Date":
            dv = DataValidation(type="date", operator="between", formula1="DATE(1900,1,1)",
                                formula2="DATE(2200,12,31)", allow_blank=True, showErrorMessage=True,
                                error="Enter a date (yyyy-mm-dd).")
            ws.add_data_validation(dv)
            dv.add(f"{col}3:{col}{last}")
            for rr in range(3, last + 1):
                ws.cell(rr, c).number_format = "yyyy-mm-dd"
        elif a["datatype"] == "Integer":
            dv = DataValidation(type="whole", operator="between", formula1="-1000000000",
                                formula2="1000000000", allow_blank=True, showErrorMessage=True,
                                error="Enter a whole number.")
            ws.add_data_validation(dv)
            dv.add(f"{col}3:{col}{last}")
    return ws


def build_entities(folder: Path, entities, rels, model_id, model_name):
    wb = Workbook()
    home = wb.active
    home.title = "Home"
    by_id = {e["EntityID"]: e for e in entities}
    order = ["Strategy", "Motivation", "Business", "Application", "Technology", "Common",
             "Implementation & Migration"]
    types_sorted = sorted(seed.ENTITY_TYPES, key=lambda t: (order.index(t[2]), t[1]))
    for t in types_sorted:
        paint_type_sheet(wb, t[0], entities, rels, by_id)

    # caches
    ws = wb.create_sheet("SYS_Lists")
    extra = {"TypeCodes": [t[0] for t in seed.ENTITY_TYPES], "RelCodes": [r[0] for r in seed.REL_TYPES],
             "SyncKeep": SYNC_KEEP}
    write_list_cache(wb, ws, {k: v for k, v in seed.LISTS.items()}, extra)

    ws = wb.create_sheet("SYS_RuleCache")
    ws.sheet_state = "hidden"
    rows = []
    for t in meta.map:
        for m in meta.attrs_for_type(t):
            a = meta.attrs[m["attr"]]
            rows.append([t, m["attr"], a["name"], a["datatype"], a["list"], m["mandatory"], m["order"], m["section"]])
    add_table(ws, "tblCacheAttrs", ["TypeCode", "AttrCode", "Name", "DataType", "ListName", "Mandatory",
                                    "DisplayOrder", "Section"], rows)
    start = len(rows) + 4
    rrows = [list(r) for r in seed.RULES]
    add_table(ws, "tblCacheRules", ["AttrCode", "TypeCode", "RuleType", "Param1", "Param2", "Severity",
                                    "Message"], rrows, start_row=start)

    ws = wb.create_sheet("SYS_Pending")
    ws.sheet_state = "hidden"
    add_table(ws, "tblPending", ["Sheet", "Row", "EntityID", "AttrCode", "NewValue", "EditedAt"], [])

    ws = wb.create_sheet("SYS_SyncState")
    ws.sheet_state = "hidden"
    add_table(ws, "tblSyncState", ["Key", "Value"], [
        ["ModelID", model_id], ["CoreFile", "EA_Core.xlsm"], ["SchemaVersion", SCHEMA_VERSION],
        ["LastSync", ""], ["LastPush", ""], ["WBSaveStamp", ""],
    ])

    stats = [(f"{t[1]}", f"=COUNTIF('T_{t[0]}'!B:B,\"E\")") for t in types_sorted
             if any(e["TypeCode"] == t[0] for e in entities)]
    build_home(home, "EA_Entities", model_name, stats)
    home["E7"] = "Type sheets (click to open):"
    home["E7"].font = Font(bold=True)
    for i, t in enumerate(types_sorted):
        c = home.cell(8 + i, 5, f"{t[1]}  ({t[2]})")
        c.hyperlink = f"#'T_{t[0]}'!E3"
        c.font = Font(color="0563C1", underline="single")
    out = folder / "EA_Entities.xlsx"
    wb.save(out)
    return out


def main():
    build(Path(sys.argv[1]), "--no-sample" not in sys.argv)


def build(folder: Path, use_sample: bool = True) -> str:
    """Builds EA_Core.xlsx, EA_Entities.xlsx, model.json and the draw.io files into folder; returns a summary."""
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    model_id = seed.guid("model/" + folder.name)
    model_name = folder.name.replace("EA_Model_", "").replace("_", " ")
    entities, rels = build_sample(use_sample)
    core = build_core(folder, entities, rels, model_id, model_name)
    write_drawio_files(folder, entities, rels, model_id)
    ents = build_entities(folder, entities, rels, model_id, model_name)
    (folder / "model.json").write_text(json.dumps({
        "modelId": model_id, "name": model_name, "parentModelId": None,
        "files": {"core": "EA_Core.xlsm", "entities": "EA_Entities.xlsm", "reports": "EA_Reports.xlsx"},
        "schemaVersion": SCHEMA_VERSION, "lastSync": None, "baseSnapshot": None,
    }, indent=2), encoding="utf-8")
    summary = f"built {core.name} and {ents.name}: {len(entities)} entities, {len(rels)} relationships"
    print(summary)
    return summary


if __name__ == "__main__":
    main()
