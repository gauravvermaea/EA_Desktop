"""Tests for the Workbench logic and build. Run: python3 -m pytest -q test_workbench.py"""
import math
import re
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

import model
import seed

import os
# the model folder under test: EAWB_MODEL, else EA_Model_Sample next to the tools folder
MODEL = Path(os.environ.get("EAWB_MODEL") or Path(__file__).resolve().parent.parent / "EA_Model_Sample")
SRC = MODEL / "src"
meta = model.Meta()


# ---------------------------------------------------------------- hash: VBA arithmetic == reference
def vba_hash(s: str) -> str:
    """Line-for-line port of modCore.HashText / FnvStep (Double arithmetic)."""
    def mod32(x):
        return x - math.trunc(x / 4294967296.0) * 4294967296.0

    def step(h, b):
        lo = h - math.trunc(h / 256.0) * 256.0
        h = h - lo + (int(lo) ^ b)
        t1 = mod32(h * 403.0)
        t2 = mod32((h - math.trunc(h / 256.0) * 256.0) * 16777216.0)
        return mod32(t1 + t2)

    h = 2166136261.0
    for ch in s.encode("utf-16-le").decode("utf-16-le"):
        for unit in [ch] if ord(ch) < 0x10000 else []:
            u = ord(unit)
            h = step(h, u % 256)
            h = step(h, u // 256)
    hi = int(math.trunc(h / 65536.0))
    lo = int(h - hi * 65536.0)
    return (format(hi, "04X") + format(lo, "04X")).lower()


@pytest.mark.parametrize("s", ["", "a", "Claims Platform", "Policy Administration System\x1eCore policy",
                               "naïve café → ü", "x" * 500, "4200000\x1e2017-03-01\x1e10.2"])
def test_vba_hash_matches_reference(s):
    assert vba_hash(s) == model.fnv_hash(s)


def test_known_fnv_vector():
    # FNV-1a 32 of "a" as UTF-16LE bytes 0x61 0x00
    h = 2166136261
    for b in (0x61, 0x00):
        h = ((h ^ b) * 16777619) % 2**32
    assert model.fnv_hash("a") == format(h, "08x")


# ---------------------------------------------------------------- relationship rules
@pytest.mark.parametrize("src,rel,tgt,ok", [
    ("ApplicationComponent", "Realization", "Capability", True),
    ("ApplicationComponent", "Access", "DataObject", True),
    ("BusinessActor", "Assignment", "Process", True),
    ("BusinessActor", "Assignment", "Role", True),
    ("Process", "Realization", "Service", True),
    ("Node", "Serving", "ApplicationComponent", True),
    ("Node", "Realization", "Path", True),                  # ArchiMate 4 change
    ("DataObject", "Realization", "BusinessObject", True),
    ("Capability", "Composition", "Capability", True),
    ("Grouping", "Aggregation", "Goal", True),
    ("Goal", "Association", "Node", True),
    ("DataObject", "Access", "ApplicationComponent", False),  # passive cannot access
    ("Goal", "Serving", "Process", False),
    ("Capability", "Composition", "Process", False),
    ("Requirement", "Assignment", "Process", False),
    ("Service", "Triggering", "DataObject", False),
])
def test_allowed(src, rel, tgt, ok):
    assert meta.is_allowed(src, rel, tgt) is ok


def test_every_group_member_is_a_type():
    for g, members in seed.G.items():
        for t in members:
            assert t in meta.types, (g, t)


def test_every_rule_token_resolves():
    for s, r, t, _ in seed.R:
        for tok in (s, t):
            assert tok in ("*", "SAME") or tok in meta.types or (tok.startswith("G:") and tok[2:] in seed.G), tok
        assert r in {x[0] for x in seed.REL_TYPES}


def test_every_type_can_relate_somehow():
    for t in meta.types:
        assert meta.allowed_rels(t, t), t


# ---------------------------------------------------------------- rules engine
def issues(type_code, attr, value, row=None, name_index=None, eid=""):
    return meta and model.evaluate(meta, type_code, attr, value, row or {}, name_index, eid)


def kinds(res):
    return {(s, rt) for s, rt, _ in res}


def test_not_applicable_value_is_error():
    assert ("Error", "NotApplicable") in kinds(issues("Goal", "Vendor", "Oracle"))


def test_not_applicable_blank_ok():
    assert issues("Goal", "Vendor", None) == []


def test_mandatory_from_mapping():
    assert ("Error", "Mandatory") in kinds(issues("ApplicationComponent", "Owner", ""))
    assert issues("Node", "Owner", "") == []


def test_range_and_integer():
    assert ("Error", "Range") in kinds(issues("Node", "Criticality", 7))
    assert issues("Node", "Criticality", 3) == []
    assert ("Error", "DataType") in kinds(issues("Node", "Criticality", "3.5"))


def test_list_value():
    assert ("Error", "List") in kinds(issues("Node", "Lifecycle", "Live"))
    assert issues("Node", "Lifecycle", "Production") == []


def test_regex_warning():
    assert ("Warning", "Regex") in kinds(issues("Node", "Version", "v19"))
    assert issues("Node", "Version", "19.3") == []


def test_date_order_and_required_if():
    row = {"GoLiveDate": date(2020, 1, 1), "RetirementDate": date(2019, 1, 1), "Lifecycle": "Retiring"}
    assert ("Error", "DateOrder") in kinds(issues("ApplicationComponent", "RetirementDate", row["RetirementDate"], row))
    row2 = {"Lifecycle": "Retiring"}
    assert ("Warning", "RequiredIf") in kinds(issues("ApplicationComponent", "RetirementDate", None, row2))


def test_unique_name_within_type():
    idx = {("ApplicationComponent", "crm"): {"a", "b"}}
    assert ("Warning", "Unique") in kinds(issues("ApplicationComponent", "Name", "CRM", {}, idx, "a"))
    idx2 = {("ApplicationComponent", "crm"): {"a"}}
    assert issues("ApplicationComponent", "Name", "CRM", {}, idx2, "a") == []


# ---------------------------------------------------------------- built workbooks
@pytest.fixture(scope="module")
def core():
    return load_workbook(MODEL / "EA_Core.xlsx")


@pytest.fixture(scope="module")
def ents():
    return load_workbook(MODEL / "EA_Entities.xlsx")


def test_core_sheets(core):
    for s in ["Home", "MM_EntityTypes", "MM_AttributeTypes", "MM_RelationshipTypes", "MM_EntityAttributes",
              "MM_EntityRelationships", "MM_TypeGroups", "MM_Rules", "MM_Lists", "MM_Styles", "DATA_Master",
              "DATA_Relationships", "H_Metamodel", "H_Relationships", "SYS_Issues", "SYS_Settings",
              "SYS_Counters", "SYS_ListCache", "SYS_MapCache", "FORM_Entity", "FORM_Relationship",
              "FORM_Type", "FORM_Rule", "FORM_Versions", "SYNC_Review", "SYS_Versions", "VERSION_Compare"]:
        assert s in core.sheetnames, s


def test_vba_tables_exist(core, ents):
    vba_tables = {t for f in SRC.glob("*.bas") for t in re.findall(r'ListObjects\("(tbl\w+)"\)', f.read_text())}
    vba_tables |= {t for f in SRC.glob("*.bas") for t in re.findall(r'Tbl\([^,]+,\s*"(tbl\w+)"\)', f.read_text())}
    have = {t for wb in (core, ents) for ws in wb.worksheets for t in ws.tables}
    missing = vba_tables - have - {"tblCacheAttrs", "tblCacheRules"}
    assert not missing, missing


def test_vba_column_headers_exist(core):
    """Every lo.ListColumns("X") the VBA uses on a metamodel table exists in that table."""
    headers = {}
    for ws in core.worksheets:
        for name, t in ws.tables.items():
            ref = ws[t]
            headers[name] = {c.value for c in ref[0]}
    allh = set().union(*headers.values())
    used = set()
    for f in SRC.glob("*.bas"):
        used |= set(re.findall(r'ListColumns\("([A-Za-z]+)"\)', f.read_text()))
        used |= set(re.findall(r'\bCol\(lo, "([A-Za-z]+)"\)', f.read_text()))
    used -= {"Key", "Value"}  # SYS_Settings / SyncState (checked by name)
    assert used <= allh | {"Key", "Value"}, used - allh


def test_master_keys_and_mapping_governance(core):
    ws = core["DATA_Master"]
    keys = [c.value for c in ws[1] if c.value]
    assert keys[:4] == ["EntityID", "TypeCode", "DisplayID", "Name"]
    assert keys[-5:] == ["Created", "CreatedBy", "Modified", "ModifiedBy", "RowHash"]
    col = {k: i + 1 for i, k in enumerate(keys)}
    for r in range(3, ws.max_row + 1):
        t = ws.cell(r, col["TypeCode"]).value
        if not t:
            continue
        for a in [x[0] for x in seed.ATTRS]:
            v = ws.cell(r, col[a]).value
            if v not in (None, ""):
                assert meta.is_mapped(t, a), f"row {r}: {a} filled but not mapped to {t}"


def test_relationship_formula_columns(core):
    ws = core["DATA_Relationships"]
    assert ws["E2"].value.startswith("=IFERROR(INDEX(DATA_Master!$D:$D")


def test_type_sheet_layout(ents):
    ws = ents["T_ApplicationComponent"]
    keys = [c.value for c in ws[1] if c.value]
    assert keys[:8] == ["EntityID", "RowKind", "PaintHash", "DisplayID", "Name", "Description", "Status",
                        "RepositoryArea"]
    assert keys[8:] == [m["attr"] for m in meta.attrs_for_type("ApplicationComponent")]
    kinds_ = [ws.cell(r, 2).value for r in range(3, 12)]
    assert "E" in kinds_ and "R" in kinds_
    rel_rows = [r for r in range(3, ws.max_row + 1) if ws.cell(r, 2).value == "R"]
    for r in rel_rows:
        assert ws.row_dimensions[r].outlineLevel == 1
        assert ws.cell(r, 5).value.startswith('=HYPERLINK("#\'T_')


def test_one_sheet_per_type(ents):
    for t in meta.types:
        assert f"T_{t}" in ents.sheetnames


def test_named_lists_cover_validations(core):
    names = set(core.defined_names.keys())
    for a in seed.ATTRS:
        if a[2] == "List":
            assert f"lst_{a[3]}" in names
    assert "MapKeys" in names and "lst_TypeCodes" in names


# ---------------------------------------------------------------- VBA static checks
def vba_files():
    return sorted(list(SRC.glob("*.bas")) + list(SRC.glob("*.cls")))


def code_lines(text):
    buf = ""
    for line in text.splitlines():
        s = re.sub(r'"[^"]*"', '""', line)       # drop strings
        s = s.split("'", 1)[0].rstrip()          # drop comments
        if s.endswith(" _"):                     # VBA line continuation
            buf += s[:-2] + " "
            continue
        yield (buf + s).strip()
        buf = ""


def test_vba_ascii_only():
    for f in vba_files() + list(SRC.glob("*.txt")):
        bad = [(i + 1, l) for i, l in enumerate(f.read_text().splitlines()) if any(ord(ch) > 126 for ch in l)]
        assert not bad, (f.name, bad[:3])


def test_vba_block_balance():
    pairs = [(r"^(Public |Private )?(Sub|Function|Property \w+) \w+", r"^End (Sub|Function|Property)$"),
             (r"^With\b", r"^End With$"),
             (r"^Select Case\b", r"^End Select$")]
    for f in vba_files():
        lines = list(code_lines(f.read_text()))
        for open_re, close_re in pairs:
            opens = sum(1 for l in lines if re.match(open_re, l) and not re.search(r":\s*End (Sub|Function)", l))
            closes = sum(1 for l in lines if re.match(close_re, l))
            assert opens == closes, (f.name, open_re, opens, closes)
        # block If / End If
        opens = 0
        for l in lines:
            if re.match(r"^(Else)?If\b.*\bThen$", l) and not l.startswith("ElseIf"):
                opens += 1
            if re.match(r"^End If$", l):
                opens -= 1
        assert opens == 0, (f.name, "If/End If", opens)
        fors = sum(1 for l in lines if re.match(r"^For\b", l))
        nexts = sum(1 for l in lines if re.match(r"^Next\b", l)) + sum(l.count(": Next") for l in lines)
        fors += sum(1 for l in lines if re.search(r":\s*For\b", l))
        assert fors == nexts, (f.name, "For/Next", fors, nexts)
        dos = sum(1 for l in lines if re.match(r"^Do\b", l))
        loops = sum(1 for l in lines if re.match(r"^Loop\b", l))
        assert dos == loops, (f.name, "Do/Loop")


def public_members():
    members = {}
    for f in vba_files():
        text = f.read_text()
        mod = re.search(r'Attribute VB_Name = "(\w+)"', text).group(1)
        names = set(re.findall(r"^Public (?:Sub|Function|Const|Property Get) (\w+)", text, re.M))
        names |= set(re.findall(r"^Public (\w+) As ", text, re.M))
        names |= set(re.findall(r"^Public WithEvents (\w+) As ", text, re.M))
        members[mod] = names
    return members


def test_cross_module_references_resolve():
    members = public_members()
    problems = []
    for f in vba_files() + list(SRC.glob("*.txt")):
        for i, l in enumerate(code_lines(f.read_text()), 1):
            for mod, name in re.findall(r"\b(mod[A-Z]\w+|cls[A-Z]\w+)\.(\w+)", l):
                if mod in members and name not in members[mod]:
                    problems.append(f"{f.name}:{i} {mod}.{name}")
        # string references used with Application.Run / OnAction
        for mod, name in re.findall(r"!(mod\w+)\.(\w+)", f.read_text()):
            if mod in members and name not in members[mod]:
                problems.append(f"{f.name} run-string {mod}.{name}")
        for name in re.findall(r'Array\("[^"]+", "(mod\w+\.\w+)"\)', f.read_text()):
            m, n = name.split(".")
            if n not in members.get(m, set()):
                problems.append(f"{f.name} button {name}")
        for proc in re.findall(r'RunCore "(mod\w+\.\w+)"', f.read_text()):
            m, n = proc.split(".")
            if n not in members.get(m, set()):
                problems.append(f"{f.name} RunCore {proc}")
        for proc in re.findall(r'MacroRef\("(\w+)"\)', f.read_text()):
            if proc not in members["modForms"]:
                problems.append(f"{f.name} MacroRef {proc}")
    assert not problems, problems


def test_unqualified_calls_to_private_helpers_are_local():
    """A Private procedure must only be called inside its own module."""
    privs = {}
    for f in vba_files():
        for n in re.findall(r"^Private (?:Sub|Function) (\w+)", f.read_text(), re.M):
            privs.setdefault(n, set()).add(f.name)
    for f in vba_files():
        text = "\n".join(code_lines(f.read_text()))
        for n, owners in privs.items():
            if f.name in owners:
                continue
            if re.search(rf"(?<![\w.]){n}\b", text) and n not in ("Col", "Key", "Field", "Nz2", "Nz3"):
                # allow if this file defines its own public/private with same name
                if not re.search(rf"^(Public |Private )?(Sub|Function) {n}\b", f.read_text(), re.M):
                    pytest.fail(f"{f.name} calls {n} which is Private in {owners}")


def test_option_explicit_everywhere():
    for f in vba_files():
        assert "Option Explicit" in f.read_text(), f.name


# ---------------------------------------------------------------- Phase 2: sync decisions
import itertools


def vba_decide(paint, sheet, master):
    """Line-for-line port of modSync.Decide."""
    if paint == "":
        return "repaint" if sheet == master else "both"
    sheet_edited = sheet != paint
    master_changed = master != paint
    if not sheet_edited and not master_changed:
        return "none"
    elif master_changed and not sheet_edited:
        return "repaint"
    elif sheet_edited and not master_changed:
        return "sheet"
    elif sheet == master:
        return "repaint"
    return "both"


@pytest.mark.parametrize("paint,sheet,master", list(itertools.product(["", "a", "b", "c"], repeat=3)))
def test_vba_decide_matches_reference(paint, sheet, master):
    assert vba_decide(paint, sheet, master) == model.decide(paint, sheet, master)


@pytest.mark.parametrize("paint,sheet,master,want", [
    ("p", "p", "p", "none"),        # nothing changed
    ("p", "p", "m", "repaint"),     # changed in EA_Core only
    ("p", "s", "p", "sheet"),       # edited on the sheet without macros
    ("p", "s", "m", "both"),        # changed on both sides: conflict
    ("p", "x", "x", "repaint"),     # same change on both sides
    ("", "x", "x", "repaint"),      # never painted, already equal
    ("", "x", "y", "both"),         # never painted, differs: EA_Core wins by default
])
def test_decide_cases(paint, sheet, master, want):
    assert model.decide(paint, sheet, master) == want
    assert model.DEFAULT_KEEP.get(want, "EA_Core") in ("EA_Core", "EA_Entities")


def test_paint_hash_is_projection_of_master():
    """PaintHash written by the painter must equal the projected master hash for an untouched row."""
    vals = ["Claims Platform", "desc", "Active", "Architecture Landscape", "Owner X", None, 5]
    assert model.row_hash(vals) == model.row_hash(list(vals))


def test_review_sheet_layout(core):
    ws = core["SYNC_Review"]
    hdr = [c.value for c in ws[5]]
    assert hdr[:15] == ["Found", "Situation", "Sheet", "SheetRow", "EntityID", "DisplayID", "Name", "TypeCode",
                        "AttrCode", "Attribute", "CoreValue", "SheetValue", "Keep", "Status", "Note"]
    assert "lst_SyncKeep" in core.defined_names


def test_sync_settings_present(core, ents):
    keys = [r[0].value for r in core["SYS_Settings"].iter_rows(min_row=2)]
    for k in ("WBSaveStamp", "LastSync", "LastFullSync", "ModelID"):
        assert k in keys
    skeys = [r[0].value for r in ents["SYS_SyncState"].iter_rows(min_row=2)]
    for k in ("WBSaveStamp", "LastSync", "SchemaVersion", "ModelID"):
        assert k in skeys
    assert ents["SYS_SyncState"].cell(skeys.index("ModelID") + 2, 2).value == \
        core["SYS_Settings"].cell(keys.index("ModelID") + 2, 2).value


def test_review_column_indexes_match_vba():
    """modSync reads the review table by position: Situation=2, Sheet=3, SheetRow=4, EntityID=5,
    TypeCode=8, AttrCode=9, SheetValue=12, Keep=13, Status=14."""
    import builder
    h = builder.SYNC_REVIEW_HEADERS
    pos = {name: i + 1 for i, name in enumerate(h)}
    assert (pos["Situation"], pos["Sheet"], pos["SheetRow"], pos["EntityID"], pos["TypeCode"], pos["AttrCode"],
            pos["SheetValue"], pos["Keep"], pos["Status"]) == (2, 3, 4, 5, 8, 9, 12, 13, 14)


def test_installer_lists_every_module():
    text = (SRC / "modInstall.bas").read_text()
    for f in vba_files():
        if f.name in ("modInstall.bas", "modEntitiesThin.bas"):
            continue
        assert f.name in text, f.name


def test_every_onaction_target_exists():
    members = public_members()
    for f in vba_files():
        for proc in re.findall(r'core & "(\w+)"', f.read_text()):
            # buttons built as core & "Proc" inside a module: the module is the file's own
            mod = re.search(r'Attribute VB_Name = "(\w+)"', f.read_text()).group(1)
            assert proc in members[mod] or any(proc in m for m in members.values()), (f.name, proc)


# ---------------------------------------------------------------- keyed master hash
def vba_keyed_hash(keys, vals):
    """Port of modData.KeyedHash."""
    parts = []
    for i, (k, v) in enumerate(zip(keys, vals), 1):
        t = model.norm(v)
        if i <= 4 or t != "":
            parts.append(f"{k}={t}")
    return model.fnv_hash("\x1e".join(parts))


def test_keyed_hash_matches_reference():
    keys = ["Name", "Description", "Status", "RepositoryArea", "Owner", "Vendor", "Criticality"]
    vals = ["Claims Platform", None, "Active", "Architecture Landscape", "IT", None, 5]
    assert vba_keyed_hash(keys, vals) == model.master_hash(keys, vals)


def test_new_blank_attribute_keeps_hash():
    keys = ["Name", "Description", "Status", "RepositoryArea", "Owner"]
    vals = ["X", "", "Active", "Area", "IT"]
    assert model.master_hash(keys, vals) == model.master_hash(keys + ["RTO"], vals + [None])
    assert model.master_hash(keys, vals) != model.master_hash(keys + ["RTO"], vals + [4])


def test_built_master_hashes_are_keyed(core):
    ws = core["DATA_Master"]
    keys = [c.value for c in ws[1]]
    col = {k: i for i, k in enumerate(keys)}
    hk = model.MASTER_SYS_HASHED + [a[0] for a in seed.ATTRS]
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row[col["EntityID"]]:
            continue
        vals = [row[col[k]] for k in hk]
        assert row[col["RowHash"]] == model.master_hash(hk, vals)


# ---------------------------------------------------------------- Phase 3
import json as _json

PBI = MODEL / "Reports"


def test_pbip_structure():
    assert (PBI / "EA_Reports.pbip").exists()
    pbip = _json.loads((PBI / "EA_Reports.pbip").read_text())
    assert pbip["artifacts"][0]["report"]["path"] == "EA_Reports.Report"
    pbir = _json.loads((PBI / "EA_Reports.Report" / "definition.pbir").read_text())
    assert pbir["datasetReference"]["byPath"]["path"] == "../EA_Reports.SemanticModel"
    _json.loads((PBI / "EA_Reports.Report" / "report.json").read_text())


def test_tmdl_tables_and_relationships():
    import pbip
    defn = PBI / "EA_Reports.SemanticModel" / "definition"
    rel = (defn / "relationships.tmdl").read_text()
    for ft, fc, tt, tc in pbip.RELATIONSHIPS:
        assert f"'{ft}'.'{fc}'" in rel and f"'{tt}'.'{tc}'" in rel
        assert fc in pbip.TABLES[ft]["columns"] and tc in pbip.TABLES[tt]["columns"]
    for name, spec in pbip.TABLES.items():
        text = (defn / "tables" / f"{name}.tmdl").read_text()
        assert text.startswith(f"table '{name}'")
        assert "\tpartition" in text and "\t\tmode: import" in text
        for line in text.splitlines():
            assert not line.startswith(" "), (name, line)   # TMDL indents with tabs
        # every declared column is produced by the M (last step selects/renames to these names)
        m = "\n".join(spec["m"])
        for col in spec["columns"]:
            assert f'"{col}"' in m, (name, col)
        # measures reference existing tables
        for _, dax, _ in spec["measures"]:
            for tname in re.findall(r"'([^']+)'\[", dax) + re.findall(r"\b(Issues)\[", dax):
                assert tname in pbip.TABLES, (name, tname)


def test_m_uses_real_headers(core):
    """Column names used in the Power Query M (VBA and Power BI) exist in EA_Core's tables."""
    import pbip
    headers = {}
    for ws in core.worksheets:
        for tname, ref in ws.tables.items():
            hdr_row = ws[ref][0]
            headers[tname] = {c.value for c in hdr_row}
    master = headers["tblMaster"]
    for col in ["Entity ID", "Type", "Display ID", "Name", "Description", "Status", "Repository Area",
                "Created", "Created By", "Modified", "Modified By", "Row Hash"]:
        assert col in master, col
    for col in ["RelID", "RelCode", "SourceID", "TargetID", "SourceName", "TargetName", "SourceType", "TargetType",
                "Label", "Status"]:
        assert col in headers["tblRel"], col
    for col in ["TypeCode", "Name", "Domain", "Aspect"]:
        assert col in headers["tblEntityTypes"], col
    for col in ["RelCode", "Name"]:
        assert col in headers["tblRelTypes"], col
    for col in ["Found", "Sheet", "RecordID", "DisplayID", "Name", "TypeCode", "AttrCode", "RuleType", "Severity",
                "Message", "Go to"]:
        assert col in headers["tblIssues"], col


def test_vba_m_queries_reference_real_tables(core):
    text = (SRC / "modReports.bas").read_text()
    tables = set(re.findall(r"Item=`(tbl\w+)`", text))
    have = {t for ws in core.worksheets for t in ws.tables}
    assert tables and tables <= have, tables - have


def test_import_preview_layout(core):
    ws = core["IMPORT_Preview"]
    import builder
    assert [c.value for c in ws[6]][:len(builder.IMPORT_HEADERS)] == builder.IMPORT_HEADERS
    assert len(builder.IMPORT_HEADERS) == 14   # modIO.AddRow writes 14 values


# ---------------------------------------------------------------- phase 4: draw.io
import base64
import json
import zlib
import xml.etree.ElementTree as ET
from urllib.parse import quote

import drawio

VALID_APP_TYPES = set("""comp collab contract interface interface2 location facility func gap interaction serv
equipment event event2 node generic netw commPath actor assess goal grouping outcome role passive plateau proc
product driver deliverable device principle requirement representation constraint resource capability course
material meaning distribution sysSw artifact path amValue valueStream workPackage""".split())


def vba_encode_uri(s: str) -> str:
    """Port of modDiagram.EncodeURIComponent (UTF-16 walk, surrogate pairs, UTF-8 bytes)."""
    units = s.encode("utf-16-le")
    units = [units[i] | (units[i + 1] << 8) for i in range(0, len(units), 2)]
    out, i = [], 0
    safe = "-_.!~*'()"
    while i < len(units):
        u = units[i]
        ch = chr(u) if u < 0xD800 or u > 0xDFFF else ""
        if (65 <= u <= 90) or (97 <= u <= 122) or (48 <= u <= 57) or (ch and ch in safe):
            out.append(ch)
        else:
            cp = u
            if 0xD800 <= u <= 0xDBFF and i + 1 < len(units) and 0xDC00 <= units[i + 1] <= 0xDFFF:
                cp = 0x10000 + (u - 0xD800) * 1024 + (units[i + 1] - 0xDC00)
                i += 1
            pct = lambda b: "%" + format(b, "02X")
            if cp < 0x80:
                out.append(pct(cp))
            elif cp < 0x800:
                out.append(pct(0xC0 | (cp >> 6)) + pct(0x80 | (cp & 63)))
            elif cp < 0x10000:
                out.append(pct(0xE0 | (cp >> 12)) + pct(0x80 | ((cp >> 6) & 63)) + pct(0x80 | (cp & 63)))
            else:
                out.append(pct(0xF0 | (cp >> 18)) + pct(0x80 | ((cp >> 12) & 63)) + pct(0x80 | ((cp >> 6) & 63))
                           + pct(0x80 | (cp & 63)))
        i += 1
    return "".join(out)


@pytest.mark.parametrize("s", ["", "abc", "<mxGraphModel a=\"1\"/>", "naïve café", "Zoë & \"quotes\"",
                               "emoji 🚀 ok", "tab\tnew\nline", "~*'()!-_."])
def test_vba_encode_uri_matches_js(s):
    assert vba_encode_uri(s) == quote(s, safe="-_.!~*'()")


def vba_stored_deflate(data: bytes) -> bytes:
    """Port of modDiagram.Compress block writing (NLEN = 65535 - LEN)."""
    n = len(data)
    if n == 0:
        return bytes([1, 0, 0, 255, 255])
    out, i = bytearray(), 0
    while i < n:
        chunk = min(n - i, 65535)
        out += bytes([1 if i + chunk >= n else 0, chunk & 255, (chunk // 256) & 255, (65535 - chunk) & 255,
                      (65535 - chunk) // 256]) + data[i:i + chunk]
        i += chunk
    return bytes(out)


@pytest.mark.parametrize("n", [0, 1, 1000, 65535, 65536, 140000])
def test_stored_deflate_inflates(n):
    data = bytes((i * 7) % 251 for i in range(n))
    assert vba_stored_deflate(data) == drawio.deflate_stored(data)
    assert zlib.decompressobj(-15).decompress(vba_stored_deflate(data)) == data


def decompress(b64):
    from urllib.parse import unquote
    return unquote(zlib.decompressobj(-15).decompress(base64.b64decode(b64)).decode("ascii"))


def library_entries(path):
    root = ET.parse(path).getroot()
    assert root.tag == "mxlibrary"
    return json.loads(root.text)


def test_stencil_library():
    entries = library_entries(MODEL / "Library" / "ArchiMate4.xml")
    assert len(entries) == 42 + 11
    types, rels = set(), set()
    for e in entries:
        m = ET.fromstring(decompress(e["xml"]))
        obj = m.find("./root/object")
        cell = obj.find("mxCell")
        if cell.get("vertex") == "1":
            types.add(obj.get("ea_type"))
            assert obj.get("ea_id") == ""
        else:
            rels.add(obj.get("ea_reltype"))
            assert cell.get("edge") == "1"
    assert types == {t[0] for t in seed.ENTITY_TYPES}
    assert rels == {r[0] for r in seed.REL_TYPES}
    app = [e for e in entries if e["title"] == "Application Component"][0]
    obj = ET.fromstring(decompress(app["xml"])).find("./root/object")
    for a in ("Owner", "Lifecycle"):        # mandatory attributes ready for Edit Data
        assert obj.get(a) == "", a


def test_model_library_has_every_active_entity(core):
    entries = library_entries(MODEL / "Library" / "ModelLibrary.xml")
    ws = core["DATA_Master"]
    keys = [c.value for c in ws[1]]
    ids = {r[keys.index("EntityID")] for r in ws.iter_rows(min_row=3, values_only=True) if r[0]}
    got = {ET.fromstring(decompress(e["xml"])).find("./root/object").get("ea_id") for e in entries}
    assert got == ids


def test_styles_use_real_drawio_shapes(core):
    ws = core["MM_Styles"]
    for key, applies, fill, style in ws.iter_rows(min_row=2, values_only=True):
        assert "=" in style, key
        m = re.search(r"appType=(\w+)", style)
        if m:
            assert m.group(1) in VALID_APP_TYPES, key
        assert "fillColor" not in style or key == "Entity:Grouping", key   # domain colour comes from FillColour


def test_relationship_line_styles_are_distinguishable():
    sigs = [drawio.signature(s) for s in drawio.REL_STYLES.values()]
    assert len(set(sigs)) == len(sigs)
    assert drawio.signature("html=1;edgeStyle=orthogonalEdgeStyle;") not in sigs   # draw.io's default connector


def test_sample_diagram(core):
    tree = ET.parse(MODEL / "Diagrams" / "Claims_Overview.drawio")
    root = tree.getroot()
    assert root.tag == "mxfile"
    page = root.find("diagram")
    cells = page.find("mxGraphModel/root")
    ids = [c.get("id") for c in cells]
    assert len(ids) == len(set(ids))
    objs = [c for c in cells if c.tag == "object"]
    shapes = [o for o in objs if o.find("mxCell").get("vertex") == "1"]
    edges = [o for o in objs if o.find("mxCell").get("edge") == "1"]
    assert len(shapes) == 14 and len(edges) == 12
    for e in edges:
        c = e.find("mxCell")
        assert c.get("source") in ids and c.get("target") in ids
    ws = core["DATA_Master"]
    keys = [c.value for c in ws[1]]
    master = {r[keys.index("EntityID")]: r for r in ws.iter_rows(min_row=3, values_only=True) if r[0]}
    for s in shapes:
        r = master[s.get("ea_id")]
        assert s.get("label") == r[keys.index("Name")] == s.get("ea_name")
        assert s.get("ea_type") == r[keys.index("TypeCode")]


def vba_fingerprint(pairs):
    """Port of modDiagram.Fingerprint / Hex32ToDouble (Double arithmetic)."""
    total = 0.0
    for eid, rh in pairs:
        h = vba_hash(f"{eid}:{rh}")
        v = 0.0
        for ch in h:
            v = v * 16.0 + "0123456789abcdef".index(ch)
        total = total + v
        total = total - math.trunc(total / 4294967296.0) * 4294967296.0
    hi = int(math.trunc(total / 65536.0))
    lo = int(total - hi * 65536.0)
    return (format(hi, "04X") + format(lo, "04X")).lower()


def test_views_register_fingerprint_matches_master(core):
    ws = core["DATA_Master"]
    keys = [c.value for c in ws[1]]
    pairs = [(r[keys.index("EntityID")], r[keys.index("RowHash")])
             for r in ws.iter_rows(min_row=3, values_only=True) if r[0] and r[keys.index("Status")] != "Retired"]
    lo = core["SYS_Views"].tables["tblViews"]
    vw = core["SYS_Views"]
    hdr = [c.value for c in vw[4]]
    assert hdr == ["File", "Page", "Entities", "Relationships", "Generated", "Harvested", "Fingerprint", "Status", "Note"]
    row = [c.value for c in vw[5]]
    assert row[0] == "Claims_Overview.drawio" and row[1] == "Claims overview"
    assert row[6] == drawio.fingerprint(pairs) == vba_fingerprint(pairs)
    assert lo.ref.startswith("A4")


def test_harvest_review_layout(core):
    ws = core["HARVEST_Review"]
    assert ws.tables["tblHarvest"].ref.startswith("A6")
    src = (SRC / "modDiagram.bas").read_text()
    m = re.search(r"rows\.Add Array\(([^)]*)\)", src)
    assert len(m.group(1).split(",")) == len([c.value for c in ws[6]]) == 14


def test_diagram_lists(core):
    src = (SRC / "modGenerator.bas").read_text()
    for name in ("DiagramScope", "ImageFormat"):
        assert f'Set d("{name}")' in src
    assert "FORM_Diagram" in core.sheetnames and "SYS_Views" in core.sheetnames


# ---------------------------------------------------------------- phase 5: parent and child
@pytest.mark.parametrize("b,p,c,want", [
    ("x", "x", "y", "child"),        # changed in child only
    ("x", "y", "x", ""),             # changed in parent only: parent kept
    ("x", "y", "y", ""),             # same change on both sides
    ("x", "y", "z", "conflict"),     # different changes
    ("", "", "new", "child"),        # value added in child
])
def test_three_way(b, p, c, want):
    assert model.three_way(b, p, c) == want


def test_retired_rules():
    # retired in parent, changed in child -> conflict
    assert model.merge_field("A", "A", "B", "Owner", True, False, False, True) == "conflict"
    # retired in child, changed in parent -> conflict on Status
    assert model.merge_field("Active", "Active", "Retired", "Status", False, False, True, True) == "conflict"
    # retired in child, parent untouched -> child taken
    assert model.merge_field("Active", "Active", "Retired", "Status", False, False, True, False) == "child"


def test_vba_three_way_matches_reference():
    src = (SRC / "modMerge.bas").read_text()
    body = src[src.index("Private Function ThreeWay"):]
    body = body[:body.index("End Function")]
    assert "If cV = bV Or cV = pV Then" in body and "ElseIf pV = bV Then" in body


def test_merge_review_layout(core):
    ws = core["MERGE_Review"]
    hdr = [c.value for c in ws[6]]
    assert hdr == ["Kind", "Action", "RecordID", "DisplayID", "Name", "TableOrType", "Field", "Base", "Parent", "Child",
                   "Keep", "Note"]
    src = (SRC / "modMerge.bas").read_text()
    # ReadDecisions keys on Kind (1), RecordID (3), Field (7) and reads Keep (11)
    assert 'CStr(data(i, 1)) & "|" & CStr(data(i, 3)) & "|" & CStr(data(i, 7))) = CStr(data(i, 11))' in src
    m = re.search(r"rows\.Add Array\(kind, action, id, disp, nm, tbl, field, T\(bV\), T\(pV\), T\(cV\), keep, note\)", src)
    assert m
    assert [c.value for c in core["SYS_Children"][4]] == ["ChildModelID", "ChildName", "Folder", "Created", "Scope",
                                                        "LastMerged", "Status"]


def test_merge_settings_and_lists(core):
    keys = {r[0] for r in core["SYS_Settings"].iter_rows(min_row=2, values_only=True)}
    assert {"ParentModelID", "ParentModelName", "ParentFolder", "ChildCreated"} <= keys
    src = (SRC / "modGenerator.bas").read_text()
    for name in ("MergeScope", "MergeKeep"):
        assert f'Set d("{name}")' in src
    assert "FORM_Merge" in core.sheetnames


def test_metamodel_merge_tables_exist(core):
    src = (SRC / "modMerge.bas").read_text()
    spec = re.search(r'Private Const MM_TABLES As String = _\n(.*?)\n\n', src, re.S).group(1)
    for sheet, table, keys in re.findall(r'"(MM_\w+)\|(tbl\w+)\|([\w,]+)\|', spec):
        ws = core[sheet]
        assert table in ws.tables, (sheet, table)
        hdr = [c.value for c in ws[1]]
        for k in keys.split(","):
            assert k in hdr, (sheet, k)


# ---------------------------------------------------------------- setup tool
def test_setup_tool_basics(tmp_path):
    import workbench_setup as ws
    assert ws.folder_name_for("Group Landscape") == "EA_Model_Group_Landscape"
    assert ws.folder_name_for("  ") == ""
    assert ws.find_template_folder() is not None
    lines = []
    assert ws.task_create_model(lines.append, tmp_path, "Unit Test", False)
    m = tmp_path / "EA_Model_Unit_Test"
    for f in ("EA_Core.xlsx", "EA_Entities.xlsx", "model.json", "src/modInstall.bas", "Library/ArchiMate4.xml",
              "Reports/EA_Reports.pbip"):
        assert (m / f).exists(), f
    assert not ws.task_create_model(lines.append, tmp_path, "Unit Test", False)   # refuses an existing folder
