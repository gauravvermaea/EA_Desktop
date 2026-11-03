"""Reference implementation of the Workbench logic.

The VBA engine (src/*.bas) mirrors these functions one for one, so they are
the specification the tests run against:
  - fnv_hash          <-> modCore.HashText
  - token_matches     <-> modMeta.TokenMatches
  - is_allowed        <-> modMeta.IsRelAllowed
  - attrs_for_type    <-> modMeta.AttrsForType
  - evaluate          <-> modRules.EvaluateValue
"""
import re
from datetime import date, datetime

import seed

FNV_OFFSET = 2166136261
FNV_PRIME = 16777619
MOD32 = 2 ** 32
SEP = "\x1e"


def fnv_hash(text: str) -> str:
    """FNV-1a 32-bit over UTF-16 code units, low byte then high byte."""
    h = FNV_OFFSET
    data = text.encode("utf-16-le")
    for i in range(0, len(data), 2):
        for b in (data[i], data[i + 1]):
            h ^= b
            h = (h * FNV_PRIME) % MOD32
    return format(h, "08x")


def norm(v) -> str:
    """Canonical text form of a cell value, shared with VBA CellText()."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def row_hash(values) -> str:
    """Positional hash (type-sheet PaintHash): every column counts, blank or not."""
    return fnv_hash(SEP.join(norm(v) for v in values))


MASTER_SYS_HASHED = ["Name", "Description", "Status", "RepositoryArea"]


def master_hash(keys, values) -> str:
    """Keyed hash of a DATA_Master row (mirrors modData.KeyedHash).

    keys/values: the four system fields, then attributes in master column order.
    System fields always count; attributes only when not blank, so adding an
    attribute column does not change existing hashes.
    """
    parts = []
    for i, (k, v) in enumerate(zip(keys, values)):
        s = norm(v)
        if i < 4 or s != "":
            parts.append(f"{k}={s}")
    return fnv_hash(SEP.join(parts))


# ---------------------------------------------------------------- metamodel queries
class Meta:
    def __init__(self):
        self.types = {t[0]: dict(code=t[0], name=t[1], domain=t[2], aspect=t[3], prefix=t[4])
                      for t in seed.ENTITY_TYPES}
        self.attrs = {a[0]: dict(code=a[0], name=a[1], datatype=a[2], list=a[3], unit=a[4], desc=a[5])
                      for a in seed.ATTRS}
        self.sys_attrs = {a[0]: dict(code=a[0], name=a[1], datatype=a[2], list=a[3]) for a in seed.SYSTEM_ATTRS}
        self.groups = {k: set(v) for k, v in seed.G.items()}
        self.allowed = list(seed.R)
        self.map = {t: [dict(attr=a, mandatory=m, section=s, order=i + 1)
                        for i, (a, m, s) in enumerate(seed.MAP[t])] for t in seed.MAP}
        self.rules = [dict(attr=r[0], type=r[1], ruletype=r[2], p1=r[3], p2=r[4], severity=r[5], msg=r[6])
                      for r in seed.RULES]
        self.lists = seed.LISTS

    def token_matches(self, token: str, type_code: str, source_type: str = "") -> bool:
        token = token.strip()
        if token == "*":
            return True
        if token == "SAME":
            return type_code == source_type
        if token.startswith("G:"):
            return type_code in self.groups.get(token[2:], set())
        return token == type_code

    def is_allowed(self, src: str, rel: str, tgt: str) -> bool:
        for s, r, t, _ in self.allowed:
            if r == rel and self.token_matches(s, src) and self.token_matches(t, tgt, src):
                return True
        return False

    def allowed_rels(self, src: str, tgt: str):
        return sorted({r for s, r, t, _ in self.allowed
                       if self.token_matches(s, src) and self.token_matches(t, tgt, src)})

    def attrs_for_type(self, type_code: str):
        return sorted(self.map.get(type_code, []), key=lambda m: m["order"])

    def is_mapped(self, type_code: str, attr: str) -> bool:
        return any(m["attr"] == attr for m in self.map.get(type_code, []))

    def mandatory(self, type_code: str, attr: str) -> bool:
        return any(m["attr"] == attr and m["mandatory"] == "Y" for m in self.map.get(type_code, []))


# ---------------------------------------------------------------- rules engine
def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def evaluate(meta: Meta, type_code: str, attr: str, value, row: dict, name_index=None, entity_id=""):
    """Return list of (severity, ruletype, message) for one value.

    row: attr -> value for the whole entity row (for DateOrder, RequiredIf).
    name_index: {(type, lower name): set(entity ids)} for Unique.
    """
    issues = []
    blank = value is None or str(value).strip() == ""
    adef = meta.attrs.get(attr) or meta.sys_attrs.get(attr)
    if adef is None:
        return issues

    if attr in meta.attrs and not meta.is_mapped(type_code, attr):
        if not blank:
            issues.append(("Error", "NotApplicable", f"{adef['name']} does not apply to {type_code}"))
        return issues

    if blank and meta.mandatory(type_code, attr):
        issues.append(("Error", "Mandatory", f"{adef['name']} is required for {type_code}"))

    # built-in data type checks
    if not blank:
        dt = adef["datatype"]
        if dt == "Integer":
            n = _num(value)
            if n is None or not float(n).is_integer():
                issues.append(("Error", "DataType", f"{adef['name']} must be a whole number"))
        elif dt in ("Decimal", "Currency", "Percent"):
            if _num(value) is None:
                issues.append(("Error", "DataType", f"{adef['name']} must be a number"))
        elif dt == "Date":
            if _date(value) is None:
                issues.append(("Error", "DataType", f"{adef['name']} must be a date"))
        elif dt == "Boolean":
            if str(value).upper() not in ("Y", "N", "TRUE", "FALSE"):
                issues.append(("Error", "DataType", f"{adef['name']} must be Y or N"))
        elif dt == "List" and adef.get("list"):
            if str(value) not in meta.lists.get(adef["list"], []):
                issues.append(("Error", "List", f"{adef['name']} must be one of the {adef['list']} values"))

    for r in meta.rules:
        if r["attr"] != attr or (r["type"] and r["type"] != type_code):
            continue
        rt, p1, p2 = r["ruletype"], r["p1"], r["p2"]
        fail = False
        if rt == "Mandatory":
            fail = blank
        elif rt == "RequiredIf":
            fail = blank and norm(row.get(p1)) == p2
        elif blank:
            continue
        elif rt == "List":
            fail = str(value) not in meta.lists.get(p1, [])
        elif rt in ("MinValue", "MaxValue", "Range"):
            n = _num(value)
            lo = _num(p1) if rt in ("MinValue", "Range") else None
            hi = _num(p1) if rt == "MaxValue" else (_num(p2) if rt == "Range" else None)
            fail = n is None or (lo is not None and n < lo) or (hi is not None and n > hi)
        elif rt == "MinLength":
            fail = len(str(value)) < int(p1)
        elif rt == "MaxLength":
            fail = len(str(value)) > int(p1)
        elif rt == "Regex":
            fail = re.search(p1, str(value)) is None
        elif rt == "DateRange":
            d = _date(value)
            fail = d is None or d < _date(p1) or d > _date(p2)
        elif rt == "DateOrder":
            d, other = _date(value), _date(row.get(p1))
            fail = d is not None and other is not None and d <= other
        elif rt == "Unique" and name_index is not None:
            ids = name_index.get((type_code if p1 == "Type" else "*", str(value).strip().lower()), set())
            fail = len(ids - {entity_id}) > 0
        if fail:
            issues.append((r["severity"], rt, r["msg"]))
    return issues


# ---------------------------------------------------------------- sync decision (mirrors modSync.Decide)
def decide(paint_hash: str, sheet_hash: str, master_hash: str) -> str:
    """Row decision for a full sync.

    paint_hash  - hash of the row as painted on the type sheet ("" = never painted)
    sheet_hash  - hash of the row as it is on the sheet now
    master_hash - hash of the master values projected onto the sheet's columns
    Returns: "none" | "repaint" | "sheet" (edited on the sheet only) | "both" (conflict)
    """
    if paint_hash == "":
        return "repaint" if sheet_hash == master_hash else "both"
    sheet_edited = sheet_hash != paint_hash
    master_changed = master_hash != paint_hash
    if not sheet_edited and not master_changed:
        return "none"
    if master_changed and not sheet_edited:
        return "repaint"
    if sheet_edited and not master_changed:
        return "sheet"
    return "repaint" if sheet_hash == master_hash else "both"


DEFAULT_KEEP = {"sheet": "EA_Entities", "both": "EA_Core", "new": "EA_Entities", "unknown": "EA_Core"}


# ---------------------------------------------------------------- merge (mirrors modMerge.ThreeWay)
def three_way(base: str, parent: str, child: str) -> str:
    """'' = nothing to do, 'child' = take the child's value, 'conflict' = both changed differently."""
    if child == base or child == parent:
        return ""
    if parent == base:
        return "child"
    return "conflict"


def merge_field(base, parent, child, field, parent_retired, base_retired, child_retired, parent_changed):
    """Field outcome including the retired rules of spec section 12 (mirrors modMerge.MergeEntities)."""
    out = three_way(base, parent, child)
    if out == "child" and parent_retired and not base_retired:
        out = "conflict"
    if out == "child" and field == "Status" and child_retired and not base_retired and parent_changed:
        out = "conflict"
    return out
