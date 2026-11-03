"""draw.io side of the EA Repository Workbench (reference implementation).

The VBA module modDiagram mirrors these functions:
  - entity_style / REL_STYLES    <-> MM_Styles rows read by modDiagram.EntityStyle / RelStyle
  - encode_uri_component         <-> modDiagram.EncodeURIComponent
  - compress                     <-> modDiagram.Compress  (stored deflate blocks, so no zlib is needed in VBA)
  - stencil / model_library      <-> modDiagram.BuildStencil / BuildModelLibrary
  - generate_diagram (layout)    <-> modDiagram.GenerateDiagram
  - fingerprint                  <-> modDiagram.Fingerprint

draw.io reads a library entry's "xml" as base64(deflateRaw(encodeURIComponent(xml))). Stored
(uncompressed) deflate blocks are valid deflate, so both sides can write them without a compressor.
"""
import base64
import json
import uuid
from urllib.parse import quote
from xml.sax.saxutils import escape, quoteattr

import model
import seed

# ---------------------------------------------------------------- shape styles
BASE = "html=1;outlineConnect=0;whiteSpace=wrap;"
APP_SHAPE = "shape=mxgraph.archimate3.application;"

# TypeCode -> (appType, archiType) for draw.io's ArchiMate 3 box-with-icon shape.
# Square = structure, rounded = behaviour, oct = motivation (ArchiMate notation).
APP_TYPES = {
    "Role": ("role", "square"), "Collaboration": ("collab", "square"), "Path": ("path", "square"),
    "Service": ("serv", "rounded"), "Process": ("proc", "rounded"), "Function": ("func", "rounded"),
    "Event": ("event", "rounded"), "Location": ("location", "square"),
    "BusinessActor": ("actor", "square"), "BusinessInterface": ("interface", "square"),
    "ApplicationComponent": ("comp", "square"), "ApplicationInterface": ("interface", "square"),
    "Node": ("node", "square"), "Device": ("device", "square"), "SystemSoftware": ("sysSw", "square"),
    "TechnologyInterface": ("interface", "square"), "CommunicationNetwork": ("netw", "square"),
    "DistributionNetwork": ("distribution", "square"), "Equipment": ("equipment", "square"),
    "Facility": ("facility", "square"), "Artifact": ("artifact", "square"), "Material": ("material", "square"),
    "Resource": ("resource", "square"), "Capability": ("capability", "rounded"),
    "ValueStream": ("valueStream", "rounded"), "CourseOfAction": ("course", "rounded"),
    "Stakeholder": ("role", "oct"), "Driver": ("driver", "oct"), "Assessment": ("assess", "oct"),
    "Goal": ("goal", "oct"), "Outcome": ("outcome", "oct"), "Principle": ("principle", "oct"),
    "Requirement": ("requirement", "oct"), "Meaning": ("meaning", "oct"), "Value": ("amValue", "oct"),
    "WorkPackage": ("workPackage", "rounded"), "Deliverable": ("deliverable", "square"),
    "Plateau": ("plateau", "square"),
}
SPECIAL = {
    "BusinessObject": BASE + "shape=mxgraph.archimate3.businessObject;overflow=fill;",
    "DataObject": BASE + "shape=mxgraph.archimate3.businessObject;overflow=fill;",
    "Product": BASE + "shape=mxgraph.archimate3.product;",
    "Grouping": ("shape=folder;html=1;whiteSpace=wrap;dashed=1;tabWidth=110;tabHeight=20;tabPosition=left;"
                 "verticalAlign=top;align=left;spacingTop=20;spacingLeft=6;fillColor=none;"),
}
DEFAULT_STYLE = "rounded=1;whiteSpace=wrap;html=1;"   # custom types added with the Add Type wizard


def entity_style(code: str) -> str:
    if code in SPECIAL:
        return SPECIAL[code]
    if code in APP_TYPES:
        a, k = APP_TYPES[code]
        return f"{BASE}{APP_SHAPE}appType={a};archiType={k};"
    return DEFAULT_STYLE


EDGE = "html=1;rounded=0;"
REL_STYLES = {
    "Composition": EDGE + "startArrow=diamondThin;startFill=1;startSize=10;endArrow=none;endFill=0;",
    "Aggregation": EDGE + "startArrow=diamondThin;startFill=0;startSize=10;endArrow=none;endFill=0;",
    "Assignment": EDGE + "startArrow=oval;startFill=1;endArrow=block;endFill=1;",
    "Realization": EDGE + "endArrow=block;endFill=0;dashed=1;",
    "Serving": EDGE + "endArrow=open;endFill=0;",
    "Access": EDGE + "endArrow=open;endFill=0;endSize=4;dashed=1;dashPattern=1 4;",
    "Influence": EDGE + "endArrow=open;endFill=0;dashed=1;dashPattern=6 4;",
    "Association": EDGE + "endArrow=none;",
    "Triggering": EDGE + "endArrow=block;endFill=1;",
    "Flow": EDGE + "endArrow=block;endFill=1;dashed=1;dashPattern=6 4;",
    "Specialization": EDGE + "endArrow=block;endFill=0;",
}
# style keys that identify a relationship when a connector has no ea_reltype (copied style)
SIGNATURE_KEYS = ["startArrow", "startFill", "endArrow", "endFill", "dashed", "dashPattern"]


def parse_style(style: str) -> dict:
    out = {}
    for part in style.split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            out[k.strip()] = v.strip()
    return out


def signature(style: str) -> str:
    """Mirrors modDiagram.Signature: the arrow and dash keys, with draw.io's defaults filled in."""
    d = parse_style(style)
    defaults = {"startArrow": "none", "startFill": "1", "endArrow": "classic", "endFill": "1", "dashed": "0",
                "dashPattern": ""}
    vals = []
    for k in SIGNATURE_KEYS:
        v = d.get(k, defaults[k])
        if k in ("startFill",) and d.get("startArrow", "none") == "none":
            v = "-"
        if k in ("endFill",) and d.get("endArrow", "classic") == "none":
            v = "-"
        if k == "dashPattern" and d.get("dashed", "0") != "1":
            v = ""
        vals.append(f"{k}={v}")
    return ";".join(vals)


# ---------------------------------------------------------------- layout
BAND_ORDER = ["Motivation", "Strategy", "Business", "Common", "Application", "Technology",
              "Implementation & Migration"]
W, H, GAP_X, GAP_Y, COLS = 160, 70, 40, 30, 6
LEFT, TOP = 40, 40
BAND_STYLE = "text;html=1;align=left;verticalAlign=middle;fontStyle=1;fontColor=#595959;ea_band=1;"
MISSING_STYLE = {"strokeColor": "#FF0000", "strokeWidth": "3"}


def band_of(type_code):
    return seed_types()[type_code][1]


_TYPES = None


def seed_types():
    global _TYPES
    if _TYPES is None:
        _TYPES = {t[0]: (t[1], t[2]) for t in seed.ENTITY_TYPES}
    return _TYPES


def type_order():
    return {t[0]: i for i, t in enumerate(seed.ENTITY_TYPES)}


def layout(entities, y0=TOP):
    """Return [(entity, x, y)] and [(band, y)] in band order, then type order, then name."""
    order = type_order()
    placed, bands = [], []
    y = y0
    for band in BAND_ORDER:
        members = [e for e in entities if band_of(e["TypeCode"]) == band]
        if not members:
            continue
        members.sort(key=lambda e: (order[e["TypeCode"]], e["Name"].lower()))
        bands.append((band, y))
        for i, e in enumerate(members):
            col, row = i % COLS, i // COLS
            placed.append((e, LEFT + col * (W + GAP_X), y + 30 + row * (H + GAP_Y)))
        rows = (len(members) + COLS - 1) // COLS
        y = y + 30 + rows * (H + GAP_Y) + 20
    return placed, bands


# ---------------------------------------------------------------- compression
def encode_uri_component(s: str) -> str:
    return quote(s, safe="-_.!~*'()")


def deflate_stored(data: bytes) -> bytes:
    out = bytearray()
    if not data:
        return bytes([1, 0, 0, 0xFF, 0xFF])
    for i in range(0, len(data), 65535):
        chunk = data[i:i + 65535]
        final = 1 if i + 65535 >= len(data) else 0
        n = len(chunk)
        out += bytes([final, n & 0xFF, n >> 8, (~n) & 0xFF, ((~n) >> 8) & 0xFF]) + chunk
    return bytes(out)


def compress(xml: str) -> str:
    return base64.b64encode(deflate_stored(encode_uri_component(xml).encode("ascii"))).decode("ascii")


# ---------------------------------------------------------------- xml helpers
def obj_xml(cell_id, label, attrs, style, x, y, w, h, vertex=True, source=None, target=None):
    a = f'<object id={quoteattr(cell_id)} label={quoteattr(escape(label))}'   # html=1 labels are HTML
    for k, v in attrs.items():
        a += f" {k}={quoteattr(str(v))}"
    a += ">"
    if vertex:
        a += (f'<mxCell style={quoteattr(style)} vertex="1" parent="1">'
              f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/></mxCell>')
    else:
        st = f' source={quoteattr(source)}' if source else ""
        tg = f' target={quoteattr(target)}' if target else ""
        if source and target:
            geo = '<mxGeometry relative="1" as="geometry"/>'
        else:
            geo = (f'<mxGeometry width="{w}" relative="1" as="geometry"><mxPoint x="0" y="{h}" as="sourcePoint"/>'
                   f'<mxPoint x="{w}" y="{h}" as="targetPoint"/></mxGeometry>')
        a += f'<mxCell style={quoteattr(style)} edge="1" parent="1"{st}{tg}>{geo}</mxCell>'
    return a + "</object>"


def graph_model(cells_xml):
    return ('<mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/>' + "".join(cells_xml)
            + "</root></mxGraphModel>")


def library(entries):
    return "<mxlibrary>" + escape(json.dumps(entries, ensure_ascii=False, separators=(",", ":"))) + "</mxlibrary>"


def with_fill(style, fill_hex):
    return style if "fillColor=" in style else style + f"fillColor=#{fill_hex};"


# ---------------------------------------------------------------- libraries
def stencil(model_id, mandatory_by_type):
    """ArchiMate 4 palette: one tagged shape per entity type, then one connector per relationship."""
    entries = []
    for code, name, dom, _, _ in seed.ENTITY_TYPES:
        attrs = {"ea_type": code, "ea_id": "", "ea_model": model_id}
        for a in mandatory_by_type.get(code, []):
            attrs[a] = ""
        xml = graph_model([obj_xml("2", name, attrs, with_fill(entity_style(code), seed.DOMAIN_COLOURS[dom]),
                                   0, 0, W, H)])
        entries.append({"xml": compress(xml), "w": W, "h": H, "aspect": "fixed", "title": name})
    for code, verb, *_ in seed.REL_TYPES:
        xml = graph_model([obj_xml("2", "", {"ea_reltype": code, "ea_rel_id": ""}, REL_STYLES[code], 0, 0,
                                   W, 0, vertex=False)])
        entries.append({"xml": compress(xml), "w": W, "h": 0, "aspect": "fixed", "title": f"{code} ({verb})"})
    return library(entries)


def entity_attrs(e, model_id):
    return {"ea_type": e["TypeCode"], "ea_id": e["EntityID"], "ea_display": e["DisplayID"], "ea_name": e["Name"],
            "ea_model": model_id}


def model_library(entities, model_id):
    order = type_order()
    entries = []
    for e in sorted((e for e in entities if e.get("Status") != "Retired"),
                    key=lambda e: (order[e["TypeCode"]], e["Name"].lower())):
        dom = band_of(e["TypeCode"])
        xml = graph_model([obj_xml("2", e["Name"], entity_attrs(e, model_id),
                                   with_fill(entity_style(e["TypeCode"]), seed.DOMAIN_COLOURS[dom]), 0, 0, W, H)])
        entries.append({"xml": compress(xml), "w": W, "h": H, "aspect": "fixed",
                        "title": f"{seed_types()[e['TypeCode']][0]}: {e['Name']} ({e['DisplayID']})"})
    return library(entries)


# ---------------------------------------------------------------- diagrams
def cell_id(text):
    return "ea" + uuid.uuid5(uuid.NAMESPACE_URL, "ea-workbench/cell/" + text).hex[:16]


def generate_diagram(entities, rels, page_name, model_id):
    """One-page .drawio file, uncompressed, laid out in domain bands."""
    placed, bands = layout([e for e in entities if e.get("Status") != "Retired"])
    cells, ids = [], {}
    for band, y in bands:
        cells.append(f'<mxCell id={quoteattr(cell_id("band/" + page_name + "/" + band))} '
                     f'value={quoteattr(band)} style={quoteattr(BAND_STYLE)} vertex="1" parent="1">'
                     f'<mxGeometry x="{LEFT}" y="{y}" width="300" height="24" as="geometry"/></mxCell>')
    for e, x, y in placed:
        cid = cell_id("entity/" + e["EntityID"])
        ids[e["EntityID"]] = cid
        dom = band_of(e["TypeCode"])
        cells.append(obj_xml(cid, e["Name"], entity_attrs(e, model_id),
                             with_fill(entity_style(e["TypeCode"]), seed.DOMAIN_COLOURS[dom]), x, y, W, H))
    for r in rels:
        if r.get("Status") == "Retired" or r["SourceID"] not in ids or r["TargetID"] not in ids:
            continue
        cells.append(obj_xml(cell_id("rel/" + r["RelID"]), r.get("Label", "") or "",
                             {"ea_rel_id": r["RelID"], "ea_reltype": r["RelCode"]}, REL_STYLES[r["RelCode"]],
                             0, 0, 0, 0, vertex=False, source=ids[r["SourceID"]], target=ids[r["TargetID"]]))
    page_id = cell_id("page/" + page_name)
    return ('<?xml version="1.0" encoding="UTF-8"?>\n<mxfile host="EA Workbench">'
            f'<diagram id={quoteattr(page_id)} name={quoteattr(page_name)}>'
            '<mxGraphModel dx="1400" dy="900" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" '
            'arrows="1" fold="1" page="1" pageScale="1" pageWidth="1654" pageHeight="1169" math="0" shadow="0">'
            + graph_model(cells)[len("<mxGraphModel>"):] + "</diagram></mxfile>\n")


def fingerprint(ids_with_hash):
    """Order-independent: sum of FNV(id:rowhash) mod 2^32 over the entities on a page."""
    total = 0
    for eid, rh in ids_with_hash:
        total = (total + int(model.fnv_hash(f"{eid}:{rh}"), 16)) % model.MOD32
    return format(total, "08x")
