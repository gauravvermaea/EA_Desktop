"""Generate the Power BI project (PBIP, TMDL semantic model) for a model folder.

Usage: python3 pbip.py <model_folder> [default_model_folder_path]
Writes <model_folder>/Reports/EA_Reports.pbip, EA_Reports.SemanticModel/, EA_Reports.Report/.
The model reads EA_Core.xlsm from the folder named by the ModelFolder parameter.
"""
import json
import sys
import uuid
from pathlib import Path

T = "\t"
CORE_FILE = "EA_Core.xlsm"


def tag(name: str) -> str:
    return str(uuid.uuid5(uuid.NAMESPACE_URL, "ea-workbench/pbi/" + name))


def q(name: str) -> str:
    """Quote every TMDL object name (also keeps names such as Relationship clear of keywords)."""
    return "'" + name.replace("'", "''") + "'"


def m_block(lines):
    """M expression lines indented for a partition source (4 tabs)."""
    return "\n".join(T * 4 + ln for ln in lines)


SOURCE = [
    "let",
    f'    Source = Excel.Workbook(File.Contents(ModelFolder & "{CORE_FILE}"), null, true),',
]

TABLES = {
    "Entity": {
        "m": SOURCE + [
            '    Master = Source{[Item="tblMaster",Kind="Table"]}[Data],',
            '    Rows = Table.SelectRows(Master, each [Entity ID] <> null and [Entity ID] <> ""),',
            '    Picked = Table.SelectColumns(Rows, {"Entity ID", "Display ID", "Name", "Description", "Status", '
            '"Repository Area", "Type", "Modified"}),',
            '    Renamed = Table.RenameColumns(Picked, {{"Type", "TypeCode"}}),',
            '    Typed = Table.TransformColumnTypes(Renamed, {{"Entity ID", type text}, {"Display ID", type text}, '
            '{"Name", type text}, {"Description", type text}, {"Status", type text}, {"Repository Area", type text}, '
            '{"TypeCode", type text}, {"Modified", type text}})',
            "in",
            "    Typed",
        ],
        "columns": ["Entity ID", "Display ID", "Name", "Description", "Status", "Repository Area", "TypeCode",
                    "Modified"],
        "measures": [
            ("Entities", "COUNTROWS('Entity')", "0"),
            ("Retired entities", "CALCULATE(COUNTROWS('Entity'), 'Entity'[Status] = \"Retired\")", "0"),
            ("Data quality %", "1 - DIVIDE(CALCULATE(DISTINCTCOUNT(Issues[RecordID]), Issues[Severity] = \"Error\", "
                               "Issues[Sheet] = \"DATA_Master\"), COUNTROWS('Entity'), 0)", "0%"),
        ],
    },
    "EntityAttribute": {
        "m": SOURCE + [
            '    Master = Source{[Item="tblMaster",Kind="Table"]}[Data],',
            '    Rows = Table.SelectRows(Master, each [Entity ID] <> null and [Entity ID] <> ""),',
            '    System = {"Entity ID", "Type", "Display ID", "Name", "Description", "Status", "Repository Area", '
            '"Created", "Created By", "Modified", "Modified By", "Row Hash"},',
            '    Long = Table.UnpivotOtherColumns(Rows, System, "Attribute", "Value"),',
            '    AsText = Table.TransformColumns(Long, {{"Value", each if _ is date or _ is datetime '
            'then Date.ToText(Date.From(_), "yyyy-MM-dd") else Text.From(_), type text}}),',
            '    Result = Table.TransformColumnTypes(Table.SelectColumns(AsText, {"Entity ID", "Attribute", "Value"}), '
            '{{"Entity ID", type text}, {"Attribute", type text}})',
            "in",
            "    Result",
        ],
        "columns": ["Entity ID", "Attribute", "Value"],
        "measures": [],
    },
    "EntityType": {
        "m": SOURCE + [
            '    Types = Source{[Item="tblEntityTypes",Kind="Table"]}[Data],',
            '    Picked = Table.SelectColumns(Types, {"TypeCode", "Name", "Domain", "Aspect"}),',
            '    Renamed = Table.RenameColumns(Picked, {{"Name", "Type Name"}}),',
            '    Typed = Table.TransformColumnTypes(Renamed, {{"TypeCode", type text}, {"Type Name", type text}, '
            '{"Domain", type text}, {"Aspect", type text}})',
            "in",
            "    Typed",
        ],
        "columns": ["TypeCode", "Type Name", "Domain", "Aspect"],
        "measures": [],
    },
    "Relationships": {
        "m": SOURCE + [
            '    Rel = Table.SelectRows(Source{[Item="tblRel",Kind="Table"]}[Data], each [RelID] <> null and [RelID] <> ""),',
            '    Verbs = Table.SelectColumns(Source{[Item="tblRelTypes",Kind="Table"]}[Data], {"RelCode", "Name"}),',
            '    Joined = Table.NestedJoin(Rel, {"RelCode"}, Verbs, {"RelCode"}, "V", JoinKind.LeftOuter),',
            '    Expanded = Table.ExpandTableColumn(Joined, "V", {"Name"}, {"Relationship"}),',
            '    Picked = Table.SelectColumns(Expanded, {"RelID", "RelCode", "Relationship", "SourceID", "TargetID", '
            '"Label", "Status"}),',
            '    Typed = Table.TransformColumnTypes(Picked, {{"RelID", type text}, {"RelCode", type text}, '
            '{"Relationship", type text}, {"SourceID", type text}, {"TargetID", type text}, {"Label", type text}, '
            '{"Status", type text}})',
            "in",
            "    Typed",
        ],
        "columns": ["RelID", "RelCode", "Relationship", "SourceID", "TargetID", "Label", "Status"],
        "measures": [("Relationship count", "COUNTROWS('Relationships')", "0")],
    },
    "Target Entity": {
        "m": SOURCE + [
            '    Master = Source{[Item="tblMaster",Kind="Table"]}[Data],',
            '    Rows = Table.SelectRows(Master, each [Entity ID] <> null and [Entity ID] <> ""),',
            '    Types = Table.SelectColumns(Source{[Item="tblEntityTypes",Kind="Table"]}[Data], {"TypeCode", "Name", "Domain"}),',
            '    Joined = Table.NestedJoin(Rows, {"Type"}, Types, {"TypeCode"}, "T", JoinKind.LeftOuter),',
            '    Expanded = Table.ExpandTableColumn(Joined, "T", {"Name", "Domain"}, {"Target Type Name", "Target Domain"}),',
            '    Picked = Table.SelectColumns(Expanded, {"Entity ID", "Display ID", "Name", "Status", "Target Type Name", '
            '"Target Domain"}),',
            '    Renamed = Table.RenameColumns(Picked, {{"Entity ID", "Target ID"}, {"Display ID", "Target Display ID"}, '
            '{"Name", "Target Name"}, {"Status", "Target Status"}}),',
            '    Typed = Table.TransformColumnTypes(Renamed, {{"Target ID", type text}, {"Target Display ID", type text}, '
            '{"Target Name", type text}, {"Target Status", type text}, {"Target Type Name", type text}, '
            '{"Target Domain", type text}})',
            "in",
            "    Typed",
        ],
        "columns": ["Target ID", "Target Display ID", "Target Name", "Target Status", "Target Type Name",
                    "Target Domain"],
        "measures": [],
    },
    "Issues": {
        "m": SOURCE + [
            '    Issues = Source{[Item="tblIssues",Kind="Table"]}[Data],',
            '    Rows = Table.SelectRows(Issues, each [Severity] <> null and [Severity] <> ""),',
            '    Picked = Table.SelectColumns(Rows, {"Found", "Sheet", "RecordID", "DisplayID", "Name", "TypeCode", '
            '"AttrCode", "RuleType", "Severity", "Message"}),',
            '    Typed = Table.TransformColumnTypes(Picked, {{"Found", type text}, {"Sheet", type text}, '
            '{"RecordID", type text}, {"DisplayID", type text}, {"Name", type text}, {"TypeCode", type text}, '
            '{"AttrCode", type text}, {"RuleType", type text}, {"Severity", type text}, {"Message", type text}})',
            "in",
            "    Typed",
        ],
        "columns": ["Found", "Sheet", "RecordID", "DisplayID", "Name", "TypeCode", "AttrCode", "RuleType",
                    "Severity", "Message"],
        "measures": [
            ("Errors", "CALCULATE(COUNTROWS(Issues), Issues[Severity] = \"Error\")", "0"),
            ("Warnings", "CALCULATE(COUNTROWS(Issues), Issues[Severity] = \"Warning\")", "0"),
        ],
    },
}

RELATIONSHIPS = [
    ("EntityAttribute", "Entity ID", "Entity", "Entity ID"),
    ("Entity", "TypeCode", "EntityType", "TypeCode"),
    ("Relationships", "SourceID", "Entity", "Entity ID"),
    ("Relationships", "TargetID", "Target Entity", "Target ID"),
]


def table_tmdl(name, spec):
    out = [f"table {q(name)}", f"{T}lineageTag: {tag('table/' + name)}", ""]
    for mname, dax, fmt in spec["measures"]:
        out += [f"{T}measure {q(mname)} = {dax}", f"{T}{T}formatString: {fmt}",
                f"{T}{T}lineageTag: {tag('measure/' + name + '/' + mname)}", ""]
    for col in spec["columns"]:
        out += [f"{T}column {q(col)}", f"{T}{T}dataType: string",
                f"{T}{T}lineageTag: {tag('column/' + name + '/' + col)}",
                f"{T}{T}summarizeBy: none", f"{T}{T}sourceColumn: {col}", "",
                f"{T}{T}annotation SummarizationSetBy = Automatic", ""]
    out += [f"{T}partition {q(name)} = m", f"{T}{T}mode: import", f"{T}{T}source =", m_block(spec["m"]), "",
            f"{T}annotation PBI_ResultType = Table", ""]
    return "\n".join(out)


def write(folder: Path, default_path: str):
    rep = folder / "Reports"
    sm = rep / "EA_Reports.SemanticModel"
    rp = rep / "EA_Reports.Report"
    (sm / "definition" / "tables").mkdir(parents=True, exist_ok=True)
    rp.mkdir(parents=True, exist_ok=True)

    (rep / "EA_Reports.pbip").write_text(json.dumps({
        "version": "1.0",
        "artifacts": [{"report": {"path": "EA_Reports.Report"}}],
        "settings": {"enableAutoRecovery": True},
    }, indent=2), encoding="utf-8")
    (sm / "definition.pbism").write_text(json.dumps({"version": "4.0", "settings": {}}, indent=2), encoding="utf-8")
    (sm / "definition" / "database.tmdl").write_text("database\n\tcompatibilityLevel: 1567\n\n", encoding="utf-8")
    order = ["ModelFolder"] + list(TABLES)
    model = ["model Model", f"{T}culture: en-US", f"{T}defaultPowerBIDataSourceVersion: powerBI_V3",
             f"{T}sourceQueryCulture: en-ZA", f"{T}dataAccessOptions", f"{T}{T}legacyRedirects",
             f"{T}{T}returnErrorValuesAsNull", "",
             f"annotation PBI_QueryOrder = {json.dumps(order)}", "",
             "annotation __PBI_TimeIntelligenceEnabled = 0", ""]
    model += [f"ref table {q(n)}" for n in TABLES] + [""]
    (sm / "definition" / "model.tmdl").write_text("\n".join(model), encoding="utf-8")
    m_path = default_path.replace('"', '""')
    (sm / "definition" / "expressions.tmdl").write_text("\n".join([
        f'expression ModelFolder = "{m_path}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]',
        f"{T}lineageTag: {tag('expression/ModelFolder')}", "",
        f"{T}annotation PBI_ResultType = Text", "",
    ]), encoding="utf-8")
    for name, spec in TABLES.items():
        (sm / "definition" / "tables" / f"{name}.tmdl").write_text(table_tmdl(name, spec), encoding="utf-8")
    rels = []
    for i, (ft, fc, tt, tc) in enumerate(RELATIONSHIPS):
        rels += [f"relationship {tag('rel/' + str(i))}", f"{T}fromColumn: {q(ft)}.{q(fc)}",
                 f"{T}toColumn: {q(tt)}.{q(tc)}", ""]
    (sm / "definition" / "relationships.tmdl").write_text("\n".join(rels), encoding="utf-8")

    (rp / "definition.pbir").write_text(json.dumps({
        "version": "4.0",
        "datasetReference": {"byPath": {"path": "../EA_Reports.SemanticModel"}, "byConnection": None},
    }, indent=2), encoding="utf-8")
    section = {"config": "{}", "displayName": "Start here", "displayOption": 1, "filters": "[]",
               "height": 720.0, "name": "ReportSection", "ordinal": 0, "visualContainers": [], "width": 1280.0}
    (rp / "report.json").write_text(json.dumps({
        "config": json.dumps({"version": "5.43", "themeCollection": {}}),
        "layoutOptimization": 0, "resourcePackages": [], "sections": [section], "filters": "[]",
    }, indent=2), encoding="utf-8")
    return rep


if __name__ == "__main__":
    folder = Path(sys.argv[1])
    default = sys.argv[2] if len(sys.argv) > 2 else str(folder).replace("/", "\\") + "\\"
    print("written", write(folder, default))
