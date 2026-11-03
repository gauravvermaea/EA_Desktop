# Power BI reports for the EA Workbench

`EA_Reports.pbip` is a Power BI project: a semantic model (tables, relationships and measures) that reads the
saved `EA_Core.xlsm` in this model folder. It needs Power BI Desktop, which is a free install. The project
opens with an empty "Start here" page. The model is ready, and the suggested pages below take about
15 minutes to lay out.

## First open

1. Open `Reports\EA_Reports.pbip` in Power BI Desktop.
2. **Home → Transform data → Edit parameters → ModelFolder:** set it to this model folder, ending with a backslash,
   for example `C:\EA\EA_Model_Sample\`. It is preset to the folder the sample was installed into.
3. **Home → Refresh.** Power BI reads the last *saved* EA_Core, so save EA_Core first (or use Refresh reports in EA_Core).
4. If Power BI asks about privacy levels, choose **Organizational** for the file, or ignore privacy levels for this file.

## The model

| Table | One row per | Key columns |
| --- | --- | --- |
| Entity | entity | Entity ID, Display ID, Name, Status, Repository Area, TypeCode |
| EntityAttribute | entity × filled attribute | Entity ID, Attribute (display name), Value (text) |
| EntityType | entity type | TypeCode, Type Name, Domain, Aspect |
| Relationships | relationship | SourceID → Entity, TargetID → Target Entity, Relationship (verb), RelCode |
| Target Entity | entity (as a relationship target) | Target ID, Target Name, Target Type Name, Target Domain |
| Issues | Validate All issue | RecordID, Severity, Message, AttrCode |

Measures: Entities, Retired entities, Data quality %, Relationship count, Errors, Warnings.

Attributes are in the long EntityAttribute table. Because of that, new attributes and new entity types appear after a
refresh without any change to the model.

## Suggested pages

| Page | Visuals |
| --- | --- |
| Portfolio overview | Cards: Entities, Relationship count, Errors, Data quality %. Stacked bar: Entities by EntityType[Domain] and EntityType[Type Name]. |
| Application × Capability | Matrix: rows Entity[Name], columns Target Entity[Target Name], values Relationship count. Filters: EntityType[TypeCode] = ApplicationComponent, Target Entity[Target Type Name] = Capability. |
| Lifecycle and risk | Matrix: rows Entity[Name], columns EntityAttribute[Value] filtered to Attribute = Lifecycle, values Entities. Add a slicer on Attribute = Criticality for a heatmap. |
| Data quality | Table: Issues[DisplayID], [Name], [Severity], [Message]. Slicer: Issues[Severity]. Card: Data quality %. |
| Drill-through | A page with Entity[Name] as the drill-through field, showing its attributes (EntityAttribute) and relationships (Relationships with Target Entity). |

Save the project after laying out the pages. The `.pbip` folder is plain text, so it can live in source control.
