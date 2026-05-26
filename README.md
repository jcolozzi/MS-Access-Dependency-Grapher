# Access DB Grapher 

This project exports a Microsoft Access database to `graph.json` and loads that JSON into an interactive HTML node graph.

## Files

- `README.MD`
- `Export-AccessGraph.ps1` - extractor.
- `access-graph-viewer.html` - viewer.
- `out\graph.json` - Northwind 2.0 sample payload for testing the viewer.
- `Northwind-2.accdb` - Northwind 2.0

## What the extractor does

- inventories tables, queries, forms, reports, macros, and modules
- exports raw object text with `SaveAsText`
- writes a normalized `graph.json`
- adds table relation edges (foreign keys)
- adds query-to-table and query-to-query SQL reference edges
- parses form and report exports for:
  - `RecordSource`
  - bound `ControlSource`
  - subform or subreport `SourceObject`
  - `LinkMasterFields` and `LinkChildFields`
- VBA heuristics from form code, report code, and modules:
  - `DoCmd.OpenForm`, `DoCmd.OpenReport`, `DoCmd.OpenQuery`, `DoCmd.OpenTable`
  - `DoCmd.RunSQL` (extracts inline SQL and scans for table/query references)
  - `CurrentDb().QueryDefs("...")`, `DBEngine(0)(0).QueryDefs("...")`
- VBA type-dependency edges:
  - `Dim x As ClassName` / `Function() As ClassName`
  - `Set x = New ClassName`
  - `ClassName.Method` (qualified member access)
- VBA data-reference edges (scans string literals in VBA code for known table/query names):
  - catches `OpenRecordset("TableName")`, `Execute "INSERT INTO ..."`, `DLookup("Field", "TableName", ...)`, domain aggregates, and any other string literal that contains a table or query name
- macro heuristics:
  - `OpenForm`, `OpenReport`, `OpenQuery`, `OpenTable`, `RunSQL`

### Not implemented yet

- full VBA parser
- complete macro argument decoding
- dependency-engine reconciliation
- robust parsing of every `SaveAsText` edge case
- query output field discovery for all query types

## Viewer features

The viewer is a single HTML file using `vis-network` from a CDN.

### Group colors and shapes

| Group  | Color        | Shape    |
|--------|--------------|----------|
| table  | Green        | database |
| query  | Blue         | diamond  |
| form   | Teal         | box      |
| report | Gold         | box      |
| macro  | Purple       | hexagon  |
| module | Slate        | triangle |
| sql    | Light blue   | ellipse  |
| field  | Gray         | dot      |

### Edge styling by kind

Edges are color-coded and styled by their relationship type:

| Edge Kind              | Style  | Color  |
|------------------------|--------|--------|
| relation / recordsource | Solid  | Gray   |
| vba-type-ref           | Dashed | Purple |
| vba-data-ref           | Dotted | Green  |
| vba-open* / navigation | Solid  | Blue   |
| controlsource / field  | Solid  | Gold   |
| sourceobject           | Solid  | Teal   |
| macro-*                | Solid  | Purple |

### Layout modes

- **Physics** (default) - `forceAtlas2Based` solver; nodes settle via physics simulation. Toggle on/off with the Physics checkbox.
- **Hierarchical** - top-down layered layout. Tables at top, queries next, forms/reports below, modules at bottom. Toggle with the Hierarchical checkbox.

### Toolbar controls

- **Load graph.json** - pick a file from disk
- **Try graph.json** - auto-load `graph.json` from the same folder
- **Find** - search nodes by label or ID
- **Fit** - zoom to fit all visible nodes
- **Physics** - toggle physics simulation
- **Hierarchical** - toggle hierarchical layout
- **Edge labels** - show/hide edge labels

### Sidebar

- **Filters** - show/hide node groups with checkboxes
- **Summary** - node and edge counts by group
- **Legend** - node shapes and edge styles
- **Details** - click any node or edge to see its full JSON

## Requirements

Run the extractor on Windows with desktop Microsoft Access installed. The script automates Access through COM:

```powershell
New-Object -ComObject Access.Application
```

## Typical usage

Open PowerShell in this folder and run:

```powershell
.\Export-AccessGraph.ps1 -DatabasePath 'C:\path\to\YourDb.accdb' -OutDir '.\out'
```

Or to save the graph in the the project folder

```powershell
.\Export-AccessGraph.ps1 -DatabasePath 'C:\path\to\YourDb.accdb' -OutDir 'C:\path\to\DB folder\out'
```

That produces:

```text
out\graph.json
out\index.html
out\raw\tables\...
out\raw\queries\...
out\raw\forms\...
out\raw\reports\...
out\raw\macros\...
out\raw\modules\...
out\raw\sql\...
```

Open `out\index.html` in a browser.

If the browser refuses to auto-load `graph.json` from a local file, click **Load graph.json** and choose the file manually.

## Useful options

Default mode creates field nodes only when a form or report actually references a field:

```powershell
.\Export-AccessGraph.ps1 -DatabasePath 'C:\path\to\YourDb.accdb' -FieldNodeMode ReferencedOnly
```

Create field nodes for every table field:

```powershell
.\Export-AccessGraph.ps1 -DatabasePath 'C:\path\to\YourDb.accdb' -FieldNodeMode AllTableFields
```

Disable VBA heuristics:

```powershell
.\Export-AccessGraph.ps1 -DatabasePath 'C:\path\to\YourDb.accdb' -DisableCodeHeuristics
```

Disable macro heuristics:

```powershell
.\Export-AccessGraph.ps1 -DatabasePath 'C:\path\to\YourDb.accdb' -DisableMacroHeuristics
```

## Graph shape

`graph.json` is written as:

```json
{
  "meta": { "database": "...", "generatedAt": "..." },
  "nodes": [
    { "id": "table:Customers", "label": "Customers", "group": "table", "meta": {} }
  ],
  "edges": [
    { "id": "e1", "from": "form:frmCustomers", "to": "table:Customers", "label": "uses data", "kind": "vba-data-ref" }
  ]
}
```

### Edge kinds reference

| Kind                 | Source                                       |
|----------------------|----------------------------------------------|
| `relation`           | Database foreign key relationship            |
| `recordsource`       | Form/report RecordSource → table/query       |
| `recordsource-sql`   | Form/report RecordSource → inline SQL        |
| `controlsource`      | Control bound to a field                     |
| `control-expression` | Control with a calculated expression         |
| `sourceobject`       | Subform/subreport SourceObject               |
| `query-sql-reference`| Query SQL referencing a table/query          |
| `sql-reference`      | SQL node referencing a table/query           |
| `vba-type-ref`       | VBA `Dim As` / `New` / qualified access      |
| `vba-data-ref`       | VBA string literal containing table/query    |
| `vba-openform`       | `DoCmd.OpenForm`                             |
| `vba-openreport`     | `DoCmd.OpenReport`                           |
| `vba-openquery`      | `DoCmd.OpenQuery`                            |
| `vba-opentable`      | `DoCmd.OpenTable`                            |
| `vba-querydefs`      | `CurrentDb().QueryDefs("...")`               |
| `vba-runsql`         | `DoCmd.RunSQL`                               |
| `macro-openform`     | Macro OpenForm action                        |
| `macro-openreport`   | Macro OpenReport action                      |
| `macro-openquery`    | Macro OpenQuery action                       |
| `macro-opentable`    | Macro OpenTable action                       |
| `macro-runsql`       | Macro RunSQL action                          |
| `field-owner`        | Field node owned by a table/query            |

## Viewer notes

The viewer uses `vis-network` from a CDN. That means:

- internet access is needed the first time unless you replace the script tag with a local copy
- the graph data itself is always loaded from a local `graph.json`
