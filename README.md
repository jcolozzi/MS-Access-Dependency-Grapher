# Microsoft Access Dependency Grapher

Export a Microsoft Access database to `graph.json` and explore it in an interactive HTML node graph.

## Files

| File | Purpose |
|------|---------|
| `Export-AccessGraph.ps1` | PowerShell extractor — scans the database and writes `graph.json` |
| `access-graph-viewer.html` | Single-file viewer using vis-network |
| `sample-graph.json` | Small sample payload for testing the viewer |

## What the extractor does

- inventories tables, queries, forms, reports, macros, and modules
- exports raw object text with `SaveAsText`
- writes a normalized `graph.json` with nodes, edges, and metadata
- copies the viewer HTML into the output folder with the graph data embedded (no separate file needed)
- shows a **progress bar** (`Write-Progress`) throughout the export so large databases don't appear stalled

### Object and edge discovery

- table relation edges (foreign keys)
- query-to-table and query-to-query SQL reference edges
- form and report exports parsed for:
  - `RecordSource` (table/query binding or inline SQL)
  - bound `ControlSource`
  - subform or subreport `SourceObject`
  - `LinkMasterFields` and `LinkChildFields`
  - `RowSource` on ComboBox/ListBox controls (table/query binding or inline SQL)
  - calculated control expressions (`=DLookup(...)`, etc.)
- VBA heuristics from form code, report code, and modules:
  - `DoCmd.OpenForm`, `DoCmd.OpenReport`, `DoCmd.OpenQuery`, `DoCmd.OpenTable`
  - `DoCmd.RunMacro` (calls to named macros)
  - `DoCmd.RunSQL` (extracts inline SQL and scans for table/query references)
  - `CurrentDb().QueryDefs("...")`, `DBEngine(0)(0).QueryDefs("...")`
  - `.SourceObject = "FormOrReportName"` (dynamic subform/subreport binding at runtime)
  - cross-module procedure calls (`Sub`/`Function`/`Property` — indexed with `Build-ProcIndex`, filtered against 87 built-in VBA/Access names to avoid false positives)
- VBA type-dependency edges:
  - `Dim x As ClassName` / `Function() As ClassName`
  - `Set x = New ClassName`
  - `ClassName.Method` (qualified member access)
- VBA data-reference edges (scans string literals in VBA code for known table/query names):
  - catches `OpenRecordset("TableName")`, `Execute "INSERT INTO ..."`, `DLookup("Field", "TableName", ...)`, domain aggregates, and any other string literal that contains a table or query name
- macro heuristics:
  - `OpenForm`, `OpenReport`, `OpenQuery`, `OpenTable`, `RunSQL`

### Progress reporting

The export shows a PowerShell progress bar with per-item detail for large phases:

| Phase | Status | % Range |
|-------|--------|---------|
| Open database | Opening database... | 0 |
| Tables | Scanning table: *name* | 5 |
| Relationships | Scanning relationships... | 10 |
| Queries | Exporting query (N): *name* | 15 |
| Forms | Exporting form (N/M): *name* | 25–40 |
| Reports | Exporting report (N/M): *name* | 40–50 |
| Macros | Exporting macros... | 50 |
| Modules | Exporting modules... | 55 |
| Query edges | Analyzing query edges... | 60 |
| Form edges | Analyzing form edges (N/M): *name* | 65–80 |
| Report edges | Analyzing report edges... | 82 |
| Macro edges | Analyzing macro edges... | 87 |
| Proc index | Building procedure index... | 88 |
| Module code | Analyzing module code... | 90 |
| Output | Writing graph output... | 95 |

### Not implemented yet

- full VBA parser
- complete macro argument decoding
- dependency-engine reconciliation
- robust parsing of every `SaveAsText` edge case
- query output field discovery for all query types

## Viewer features

The viewer is a single HTML file using `vis-network` from a CDN.

### Toolbar controls

| Control | Description |
|---------|-------------|
| **Load graph.json** | Pick a file from disk |
| **Find** | Search nodes by label or ID (Enter key also works) |
| **Fit** | Zoom to fit all visible nodes |
| **Physics** | Toggle force-directed physics simulation (auto-disabled after initial layout stabilizes) |
| **Hierarchical** | Toggle top-down hierarchical layout |
| **Edge labels** | Show/hide edge labels on the graph |
| **☽ / ☀** | Toggle light/dark mode (persisted to localStorage) |
| **☰** | Toggle sidebar visibility (persisted to localStorage) |

### Sidebar panels

All sidebar sections are collapsible accordions. Click any section header to expand or collapse it.

| Panel | Default | Contents |
|-------|---------|----------|
| **Filters** | Open | Checkboxes to show/hide each node group (table, query, form, etc.). State persisted to localStorage. |
| **Edge Filters** | Open | Checkboxes for each edge kind with color swatches and line styles. State persisted to localStorage. |
| **Summary** | Collapsed | Node and edge counts by type |
| **Legend** | Collapsed | Visual key: shapes and colors for each node group |
| **Details** | Open | Rich detail card for the selected node or edge |
| **Reports** | Collapsed | Ten collapsible analysis reports (see [Reports panel](#reports-panel)) |

The sidebar can be collapsed entirely with the **☰** button. The collapsed state persists across sessions via localStorage.

### Node interaction

- **Click** a node to see its details in the sidebar and **pin** a popup card beside it that stays open and follows the node as you pan or zoom
- **Double-click** a node to zoom and focus on it
- **Hover** over a node to see a rich tooltip with type-specific fields, connection summary, and expandable internals
- **Ctrl+click** to multi-select nodes
- **Focus Neighborhood** button in the details panel dims everything except the selected node and its direct neighbors

The pinned popup is dismissed by clicking its **✕** button, clicking an edge, clicking empty canvas space, or loading a new graph.

### Details panel

When a **node** is selected, the panel shows:
- group badge and label
- metadata (field count, raw size, SQL preview, etc.)
- incoming and outgoing neighbor lists with clickable links
- **Focus Neighborhood** button to highlight the node's connections
- expandable internals (rawHash, rawPath, etc.)

When an **edge** is selected, the panel shows:
- edge kind badge and label
- clickable From/To node links
- edge metadata

### Reports panel

The **Reports** panel (collapsed by default) holds ten on-demand analysis reports, each a collapsible sub-section that populates when a graph is loaded. Report rows that reference a node are clickable — selecting one focuses that node in the graph.

| Report | What it surfaces |
|--------|------------------|
| **⚠ Broken Objects** | Warnings emitted during export (`graph.meta.warnings`), with the owning object and group |
| **Orphan Candidates** | Non-field nodes with no inbound edges (potentially unused objects) |
| **Inline SQL Inventory** | All inline-SQL nodes with their origin and a SQL preview |
| **Linked Table Inventory** | Linked tables with their connection string, source table, and field count |
| **High Fan-In Objects** | Objects ranked by number of inbound edges (most depended-upon) |
| **Duplicate Inline SQL** | Inline-SQL nodes referenced by two or more consumers |
| **Unverified Field Bindings** | Field bindings the extractor could not verify against a real field |
| **Circular Dependencies** | Dependency cycles (strongly-connected components) between objects |
| **Complexity Hotspots** | Objects ranked by raw export size |
| **Tables Without Relationships** | Tables not participating in any relationship edge |

### Tooltips

Tooltips are built lazily on hover and cached for performance:
- **Node tooltips** — colored header with group/label, type-specific fields (SQL preview for queries, field count for tables, linked source for linked tables), connection summary, expandable internals
- **Edge tooltips** — colored header with edge kind, source → destination route, edge metadata

### Dark mode

Toggle with the **☽** button in the toolbar. Applies a dark color scheme to the entire UI including the graph canvas, sidebar, tooltips, and edge labels. The preference is saved to localStorage (`access-graph-theme`) and restored on next visit.

### Loading overlay

A spinner with progress text appears over the graph area while the data is being processed and the layout stabilizes. The UI shell (toolbar, sidebar, status bar) renders immediately — the graph work is deferred so the interface never appears frozen.

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
| vba-runmacro           | Solid  | Purple |
| vba-call               | Dashed | Gray   |
| vba-sourceobject       | Dashed | Teal   |
| controlsource / field  | Solid  | Gold   |
| rowsource              | Dashed | Gold   |
| sourceobject           | Solid  | Teal   |
| macro-*                | Solid  | Purple |

### Layout modes

- **Physics** (default) — `forceAtlas2Based` solver; nodes settle via physics simulation. On initial load the simulation runs through 400 stabilization iterations, then physics is **automatically disabled** so nodes stay in place without bouncing. The Physics checkbox unchecks itself after stabilization. Re-enable physics manually with the checkbox if needed.
- **Hierarchical** — top-down layered layout. Tables at top, queries next, forms/reports below, modules at bottom. Toggle with the Hierarchical checkbox. Stabilize-then-freeze also applies here.

### Embedded graph data

When the extractor copies the viewer into the output folder, it embeds the graph JSON directly into the HTML file. This means `out\index.html` is fully self-contained — no separate `graph.json` fetch is needed and it works on `file://` without a web server.

## Requirements

Run the extractor on Windows with desktop Microsoft Access installed. The script automates Access through COM:

```powershell
New-Object -ComObject Access.Application
```

## Typical usage

### No arguments (file picker)

Run the script with no arguments and a file picker opens to select an Access database. The output is written to an `access-graph-out` folder created in the **same location as the selected database**.

```powershell
.\Export-AccessGraph.ps1
```

Add `-NestInNamedFolder` to nest the output one level deeper, inside a folder named after the database. Picking `example.accdb` produces `example\access-graph-out`:

```powershell
.\Export-AccessGraph.ps1 -NestInNamedFolder
```

### With arguments

Open PowerShell in this folder and run:

```powershell
.\Export-AccessGraph.ps1 -DatabasePath 'C:\path\to\YourDb.accdb' -OutDir '.\out'
```

or to save the graph alongside the database:

```powershell
.\Export-AccessGraph.ps1 -DatabasePath 'C:\path\to\YourDb.accdb' -OutDir 'C:\path\to\DB folder\out'
```

That produces:

```text
out\
├── graph.json              # graph data
├── index.html              # self-contained viewer with embedded data
└── raw\
    ├── queries\            # SaveAsText exports
    ├── forms\
    ├── reports\
    ├── macros\
    ├── modules\
    └── sql\                # extracted SQL snippets from forms/reports
```

Open `out\index.html` in a browser.

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-DatabasePath` | *(file picker)* | Path to the `.accdb` or `.mdb` file. If omitted, a file picker opens to select one. |
| `-OutDir` | `access-graph-out` beside the database | Output directory. If omitted, an `access-graph-out` folder is created in the same location as the database. |
| `-NestInNamedFolder` | off | Place the output inside a subfolder named after the database. e.g. picking `example.accdb` creates `example\access-graph-out`. Ignored when `-OutDir` is specified. |
| `-FieldNodeMode` | `ReferencedOnly` | `None`, `ReferencedOnly`, or `AllTableFields` |
| `-DisableCodeHeuristics` | off | Skip VBA code analysis |
| `-DisableMacroHeuristics` | off | Skip macro analysis |
| `-SkipViewerCopy` | off | Don't copy the viewer HTML into the output |

### Examples

Create field nodes for every table field (uses friendly DAO type names like `Text`, `Long`, `DateTime`):

```powershell
.\Export-AccessGraph.ps1 -DatabasePath '.\YourDb.accdb' -FieldNodeMode AllTableFields
```

Disable VBA heuristics:

```powershell
.\Export-AccessGraph.ps1 -DatabasePath '.\YourDb.accdb' -DisableCodeHeuristics
```

Disable macro heuristics:

```powershell
.\Export-AccessGraph.ps1 -DatabasePath '.\YourDb.accdb' -DisableMacroHeuristics
```

## Graph shape

`graph.json` is structured as:

```json
{
  "meta": {
    "database": "...",
    "generatedAt": "...",
    "fieldNodeMode": "ReferencedOnly",
    "stats": { "nodeCount": 424, "edgeCount": 996, "tables": 20, "warnings": 0 },
    "warnings": []
  },
  "nodes": [
    { "id": "table:Customers", "label": "Customers", "group": "table", "meta": { "fieldCount": 12, "rawSize": 1024 } }
  ],
  "edges": [
    { "id": "e1", "from": "form:frmCustomers", "to": "table:Customers", "label": "uses data", "kind": "vba-data-ref", "meta": {} }
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
| `vba-sourceobject`   | VBA `.SourceObject = "..."` assignment        |
| `vba-runmacro`       | `DoCmd.RunMacro`                              |
| `vba-call`           | Cross-module procedure call                   |
| `rowsource`          | Control RowSource → table/query/SQL           |
| `macro-openform`     | Macro OpenForm action                        |
| `macro-openreport`   | Macro OpenReport action                      |
| `macro-openquery`    | Macro OpenQuery action                       |
| `macro-opentable`    | Macro OpenTable action                       |
| `macro-runsql`       | Macro RunSQL action                          |
| `field-owner`        | Field node owned by a table/query            |

## Viewer notes

The viewer uses `vis-network` from a CDN (`unpkg.com/vis-network`). Internet access is needed on first load unless you replace the script tag with a local copy. The graph data itself is always local — either embedded in the HTML or loaded from `graph.json`.

### Persisted settings

The viewer saves the following to `localStorage`:

| Key | Purpose |
|-----|---------|
| `access-graph-theme` | Light or dark mode |
| `access-graph-sidebar` | Sidebar open or collapsed |
| `access-graph-groups` | Which node groups are checked in the Filters panel |
| `access-graph-edge-kinds` | Which edge kinds are checked in the Edge Filters panel |

Filter state is restored on the next visit so only previously-visible groups and edge kinds are rendered, keeping load times fast for large graphs.

## Troubleshooting

### "Running scripts is disabled on this system" / having to use `Bypass`

If the script only runs when you set the execution policy to `Bypass`, the cause is almost always the **Mark of the Web (MOTW)** — not the script's contents. When `Export-AccessGraph.ps1` arrives via download, email, or a network share, Windows tags it with a `Zone.Identifier` alternate data stream marking it "from the Internet." Under the `RemoteSigned` policy, any script flagged that way must be digitally signed, so it is blocked. `Bypass` ignores the flag entirely, which is why it works.

You can keep `RemoteSigned` and still run the script in any of three ways.

#### Option 1 — Unblock the file (simplest)

Removes the MOTW tag. Run once per copy of the file:

```powershell
Unblock-File -Path .\Export-AccessGraph.ps1
```

Files run directly from a UNC/network path can still be treated as "remote" even after unblocking. If unblocking alone doesn't help, copy the script to a local folder first:

```powershell
Copy-Item '.\Export-AccessGraph.ps1' "$env:USERPROFILE\Scripts\"
Unblock-File "$env:USERPROFILE\Scripts\Export-AccessGraph.ps1"
```

#### Option 2 — Per-invocation bypass (no policy change)

Run a single session at `Bypass` without changing the machine-wide policy:

```powershell
powershell -ExecutionPolicy Bypass -File .\Export-AccessGraph.ps1 -DatabasePath 'C:\path\to\YourDb.accdb'
```

#### Option 3 — Sign the script (permanent, survives copying)

A signed script runs under `RemoteSigned` even with the MOTW present, and it keeps working when the file is copied or moved. This is the best option if you distribute the script to others.

### Creating a code-signing certificate

You don't need a paid certificate for internal use — a **self-signed** code-signing certificate works under `RemoteSigned` as long as it is trusted on the machines that run the script.

**1. Create the certificate** (PowerShell 5.1+, run once):

```powershell
$cert = New-SelfSignedCertificate `
    -Subject 'CN=Access Graph Tools' `
    -Type CodeSigningCert `
    -CertStoreLocation Cert:\CurrentUser\My `
    -KeyUsage DigitalSignature `
    -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears(5)
```

**2. Trust the certificate.** A self-signed cert is only trusted once it lives in the *Trusted Root Certification Authorities* and *Trusted Publishers* stores. Copy it from your personal store:

```powershell
$store = Get-Item "Cert:\CurrentUser\My\$($cert.Thumbprint)"
foreach ($path in 'Cert:\CurrentUser\Root', 'Cert:\CurrentUser\TrustedPublisher') {
    $s = New-Object System.Security.Cryptography.X509Certificates.X509Store(
        ($path -replace 'Cert:\\CurrentUser\\', ''), 'CurrentUser')
    $s.Open('ReadWrite'); $s.Add($store); $s.Close()
}
```

**3. Sign the script:**

```powershell
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Select-Object -First 1
Set-AuthenticodeSignature -FilePath .\Export-AccessGraph.ps1 -Certificate $cert
```

`Set-AuthenticodeSignature` appends a signature block to the bottom of the file. The script now runs under `RemoteSigned` on this machine.

> **Note:** A self-signed certificate is only trusted on machines where you have installed it into the trusted stores (step 2). To run the signed script on other PCs, export the public certificate (`Export-Certificate`) and import it into the *Trusted Root* and *Trusted Publishers* stores there. For wider distribution outside a managed environment, use a certificate from a public Certificate Authority instead.

**Re-sign after edits.** Any change to the file invalidates the signature, so re-run `Set-AuthenticodeSignature` (step 3) after editing the script.

## Screenshots - Northwind 2.0

### Dependency Graph - Object Filters and Edge Filters

![Dependency Graph](https://github.com/jcolozzi/MS-Access-Dependency-Grapher/blob/main/images/DependencyGraph.png)

### Summary / Legend

![Summary Legend](https://github.com/jcolozzi/MS-Access-Dependency-Grapher/blob/main/images/SummaryLegend.png)

### Object Details - Focus Neighborhood

![Focus Neighborhood](https://github.com/jcolozzi/MS-Access-Dependency-Grapher/blob/main/images/FocusNeighborhood.png)

### Light / Dark Mode

![Light Dark Mode](https://github.com/jcolozzi/MS-Access-Dependency-Grapher/blob/main/images/Light-Dark.png)

### Reports

![Reports](https://github.com/jcolozzi/MS-Access-Dependency-Grapher/blob/main/images/Reports.png)
