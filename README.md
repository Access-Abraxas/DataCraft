<p align="center">
  <a href="https://github.com/storytold/datacraft">
    <img alt="DataCraft" src="assets/app-icon/hicolor/256x256/apps/DataCraft_Icon.png" width="200" alt="DataCraft icon: An Microsoft Access database compatible desktop database development tool">
  </a>
</p>

<p align="center">
  <img src="assets/app-icon/hicolor/256x256/apps/DataCraft_Icon.png" width="128" height="128" alt="DataCraft icon: An Microsoft Access database compatible desktop database development tool">
</p>

<h1 align="center">DataCraft</h1>

<p align="center">
  <b>Databases and data; an open-source, modern Microsoft Access-compatible desktop database development tool, built in pure Rust.<b>
</p>

<p align="center">
  A fast, open-source, clean-room take on desktop database workflows. It runs natively on Windows, providing table, query, form, report, macro, and module designers plus a Visual Basic Editor with VBA support.<br>
  <i>By the DataCraft community.<i>
</p>

<p align="center">
  <img alt="Written in Rust" src="https://img.shields.io/badge/written%20in-C%23-662e9a?style=flat-square&logo=rust&logoColor=white">
  <img alt="Runs on Windows" src="https://img.shields.io/badge/runs%20on-Windows-1f9d55?style=flat-square&logo=windows&logoColor=white">
  <img alt="License: MIT" src="https://img.shields.io/badge/license-MIT-147a40?style=flat-square">
  <img alt="Agent-drivable over MCP" src="https://img.shields.io/badge/agents-MCP-1f9d55?style=flat-square">
  <br>
  <a href="https://sonarcloud.io/project/overview?id=storytold_datacraft"><img alt="Scanned by SonarCloud" src="https://img.shields.io/badge/Scanned_by-SonarCloud-4b9fd5?style=flat-square&logo=sonarqubecloud&logoColor=white"></a>
  <a href="https://coderabbit.ai"><img alt="Reviewed by CodeRabbit" src="https://img.shields.io/coderabbit/prs/github/storytold/datacraft?style=flat-square&label=Reviewed%20by%20CodeRabbit&logo=coderabbit"></a>
</p>

<p align="center">
  <a href="https://github.com/storytold/datacraft"><b>DataCraft on GitHub</b></a>  
  <a href="https://github.com/storytold/datacraft/releases">Releases</a>  
  <a href="https://github.com/storytold/datacraft/discussions">Discussions</a>
</p>

<br>

<p align="center">
  <img src="docs/images/hero-sales.png" alt="DataCraft showing a quarterly sales dashboard: a styled table with banded rows, filter buttons, a totals row, data bars, arrow icons and sparklines, next to a clustered column chart and a pie chart" width="100%">
  <br><sub><b>Quarterly Sales Dashboard</b>: a table with a totals row, conditional formatting, sparklines and two live charts   built entirely from DataCraft designers.<sub>
</p>

> [!NOTE]
> **DataCraft is a community project building a true open-source Microsoft Access replacement.** Developers, database designers, VBA power users, and people who picked up Access last week. If you build desktop databases, you"re one of us. **[Come say hi on Discord](https://discord.gg/datacraft).**

<p align="center">
  <a href="#a-look-around">A look around</a>  
  <a href="#why-datacraft">Why DataCraft</a>  
  <a href="#quick-start">Quick start</a>  
  <a href="#for-agents-mcp-cli-and-the-control-channel">For agents</a>  
  <a href="#architecture">Architecture</a>  
  <a href="#roadmap">Roadmap</a>  
  <a href="#license-and-credits">License and credits</a>
</p>

## A look around

<table>
<tr>
<td width="50%" valign="top"><img src="docs/images/budget.png" alt="A household budget with accounting number formats, red highlights for overspent categories, a green-yellow-red colour scale on the percent-used column, a totals row with a double border, and a horizontal bar chart comparing budget with actual" width="100%"><p align="center"><sub><b>Table Designer.</b> Field properties, data types, indexes, validation rules, and relationships   built on the Access Table Designer pattern.<sub>
<td width="50%" valign="top"><img src="docs/images/grades-formulas.png" alt="A query grid with criteria, joins, calculated fields, and the SQL view" width="100%"><p align="center"><sub><b>Query Designer.</b> Visual query building with grid, SQL view, parameter queries, and crosstab queries.<sub>
</tr>
<tr>
<td width="50%" valign="top"><img src="docs/images/dark-insert.png" alt="A form in Design view with the ribbon open showing controls, layout, and formatting options" width="100%"><p align="center"><sub><b>Form Designer.</b> Drag-and-drop controls, sections, subforms, and event-driven programming.<sub>
<td width="50%" valign="top"><img src="docs/images/pivot.png" alt="A report in Design view with grouping, sorting, and the report footer" width="100%"><p align="center"><sub><b>Report Designer.</b> Grouped and sorted reports, page headers/footers, and export to PDF.<sub>
</tr>
<tr>
<td width="50%" valign="top"><img src="docs/images/format-chart.png" alt="The Visual Basic Editor with a code module open, the Immediate window, and the Object Browser" width="100%"><p align="center"><sub><b>Visual Basic Editor.</b> VBA, VB.NET, and C# code editing with syntax coloring, breakpoints, stepping, and IntelliSense.<sub>
<td width="50%" valign="top"><img src="docs/images/page-break.png" alt="The Navigation Pane showing tables, queries, forms, reports, macros, and modules organized by object type" width="100%"><p align="center"><sub><b>Navigation Pane.</b> Access-style object browser with groups, filter, and search.<sub>
</tr>
</table>

## Why DataCraft

- **Familiar.** Ribbon, Navigation Pane, property sheets, datasheets, keyboard shortcuts, the Expression Builder, and the Visual Basic Editor   everything is where Access users expect it.
- **Compatible.** Opens and saves ACCDB and MDB databases via ADO.NET, reads and writes SQLite databases, and uses the native single-file ADX format (SQLite-based) for new databases. System tables (`ASysObjects`, `ASysObjDefs`) store database object metadata just like Access.
- **Capable.** Table, query, form, report, macro, and module designers; a full Visual Basic Editor with VBA, VB.NET, and C# language services; syntax coloring, breakpoints, stepping, Locals/Watch/Immediate windows, and a Primary Interop Assembly for automation.
- **Fast.** .NET 10 Windows Forms application, virtualised datasheets that only render visible rows, and incremental recalculation for computed fields.
- **Agent-native.** Every ribbon button, menu item, dialog, keystroke, and mouse gesture is a command that agents can run over a JSON control channel, an **MCP server**, or the CLI.
- **Free.** MIT licensed, no subscription, no licence server, no telemetry, no requirement that Microsoft Access be installed.

## Quick start

DataCraft requires .NET 10 or later.

```sh
git clone https://github.com/storytold/datacraft
cd datacraft
dotnet run --project cs_code/AccessClone.sln --configuration Release
```

Open an existing ACCDB, MDB, SQLite, or ADX database, or create a new ADX database from the File > New backstage.

### Ribbon access keys

Press and release **left Alt** on Windows, or press **F10**, then type a sequence one key at a time. Keytips appear on the supported Home ribbon controls and menu items.

## For agents: MCP, CLI, and the control channel

DataCraft is built to be agent-drivable. Every command in the UI is exposed through:
- An **MCP server** for LLM agent tool use
- A **JSON control channel** over TCP/stdin
- A **CLI** (`datacraft-cli`) for info, convert, run, commands, and mcp subcommands

See [docs/agent-integration.md](docs/agent-integration.md) for details.

## Architecture

DataCraft is a .NET 10 Windows Forms solution:

| Layer | Responsibility |
|---|---|
| `cs_code/src/AccessClone` | Main application shell, ribbon, docking workspace, Navigation Pane, Field List, Properties |
| `Controls` | Reusable Access-style controls: datasheet, code editor, property sheet |
| `Editor` | Visual Basic Editor shell, code pane, debug windows, Object Browser |
| `Models` | Database object models: tables, queries, forms, reports, macros, modules |
| `Services` | ADO.NET access, SQLite engine, ADX persistence, system-table catalog |
| `Resources` | Embedded database templates, card images, ribbon icons |

The ADX file format is SQLite-based with `ASysObjects` and `ASysObjDefs` system tables for object metadata.

## Roadmap

See [ROADMAP.md](ROADMAP.md) for the implementation plan, parity targets, and release schedule.

## License and credits

DataCraft is [MIT licensed](LICENSE).

Built with:
- [.NET 10](https://dotnet.microsoft.com/)   application framework
- [Microsoft.Data.Sqlite](https://learn.microsoft.com/en-us/dotnet/standard/data/sqlite/)   SQLite database engine
- [System.Data.OleDb](https://learn.microsoft.com/en-us/dotnet/framework/data/adonet/ole-db/)   ACCDB/MDB access via ADO.NET
- [Roslyn](https://github.com/dotnet/roslyn)   C# and VB.NET compilation and language services
