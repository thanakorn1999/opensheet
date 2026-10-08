# XLSX Editor

A small desktop app for opening, editing and saving `.xlsx` files on macOS and Windows, with no Microsoft Excel needed. Built with [Avalonia](https://avaloniaui.net/) and [ClosedXML](https://github.com/ClosedXML/ClosedXML). Everything runs locally and nothing is uploaded.

## Features

- Open (menu, Cmd/Ctrl+O or drag & drop), Save and Save As
- Edit cells in place. Formulas work too (`=SUM(A1:A10)`): while typing one, click a cell to insert its address
- Undo / redo, cut / copy / paste (paste a tab-separated block from Excel or Google Sheets)
- Find text across the sheet
- Find duplicate rows by one or more columns, then delete them, keep one of each, or copy or move them to a new sheet
- Rename sheets (double-click the tab), duplicate sheets, delete columns

## Run

Needs the [.NET 10 SDK](https://dotnet.microsoft.com/download).

```bash
dotnet run --project src/XlsxEditor.App
dotnet test   # core tests
```

## Keyboard shortcuts

Cmd on macOS, Ctrl on Windows.

| Keys | Action |
|---|---|
| Cmd+O / Cmd+S / Cmd+Shift+S | Open / Save / Save As |
| Cmd+Z / Cmd+Shift+Z (or Cmd+Y) | Undo / Redo |
| Cmd+X / Cmd+C / Cmd+V | Cut / Copy / Paste |
| Delete or Backspace | Clear cell |
| Type, F2, or double-click | Edit cell |
| Enter / Tab (+Shift to go back) | Save the edit and move down / right |
| Esc | Cancel the edit or close a bar |
| Arrows, Page Up/Down, Home | Move |
| Cmd+Arrow | Jump to the edge of the data |
| Cmd+Home | Go to A1 |
| Cmd+Page Down / Page Up | Next / previous sheet |
| Cmd+F, then Enter / Shift+Enter | Find next / previous |

## Project layout

- `src/XlsxEditor.Core`: workbook logic (no UI). Wraps ClosedXML.
- `src/XlsxEditor.App`: Avalonia desktop app.
- `tests/XlsxEditor.Core.Tests`: xUnit tests for the core.

## License

[MIT](LICENSE). Third-party packages keep their own licenses (Avalonia and ClosedXML are both MIT).
