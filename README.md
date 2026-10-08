# XLSX Editor

A small desktop app for opening, editing and saving `.xlsx` files on macOS and Windows, with no Microsoft Excel needed. Built with [Avalonia](https://avaloniaui.net/) and [ClosedXML](https://github.com/ClosedXML/ClosedXML). Everything runs locally and nothing is uploaded.

## Features

- Open (menu, Cmd/Ctrl+O or drag & drop), Save and Save As
- Edit cells in place. Formulas work too (`=SUM(A1:A10)`): while typing one, click a cell to insert its address
- Undo / redo, cut / copy / paste (paste a tab-separated block from Excel or Google Sheets)
- Find text across the sheet
- Find duplicate rows by one or more columns, then delete them, keep one of each, or copy or move them to a new sheet
- Select ranges, whole columns or rows by clicking/dragging the headers, or everything with Cmd/Ctrl+A
- Right-click cells, headers and sheet tabs for the common actions
- Insert/delete rows and columns; add (**+** next to the tabs), rename (double-click the tab), duplicate, delete and reorder sheets

## Run

Needs the [.NET 10 SDK](https://dotnet.microsoft.com/download).

```bash
dotnet run --project src/XlsxEditor.App
dotnet test   # core tests
```

## Build an app

Both builds are self-contained, so the target computer doesn't need .NET.

**macOS** produces `publish/XLSX Editor.app`:

```bash
./build-mac.sh            # build
./build-mac.sh --install  # build and copy to /Applications
```

To make it the default for `.xlsx` files, select any `.xlsx` in Finder, open Get Info (Cmd+I), set **Open with** to *XLSX Editor*, then click **Change All…**.
The app is signed ad-hoc but not notarized. On a Mac other than the one that built it, open it the first time with right-click → Open.

**Windows** produces one `.exe` in `publish/win` (this also works when run from a Mac):

```bash
dotnet publish src/XlsxEditor.App -c Release -r win-x64 --self-contained \
  -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish/win
```

To open `.xlsx` files with it, right-click a file → Open with → Choose another app → browse to `XlsxEditor.App.exe` → Always.

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
| Shift+Arrow, Shift+click, drag | Extend the selection |
| Cmd+A | Select all |
| Cmd+Page Down / Page Up | Next / previous sheet |
| Shift+F11 | New sheet |
| Cmd+F, then Enter / Shift+Enter | Find next / previous |

## Project layout

- `src/XlsxEditor.Core`: workbook logic (no UI). Wraps ClosedXML.
- `src/XlsxEditor.App`: Avalonia desktop app.
- `tests/XlsxEditor.Core.Tests`: xUnit tests for the core.

## License

[MIT](LICENSE). Third-party packages keep their own licenses (Avalonia and ClosedXML are both MIT).
