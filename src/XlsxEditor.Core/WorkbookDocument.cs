using System.Globalization;
using ClosedXML.Excel;

namespace XlsxEditor.Core;

/// <summary>Display data for one cell. Row/column are 1-based like Excel.</summary>
public readonly record struct CellView(string Text, string Content, bool IsNumber);

/// <summary>An opened .xlsx file. Thin wrapper over ClosedXML so the UI never touches it directly.</summary>
public sealed class WorkbookDocument : IDisposable
{
    XLWorkbook _wb; // replaced when undoing a structural edit

    /// <summary>Restores an earlier state and returns the step that redoes it.</summary>
    delegate UndoStep UndoStep();

    const int MaxUndo = 100;
    readonly List<UndoStep> _undo = [];
    readonly Stack<UndoStep> _redo = new();
    int _editDepth; // only the outermost edit records an undo step

    WorkbookDocument(XLWorkbook wb, string? path)
    {
        _wb = wb;
        Path = path;
    }

    public string? Path { get; private set; }

    public static WorkbookDocument Open(string path) => new(new XLWorkbook(path), path);

    public void SaveAs(string path)
    {
        _wb.SaveAs(path);
        Path = path;
    }

    public bool CanUndo => _undo.Count > 0;
    public bool CanRedo => _redo.Count > 0;

    public bool Undo()
    {
        if (_undo.Count == 0) return false;
        var step = _undo[^1];
        _undo.RemoveAt(_undo.Count - 1);
        _redo.Push(step());
        return true;
    }

    public bool Redo()
    {
        if (_redo.Count == 0) return false;
        _undo.Add(_redo.Pop()());
        return true;
    }

    /// <summary>Records <paramref name="undo"/> (unless nested in another edit) and runs <paramref name="change"/>.
    /// If the change throws, the step stays so a half-done change can still be undone.</summary>
    T Edit<T>(Func<UndoStep> undo, Func<T> change)
    {
        if (_editDepth++ == 0)
        {
            _undo.Add(undo());
            if (_undo.Count > MaxUndo) _undo.RemoveAt(0);
            _redo.Clear();
        }
        try { return change(); }
        finally { _editDepth--; }
    }

    void Edit(Func<UndoStep> undo, Action change) => Edit(undo, () => { change(); return 0; });

    /// <summary>Undo for cell edits: remembers just these cells' values/formulas (formatting isn't touched by edits).</summary>
    UndoStep CellsStep(int sheet, IReadOnlyList<(int Row, int Col)> cells)
    {
        var ws = Sheet(sheet);
        var saved = cells.Select(p => ws.Cell(p.Row, p.Col))
            .Select(c => (c.Address.RowNumber, c.Address.ColumnNumber, Formula: c.HasFormula ? c.FormulaA1 : null, Value: c.HasFormula ? (XLCellValue)Blank.Value : c.Value))
            .ToList();
        return () =>
        {
            var redo = CellsStep(sheet, cells);
            var target = Sheet(sheet);
            foreach (var (r, c, formula, value) in saved)
            {
                var cell = target.Cell(r, c);
                if (formula is not null) cell.FormulaA1 = formula;
                else cell.Value = value; // also drops a formula
            }
            return redo;
        };
    }

    /// <summary>Undo for structural edits (rows, columns, sheets): a full copy of the workbook.</summary>
    // ponytail: copies the whole file per step; fine for normal sheets, record finer steps if huge files feel slow.
    UndoStep SnapshotStep()
    {
        var ms = new MemoryStream(); // not disposed: ClosedXML reuses the last saved stream as the base for its next save
        _wb.SaveAs(ms);
        var bytes = ms.ToArray();
        return () =>
        {
            var redo = SnapshotStep();
            var old = _wb;
            _wb = new XLWorkbook(new MemoryStream(bytes));
            old.Dispose();
            return redo;
        };
    }

    public IReadOnlyList<string> SheetNames => _wb.Worksheets.OrderBy(ws => ws.Position).Select(ws => ws.Name).ToList(); // same order as Sheet(index)

    /// <summary>Last used row/column of a sheet (0 when empty). Sheet index is 0-based.</summary>
    public (int Rows, int Cols) UsedSize(int sheet)
    {
        var last = Sheet(sheet).LastCellUsed();
        return last is null ? (0, 0) : (last.Address.RowNumber, last.Address.ColumnNumber);
    }

    public CellView GetCell(int sheet, int row, int col)
    {
        var cell = Sheet(sheet).Cell(row, col);
        if (cell.IsEmpty() && !cell.HasFormula) return default;

        string text;
        try { text = cell.GetFormattedString(); }
        catch (Exception) { text = "#ERROR"; } // unsupported formula function etc.

        var content = cell.HasFormula ? "=" + cell.FormulaA1 : text;
        return new CellView(text, content, cell.DataType == XLDataType.Number);
    }

    /// <summary>
    /// Sets a cell from what the user typed, like Excel: "=…" is a formula, a number (current culture) or
    /// TRUE/FALSE becomes that type, "'…" forces text, blank clears the value (formatting stays).
    /// </summary>
    // ponytail: dates and percentages stay text until someone types one and expects otherwise.
    public void SetCell(int sheet, int row, int col, string input) =>
        Edit(() => CellsStep(sheet, [(row, col)]), () => Put(sheet, row, col, input));

    /// <summary>Pastes a block of typed values (e.g. tab-separated text from Excel) with its top-left at (row, col). One undo step.</summary>
    public void SetCells(int sheet, int row, int col, IReadOnlyList<IReadOnlyList<string>> block)
    {
        var cells = block.SelectMany((line, r) => line.Select((_, c) => (row + r, col + c))).ToList();
        Edit(() => CellsStep(sheet, cells), () =>
        {
            for (int r = 0; r < block.Count; r++)
                for (int c = 0; c < block[r].Count; c++)
                    Put(sheet, row + r, col + c, block[r][c]);
        });
    }

    /// <summary>Copies one cell's value or formula like Excel's paste: relative references shift with the move.</summary>
    public void CopyCell(int fromSheet, int fromRow, int fromCol, int toSheet, int toRow, int toCol)
    {
        var src = Sheet(fromSheet).Cell(fromRow, fromCol);
        Edit(() => CellsStep(toSheet, [(toRow, toCol)]), () =>
        {
            var dst = Sheet(toSheet).Cell(toRow, toCol);
            if (src.HasFormula) dst.FormulaR1C1 = src.FormulaR1C1; // R1C1 is relative, so A1 refs shift
            else dst.Value = src.Value;
        });
    }

    void Put(int sheet, int row, int col, string input)
    {
        var cell = Sheet(sheet).Cell(row, col);
        if (input.Length == 0) cell.Value = Blank.Value;
        else if (input.Length > 1 && input[0] == '=') cell.FormulaA1 = input[1..];
        else if (input[0] == '\'') cell.Value = input[1..];
        else if (double.TryParse(input, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out var number)) cell.Value = number;
        else if (bool.TryParse(input, out var flag)) cell.Value = flag;
        else cell.Value = input;
    }

    /// <summary>
    /// Groups of rows (1-based, ascending) whose displayed values in <paramref name="cols"/> are all equal,
    /// like Excel's "Remove Duplicates". Only groups of 2+ rows. Case-insensitive; rows blank in every
    /// one of those columns are ignored. <paramref name="hasHeader"/> skips row 1.
    /// </summary>
    public IReadOnlyList<int[]> DuplicateGroups(int sheet, IReadOnlyList<int> cols, bool hasHeader = false) =>
        Enumerable.Range(hasHeader ? 2 : 1, Math.Max(0, UsedSize(sheet).Rows - (hasHeader ? 1 : 0)))
            .Select(r => (Row: r, Values: cols.Select(c => GetCell(sheet, r, c).Text ?? "").ToArray()))
            .Where(x => x.Values.Any(v => v.Length > 0))
            .GroupBy(x => string.Join('\u001F', x.Values), StringComparer.OrdinalIgnoreCase) // unit separator: can't appear in cell text
            .Where(g => g.Count() > 1)
            .Select(g => g.Select(x => x.Row).ToArray())
            .ToList();

    public void DeleteRows(int sheet, IEnumerable<int> rows) => Edit(SnapshotStep, () => DeleteLines(sheet, rows, byRow: true));

    public void DeleteColumns(int sheet, IEnumerable<int> cols) => Edit(SnapshotStep, () => DeleteLines(sheet, cols, byRow: false));

    // ponytail: one ClosedXML delete per row/column; batch contiguous runs if deleting thousands feels slow.
    void DeleteLines(int sheet, IEnumerable<int> lines, bool byRow)
    {
        var ws = Sheet(sheet);
        var deleted = lines.Distinct().Order().ToList();

        // ClosedXML 0.105 shifts hyperlinks one at a time in storage order and throws "same key" when a link
        // moves onto one that hasn't moved yet. So detach them all, delete, then re-attach where they shifted to.
        // Detaching resets font color/underline, so remember those too.
        // ponytail: a link spanning several cells comes back on its top-left cell only.
        var links = ws.Hyperlinks
            .Select(h => (Link: h, Cell: h.Cell!)) // links attached to a sheet always have a cell
            .Select(x => (x.Link, x.Cell.Address.RowNumber, x.Cell.Address.ColumnNumber,
                x.Cell.Style.Font.FontColor, x.Cell.Style.Font.Underline))
            .ToList();
        foreach (var l in links) ws.Hyperlinks.Delete(l.Link);

        for (int i = deleted.Count - 1; i >= 0; i--) // last first keeps the other numbers valid
            if (byRow) ws.Row(deleted[i]).Delete();
            else ws.Column(deleted[i]).Delete();

        foreach (var l in links)
        {
            int i = deleted.BinarySearch(byRow ? l.RowNumber : l.ColumnNumber);
            if (i >= 0) continue; // its row/column is gone
            var cell = byRow // ~i = deleted lines before it
                ? ws.Cell(l.RowNumber - ~i, l.ColumnNumber)
                : ws.Cell(l.RowNumber, l.ColumnNumber - ~i);
            cell.SetHyperlink(l.Link);
            cell.Style.Font.FontColor = l.FontColor;
            cell.Style.Font.Underline = l.Underline;
        }
    }

    /// <summary>
    /// Copies <paramref name="rows"/> (in order, after row 1 when <paramref name="withHeader"/>) to a new sheet
    /// placed right after this one. Blank <paramref name="name"/> picks "Duplicates", "Duplicates 2", …
    /// Throws if <paramref name="name"/> fails <see cref="CheckNewSheetName"/>. Returns the new sheet's name.
    /// </summary>
    public string CopyRowsToNewSheet(int sheet, IEnumerable<int> rows, bool withHeader, string? name = null)
    {
        name = NewSheetName(name, DuplicatesName); // validate before recording undo
        return Edit(SnapshotStep, () => CopyRowsTo(sheet, rows, withHeader, name));
    }

    static string DuplicatesName(int i) => i == 1 ? "Duplicates" : $"Duplicates {i}";

    string CopyRowsTo(int sheet, IEnumerable<int> rows, bool withHeader, string name)
    {
        var src = Sheet(sheet);
        var dst = _wb.AddWorksheet(name);
        dst.Position = sheet + 2; // 1-based: right after the source
        int n = 1;
        if (withHeader) src.Row(1).CopyTo(dst.Row(n++));
        foreach (var r in rows.Distinct().Order()) src.Row(r).CopyTo(dst.Row(n++));
        return name;
    }

    /// <summary>Same as <see cref="CopyRowsToNewSheet"/>, then deletes the rows here.</summary>
    public string MoveRowsToNewSheet(int sheet, IReadOnlyCollection<int> rows, bool withHeader, string? name = null)
    {
        name = NewSheetName(name, DuplicatesName);
        return Edit(SnapshotStep, () => // one undo step for both halves
        {
            CopyRowsTo(sheet, rows, withHeader, name);
            DeleteLines(sheet, rows, byRow: true);
            return name;
        });
    }

    /// <summary>
    /// Copies a whole sheet (values, formulas, styles, merges…) to a new sheet right after it, like Excel's
    /// "Move or Copy → Create a copy". Blank <paramref name="name"/> picks "Sheet1 (2)", "Sheet1 (3)", …
    /// Throws if <paramref name="name"/> fails <see cref="CheckNewSheetName"/>. Returns the new sheet's name.
    /// </summary>
    public string CopySheet(int sheet, string? name = null)
    {
        var src = Sheet(sheet);
        name = NewSheetName(name, i =>
        {
            var suffix = $" ({i + 1})";
            return src.Name[..Math.Min(src.Name.Length, 31 - suffix.Length)] + suffix; // stay within 31 chars
        });
        return Edit(SnapshotStep, () =>
        {
            src.CopyTo(name).Position = sheet + 2; // 1-based: right after the source
            return name;
        });
    }

    /// <summary>The trimmed <paramref name="requested"/> name if valid (throws if not), or when blank the first free <paramref name="auto"/>(1, 2, …).</summary>
    string NewSheetName(string? requested, Func<int, string> auto)
    {
        if (string.IsNullOrWhiteSpace(requested))
        {
            for (int i = 1; ; i++)
                if (CheckNewSheetName(auto(i)) is null) return auto(i);
        }
        var name = requested.Trim();
        if (CheckNewSheetName(name) is { } error) throw new ArgumentException(error, nameof(requested));
        return name;
    }

    /// <summary>Renames a sheet; throws if <paramref name="name"/> fails <see cref="CheckNewSheetName"/>. Returns the trimmed name.</summary>
    public string RenameSheet(int sheet, string name)
    {
        name = name.Trim();
        if (CheckNewSheetName(name, except: sheet) is { } error) throw new ArgumentException(error, nameof(name));
        Edit(SnapshotStep, () => Sheet(sheet).Name = name); // snapshot: rename also rewrites formulas elsewhere
        return name;
    }

    /// <summary>
    /// Why <paramref name="name"/> can't be a sheet's name (Excel's rules), or null if it can.
    /// <paramref name="except"/> is a sheet being renamed, so its own current name doesn't count as taken.
    /// </summary>
    public string? CheckNewSheetName(string name, int? except = null) =>
        name.Length == 0 ? "Sheet name is empty"
        : name.Length > 31 ? "Sheet name is longer than 31 characters"
        : name.IndexOfAny(['\\', '/', '?', '*', '[', ']', ':']) >= 0 ? "Sheet name can't contain \\ / ? * [ ] :"
        : name.StartsWith('\'') || name.EndsWith('\'') ? "Sheet name can't start or end with '"
        : SheetNames.Where((_, i) => i != except).Contains(name, StringComparer.OrdinalIgnoreCase) ? $"A sheet named \"{name}\" already exists" // Excel ignores case
        : null;

    /// <summary>Parses "A", "A,C", "B:D" or a mix like "a, c:e" into 1-based column numbers. Null if invalid.</summary>
    public static IReadOnlyList<int>? ParseColumns(string text)
    {
        var cols = new SortedSet<int>();
        foreach (var part in text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var ends = part.Split(':', StringSplitOptions.TrimEntries);
            if (ends.Length > 2 || !ends.All(XLHelper.IsValidColumn)) return null;
            int a = XLHelper.GetColumnNumberFromLetter(ends[0]), b = XLHelper.GetColumnNumberFromLetter(ends[^1]);
            for (int c = Math.Min(a, b); c <= Math.Max(a, b); c++) cols.Add(c);
        }
        return cols.Count == 0 ? null : cols.ToList();
    }

    /// <summary>
    /// Next cell after (row, col) whose displayed text contains <paramref name="query"/>, case-insensitive.
    /// Goes row by row and wraps around; the start cell itself is checked last.
    /// </summary>
    // ponytail: scans the used range cell by cell; index non-empty cells if huge sparse sheets feel slow.
    public (int Row, int Col)? Find(int sheet, string query, int row, int col, bool backward = false)
    {
        var (rows, cols) = UsedSize(sheet);
        if (rows == 0 || query.Length == 0) return null;
        long total = (long)rows * cols;
        long start = (long)(Math.Clamp(row, 1, rows) - 1) * cols + Math.Clamp(col, 1, cols) - 1;
        int step = backward ? -1 : 1;
        for (long i = 1; i <= total; i++)
        {
            long idx = ((start + step * i) % total + total) % total;
            int r = (int)(idx / cols) + 1, c = (int)(idx % cols) + 1;
            if (GetCell(sheet, r, c).Text is { } text && text.Contains(query, StringComparison.OrdinalIgnoreCase))
                return (r, c);
        }
        return null;
    }

    /// <summary>
    /// Where Ctrl/Cmd+Arrow lands from (row, col) moving by (dRow, dCol), like Excel: to the end of the current
    /// block of filled cells, else to the next filled cell, else to the edge of the used range.
    /// </summary>
    public (int Row, int Col) JumpEdge(int sheet, int row, int col, int dRow, int dCol)
    {
        var (rows, cols) = UsedSize(sheet);
        int maxR = Math.Max(rows, row), maxC = Math.Max(cols, col);
        bool Inside(int r, int c) => r >= 1 && c >= 1 && r <= maxR && c <= maxC;
        bool Filled(int r, int c) => !string.IsNullOrEmpty(GetCell(sheet, r, c).Text);

        int nr = row + dRow, nc = col + dCol;
        if (!Inside(nr, nc)) return (Math.Clamp(nr, 1, maxR), Math.Clamp(nc, 1, maxC));
        if (Filled(row, col) && Filled(nr, nc))
        {
            while (Inside(nr + dRow, nc + dCol) && Filled(nr + dRow, nc + dCol)) (nr, nc) = (nr + dRow, nc + dCol);
            return (nr, nc);
        }
        while (Inside(nr, nc) && !Filled(nr, nc)) (nr, nc) = (nr + dRow, nc + dCol);
        return Inside(nr, nc) ? (nr, nc) : (Math.Clamp(nr, 1, maxR), Math.Clamp(nc, 1, maxC));
    }

    public static string Address(int row, int col) => ColumnLetter(col) + row;

    public static string ColumnLetter(int col) => XLHelper.GetColumnLetterFromNumber(col);

    IXLWorksheet Sheet(int index) => _wb.Worksheet(index + 1);

    public void Dispose() => _wb.Dispose();
}
