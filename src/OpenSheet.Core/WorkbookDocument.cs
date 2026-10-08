using System.Globalization;
using ClosedXML.Excel;

namespace OpenSheet.Core;

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

    char _csvDelimiter = ','; // kept from an opened .csv so saving writes the same format back

    public static WorkbookDocument Open(string path)
    {
        var wb = Load(path, out var delimiter);
        // Files from tools that don't store formula results (openpyxl, some exports) come with blank results:
        // compute those. Results Excel stored are kept, so functions ClosedXML can't evaluate still show right.
        foreach (var ws in wb.Worksheets)
            foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.All, c => c.HasFormula && c.CachedValue.IsBlank))
                cell.InvalidateFormula();
        return new(wb, path) { _csvDelimiter = delimiter };
    }

    /// <summary>Loads an .xlsx, or a .csv as a one-sheet workbook named after the file.</summary>
    static XLWorkbook Load(string path, out char delimiter)
    {
        delimiter = ',';
        if (!Csv.IsCsv(path)) return new XLWorkbook(path);

        var text = Csv.ReadText(path);
        delimiter = Csv.DetectDelimiter(text);
        var wb = new XLWorkbook();
        var name = string.Concat(System.IO.Path.GetFileNameWithoutExtension(path).Select(ch => "\\/?*[]:".Contains(ch) ? '_' : ch)).Trim('\'', ' ');
        var ws = wb.AddWorksheet(name.Length == 0 ? "Sheet1" : name[..Math.Min(name.Length, 31)]);
        var rows = Csv.Parse(text, delimiter);
        for (int r = 0; r < rows.Count; r++)
            for (int c = 0; c < rows[r].Count; c++)
                Put(ws.Cell(r + 1, c + 1), rows[r][c]); // "42" becomes a number, "=A1*2" a formula, like Excel
        return wb;
    }

    /// <summary>
    /// Saves as .xlsx, or as .csv when the path ends in .csv. A CSV holds one sheet (<paramref name="sheet"/>)
    /// and only its displayed values, like Excel.
    /// </summary>
    public void SaveAs(string path, int sheet = 0)
    {
        if (Csv.IsCsv(path))
        {
            var (rows, cols) = UsedSize(sheet);
            Csv.Write(path, Enumerable.Range(1, rows).Select(r => Enumerable.Range(1, cols).Select(c => GetCell(sheet, r, c).Text ?? "")), _csvDelimiter);
        }
        else _wb.SaveAs(path);
        Path = path;
    }

    public bool IsCsv => Path is not null && Csv.IsCsv(Path);

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
    UndoStep CellsStep(IReadOnlyList<(int Sheet, int Row, int Col)> cells)
    {
        var saved = cells.Select(p => (p, Cell: Sheet(p.Sheet).Cell(p.Row, p.Col)))
            .Select(x => (x.p, Formula: x.Cell.HasFormula ? x.Cell.FormulaA1 : null, Value: x.Cell.HasFormula ? (XLCellValue)Blank.Value : x.Cell.Value))
            .ToList();
        return () =>
        {
            var redo = CellsStep(cells);
            foreach (var (p, formula, value) in saved)
            {
                var cell = Sheet(p.Sheet).Cell(p.Row, p.Col);
                if (formula is not null) cell.FormulaA1 = formula;
                else cell.Value = value; // also drops a formula
            }
            return redo;
        };
    }

    static List<(int Sheet, int Row, int Col)> Cells(int sheet, int top, int left, int bottom, int right) =>
        [.. from r in Enumerable.Range(top, Math.Max(0, bottom - top + 1))
            from c in Enumerable.Range(left, Math.Max(0, right - left + 1))
            select (sheet, r, c)];

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
        Edit(() => CellsStep([(sheet, row, col)]), () => Put(sheet, row, col, input));

    /// <summary>Pastes a block of typed values (e.g. tab-separated text from Excel) with its top-left at (row, col). One undo step.</summary>
    public void SetCells(int sheet, int row, int col, IReadOnlyList<IReadOnlyList<string>> block)
    {
        var cells = block.SelectMany((line, r) => line.Select((_, c) => (sheet, row + r, col + c))).ToList();
        Edit(() => CellsStep(cells), () =>
        {
            for (int r = 0; r < block.Count; r++)
                for (int c = 0; c < block[r].Count; c++)
                    Put(sheet, row + r, col + c, block[r][c]);
        });
    }

    /// <summary>The part of a range that has data (whole columns/rows shrink to the used area). Null if none.</summary>
    public (int Top, int Left, int Bottom, int Right)? UsedPart(int sheet, int top, int left, int bottom, int right)
    {
        var (rows, cols) = UsedSize(sheet);
        bottom = Math.Min(bottom, rows);
        right = Math.Min(right, cols);
        return bottom < top || right < left ? null : (top, left, bottom, right);
    }

    /// <summary>Clears values/formulas in a range (formatting stays). One undo step.</summary>
    public void ClearRange(int sheet, int top, int left, int bottom, int right)
    {
        if (UsedPart(sheet, top, left, bottom, right) is not { } u) return;
        var cells = Cells(sheet, u.Top, u.Left, u.Bottom, u.Right);
        Edit(() => CellsStep(cells), () =>
        {
            foreach (var (_, r, c) in cells) Sheet(sheet).Cell(r, c).Value = Blank.Value;
        });
    }

    /// <summary>A range's displayed text as tab-separated lines, the way Excel puts it on the clipboard.</summary>
    public string RangeText(int sheet, int top, int left, int bottom, int right)
    {
        if (UsedPart(sheet, top, left, bottom, right) is not { } u) return "";
        return string.Join("\n", Enumerable.Range(u.Top, u.Bottom - u.Top + 1).Select(r =>
            string.Join('\t', Enumerable.Range(u.Left, u.Right - u.Left + 1).Select(c => GetCell(sheet, r, c).Text ?? ""))));
    }

    /// <summary>
    /// Pastes a range with its top-left at (toRow, toCol), like Excel. Copy shifts relative references;
    /// <paramref name="move"/> (cut) keeps formulas as written and clears the source. Values keep their type.
    /// One undo step. Returns the pasted size.
    /// </summary>
    // ponytail: values and formulas only; formatting isn't copied.
    public (int Rows, int Cols) CopyRange(int fromSheet, int top, int left, int bottom, int right, int toSheet, int toRow, int toCol, bool move = false)
    {
        if (UsedPart(fromSheet, top, left, bottom, right) is not { } u) return (0, 0);
        int rows = u.Bottom - u.Top + 1, cols = u.Right - u.Left + 1;
        var src = Cells(fromSheet, u.Top, u.Left, u.Bottom, u.Right);
        var dst = Cells(toSheet, toRow, toCol, toRow + rows - 1, toCol + cols - 1);
        Edit(() => CellsStep(move ? [.. src, .. dst] : dst), () =>
        {
            // Read everything first so overlapping source and destination don't clobber each other.
            var data = src.Select(p => Sheet(p.Sheet).Cell(p.Row, p.Col))
                .Select(c => (Formula: c.HasFormula ? (move ? c.FormulaA1 : c.FormulaR1C1) : null, c.Value))
                .ToList();
            if (move) foreach (var (s, r, c) in src) Sheet(s).Cell(r, c).Value = Blank.Value;
            for (int i = 0; i < dst.Count; i++)
            {
                var cell = Sheet(toSheet).Cell(dst[i].Row, dst[i].Col);
                if (data[i].Formula is not { } f) cell.Value = data[i].Value;
                else if (move) cell.FormulaA1 = f;
                else cell.FormulaR1C1 = f; // R1C1 is relative, so A1 references shift
            }
        });
        return (rows, cols);
    }

    void Put(int sheet, int row, int col, string input) => Put(Sheet(sheet).Cell(row, col), input);

    static void Put(IXLCell cell, string input)
    {
        if (input.Length == 0) cell.Value = Blank.Value;
        else if (input.Length > 1 && input[0] == '=') cell.FormulaA1 = input[1..];
        else if (input[0] == '\'') cell.Value = input[1..];
        else if (TryParseNumber(input, out var number)) cell.Value = number;
        else if (bool.TryParse(input, out var flag)) cell.Value = flag;
        else cell.Value = input;
    }

    /// <summary>
    /// A number in the user's format. Thousands separators count only when they really group by threes
    /// ("1,234.5"), so "1,5" stays text instead of turning into 15.
    /// </summary>
    static bool TryParseNumber(string input, out double number)
    {
        var nf = CultureInfo.CurrentCulture.NumberFormat;
        if (double.TryParse(input, NumberStyles.Float, nf, out number)) return true;
        string g = System.Text.RegularExpressions.Regex.Escape(nf.NumberGroupSeparator), d = System.Text.RegularExpressions.Regex.Escape(nf.NumberDecimalSeparator);
        return System.Text.RegularExpressions.Regex.IsMatch(input.Trim(), $@"^[+-]?\d{{1,3}}({g}\d{{3}})+({d}\d+)?$")
            && double.TryParse(input, NumberStyles.Float | NumberStyles.AllowThousands, nf, out number);
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

    /// <summary>Inserts <paramref name="count"/> empty rows above row <paramref name="before"/>; formulas adjust.</summary>
    public void InsertRows(int sheet, int before, int count) => Edit(SnapshotStep, () =>
        WithLinksDetached(Sheet(sheet), byRow: true,
            ws => ws.Row(before).InsertRowsAbove(count),
            line => line >= before ? line + count : line));

    /// <summary>Inserts <paramref name="count"/> empty columns left of column <paramref name="before"/>; formulas adjust.</summary>
    public void InsertColumns(int sheet, int before, int count) => Edit(SnapshotStep, () =>
        WithLinksDetached(Sheet(sheet), byRow: false,
            ws => ws.Column(before).InsertColumnsBefore(count),
            line => line >= before ? line + count : line));

    // ponytail: one ClosedXML delete per row/column; batch contiguous runs if deleting thousands feels slow.
    void DeleteLines(int sheet, IEnumerable<int> lines, bool byRow)
    {
        var deleted = lines.Distinct().Order().ToList();
        WithLinksDetached(Sheet(sheet), byRow,
            ws =>
            {
                for (int i = deleted.Count - 1; i >= 0; i--) // last first keeps the other numbers valid
                    if (byRow) ws.Row(deleted[i]).Delete();
                    else ws.Column(deleted[i]).Delete();
            },
            line =>
            {
                int i = deleted.BinarySearch(line);
                return i >= 0 ? null : line - ~i; // gone, or moved up/left by the deleted lines before it (~i)
            });
    }

    /// <summary>
    /// Runs a row/column insert or delete with hyperlinks detached. ClosedXML 0.105 shifts hyperlinks one at a
    /// time in storage order and throws "same key" when a link moves onto one that hasn't moved yet. So detach
    /// them all, make the change, then re-attach each at <paramref name="newLine"/>(its row or column), or drop it
    /// when that's null. Detaching resets font color/underline, so those are restored too.
    /// </summary>
    // ponytail: a link spanning several cells comes back on its top-left cell only.
    static void WithLinksDetached(IXLWorksheet ws, bool byRow, Action<IXLWorksheet> change, Func<int, int?> newLine)
    {
        var links = ws.Hyperlinks
            .Select(h => (Link: h, Cell: h.Cell!)) // links attached to a sheet always have a cell
            .Select(x => (x.Link, x.Cell.Address.RowNumber, x.Cell.Address.ColumnNumber,
                x.Cell.Style.Font.FontColor, x.Cell.Style.Font.Underline))
            .ToList();
        foreach (var l in links) ws.Hyperlinks.Delete(l.Link);

        change(ws);

        foreach (var l in links)
        {
            if (newLine(byRow ? l.RowNumber : l.ColumnNumber) is not { } n) continue;
            var cell = byRow ? ws.Cell(n, l.ColumnNumber) : ws.Cell(l.RowNumber, n);
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

    /// <summary>Adds an empty sheet at <paramref name="index"/> (0-based). Blank name picks "Sheet2", "Sheet3", … Returns the name.</summary>
    public string AddSheet(int index, string? name = null)
    {
        name = NewSheetName(name, i => $"Sheet{i}");
        return Edit(SnapshotStep, () =>
        {
            _wb.AddWorksheet(name).Position = index + 1;
            return name;
        });
    }

    /// <summary>
    /// Copies every sheet of another .xlsx into this workbook, starting at <paramref name="index"/> (0-based).
    /// Names that are taken get " (2)", " (3)", … One undo step. Returns the new sheets' names.
    /// </summary>
    public IReadOnlyList<string> ImportSheets(string path, int index)
    {
        using var other = Load(path, out _);
        return Edit(SnapshotStep, () =>
        {
            var names = new List<string>();
            foreach (var src in other.Worksheets.OrderBy(ws => ws.Position))
            {
                var name = NewSheetName(null, i =>
                {
                    if (i == 1) return src.Name;
                    var suffix = $" ({i})";
                    return src.Name[..Math.Min(src.Name.Length, 31 - suffix.Length)] + suffix;
                });
                src.CopyTo(_wb, name).Position = index + names.Count + 1;
                names.Add(name);
            }
            return (IReadOnlyList<string>)names;
        });
    }

    /// <summary>
    /// Appends the rows of every sheet in another .xlsx below the data of <paramref name="sheet"/>, column A to A.
    /// <paramref name="skipHeader"/> leaves out each source sheet's first row. One undo step.
    /// Returns the first appended row and how many rows were added.
    /// </summary>
    // ponytail: matches columns by position, not by header name.
    public (int FirstRow, int Count) AppendRows(string path, int sheet, bool skipHeader)
    {
        using var other = Load(path, out _);
        return Edit(SnapshotStep, () =>
        {
            var dst = Sheet(sheet);
            int first = UsedSize(sheet).Rows + 1, next = first;
            foreach (var src in other.Worksheets.OrderBy(ws => ws.Position))
            {
                if (src.LastCellUsed() is not { } last) continue;
                int top = skipHeader ? 2 : 1, rows = last.Address.RowNumber - top + 1;
                if (rows <= 0) continue;
                src.Range(top, 1, last.Address.RowNumber, last.Address.ColumnNumber).CopyTo(dst.Cell(next, 1));
                next += rows;
            }
            return (first, next - first);
        });
    }

    /// <summary>Deletes a sheet. A workbook must keep at least one, like Excel.</summary>
    public void DeleteSheet(int sheet)
    {
        if (_wb.Worksheets.Count <= 1) throw new InvalidOperationException("A workbook needs at least one sheet.");
        Edit(SnapshotStep, () => Sheet(sheet).Delete());
    }

    /// <summary>Moves a sheet to <paramref name="index"/> (0-based) in the tab order.</summary>
    public void MoveSheet(int sheet, int index) =>
        Edit(SnapshotStep, () => Sheet(sheet).Position = Math.Clamp(index, 0, _wb.Worksheets.Count - 1) + 1);

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
