using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using OpenSheet.Core;

namespace OpenSheet.App.Controls;

public enum GridArea { Cell, ColumnHeader, RowHeader, Corner }

/// <summary>
/// Spreadsheet surface. Draws only the visible cells, so sheet size doesn't matter.
/// Scrolls by whole rows/columns like Excel. Row/column numbers are 1-based.
/// </summary>
public class SheetGrid : Control
{
    // ponytail: uniform sizes; read real column widths / row heights from the file in Phase 3.
    const double RowHeight = 22, ColWidth = 90, HeaderWidth = 50, Pad = 4;
    public const int MaxRows = 1_048_576, MaxCols = 16_384; // Excel limits

    static readonly Typeface Font = new(FontFamily.Default);
    static readonly IBrush HeaderBg = new SolidColorBrush(Color.Parse("#F3F3F3"));
    static readonly IBrush HeaderText = new SolidColorBrush(Color.Parse("#555555"));
    static readonly IPen GridLine = new Pen(new SolidColorBrush(Color.Parse("#E1E1E1")));
    static readonly IPen SelectionPen = new Pen(new SolidColorBrush(Color.Parse("#107C41")), 2);
    static readonly IBrush SelectionBg = new SolidColorBrush(Color.Parse("#107C41"), 0.12);
    static readonly IBrush SelectedHeaderBg = new SolidColorBrush(Color.Parse("#D3E9DB"));
    static readonly IBrush SelectedHeaderText = new SolidColorBrush(Color.Parse("#0B5A2F"));
    static readonly IPen ReferencePen = new Pen(new SolidColorBrush(Color.Parse("#1E6FD9")), 2, new DashStyle([3, 2], 0));
    static readonly IBrush ReferenceBg = new SolidColorBrush(Color.Parse("#1E6FD9"), 0.10);
    static readonly IBrush DuplicateBg = new SolidColorBrush(Color.Parse("#FFC7CE")); // Excel's light red fill

    WorkbookDocument? _doc;
    int _sheet;
    (IReadOnlyList<int> Cols, IReadOnlySet<int> Rows)? _duplicates;

    /// <summary>Cells to highlight as duplicates, or null for none. Cleared when the sheet changes.</summary>
    public (IReadOnlyList<int> Cols, IReadOnlySet<int> Rows)? Duplicates
    {
        get => _duplicates;
        set { _duplicates = value; InvalidateVisual(); }
    }

    public SheetGrid()
    {
        Focusable = true;
        ClipToBounds = true;
    }

    public int FirstRow { get; private set; } = 1;
    public int FirstCol { get; private set; } = 1;
    /// <summary>The active cell (where typing goes). Also one corner of the selected range.</summary>
    public int SelectedRow { get; private set; } = 1;
    public int SelectedCol { get; private set; } = 1;
    /// <summary>The other corner of the selected range (same as the active cell when one cell is selected).</summary>
    public int ExtentRow { get; private set; } = 1;
    public int ExtentCol { get; private set; } = 1;

    /// <summary>The selected range, normalized. Whole columns/rows run to <see cref="MaxRows"/>/<see cref="MaxCols"/>.</summary>
    public (int Top, int Left, int Bottom, int Right) Selection =>
        (Math.Min(SelectedRow, ExtentRow), Math.Min(SelectedCol, ExtentCol), Math.Max(SelectedRow, ExtentRow), Math.Max(SelectedCol, ExtentCol));

    /// <summary>The selection as Excel writes it: "B3", "B3:D9", "B:D" (whole columns), "3:5" (whole rows).</summary>
    public string SelectionText
    {
        get
        {
            var (t, l, b, r) = Selection;
            bool allRows = t == 1 && b == MaxRows, allCols = l == 1 && r == MaxCols;
            if (allRows && !allCols) return $"{WorkbookDocument.ColumnLetter(l)}:{WorkbookDocument.ColumnLetter(r)}";
            if (allCols && !allRows) return $"{t}:{b}";
            return t == b && l == r ? WorkbookDocument.Address(t, l) : $"{WorkbookDocument.Address(t, l)}:{WorkbookDocument.Address(b, r)}";
        }
    }

    (int Top, int Left, int Bottom, int Right)? _reference;

    /// <summary>A range being picked for a formula (blue dashed box, like Excel), or null.</summary>
    public (int Top, int Left, int Bottom, int Right)? Reference
    {
        get => _reference;
        set { _reference = value; InvalidateVisual(); }
    }

    enum Drag { None, Cells, Rows, Cols }
    Drag _drag;

    /// <summary>Raised after scroll, selection or resize.</summary>
    public event Action? ViewChanged;

    public int VisibleRows => Math.Max(1, (int)((Bounds.Height - RowHeight) / RowHeight));
    public int VisibleCols => Math.Max(1, (int)((Bounds.Width - HeaderWidth) / ColWidth));

    public CellView SelectedCell => _doc?.GetCell(_sheet, SelectedRow, SelectedCol) ?? default;

    public void Show(WorkbookDocument? doc, int sheet)
    {
        _doc = doc;
        _sheet = sheet;
        _duplicates = null;
        _reference = null;
        FirstRow = FirstCol = SelectedRow = SelectedCol = ExtentRow = ExtentCol = 1;
        Changed();
    }

    public void ScrollTo(int row, int col)
    {
        FirstRow = Math.Clamp(row, 1, MaxRows);
        FirstCol = Math.Clamp(col, 1, MaxCols);
        Changed();
    }

    /// <summary>Selects one cell, or with <paramref name="extend"/> stretches the range to it (keeping the active cell).</summary>
    public void Select(int row, int col, bool extend = false)
    {
        ExtentRow = Math.Clamp(row, 1, MaxRows);
        ExtentCol = Math.Clamp(col, 1, MaxCols);
        if (!extend) (SelectedRow, SelectedCol) = (ExtentRow, ExtentCol);
        Reveal(ExtentRow, ExtentCol); // keep the moved corner on screen
    }

    /// <summary>Scrolls just enough to show (row, col).</summary>
    public void Reveal(int row, int col)
    {
        row = Math.Clamp(row, 1, MaxRows);
        col = Math.Clamp(col, 1, MaxCols);
        if (row < FirstRow) FirstRow = row;
        else if (row >= FirstRow + VisibleRows) FirstRow = row - VisibleRows + 1;
        if (col < FirstCol) FirstCol = col;
        else if (col >= FirstCol + VisibleCols) FirstCol = col - VisibleCols + 1;
        Changed();
    }

    /// <summary>Selects whole columns <paramref name="from"/>..<paramref name="to"/>; the active cell goes to the top of <paramref name="from"/>.</summary>
    public void SelectColumns(int from, int to)
    {
        (SelectedRow, SelectedCol, ExtentRow, ExtentCol) = (1, from, MaxRows, to);
        Changed();
    }

    public void SelectRows(int from, int to)
    {
        (SelectedRow, SelectedCol, ExtentRow, ExtentCol) = (from, 1, to, MaxCols);
        Changed();
    }

    public void SelectAll()
    {
        (SelectedRow, SelectedCol, ExtentRow, ExtentCol) = (1, 1, MaxRows, MaxCols);
        Changed();
    }

    void Changed()
    {
        InvalidateVisual();
        ViewChanged?.Invoke();
    }

    protected override void OnSizeChanged(SizeChangedEventArgs e)
    {
        base.OnSizeChanged(e);
        ViewChanged?.Invoke();
    }

    public override void Render(DrawingContext ctx)
    {
        var (w, h) = (Bounds.Width, Bounds.Height);
        int rows = (int)Math.Ceiling((h - RowHeight) / RowHeight);
        int cols = (int)Math.Ceiling((w - HeaderWidth) / ColWidth);

        ctx.FillRectangle(Brushes.White, new Rect(0, 0, w, h));
        ctx.FillRectangle(HeaderBg, new Rect(0, 0, w, RowHeight));
        ctx.FillRectangle(HeaderBg, new Rect(0, 0, HeaderWidth, h));

        var (top, left, bottom, right) = Selection;
        // Highlight headers of selected rows/columns, like Excel.
        for (int c = Math.Max(left, FirstCol); c <= Math.Min(right, FirstCol + cols); c++)
            ctx.FillRectangle(SelectedHeaderBg, new Rect(CellRect(1, c).X, 0, ColWidth, RowHeight));
        for (int r = Math.Max(top, FirstRow); r <= Math.Min(bottom, FirstRow + rows); r++)
            ctx.FillRectangle(SelectedHeaderBg, new Rect(0, CellRect(r, 1).Y, HeaderWidth, RowHeight));

        if (_duplicates is { } dup)
            for (int r = 0; r < rows; r++)
                if (dup.Rows.Contains(FirstRow + r))
                    foreach (var col in dup.Cols)
                        if (col >= FirstCol && col < FirstCol + cols)
                            ctx.FillRectangle(DuplicateBg, CellRect(FirstRow + r, col));

        for (int r = 0; r <= rows; r++)
        {
            double y = RowHeight * (r + 1);
            ctx.DrawLine(GridLine, new Point(0, y), new Point(w, y));
            bool sel = FirstRow + r >= top && FirstRow + r <= bottom;
            if (r < rows) DrawText(ctx, (FirstRow + r).ToString(), new Rect(0, y, HeaderWidth, RowHeight), sel ? SelectedHeaderText : HeaderText, center: true);
        }
        for (int c = 0; c <= cols; c++)
        {
            double x = HeaderWidth + ColWidth * c;
            ctx.DrawLine(GridLine, new Point(x, 0), new Point(x, h));
            if (c < cols)
            {
                var letter = WorkbookDocument.ColumnLetter(FirstCol + c);
                bool sel = FirstCol + c >= left && FirstCol + c <= right;
                DrawText(ctx, letter, new Rect(x, 0, ColWidth, RowHeight), sel ? SelectedHeaderText : HeaderText, center: true);
            }
        }

        if (_doc is not null)
        {
            for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                var cell = _doc.GetCell(_sheet, FirstRow + r, FirstCol + c);
                if (cell.Text is not { Length: > 0 }) continue;
                // ponytail: text is clipped to its cell; Excel-style overflow into empty neighbours comes later.
                DrawText(ctx, cell.Text, CellRect(FirstRow + r, FirstCol + c), Brushes.Black, rightAlign: cell.IsNumber);
            }
        }

        // Ranges: tinted fill + border; clamped to just past the screen so huge ranges stay cheap to draw.
        var area = new Rect(0, RowHeight, w, h - RowHeight).Inflate(4);
        Rect RangeRect(int t, int l, int b, int r) =>
            CellRect(Math.Max(t, FirstRow - 1), Math.Max(l, FirstCol - 1))
                .Union(CellRect(Math.Min(b, FirstRow + rows + 1), Math.Min(r, FirstCol + cols + 1)));
        var range = RangeRect(top, left, bottom, right);
        using (ctx.PushClip(new Rect(HeaderWidth, RowHeight, Math.Max(0, w - HeaderWidth), Math.Max(0, h - RowHeight))))
        {
            if (top != bottom || left != right)
            {
                ctx.FillRectangle(SelectionBg, range.Intersect(area));
                ctx.DrawRectangle(null, SelectionPen, range);
                ctx.DrawRectangle(null, GridLine, CellRect(SelectedRow, SelectedCol).Deflate(1)); // active cell
            }
            else ctx.DrawRectangle(null, SelectionPen, range);

            if (_reference is { } f)
            {
                var refRect = RangeRect(f.Top, f.Left, f.Bottom, f.Right);
                ctx.FillRectangle(ReferenceBg, refRect.Intersect(area));
                ctx.DrawRectangle(null, ReferencePen, refRect.Deflate(1));
            }
        }
    }

    /// <summary>Where a cell is drawn, in this control's coordinates (may be off-screen).</summary>
    public Rect CellRect(int row, int col) =>
        new(HeaderWidth + (col - FirstCol) * ColWidth, RowHeight * (row - FirstRow + 1), ColWidth, RowHeight);

    static void DrawText(DrawingContext ctx, string text, Rect cell, IBrush brush, bool center = false, bool rightAlign = false)
    {
        var ft = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, Font, 13, brush);
        double x = center ? cell.X + (cell.Width - ft.Width) / 2
                 : rightAlign ? cell.Right - Pad - ft.Width
                 : cell.X + Pad;
        using (ctx.PushClip(cell.Deflate(1)))
            ctx.DrawText(ft, new Point(x, cell.Y + (cell.Height - ft.Height) / 2));
    }

    public GridArea AreaAt(Point p) => (p.Y < RowHeight, p.X < HeaderWidth) switch
    {
        (true, true) => GridArea.Corner,
        (true, false) => GridArea.ColumnHeader,
        (false, true) => GridArea.RowHeader,
        _ => GridArea.Cell,
    };

    /// <summary>Whether a right-click here should keep the current selection (it's inside it), like Excel.</summary>
    bool InSelection(Point p)
    {
        var (t, l, b, r) = Selection;
        return AreaAt(p) switch
        {
            GridArea.ColumnHeader => t == 1 && b == MaxRows && ColAt(p.X) >= l && ColAt(p.X) <= r,
            GridArea.RowHeader => l == 1 && r == MaxCols && RowAt(p.Y) >= t && RowAt(p.Y) <= b,
            GridArea.Cell => RowAt(p.Y) >= t && RowAt(p.Y) <= b && ColAt(p.X) >= l && ColAt(p.X) <= r,
            _ => true,
        };
    }

    // Click a cell / column header / row header / corner; drag or Shift+click to extend.
    // Right-click selects what's under the pointer (unless it's already selected) for the context menu.
    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        Focus();
        var props = e.GetCurrentPoint(this).Properties;
        bool right = props.IsRightButtonPressed;
        if (!props.IsLeftButtonPressed && !right) return;
        var p = e.GetPosition(this);
        if (right && InSelection(p)) return; // let ContextRequested fire with the selection intact
        bool shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift) && !right;
        bool inColHeader = p.Y < RowHeight, inRowHeader = p.X < HeaderWidth;
        if (inColHeader && inRowHeader)
        {
            SelectAll();
            _drag = Drag.None;
        }
        else if (inColHeader)
        {
            int c = ColAt(p.X);
            SelectColumns(shift ? SelectedCol : c, c);
            _drag = Drag.Cols;
        }
        else if (inRowHeader)
        {
            int r = RowAt(p.Y);
            SelectRows(shift ? SelectedRow : r, r);
            _drag = Drag.Rows;
        }
        else
        {
            Select(RowAt(p.Y), ColAt(p.X), extend: shift);
            _drag = Drag.Cells;
        }
        if (right)
        {
            _drag = Drag.None;
            return; // not handled, so ContextRequested still fires
        }
        e.Pointer.Capture(this); // keep getting moves when dragging past the edge
        e.Handled = true;
    }

    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (_drag == Drag.None) return;
        var p = e.GetPosition(this);
        // Past an edge, step one beyond the last visible line so Select's keep-on-screen scrolls.
        int row = p.Y >= Bounds.Height ? FirstRow + VisibleRows : RowAt(p.Y);
        int col = p.X >= Bounds.Width ? FirstCol + VisibleCols : ColAt(p.X);
        switch (_drag)
        {
            case Drag.Cells: if ((row, col) != (ExtentRow, ExtentCol)) Select(row, col, extend: true); break;
            case Drag.Cols: if (col != ExtentCol) SelectColumns(SelectedCol, col); break;
            case Drag.Rows: if (row != ExtentRow) SelectRows(SelectedRow, row); break;
        }
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        _drag = Drag.None;
        e.Pointer.Capture(null);
    }

    // Row/column under a point; above/left of the cells it reports the line just off-screen so dragging scrolls back.
    int RowAt(double y) => Math.Clamp(FirstRow + (int)Math.Floor((y - RowHeight) / RowHeight), 1, MaxRows);
    int ColAt(double x) => Math.Clamp(FirstCol + (int)Math.Floor((x - HeaderWidth) / ColWidth), 1, MaxCols);

    /// <summary>The row and column under a point, wherever it is (headers and past the edges included).</summary>
    public (int Row, int Col) CellAt(Point p) => (RowAt(p.Y), ColAt(p.X));

    /// <summary>The cell under a point, or null over the headers.</summary>
    public (int Row, int Col)? HitTest(Point p) =>
        p.X < HeaderWidth || p.Y < RowHeight ? null : (RowAt(p.Y), ColAt(p.X));

    protected override void OnPointerWheelChanged(PointerWheelEventArgs e)
    {
        base.OnPointerWheelChanged(e);
        // Shift+wheel scrolls sideways on Windows; macOS trackpads send Delta.X directly.
        var (dx, dy) = e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? (e.Delta.Y, 0.0) : (e.Delta.X, e.Delta.Y);
        ScrollTo(FirstRow - (int)Math.Round(dy * 3), FirstCol - (int)Math.Round(dx * 3));
        e.Handled = true;
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        int page = VisibleRows;
        bool cmd = e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta);
        bool shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        if (cmd && e.Key == Key.A)
        {
            SelectAll();
            e.Handled = true;
            return;
        }
        // Shift+arrows move the range's far corner; plain arrows move the active cell.
        var (row, col) = shift ? (ExtentRow, ExtentCol) : (SelectedRow, SelectedCol);
        (int r, int c)? target = e.Key switch
        {
            // Ctrl/Cmd+Arrow: jump to the edge of the data, like Excel.
            Key.Up or Key.Down or Key.Left or Key.Right when cmd && _doc is not null => _doc.JumpEdge(_sheet, row, col,
                e.Key == Key.Up ? -1 : e.Key == Key.Down ? 1 : 0, e.Key == Key.Left ? -1 : e.Key == Key.Right ? 1 : 0),
            Key.PageUp or Key.PageDown when cmd => null, // Ctrl/Cmd+PageUp/Down switch sheets (MainWindow)
            Key.Up => (row - 1, col),
            Key.Down => (row + 1, col),
            Key.Left => (row, col - 1),
            Key.Right => (row, col + 1),
            Key.PageUp => (row - page, col),
            Key.PageDown => (row + page, col),
            Key.Home => cmd ? (1, 1) : (row, 1),
            _ => null,
        };
        if (target is { } t) Select(t.r, t.c, extend: shift);
        else if (e.Key is Key.Enter or Key.Tab) // move the active cell; Shift goes back
            Select(SelectedRow + (e.Key == Key.Enter ? (shift ? -1 : 1) : 0), SelectedCol + (e.Key == Key.Tab ? (shift ? -1 : 1) : 0));
        else return;
        e.Handled = true;
    }
}
