using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using XlsxEditor.Core;

namespace XlsxEditor.App.Controls;

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
    public int SelectedRow { get; private set; } = 1;
    public int SelectedCol { get; private set; } = 1;

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
        FirstRow = FirstCol = SelectedRow = SelectedCol = 1;
        Changed();
    }

    public void ScrollTo(int row, int col)
    {
        FirstRow = Math.Clamp(row, 1, MaxRows);
        FirstCol = Math.Clamp(col, 1, MaxCols);
        Changed();
    }

    public void Select(int row, int col)
    {
        SelectedRow = Math.Clamp(row, 1, MaxRows);
        SelectedCol = Math.Clamp(col, 1, MaxCols);
        // Keep the selection on screen.
        if (SelectedRow < FirstRow) FirstRow = SelectedRow;
        else if (SelectedRow >= FirstRow + VisibleRows) FirstRow = SelectedRow - VisibleRows + 1;
        if (SelectedCol < FirstCol) FirstCol = SelectedCol;
        else if (SelectedCol >= FirstCol + VisibleCols) FirstCol = SelectedCol - VisibleCols + 1;
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
            if (r < rows) DrawText(ctx, (FirstRow + r).ToString(), new Rect(0, y, HeaderWidth, RowHeight), HeaderText, center: true);
        }
        for (int c = 0; c <= cols; c++)
        {
            double x = HeaderWidth + ColWidth * c;
            ctx.DrawLine(GridLine, new Point(x, 0), new Point(x, h));
            if (c < cols)
            {
                var letter = WorkbookDocument.ColumnLetter(FirstCol + c);
                DrawText(ctx, letter, new Rect(x, 0, ColWidth, RowHeight), HeaderText, center: true);
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

        ctx.DrawRectangle(null, SelectionPen, CellRect(SelectedRow, SelectedCol));
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

    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        Focus();
        if (HitTest(e.GetPosition(this)) is not { } hit) return;
        Select(hit.Row, hit.Col);
        e.Handled = true;
    }

    /// <summary>The cell under a point, or null over the headers.</summary>
    public (int Row, int Col)? HitTest(Point p) =>
        p.X < HeaderWidth || p.Y < RowHeight ? null
        : (FirstRow + (int)((p.Y - RowHeight) / RowHeight), FirstCol + (int)((p.X - HeaderWidth) / ColWidth));

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
        bool jump = e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta);
        (int r, int c)? target = e.Key switch
        {
            // Ctrl/Cmd+Arrow: jump to the edge of the data, like Excel.
            Key.Up or Key.Down or Key.Left or Key.Right when jump && _doc is not null => _doc.JumpEdge(_sheet, SelectedRow, SelectedCol,
                e.Key == Key.Up ? -1 : e.Key == Key.Down ? 1 : 0, e.Key == Key.Left ? -1 : e.Key == Key.Right ? 1 : 0),
            Key.Up => (SelectedRow - 1, SelectedCol),
            Key.Down or Key.Enter => (SelectedRow + 1, SelectedCol),
            Key.Left => (SelectedRow, SelectedCol - 1),
            Key.Right or Key.Tab => (SelectedRow, SelectedCol + 1),
            Key.PageUp => (SelectedRow - page, SelectedCol),
            Key.PageDown => (SelectedRow + page, SelectedCol),
            Key.Home => e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta) ? (1, 1) : (SelectedRow, 1),
            _ => null,
        };
        if (target is not { } t) return;
        Select(t.r, t.c);
        e.Handled = true;
    }
}
