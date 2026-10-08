using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using OpenSheet.App.Controls;
using OpenSheet.Core;

namespace OpenSheet.App;

public partial class MainWindow : Window
{
    WorkbookDocument? _doc;
    (int Rows, int Cols) _used;
    bool _syncing; // guards scrollbar <-> grid feedback
    bool _dirty;   // edited since open/save
    DuplicatesWindow? _dupWindow;
    (int Sheet, int Row, int Col, string Original)? _editing; // cell open in the in-cell editor
    (int Start, int End)? _pointRef; // reference just inserted into a formula by clicking/dragging on the sheet
    (GridArea Area, int Row, int Col)? _pointAnchor; // where that click/drag started, while the button is down
    // What we last put on the clipboard, so pasting it back can keep formulas (the clipboard only holds text).
    (string Text, int Sheet, int Top, int Left, int Bottom, int Right, bool Cut)? _clip;
    Func<string, string?>? _prompt; // PromptBox's Enter action: returns an error to show, or null when done

    static readonly FilePickerFileType XlsxType = new("Excel Workbook")
    {
        Patterns = ["*.xlsx"],
        AppleUniformTypeIdentifiers = ["org.openxmlformats.spreadsheetml.sheet"],
    };
    static readonly FilePickerFileType CsvType = new("CSV (comma-separated values)")
    {
        Patterns = ["*.csv"],
        AppleUniformTypeIdentifiers = ["public.comma-separated-values-text"],
    };
    static readonly FilePickerFileType SpreadsheetTypes = new("Spreadsheets (.xlsx, .csv)")
    {
        Patterns = ["*.xlsx", "*.csv"],
        AppleUniformTypeIdentifiers = ["org.openxmlformats.spreadsheetml.sheet", "public.comma-separated-values-text"],
    };

    internal WorkbookDocument? Doc => _doc;
    internal int CurrentSheet => SheetTabs.SelectedIndex;
    /// <summary>Bumped on every edit so other windows can tell their results are stale.</summary>
    internal int Edits { get; private set; }
    internal bool HasUnsavedChanges => _dirty;
    bool _closeConfirmed;

    public MainWindow()
    {
        InitializeComponent();
        Sheet.ViewChanged += SyncView;
        VScroll.ValueChanged += (_, e) => { if (!_syncing) Sheet.ScrollTo((int)Math.Round(e.NewValue), Sheet.FirstCol); };
        HScroll.ValueChanged += (_, e) => { if (!_syncing) Sheet.ScrollTo(Sheet.FirstRow, (int)Math.Round(e.NewValue)); };
        AddHandler(DragDrop.DropEvent, OnDrop);
        // Tunnel so Enter/Esc/Tab reach us before the TextBox eats them.
        SearchBox.AddHandler(KeyDownEvent, OnSearchKeyDown, RoutingStrategies.Tunnel);
        PromptBox.AddHandler(KeyDownEvent, OnPromptKeyDown, RoutingStrategies.Tunnel);
        CellEditor.AddHandler(KeyDownEvent, OnEditorKeyDown, RoutingStrategies.Tunnel);
        CellEditor.LostFocus += (_, _) => { if (_pointAnchor is null) CommitEdit(); }; // clicking elsewhere saves, like Excel
        CellEditor.TextChanged += (_, _) => { if (_editing is not null) FormulaBox.Text = CellEditor.Text; };
        Sheet.KeyDown += OnSheetKeyDown;
        Sheet.AddHandler(TextInputEvent, OnSheetTextInput);
        Sheet.AddHandler(PointerPressedEvent, OnSheetPointerPressed, RoutingStrategies.Tunnel);
        Sheet.AddHandler(PointerMovedEvent, OnSheetPointerMoved, RoutingStrategies.Tunnel);
        Sheet.AddHandler(PointerReleasedEvent, OnSheetPointerReleased, RoutingStrategies.Tunnel);
        Sheet.DoubleTapped += (_, e) => { if (Sheet.HitTest(e.GetPosition(Sheet)) is not null) BeginEdit(null); }; // not on headers
        SheetTabs.DoubleTapped += (_, _) => RenameCurrentSheet();
        Sheet.ContextRequested += OnSheetContextRequested;
        SheetTabs.ContextRequested += OnTabsContextRequested;
        // A plain mouse wheel only scrolls vertically; turn it sideways for the tab row.
        TabScroller.AddHandler(PointerWheelChangedEvent, (_, e) =>
        {
            if (e.Delta.Y == 0 || e.Delta.X != 0) return; // trackpads already send sideways deltas
            ScrollTabs(-e.Delta.Y * 60);
            e.Handled = true;
        }, RoutingStrategies.Tunnel);
        SheetTabs.AddHandler(PointerPressedEvent, OnTabPointerPressed, RoutingStrategies.Tunnel);
        SheetTabs.AddHandler(PointerMovedEvent, OnTabPointerMoved, RoutingStrategies.Tunnel);
        SheetTabs.AddHandler(PointerReleasedEvent, OnTabPointerReleased, RoutingStrategies.Tunnel, handledEventsToo: true);

        // Menus are written with Ctrl; macOS users expect Cmd. (OnKeyDown accepts either.)
        if (OperatingSystem.IsMacOS())
            foreach (var item in MainMenu.GetLogicalDescendants().OfType<MenuItem>())
                if (item.InputGesture is { } g && g.KeyModifiers.HasFlag(KeyModifiers.Control))
                    item.InputGesture = new KeyGesture(g.Key, (g.KeyModifiers & ~KeyModifiers.Control) | KeyModifiers.Meta);
    }

    // ---- Cell editing (in-cell editor) ----

    /// <summary>Opens the editor on the selected cell; <paramref name="text"/> replaces its content (null keeps it).</summary>
    void BeginEdit(string? text)
    {
        if (_doc is null || CurrentSheet < 0) return;
        var original = Sheet.SelectedCell.Content ?? "";
        _editing = (CurrentSheet, Sheet.SelectedRow, Sheet.SelectedCol, original);
        EndPointing();
        CellEditor.Text = text ?? original;
        PlaceEditor();
        CellEditor.IsVisible = true;
        CellEditor.Focus();
        CellEditor.CaretIndex = CellEditor.Text.Length;
    }

    void PlaceEditor()
    {
        if (_editing is not { } t) return;
        var r = Sheet.CellRect(t.Row, t.Col);
        Canvas.SetLeft(CellEditor, r.X);
        Canvas.SetTop(CellEditor, r.Y);
        CellEditor.MinWidth = r.Width + 1; // grows with long text
        CellEditor.Height = r.Height + 1;
    }

    void CommitEdit(int moveRows = 0, int moveCols = 0)
    {
        if (_editing is not { } t || _doc is null) return;
        _editing = null;
        EndPointing();
        CellEditor.IsVisible = false;
        var text = CellEditor.Text ?? "";
        // Unchanged = don't touch it: a date shows as text and would come back as text.
        if (text != t.Original)
        {
            try
            {
                _doc.SetCell(t.Sheet, t.Row, t.Col, text);
                MarkEdited();
            }
            catch (Exception ex) { StatusText.Text = $"Cannot set {WorkbookDocument.Address(t.Row, t.Col)}: {ex.Message}"; }
        }
        if (moveRows != 0 || moveCols != 0) Sheet.Select(t.Row + moveRows, t.Col + moveCols);
        Sheet.Focus();
        SyncView();
    }

    void CancelEdit()
    {
        _editing = null;
        EndPointing();
        CellEditor.IsVisible = false;
        Sheet.Focus();
        SyncView();
    }

    void OnEditorKeyDown(object? sender, KeyEventArgs e)
    {
        bool shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        switch (e.Key)
        {
            case Key.Enter: CommitEdit(moveRows: shift ? -1 : 1); break;
            case Key.Tab: CommitEdit(moveCols: shift ? -1 : 1); break;
            case Key.Escape: CancelEdit(); break;
            default: EndPointing(); return; // typing ends "click to insert a reference"
        }
        e.Handled = true;
    }

    // While typing a formula, click or drag on the sheet to insert a reference, like Excel:
    // a cell (A1), a range (A1:C5), column headers (B:B, B:D) or row headers (3:3, 3:5).
    // Clicking again right away replaces the reference just inserted.
    void OnSheetPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (_editing is null || !e.GetCurrentPoint(Sheet).Properties.IsLeftButtonPressed) return;
        var pos = e.GetPosition(Sheet);
        var area = Sheet.AreaAt(pos);
        if (area == GridArea.Corner) return;
        var text = CellEditor.Text ?? "";
        int caret = CellEditor.CaretIndex;
        bool replacing = _pointRef is { } p && p.End == caret;
        bool afterOperator = caret > 0 && "=+-*/^(,:;<>&".Contains(text[caret - 1]);
        if (!text.StartsWith('=') || !(replacing || afterOperator)) return; // normal click: commit and select

        if (!replacing) _pointRef = (caret, caret);
        var (row, col) = Sheet.CellAt(pos);
        _pointAnchor = (area, row, col);
        UpdatePointRef(row, col);
        e.Pointer.Capture(Sheet); // keep getting moves while dragging
        e.Handled = true;         // the grid must not select or take focus
    }

    void OnSheetPointerMoved(object? sender, PointerEventArgs e)
    {
        if (_pointAnchor is null) return;
        var (row, col) = Sheet.CellAt(e.GetPosition(Sheet));
        // Dragging past an edge scrolls, only along the axis being picked.
        Sheet.Reveal(_pointAnchor.Value.Area == GridArea.ColumnHeader ? Sheet.FirstRow : row,
                     _pointAnchor.Value.Area == GridArea.RowHeader ? Sheet.FirstCol : col);
        UpdatePointRef(row, col);
        e.Handled = true;
    }

    void OnSheetPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        if (_pointAnchor is null) return;
        _pointAnchor = null;
        e.Pointer.Capture(null);
        if (!CellEditor.IsFocused) CellEditor.Focus(); // keep typing the formula
        CellEditor.CaretIndex = _pointRef?.End ?? CellEditor.CaretIndex;
        e.Handled = true;
    }

    /// <summary>Rewrites the reference being picked so it spans from the anchor to (row, col).</summary>
    void UpdatePointRef(int row, int col)
    {
        if (_pointAnchor is not { } a || _pointRef is not { } span) return;
        int top = Math.Min(a.Row, row), bottom = Math.Max(a.Row, row);
        int left = Math.Min(a.Col, col), right = Math.Max(a.Col, col);
        string Col(int c) => WorkbookDocument.ColumnLetter(c);
        var (reference, box) = a.Area switch
        {
            GridArea.ColumnHeader => ($"{Col(left)}:{Col(right)}", (1, left, SheetGrid.MaxRows, right)),
            GridArea.RowHeader => ($"{top}:{bottom}", (top, 1, bottom, SheetGrid.MaxCols)),
            _ => (top == bottom && left == right
                    ? WorkbookDocument.Address(top, left)
                    : $"{WorkbookDocument.Address(top, left)}:{WorkbookDocument.Address(bottom, right)}",
                (top, left, bottom, right)),
        };
        var text = CellEditor.Text ?? "";
        CellEditor.Text = text[..span.Start] + reference + text[span.End..];
        CellEditor.CaretIndex = span.Start + reference.Length;
        _pointRef = (span.Start, span.Start + reference.Length);
        Sheet.Reference = box;
    }

    void EndPointing()
    {
        _pointRef = null;
        Sheet.Reference = null;
    }

    void OnSheetKeyDown(object? sender, KeyEventArgs e)
    {
        if (_doc is null) return;
        if (e.Key == Key.F2) BeginEdit(null);
        else if (e.Key is Key.Delete or Key.Back) ClearSelected();
        else return;
        e.Handled = true;
    }

    // Typing on a cell starts editing it with that character, like Excel.
    void OnSheetTextInput(object? sender, TextInputEventArgs e)
    {
        if (_doc is null || e.Text is not { Length: > 0 } text || char.IsControl(text[0])) return;
        BeginEdit(text);
        e.Handled = true;
    }

    void ClearSelected()
    {
        if (_doc is null || CurrentSheet < 0) return;
        var (top, left, bottom, right) = Sheet.Selection;
        _doc.ClearRange(CurrentSheet, top, left, bottom, right);
        MarkEdited();
    }

    /// <summary>Selected columns as "B" or "B:D", for prefilling column inputs.</summary>
    string SelectedColumnsText()
    {
        var (_, left, _, right) = Sheet.Selection;
        return left == right ? WorkbookDocument.ColumnLetter(left) : $"{WorkbookDocument.ColumnLetter(left)}:{WorkbookDocument.ColumnLetter(right)}";
    }

    // ---- Undo / clipboard ----

    void UndoRedo(bool redo)
    {
        if (_doc is null) return;
        var before = _doc.SheetNames;
        if (!(redo ? _doc.Redo() : _doc.Undo()))
        {
            StatusText.Text = redo ? "Nothing to redo" : "Nothing to undo";
            return;
        }
        if (before.SequenceEqual(_doc.SheetNames)) MarkEdited(); // stay where we are
        else Reload(Math.Min(CurrentSheet, _doc.SheetNames.Count - 1));
        StatusText.Text = redo ? "Redo" : "Undo";
    }

    async Task CopyAsync(bool cut)
    {
        if (_doc is null || CurrentSheet < 0 || Clipboard is null) return;
        var (top, left, bottom, right) = Sheet.Selection;
        var text = _doc.RangeText(CurrentSheet, top, left, bottom, right);
        await Clipboard.SetTextAsync(text);
        // Like Excel: copy pastes with references shifted; cut moves the cells when you paste.
        _clip = (text, CurrentSheet, top, left, bottom, right, cut);
        StatusText.Text = cut ? $"Cut {Sheet.SelectionText} — paste to move it" : $"Copied {Sheet.SelectionText}";
    }

    async Task PasteAsync()
    {
        if (_doc is null || CurrentSheet < 0 || Clipboard is null) return;
        if (await Clipboard.TryGetTextAsync() is not { } text) return;
        int sheet = CurrentSheet;
        var (row, col, _, _) = Sheet.Selection; // paste at the top-left of the selection
        try
        {
            int rows, cols;
            if (_clip is { } c && c.Text == text && c.Sheet < _doc.SheetNames.Count)
            {
                (rows, cols) = _doc.CopyRange(c.Sheet, c.Top, c.Left, c.Bottom, c.Right, sheet, row, col, move: c.Cut);
                if (c.Cut) _clip = null; // a cut pastes once
            }
            else
            {
                // From another app (Excel, Sheets, a text editor): rows by line, cells by tab.
                var lines = text.Replace("\r\n", "\n").Split('\n');
                if (lines.Length > 1 && lines[^1].Length == 0) lines = lines[..^1];
                var block = lines.Select(l => (IReadOnlyList<string>)l.Split('\t')).ToList();
                _doc.SetCells(sheet, row, col, block);
                (rows, cols) = (block.Count, block.Max(l => l.Count));
            }
            MarkEdited();
            if (rows > 0)
            {
                Sheet.Select(row, col);
                Sheet.Select(row + rows - 1, col + cols - 1, extend: true); // select what was pasted
            }
            StatusText.Text = "Pasted";
        }
        catch (Exception ex) { StatusText.Text = $"Cannot paste: {ex.Message}"; }
    }

    void OnUndoClick(object? sender, RoutedEventArgs e) => UndoRedo(redo: false);
    void OnRedoClick(object? sender, RoutedEventArgs e) => UndoRedo(redo: true);
    async void OnCutClick(object? sender, RoutedEventArgs e) => await CopyAsync(cut: true);
    async void OnCopyClick(object? sender, RoutedEventArgs e) => await CopyAsync(cut: false);
    async void OnPasteClick(object? sender, RoutedEventArgs e) => await PasteAsync();
    void OnClearClick(object? sender, RoutedEventArgs e) => ClearSelected();
    void OnEditCellClick(object? sender, RoutedEventArgs e) => BeginEdit(null);
    void OnSelectAllClick(object? sender, RoutedEventArgs e) => Sheet.SelectAll();
    void OnNextSheetClick(object? sender, RoutedEventArgs e) => StepSheet(1);
    void OnPrevSheetClick(object? sender, RoutedEventArgs e) => StepSheet(-1);

    void StepSheet(int delta)
    {
        if (_doc is null) return;
        SheetTabs.SelectedIndex = Math.Clamp(CurrentSheet + delta, 0, _doc.SheetNames.Count - 1);
    }

    // ---- Sheet commands (share the PromptBox) ----

    void ShowPrompt(string placeholder, string text, Func<string, string?> onEnter)
    {
        _prompt = onEnter;
        PromptBox.PlaceholderText = placeholder;
        PromptBox.Text = text;
        PromptBox.IsVisible = true;
        PromptBox.Focus();
        PromptBox.SelectAll();
    }

    void HidePrompt()
    {
        PromptBox.IsVisible = false;
        _prompt = null;
        Sheet.Focus();
    }

    void OnPromptKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape) HidePrompt();
        else if (e.Key == Key.Enter)
        {
            string? error;
            try { error = _prompt?.Invoke(PromptBox.Text ?? ""); }
            catch (Exception ex) { error = ex.Message; }
            if (error is null) HidePrompt();
            else StatusText.Text = error;
        }
        else return;
        e.Handled = true;
    }

    void OnRenameSheetClick(object? sender, RoutedEventArgs e) => RenameCurrentSheet();

    void RenameCurrentSheet()
    {
        if (_doc is not { } doc || CurrentSheet < 0) return;
        int sheet = CurrentSheet;
        ShowPrompt("New sheet name  (Enter = rename, Esc = cancel)", doc.SheetNames[sheet], text =>
        {
            if (doc.CheckNewSheetName(text.Trim(), except: sheet) is { } error) return error;
            var name = doc.RenameSheet(sheet, text);
            Reload(sheet);
            StatusText.Text = $"Renamed sheet to \"{name}\"";
            return null;
        });
    }

    void OnDuplicateSheetClick(object? sender, RoutedEventArgs e)
    {
        if (_doc is not { } doc || CurrentSheet < 0) return;
        int sheet = CurrentSheet;
        ShowPrompt($"Name for the copy of \"{doc.SheetNames[sheet]}\"  (blank = automatic, Enter = copy, Esc = cancel)", "", text =>
        {
            if (text.Trim().Length > 0 && doc.CheckNewSheetName(text.Trim()) is { } error) return error;
            var name = doc.CopySheet(sheet, text);
            Reload(sheet + 1); // show the copy, like Excel
            StatusText.Text = $"Copied to sheet \"{name}\"";
            return null;
        });
    }

    // ---- Rows, columns, sheets (Sheet menu + right-click menus) ----

    /// <summary>Runs an edit, showing failures in the status bar instead of crashing.</summary>
    void Try(Action action)
    {
        if (_doc is null || CurrentSheet < 0) return;
        CommitEdit();
        try { action(); }
        catch (Exception ex) { StatusText.Text = ex.Message; }
    }

    void InsertRows() => Try(() =>
    {
        var (top, _, bottom, _) = Sheet.Selection;
        if (bottom - top + 1 >= SheetGrid.MaxRows) throw new InvalidOperationException("Select fewer rows to insert.");
        _doc!.InsertRows(CurrentSheet, top, bottom - top + 1); // as many as are selected, like Excel
        MarkEdited();
    });

    void InsertColumns() => Try(() =>
    {
        var (_, left, _, right) = Sheet.Selection;
        if (right - left + 1 >= SheetGrid.MaxCols) throw new InvalidOperationException("Select fewer columns to insert.");
        _doc!.InsertColumns(CurrentSheet, left, right - left + 1);
        MarkEdited();
    });

    void DeleteRows() => Try(() =>
    {
        var (top, left, bottom, _) = Sheet.Selection;
        bottom = Math.Min(bottom, _doc!.UsedSize(CurrentSheet).Rows); // rows past the data are already empty
        if (bottom < top) return;
        _doc.DeleteRows(CurrentSheet, Enumerable.Range(top, bottom - top + 1));
        Sheet.Select(top, left);
        MarkEdited();
        StatusText.Text = $"Deleted {bottom - top + 1} rows";
    });

    void DeleteColumns() => Try(() =>
    {
        var (top, left, _, right) = Sheet.Selection;
        right = Math.Min(right, _doc!.UsedSize(CurrentSheet).Cols);
        if (right < left) return;
        _doc.DeleteColumns(CurrentSheet, Enumerable.Range(left, right - left + 1));
        Sheet.Select(top, left);
        MarkEdited();
        StatusText.Text = $"Deleted {right - left + 1} columns";
    });

    void NewSheet() => Try(() =>
    {
        int index = CurrentSheet + 1;
        var name = _doc!.AddSheet(index);
        Reload(index);
        StatusText.Text = $"Added sheet \"{name}\"";
    });

    void DeleteSheet() => Try(() =>
    {
        var name = _doc!.SheetNames[CurrentSheet];
        int index = CurrentSheet;
        _doc.DeleteSheet(index);
        Reload(Math.Min(index, _doc.SheetNames.Count - 1));
        StatusText.Text = $"Deleted sheet \"{name}\" (Cmd/Ctrl+Z to undo)";
    });

    void MoveSheet(int delta) => MoveSheet(CurrentSheet, CurrentSheet + delta);

    void MoveSheet(int from, int to) => Try(() =>
    {
        to = Math.Clamp(to, 0, _doc!.SheetNames.Count - 1);
        if (to == from) return;
        _doc.MoveSheet(from, to);
        Reload(to);
        StatusText.Text = "";
    });

    // ---- Drag a sheet tab to reorder ----

    (int From, Point Start, bool Dragging)? _tabDrag;

    int TabIndexAt(Visual? v) =>
        v?.FindAncestorOfType<TabStripItem>(includeSelf: true) is { } tab ? SheetTabs.IndexFromContainer(tab) : -1;

    void OnTabPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (!e.GetCurrentPoint(SheetTabs).Properties.IsLeftButtonPressed) return;
        int index = TabIndexAt(e.Source as Visual);
        _tabDrag = index >= 0 ? (index, e.GetPosition(SheetTabs), false) : null; // tab still gets selected as usual
    }

    void OnTabPointerMoved(object? sender, PointerEventArgs e)
    {
        if (_tabDrag is not { } d || d.Dragging) return;
        var delta = e.GetPosition(SheetTabs) - d.Start;
        if (Math.Abs(delta.X) < 6 && Math.Abs(delta.Y) < 6) return; // a click, not a drag (yet)
        _tabDrag = d with { Dragging = true };
        e.Pointer.Capture(SheetTabs);
        SheetTabs.Cursor = new Cursor(StandardCursorType.DragMove);
        StatusText.Text = $"Moving \"{_doc?.SheetNames[d.From]}\" — drop it on another tab";
    }

    void OnTabPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        if (_tabDrag is not { } d) return;
        _tabDrag = null;
        if (!d.Dragging) return;
        e.Pointer.Capture(null);
        SheetTabs.Cursor = null;
        StatusText.Text = "";
        // Dropped on a tab: take its place. Past the last tab: go to the end.
        var pos = e.GetPosition(SheetTabs);
        int to = TabIndexAt(SheetTabs.InputHitTest(pos) as Visual);
        if (to < 0 && pos.X > SheetTabs.Bounds.Width) to = SheetTabs.ItemCount - 1;
        if (to >= 0) MoveSheet(d.From, to);
        e.Handled = true;
    }

    void OnNewSheetClick(object? sender, RoutedEventArgs e) => NewSheet();
    void OnDeleteSheetClick(object? sender, RoutedEventArgs e) => DeleteSheet();
    void OnMoveSheetLeftClick(object? sender, RoutedEventArgs e) => MoveSheet(-1);
    void OnMoveSheetRightClick(object? sender, RoutedEventArgs e) => MoveSheet(1);
    void OnInsertRowsClick(object? sender, RoutedEventArgs e) => InsertRows();
    void OnInsertColumnsClick(object? sender, RoutedEventArgs e) => InsertColumns();
    void OnDeleteRowsClick(object? sender, RoutedEventArgs e) => DeleteRows();
    void OnDeleteColumnsClick(object? sender, RoutedEventArgs e) => DeleteColumns();

    static ContextMenu Menu(params (string Header, Action Run)?[] items)
    {
        var menu = new ContextMenu();
        foreach (var item in items)
        {
            if (item is not { } i)
            {
                menu.Items.Add(new Separator());
                continue;
            }
            var mi = new MenuItem { Header = i.Header };
            mi.Click += (_, _) => i.Run();
            menu.Items.Add(mi);
        }
        return menu;
    }

    void OnSheetContextRequested(object? sender, ContextRequestedEventArgs e)
    {
        if (_doc is null) return;
        CommitEdit();
        var area = e.TryGetPosition(Sheet, out var p) ? Sheet.AreaAt(p) : GridArea.Cell;
        (string, Action)? cut = ("Cut", () => _ = CopyAsync(cut: true)),
            copy = ("Copy", () => _ = CopyAsync(cut: false)),
            paste = ("Paste", () => _ = PasteAsync()),
            clear = ("Clear Contents", ClearSelected),
            insertRows = ("Insert Rows Above", InsertRows),
            insertCols = ("Insert Columns Left", InsertColumns),
            deleteRows = ("Delete Rows", DeleteRows),
            deleteCols = ("Delete Columns", DeleteColumns);
        var menu = area switch
        {
            GridArea.ColumnHeader => Menu(cut, copy, paste, null, insertCols, deleteCols, clear, null,
                ("Find Duplicates in These Columns…", () => OnFindDuplicatesClick(null, new RoutedEventArgs()))),
            GridArea.RowHeader => Menu(cut, copy, paste, null, insertRows, deleteRows, clear),
            _ => Menu(cut, copy, paste, null, insertRows, insertCols, deleteRows, deleteCols, clear, null,
                ("Edit Cell", () => BeginEdit(null))),
        };
        menu.Open(Sheet);
        e.Handled = true;
    }

    void OnTabsContextRequested(object? sender, ContextRequestedEventArgs e)
    {
        if (_doc is null) return;
        // Right-clicking a tab acts on that tab, so select it first.
        if ((e.Source as Visual)?.FindAncestorOfType<TabStripItem>(includeSelf: true) is { } tab
            && SheetTabs.IndexFromContainer(tab) is >= 0 and var index)
            SheetTabs.SelectedIndex = index;
        Menu(("New Sheet", NewSheet),
             ("Rename…", RenameCurrentSheet),
             ("Duplicate…", () => OnDuplicateSheetClick(null, new RoutedEventArgs())),
             ("Delete", DeleteSheet),
             null,
             ("Move Left", () => MoveSheet(-1)),
             ("Move Right", () => MoveSheet(1))).Open(SheetTabs);
        e.Handled = true;
    }

    void ShowSearch()
    {
        SearchBox.IsVisible = true;
        SearchBox.Focus();
        SearchBox.SelectAll();
    }

    void OnFindClick(object? sender, RoutedEventArgs e) => ShowSearch();

    void OnSearchKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            SearchBox.IsVisible = false;
            Sheet.Focus();
            e.Handled = true;
        }
        else if (e.Key == Key.Enter)
        {
            FindNext(backward: e.KeyModifiers.HasFlag(KeyModifiers.Shift));
            e.Handled = true;
        }
    }

    void FindNext(bool backward)
    {
        var query = SearchBox.Text ?? "";
        if (_doc is null || query.Length == 0) return;
        if (_doc.Find(SheetTabs.SelectedIndex, query, Sheet.SelectedRow, Sheet.SelectedCol, backward) is { } hit)
        {
            Sheet.Select(hit.Row, hit.Col);
            StatusText.Text = "";
        }
        else StatusText.Text = $"No match for \"{query}\"";
    }

    public async Task OpenAsync(string path)
    {
        if (!await ConfirmDiscardAsync()) return;
        StatusText.Text = "Opening…";
        try
        {
            var doc = await Task.Run(() => WorkbookDocument.Open(path));
            var old = _doc;
            CancelEdit();
            HidePrompt();
            _doc = doc;
            _dirty = false;
            UpdateTitle();
            SheetTabs.ItemsSource = doc.SheetNames;
            SheetTabs.SelectedIndex = 0;
            ShowSheet(0);
            old?.Dispose();
            StatusText.Text = "";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Cannot open: {ex.Message}";
        }
    }

    // ---- Unsaved changes ----

    /// <summary>True when it's fine to drop the current file: nothing unsaved, or the user saved or chose not to.</summary>
    async Task<bool> ConfirmDiscardAsync()
    {
        if (!_dirty || _doc is null) return true;
        switch (await AskSaveAsync())
        {
            case true:
                await SaveAsync(pick: false);
                return !_dirty; // still dirty = save failed or was cancelled
            case false: return true;
            default: return false;
        }
    }

    static Window NewDialog(string title) => new()
    {
        Title = title,
        SizeToContent = SizeToContent.WidthAndHeight,
        CanResize = false,
        WindowStartupLocation = WindowStartupLocation.CenterOwner,
    };

    /// <summary>Save / Don't Save / Cancel → true / false / null.</summary>
    async Task<bool?> AskSaveAsync()
    {
        bool? result = null;
        var dialog = NewDialog("Unsaved changes");
        Button Choice(string text, bool? value, bool isDefault = false)
        {
            var b = new Button { Content = text, IsDefault = isDefault, IsCancel = value is null };
            b.Click += (_, _) => { result = value; dialog.Close(); };
            return b;
        }
        dialog.Content = new StackPanel
        {
            Margin = new Thickness(20),
            Spacing = 16,
            Children =
            {
                new TextBlock { Text = $"Save changes to \"{Path.GetFileName(_doc?.Path)}\" before closing it?" },
                new StackPanel
                {
                    Orientation = Avalonia.Layout.Orientation.Horizontal,
                    HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
                    Spacing = 8,
                    Children = { Choice("Don't Save", false), Choice("Cancel", null), Choice("Save", true, isDefault: true) },
                },
            },
        };
        await dialog.ShowDialog(this);
        return result;
    }

    protected override async void OnClosing(WindowClosingEventArgs e)
    {
        base.OnClosing(e);
        if (_closeConfirmed || !_dirty) return;
        e.Cancel = true; // ask first, then close for real
        if (await ConfirmDiscardAsync())
        {
            _closeConfirmed = true;
            Close();
        }
    }

    void UpdateTitle() => Title = $"{(_dirty ? "• " : "")}{Path.GetFileName(_doc?.Path)} — OpenSheet";

    /// <summary>Call after editing cells: marks the file unsaved and redraws in place.</summary>
    internal void MarkEdited()
    {
        if (_doc is null) return;
        _dirty = true;
        Edits++;
        UpdateTitle();
        _used = _doc.UsedSize(CurrentSheet);
        Sheet.Duplicates = null; // may no longer match the cells
        Sheet.InvalidateVisual();
        SyncView();
    }

    /// <summary>Call after adding/renaming sheets or deleting rows: refreshes the tabs and shows <paramref name="sheet"/>.</summary>
    internal void Reload(int sheet)
    {
        if (_doc is null) return;
        SheetTabs.ItemsSource = _doc.SheetNames;
        SheetTabs.SelectedIndex = sheet;
        ShowSheet(sheet); // SelectionChanged may not fire when the index is unchanged
        MarkEdited();
    }

    void ShowSheet(int index)
    {
        if (_doc is null || index < 0) return;
        _used = _doc.UsedSize(index);
        Sheet.Show(_doc, index);
        Sheet.Focus();
    }

    void SyncView()
    {
        _syncing = true;
        VScroll.ViewportSize = VScroll.LargeChange = Sheet.VisibleRows;
        VScroll.Maximum = Math.Max(Math.Max(_used.Rows, Sheet.FirstRow), 1);
        VScroll.Value = Sheet.FirstRow;
        HScroll.ViewportSize = HScroll.LargeChange = Sheet.VisibleCols;
        HScroll.Maximum = Math.Max(Math.Max(_used.Cols, Sheet.FirstCol), 1);
        HScroll.Value = Sheet.FirstCol;
        _syncing = false;

        AddressText.Text = Sheet.SelectionText;
        if (_editing is null) FormulaBox.Text = Sheet.SelectedCell.Content;
        else PlaceEditor(); // follow the cell when scrolling
    }

    void OnSheetChanged(object? sender, SelectionChangedEventArgs e)
    {
        ShowSheet(SheetTabs.SelectedIndex);
        // Keep the selected tab visible (after layout, so a new tab has a size).
        Avalonia.Threading.Dispatcher.UIThread.Post(() => SheetTabs.ContainerFromIndex(SheetTabs.SelectedIndex)?.BringIntoView());
    }

    void ScrollTabs(double dx) => TabScroller.Offset = TabScroller.Offset.WithX(
        Math.Clamp(TabScroller.Offset.X + dx, 0, Math.Max(0, TabScroller.Extent.Width - TabScroller.Viewport.Width)));

    void OnTabsLeftClick(object? sender, RoutedEventArgs e) => ScrollTabs(-200);
    void OnTabsRightClick(object? sender, RoutedEventArgs e) => ScrollTabs(200);

    void OnFindDuplicatesClick(object? sender, RoutedEventArgs e)
    {
        if (_dupWindow is not null)
        {
            _dupWindow.Activate();
            return;
        }
        _dupWindow = new DuplicatesWindow(this, SelectedColumnsText());
        _dupWindow.Closed += (_, _) => _dupWindow = null;
        _dupWindow.Show(this);
    }

    async void OnSaveClick(object? sender, RoutedEventArgs e) => await SaveAsync(pick: false);

    async void OnSaveAsClick(object? sender, RoutedEventArgs e) => await SaveAsync(pick: true);

    async Task SaveAsync(bool pick)
    {
        if (_doc is null) return;
        CommitEdit(); // don't leave a half-typed cell out of the file
        var path = _doc.Path;
        if (pick || path is null)
        {
            var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Save workbook",
                SuggestedFileName = Path.GetFileName(path) ?? "Book1.xlsx",
                DefaultExtension = _doc.IsCsv ? "csv" : "xlsx",
                FileTypeChoices = _doc.IsCsv ? [CsvType, XlsxType] : [XlsxType, CsvType],
            });
            path = file?.TryGetLocalPath();
            if (path is null) return;
        }
        try
        {
            // ponytail: saves on the UI thread so rendering never reads the workbook mid-save.
            _doc.SaveAs(path, CurrentSheet); // a .csv gets the current sheet
            _dirty = false;
            UpdateTitle();
            StatusText.Text = _doc.IsCsv && _doc.SheetNames.Count > 1
                ? $"Saved \"{_doc.SheetNames[CurrentSheet]}\" to {Path.GetFileName(path)}. CSV keeps one sheet and values only"
                : $"Saved {Path.GetFileName(path)}";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Cannot save: {ex.Message}";
        }
    }

    async void OnOpenClick(object? sender, RoutedEventArgs e) => await PickAndOpenAsync();

    async Task PickAndOpenAsync()
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Open workbook",
            FileTypeFilter = [SpreadsheetTypes, XlsxType, CsvType],
        });
        if (files is [var file, ..] && file.TryGetLocalPath() is { } path) await OpenAsync(path);
    }

    async void OnDrop(object? sender, DragEventArgs e)
    {
        if (e.DataTransfer.TryGetFiles() is not [var file, ..] || file.TryGetLocalPath() is not { } path) return;
        if (_doc is null) await OpenAsync(path);
        else await AddFileAsync(path); // a file is already open: ask what to do with this one
    }

    async void OnImportClick(object? sender, RoutedEventArgs e)
    {
        if (_doc is null)
        {
            await PickAndOpenAsync();
            return;
        }
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Import from workbook",
            FileTypeFilter = [SpreadsheetTypes, XlsxType, CsvType],
        });
        if (files is [var file, ..] && file.TryGetLocalPath() is { } path) await AddFileAsync(path);
    }

    enum AddMode { Open, ImportSheets, AppendRows }

    /// <summary>Another .xlsx while one is open: open it instead, import its sheets, or append its rows here.</summary>
    async Task AddFileAsync(string path)
    {
        if (_doc is null || CurrentSheet < 0) return;
        if (await AskAddModeAsync(Path.GetFileName(path), _doc.SheetNames[CurrentSheet]) is not { } choice) return;
        if (choice.Mode == AddMode.Open)
        {
            await OpenAsync(path);
            return;
        }
        CommitEdit();
        StatusText.Text = "Importing…";
        try
        {
            if (choice.Mode == AddMode.ImportSheets)
            {
                int index = CurrentSheet + 1;
                var names = _doc.ImportSheets(path, index);
                Reload(index); // show the first imported sheet
                StatusText.Text = $"Imported {names.Count} sheet(s) from {Path.GetFileName(path)}: {string.Join(", ", names)}";
            }
            else
            {
                var (first, count) = _doc.AppendRows(path, CurrentSheet, choice.SkipHeader);
                MarkEdited();
                if (count > 0)
                {
                    Sheet.Select(first, 1);
                    Sheet.SelectRows(first, first + count - 1); // show what was added
                }
                StatusText.Text = $"Appended {count} row(s) from {Path.GetFileName(path)}";
            }
        }
        catch (Exception ex) { StatusText.Text = $"Cannot import: {ex.Message}"; }
    }

    async Task<(AddMode Mode, bool SkipHeader)?> AskAddModeAsync(string fileName, string sheetName)
    {
        (AddMode, bool)? result = null;
        var dialog = NewDialog("Add workbook");
        var skipHeader = new CheckBox { Content = "Its first row is a header — don't append it", IsChecked = true, Margin = new Thickness(24, 0, 0, 0) };
        Control Option(string title, string detail, AddMode mode)
        {
            var b = new Button
            {
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Stretch,
                Padding = new Thickness(12, 8),
                Content = new StackPanel
                {
                    Children =
                    {
                        new TextBlock { Text = title, FontWeight = Avalonia.Media.FontWeight.SemiBold },
                        new TextBlock { Text = detail, Opacity = 0.7, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                    },
                },
            };
            b.Click += (_, _) => { result = (mode, skipHeader.IsChecked == true); dialog.Close(); };
            return b;
        }
        var cancel = new Button { Content = "Cancel", IsCancel = true, HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right };
        cancel.Click += (_, _) => dialog.Close();
        dialog.Content = new StackPanel
        {
            Margin = new Thickness(20),
            Spacing = 10,
            Width = 420,
            Children =
            {
                new TextBlock { Text = $"What do you want to do with \"{fileName}\"?", FontWeight = Avalonia.Media.FontWeight.SemiBold },
                Option("Open it", "Switch to that file (asks to save this one first if needed).", AddMode.Open),
                Option("Import its sheets", "Add every sheet from it as new sheets in this workbook.", AddMode.ImportSheets),
                Option($"Append its rows to \"{sheetName}\"", "Add its data below the last row of this sheet, column A to A.", AddMode.AppendRows),
                skipHeader,
                cancel,
            },
        };
        await dialog.ShowDialog(this);
        return result;
    }

    protected override async void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (!e.Handled && e.Key == Key.F11 && e.KeyModifiers == KeyModifiers.Shift)
        {
            e.Handled = true;
            NewSheet(); // Excel's shortcut
            return;
        }
        // Ctrl on Windows, Cmd on macOS (either works on both).
        if (e.Handled || (!e.KeyModifiers.HasFlag(KeyModifiers.Control) && !e.KeyModifiers.HasFlag(KeyModifiers.Meta))) return;
        bool shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        bool inTextBox = e.Source is TextBox; // text boxes keep their own undo/copy/paste
        e.Handled = true;
        switch (e.Key)
        {
            case Key.F: ShowSearch(); break;
            case Key.O: await PickAndOpenAsync(); break;
            case Key.S: await SaveAsync(pick: shift); break;
            case Key.PageDown: StepSheet(1); break;
            case Key.PageUp: StepSheet(-1); break;
            case Key.Z when !inTextBox: UndoRedo(redo: shift); break;
            case Key.Y when !inTextBox: UndoRedo(redo: true); break;
            case Key.C when !inTextBox: await CopyAsync(cut: false); break;
            case Key.X when !inTextBox: await CopyAsync(cut: true); break;
            case Key.V when !inTextBox: await PasteAsync(); break;
            default: e.Handled = false; break;
        }
    }
}
