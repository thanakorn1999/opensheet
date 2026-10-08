using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Platform.Storage;
using XlsxEditor.Core;

namespace XlsxEditor.App;

public partial class MainWindow : Window
{
    WorkbookDocument? _doc;
    (int Rows, int Cols) _used;
    bool _syncing; // guards scrollbar <-> grid feedback
    bool _dirty;   // edited since open/save
    DuplicatesWindow? _dupWindow;
    (int Sheet, int Row, int Col, string Original)? _editing; // cell open in the in-cell editor
    (int Start, int End)? _pointRef; // cell reference just inserted into a formula by clicking a cell
    // What we last put on the clipboard, so pasting it back can keep formulas (the clipboard only holds text).
    (string Text, int Sheet, int Row, int Col, string? CutContent)? _clip;
    Func<string, string?>? _prompt; // PromptBox's Enter action: returns an error to show, or null when done

    static readonly FilePickerFileType XlsxType = new("Excel Workbook")
    {
        Patterns = ["*.xlsx"],
        AppleUniformTypeIdentifiers = ["org.openxmlformats.spreadsheetml.sheet"],
    };

    internal WorkbookDocument? Doc => _doc;
    internal int CurrentSheet => SheetTabs.SelectedIndex;
    /// <summary>Bumped on every edit so other windows can tell their results are stale.</summary>
    internal int Edits { get; private set; }

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
        CellEditor.LostFocus += (_, _) => CommitEdit(); // clicking elsewhere saves, like Excel
        CellEditor.TextChanged += (_, _) => { if (_editing is not null) FormulaBox.Text = CellEditor.Text; };
        Sheet.KeyDown += OnSheetKeyDown;
        Sheet.AddHandler(TextInputEvent, OnSheetTextInput);
        Sheet.AddHandler(PointerPressedEvent, OnSheetPointerPressed, RoutingStrategies.Tunnel);
        Sheet.DoubleTapped += (_, _) => BeginEdit(null);
        SheetTabs.DoubleTapped += (_, _) => RenameCurrentSheet();

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
        _pointRef = null;
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
            default: _pointRef = null; return; // typing ends "click to insert a reference"
        }
        e.Handled = true;
    }

    // While typing a formula, clicking a cell inserts its address (clicking again swaps it), like Excel.
    void OnSheetPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (_editing is null || Sheet.HitTest(e.GetPosition(Sheet)) is not { } hit) return;
        var text = CellEditor.Text ?? "";
        int caret = CellEditor.CaretIndex;
        bool replacing = _pointRef is { } p && p.End == caret;
        bool afterOperator = caret > 0 && "=+-*/^(,:;<>&".Contains(text[caret - 1]);
        if (!text.StartsWith('=') || !(replacing || afterOperator)) return; // normal click: commit and select

        int start = replacing ? _pointRef!.Value.Start : caret;
        var address = WorkbookDocument.Address(hit.Row, hit.Col);
        CellEditor.Text = text[..start] + address + text[caret..];
        CellEditor.CaretIndex = start + address.Length;
        _pointRef = (start, start + address.Length);
        e.Handled = true; // keep the editor open and focused
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
        _doc.SetCell(CurrentSheet, Sheet.SelectedRow, Sheet.SelectedCol, "");
        MarkEdited();
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
        var cell = Sheet.SelectedCell;
        var text = cell.Text ?? "";
        await Clipboard.SetTextAsync(text);
        // Copy pastes with references shifted; cut pastes the exact content once.
        _clip = (text, CurrentSheet, Sheet.SelectedRow, Sheet.SelectedCol, cut ? cell.Content ?? "" : null);
        if (cut) ClearSelected();
        StatusText.Text = cut ? "Cut" : "Copied";
    }

    // ponytail: single-cell selection, so copy/cut take one cell; paste accepts a whole tab-separated block.
    async Task PasteAsync()
    {
        if (_doc is null || CurrentSheet < 0 || Clipboard is null) return;
        if (await Clipboard.TryGetTextAsync() is not { } text) return;
        int sheet = CurrentSheet, row = Sheet.SelectedRow, col = Sheet.SelectedCol;
        try
        {
            if (_clip is { } c && c.Text == text && c.Sheet < _doc.SheetNames.Count)
            {
                if (c.CutContent is { } content)
                {
                    _doc.SetCell(sheet, row, col, content);
                    _clip = null; // a cut pastes once
                }
                else _doc.CopyCell(c.Sheet, c.Row, c.Col, sheet, row, col);
            }
            else
            {
                // From another app (Excel, Sheets, a text editor): rows by line, cells by tab.
                var lines = text.Replace("\r\n", "\n").Split('\n');
                if (lines.Length > 1 && lines[^1].Length == 0) lines = lines[..^1];
                _doc.SetCells(sheet, row, col, lines.Select(l => (IReadOnlyList<string>)l.Split('\t')).ToList());
            }
            MarkEdited();
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

    void OnDeleteColumnsClick(object? sender, RoutedEventArgs e)
    {
        if (_doc is not { } doc || CurrentSheet < 0) return;
        int sheet = CurrentSheet;
        ShowPrompt("Columns to delete, e.g. B or B,D:F  (Enter = delete, Esc = cancel)", WorkbookDocument.ColumnLetter(Sheet.SelectedCol), text =>
        {
            if (WorkbookDocument.ParseColumns(text) is not { } cols) return "Invalid columns — use e.g. B or B,D:F";
            doc.DeleteColumns(sheet, cols);
            MarkEdited();
            StatusText.Text = $"Deleted column {string.Join(", ", cols.Select(WorkbookDocument.ColumnLetter))}";
            return null;
        });
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
        StatusText.Text = "Opening…";
        try
        {
            var doc = await Task.Run(() => WorkbookDocument.Open(path));
            var old = _doc;
            _editing = null;
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

    void UpdateTitle() => Title = $"{(_dirty ? "• " : "")}{Path.GetFileName(_doc?.Path)} — XLSX Editor";

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

        AddressText.Text = WorkbookDocument.Address(Sheet.SelectedRow, Sheet.SelectedCol);
        if (_editing is null) FormulaBox.Text = Sheet.SelectedCell.Content;
        else PlaceEditor(); // follow the cell when scrolling
    }

    void OnSheetChanged(object? sender, SelectionChangedEventArgs e) => ShowSheet(SheetTabs.SelectedIndex);

    void OnFindDuplicatesClick(object? sender, RoutedEventArgs e)
    {
        if (_dupWindow is not null)
        {
            _dupWindow.Activate();
            return;
        }
        _dupWindow = new DuplicatesWindow(this, Sheet.SelectedCol);
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
                DefaultExtension = "xlsx",
                FileTypeChoices = [XlsxType],
            });
            path = file?.TryGetLocalPath();
            if (path is null) return;
        }
        try
        {
            // ponytail: saves on the UI thread so rendering never reads the workbook mid-save.
            _doc.SaveAs(path);
            _dirty = false;
            UpdateTitle();
            StatusText.Text = $"Saved {Path.GetFileName(path)}";
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
            FileTypeFilter = [XlsxType],
        });
        if (files is [var file, ..] && file.TryGetLocalPath() is { } path) await OpenAsync(path);
    }

    async void OnDrop(object? sender, DragEventArgs e)
    {
        if (e.DataTransfer.TryGetFiles() is [var file, ..] && file.TryGetLocalPath() is { } path) await OpenAsync(path);
    }

    protected override async void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
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
