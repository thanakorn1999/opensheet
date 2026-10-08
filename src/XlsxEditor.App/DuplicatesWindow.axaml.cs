using Avalonia.Controls;
using Avalonia.Interactivity;
using XlsxEditor.Core;

namespace XlsxEditor.App;

/// <summary>Finds duplicate rows on the main window's current sheet, then deletes or moves them.</summary>
public partial class DuplicatesWindow : Window
{
    readonly MainWindow _main;
    // What the last Find ran against, so actions never hit a different file or sheet.
    WorkbookDocument? _doc;
    int _sheet;
    int _edits;
    IReadOnlyList<int[]> _groups = [];

    // For the XAML previewer only; the app uses the constructor below.
    public DuplicatesWindow() : this(null!, 1) { }

    public DuplicatesWindow(MainWindow main, int selectedCol)
    {
        InitializeComponent();
        _main = main;
        ColumnsBox.Text = WorkbookDocument.ColumnLetter(selectedCol);
        Opened += (_, _) => { ColumnsBox.Focus(); ColumnsBox.SelectAll(); };
        Closed += (_, _) => _main.Sheet.Duplicates = null;
    }

    record GroupItem(int[] Rows, string Label)
    {
        public override string ToString() => Label;
    }

    void OnFindClick(object? sender, RoutedEventArgs e) => Find();

    void Find()
    {
        _groups = [];
        GroupList.ItemsSource = null;
        DeleteButton.IsEnabled = CopyButton.IsEnabled = MoveButton.IsEnabled = false;

        if (_main.Doc is not { } doc)
        {
            SummaryText.Text = "Open a workbook first.";
            return;
        }
        if (WorkbookDocument.ParseColumns(ColumnsBox.Text ?? "") is not { } cols)
        {
            SummaryText.Text = "Invalid columns — use e.g. A,C or B:D";
            return;
        }

        _doc = doc;
        _sheet = _main.CurrentSheet;
        _edits = _main.Edits;
        // ponytail: runs on the UI thread; move to Task.Run once very large sheets feel slow.
        _groups = doc.DuplicateGroups(_sheet, cols, HeaderCheck.IsChecked == true);
        _main.Sheet.Duplicates = (cols, _groups.SelectMany(g => g).ToHashSet());

        GroupList.ItemsSource = _groups.Select(g => new GroupItem(g, Describe(doc, cols, g))).ToList();
        var label = string.Join("+", cols.Select(WorkbookDocument.ColumnLetter));
        int total = _groups.Sum(g => g.Length);
        SummaryText.Text = _groups.Count == 0
            ? $"{label}: no duplicates on \"{doc.SheetNames[_sheet]}\""
            : $"{label}: {_groups.Count} groups, {total} rows ({total - _groups.Count} extra) on \"{doc.SheetNames[_sheet]}\"";
        DeleteButton.IsEnabled = CopyButton.IsEnabled = MoveButton.IsEnabled = _groups.Count > 0;
    }

    string Describe(WorkbookDocument doc, IReadOnlyList<int> cols, int[] rows)
    {
        var values = string.Join(" | ", cols.Select(c => doc.GetCell(_sheet, rows[0], c).Text is { Length: > 0 } t ? t : "(blank)"));
        var list = string.Join(", ", rows.Take(8)) + (rows.Length > 8 ? ", …" : "");
        return $"{rows.Length}×  {values}   — rows {list}";
    }

    void OnGroupSelected(object? sender, SelectionChangedEventArgs e)
    {
        if (GroupList.SelectedItem is GroupItem g && _main.CurrentSheet == _sheet)
            _main.Sheet.Select(g.Rows[0], _main.Sheet.SelectedCol);
    }

    /// <summary>Rows the user chose to act on, or null (with a message) if the results are stale.</summary>
    List<int>? Targets()
    {
        if (_doc is null || _doc != _main.Doc || _sheet != _main.CurrentSheet || _edits != _main.Edits)
        {
            StatusText.Text = "The file or sheet changed — press Find again.";
            return null;
        }
        return (ExtrasRadio.IsChecked == true ? _groups.SelectMany(g => g.Skip(1)) : _groups.SelectMany(g => g)).ToList();
    }

    void OnDeleteClick(object? sender, RoutedEventArgs e) => Edit(rows =>
    {
        _doc!.DeleteRows(_sheet, rows);
        return $"Deleted {rows.Count} rows.";
    });

    void OnCopyClick(object? sender, RoutedEventArgs e)
    {
        if (!SheetNameOk()) return;
        Edit(rows =>
        {
            var name = _doc!.CopyRowsToNewSheet(_sheet, rows, HeaderCheck.IsChecked == true, SheetNameBox.Text);
            return $"Copied {rows.Count} rows to sheet \"{name}\".";
        });
    }

    void OnMoveClick(object? sender, RoutedEventArgs e)
    {
        if (!SheetNameOk()) return;
        Edit(rows =>
        {
            var name = _doc!.MoveRowsToNewSheet(_sheet, rows, HeaderCheck.IsChecked == true, SheetNameBox.Text);
            return $"Moved {rows.Count} rows to sheet \"{name}\".";
        });
    }

    /// <summary>Checks the typed sheet name before touching anything. Blank is fine (auto-named).</summary>
    bool SheetNameOk()
    {
        var name = SheetNameBox.Text?.Trim() ?? "";
        if (name.Length == 0 || _main.Doc?.CheckNewSheetName(name) is not { } error) return true;
        StatusText.Text = error;
        SheetNameBox.Focus();
        return false;
    }

    void Edit(Func<List<int>, string> action)
    {
        if (Targets() is not { } rows) return;
        string message;
        try
        {
            message = action(rows) + " Not saved yet — Save (Cmd/Ctrl+S) to keep it, or reopen the file to undo.";
            SheetNameBox.Text = ""; // a used name would collide next time
        }
        catch (Exception ex) { message = $"Failed: {ex.Message}. The sheet may be partly changed — reopen the file without saving to discard."; }
        _main.Reload(_sheet);
        Find(); // refresh results and highlight against the edited sheet
        StatusText.Text = message;
    }
}
