using ClosedXML.Excel;
using XlsxEditor.Core;

namespace XlsxEditor.Core.Tests;

public class WorkbookDocumentTests
{
    [Fact]
    public void Open_ReadsSheetsCellsAndFormulas()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sales");
            ws.Cell("A1").Value = "Name";
            ws.Cell("B1").Value = 25;
            ws.Cell("B2").Value = 5;
            ws.Cell("B3").FormulaA1 = "SUM(B1:B2)";
            wb.AddWorksheet("Users");
            wb.SaveAs(path);
        }

        try
        {
            using var doc = WorkbookDocument.Open(path);

            Assert.Equal(["Sales", "Users"], doc.SheetNames);
            Assert.Equal((3, 2), doc.UsedSize(0));
            Assert.Equal((0, 0), doc.UsedSize(1));

            Assert.Equal(new CellView("Name", "Name", false), doc.GetCell(0, 1, 1));
            Assert.Equal(new CellView("30", "=SUM(B1:B2)", true), doc.GetCell(0, 3, 2));
            Assert.Equal(default, doc.GetCell(0, 10, 10));
            Assert.Equal("AB12", WorkbookDocument.Address(12, 28));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DuplicateRows_IgnoresBlanksAndCase()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = "apple";
            ws.Cell("A2").Value = "Pear";
            ws.Cell("A3").Value = "APPLE";
            // A4 blank, A5 blank
            ws.Cell("A6").Value = 7;
            ws.Cell("A7").Value = 7;
            ws.Cell("B8").Value = "x"; // extends used range past column A
            wb.SaveAs(path);
        }

        try
        {
            using var doc = WorkbookDocument.Open(path);
            Assert.Equal([[1, 3], [6, 7]], doc.DuplicateGroups(0, [1]));
            Assert.Empty(doc.DuplicateGroups(0, [2]));
            Assert.Equal([[6, 7]], doc.DuplicateGroups(0, [1], hasHeader: true)); // row 1 skipped, so "APPLE" has no partner
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DuplicateRows_MultipleColumnsMatchTogether()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            // A (name) | B (city) | C (blank)
            ws.Cell("A1").Value = "Ann"; ws.Cell("B1").Value = "BKK";
            ws.Cell("A2").Value = "Ann"; ws.Cell("B2").Value = "CNX"; // same name, different city
            ws.Cell("A3").Value = "ann"; ws.Cell("B3").Value = "bkk"; // dup of row 1
            ws.Cell("A4").Value = "Bob";                               // B4 blank
            ws.Cell("A5").Value = "Bob";                               // dup of row 4 (blank counts as a value)
            ws.Cell("D6").Value = "x";                                  // row 6 blank in A:C
            ws.Cell("D7").Value = "y";                                  // row 7 blank in A:C, not a dup
            wb.SaveAs(path);
        }

        try
        {
            using var doc = WorkbookDocument.Open(path);
            Assert.Equal([[1, 3], [4, 5]], doc.DuplicateGroups(0, [1, 2]));
            Assert.Equal([[1, 2, 3], [4, 5]], doc.DuplicateGroups(0, [1]));
            Assert.Equal([[1, 3], [4, 5]], doc.DuplicateGroups(0, [1, 2, 3]));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DeleteMoveAndSave_RoundTrip()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Name";
            ws.Cell("A2").Value = "Ann";
            ws.Cell("A3").Value = "Bob";
            ws.Cell("A4").Value = "Ann";
            ws.Cell("A5").Value = "Ann";
            wb.AddWorksheet("Duplicates"); // forces the new sheet to be "Duplicates 2"
            wb.SaveAs(path);
        }

        try
        {
            using (var doc = WorkbookDocument.Open(path))
            {
                var groups = doc.DuplicateGroups(0, [1], hasHeader: true);
                Assert.Equal([[2, 4, 5]], groups);

                // Keep the first of each group, move the extras.
                var name = doc.MoveRowsToNewSheet(0, groups.SelectMany(g => g.Skip(1)).ToList(), withHeader: true);
                Assert.Equal("Duplicates 2", name);
                Assert.Equal(["Data", "Duplicates 2", "Duplicates"], doc.SheetNames);
                Assert.Equal(["Name", "Ann", "Bob"], Column(doc, 0));
                Assert.Equal(["Name", "Ann", "Ann"], Column(doc, 1));

                doc.DeleteRows(0, [2]);
                Assert.Equal(["Name", "Bob"], Column(doc, 0));
                doc.SaveAs(path); // overwrite the file it was opened from
            }

            using var reopened = WorkbookDocument.Open(path);
            Assert.Equal(["Name", "Bob"], Column(reopened, 0));
            Assert.Equal(["Name", "Ann", "Ann"], Column(reopened, 1));
        }
        finally { File.Delete(path); }

        static string[] Column(WorkbookDocument doc, int sheet) =>
            Enumerable.Range(1, doc.UsedSize(sheet).Rows).Select(r => doc.GetCell(sheet, r, 1).Text).ToArray();
    }

    [Fact]
    public void DeleteRows_WithHyperlinks()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            for (int r = 5; r >= 1; r--) // bottom-up: ClosedXML shifts links in stored order and used to collide
            {
                ws.Cell(r, 1).Value = $"row{r}";
                ws.Cell(r, 2).SetHyperlink(new XLHyperlink($"https://example.com/{r}"));
            }
            ws.Cell(5, 2).Style.Font.FontColor = XLColor.Red; // custom link formatting must survive
            wb.SaveAs(path);
        }

        try
        {
            using (var doc = WorkbookDocument.Open(path))
            {
                doc.DeleteRows(0, [2, 3]);
                Assert.Equal(["row1", "row4", "row5"], Enumerable.Range(1, 3).Select(r => doc.GetCell(0, r, 1).Text));
                doc.MoveRowsToNewSheet(0, [2], withHeader: false);
                doc.SaveAs(path);
            }

            using var wb2 = new XLWorkbook(path);
            Assert.Equal("https://example.com/1", wb2.Worksheet(1).Cell(1, 2).GetHyperlink().ExternalAddress.ToString());
            Assert.Equal("https://example.com/5", wb2.Worksheet(1).Cell(2, 2).GetHyperlink().ExternalAddress.ToString());
            Assert.Equal(XLColor.Red, wb2.Worksheet(1).Cell(2, 2).Style.Font.FontColor);
            Assert.Equal(XLFontUnderlineValues.Single, wb2.Worksheet(1).Cell(2, 2).Style.Font.Underline);
            Assert.False(wb2.Worksheet(1).Cell(3, 2).HasHyperlink);
            Assert.Equal("https://example.com/4", wb2.Worksheet(2).Cell(1, 2).GetHyperlink().ExternalAddress.ToString());
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void CopyRowsToNewSheet_KeepsSourceAndNamesSheet()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Name";
            ws.Cell("A2").Value = "Ann";
            ws.Cell("A3").Value = "Ann";
            wb.AddWorksheet("Duplicates");
            wb.SaveAs(path);
        }

        try
        {
            using var doc = WorkbookDocument.Open(path);

            Assert.Equal("Duplicates 2", doc.CopyRowsToNewSheet(0, [2, 3], withHeader: true));      // auto, skips taken name
            Assert.Equal("Duplicates 3", doc.CopyRowsToNewSheet(0, [3], withHeader: false, name: "  "));
            Assert.Equal("Ann list", doc.CopyRowsToNewSheet(0, [3], withHeader: false, name: " Ann list "));
            Assert.Equal(["Data", "Ann list", "Duplicates 3", "Duplicates 2", "Duplicates"], doc.SheetNames);
            Assert.Equal((3, 1), doc.UsedSize(0)); // source untouched
            Assert.Equal("Ann", doc.GetCell(3, 3, 1).Text);

            Assert.Throws<ArgumentException>(() => doc.CopyRowsToNewSheet(0, [2], false, "ann LIST")); // case-insensitive dup
            Assert.Equal(5, doc.SheetNames.Count); // nothing created on failure

            Assert.Null(doc.CheckNewSheetName("New"));
            Assert.NotNull(doc.CheckNewSheetName("data"));
            Assert.NotNull(doc.CheckNewSheetName("a/b"));
            Assert.NotNull(doc.CheckNewSheetName("'quoted'"));
            Assert.NotNull(doc.CheckNewSheetName(new string('x', 32)));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void CopySheet_CopiesEverythingWithUniqueName()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        var longName = new string('L', 31);
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sales");
            ws.Cell("A1").Value = 2;
            ws.Cell("A2").FormulaA1 = "A1*10";
            ws.Cell("A1").Style.Font.Bold = true;
            wb.AddWorksheet("Sales (2)");
            wb.AddWorksheet(longName);
            wb.SaveAs(path);
        }

        try
        {
            using (var doc = WorkbookDocument.Open(path))
            {
                Assert.Equal("Sales (3)", doc.CopySheet(0));               // "(2)" is taken
                Assert.Equal("My copy", doc.CopySheet(0, " My copy "));
                Assert.Equal(new string('L', 27) + " (2)", doc.CopySheet(4)); // trimmed to 31 chars
                Assert.Throws<ArgumentException>(() => doc.CopySheet(0, "sales"));
                Assert.Equal(["Sales", "My copy", "Sales (3)", "Sales (2)", longName, new string('L', 27) + " (2)"], doc.SheetNames);

                Assert.Equal(new CellView("20", "=A1*10", true), doc.GetCell(1, 2, 1));
                doc.SaveAs(path);
            }

            using var wb2 = new XLWorkbook(path);
            Assert.True(wb2.Worksheet("My copy").Cell("A1").Style.Font.Bold);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void SetCell_ParsesLikeExcelAndRecalculates()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = 1;
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A2").FormulaA1 = "A1*10";
            wb.SaveAs(path);
        }

        try
        {
            using (var doc = WorkbookDocument.Open(path))
            {
                doc.SetCell(0, 1, 1, "5");
                Assert.Equal(new CellView("50", "=A1*10", true), doc.GetCell(0, 2, 1)); // dependent formula updated
                doc.SetCell(0, 1, 2, "=A1+1");
                Assert.Equal(new CellView("6", "=A1+1", true), doc.GetCell(0, 1, 2));
                doc.SetCell(0, 3, 1, "hello");
                Assert.Equal(new CellView("hello", "hello", false), doc.GetCell(0, 3, 1));
                doc.SetCell(0, 4, 1, "'007");
                Assert.Equal(new CellView("007", "007", false), doc.GetCell(0, 4, 1));
                doc.SetCell(0, 5, 1, "TRUE");
                Assert.Equal("TRUE", doc.GetCell(0, 5, 1).Text);
                doc.SetCell(0, 3, 1, "");
                Assert.Equal(default, doc.GetCell(0, 3, 1));
                doc.SaveAs(path);
            }

            using var wb2 = new XLWorkbook(path);
            Assert.Equal(5.0, wb2.Worksheet(1).Cell("A1").Value.GetNumber());
            Assert.True(wb2.Worksheet(1).Cell("A1").Style.Font.Bold); // formatting kept
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DeleteColumns_ShiftsCellsFormulasAndHyperlinks()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            for (int c = 5; c >= 1; c--) // right-to-left: same ClosedXML link-shift trap as rows
            {
                ws.Cell(1, c).Value = c;
                ws.Cell(2, c).SetHyperlink(new XLHyperlink($"https://example.com/{c}"));
            }
            ws.Cell("F1").FormulaA1 = "SUM(A1:E1)";
            wb.SaveAs(path);
        }

        try
        {
            using var doc = WorkbookDocument.Open(path);
            doc.DeleteColumns(0, [4, 2]);
            Assert.Equal(["1", "3", "5", "9"], Enumerable.Range(1, 4).Select(c => doc.GetCell(0, 1, c).Text));
            Assert.Equal("=SUM(A1:C1)", doc.GetCell(0, 1, 4).Content);
            doc.SaveAs(path);

            using var wb2 = new XLWorkbook(path);
            Assert.Equal("https://example.com/3", wb2.Worksheet(1).Cell(2, 2).GetHyperlink().ExternalAddress.ToString());
            Assert.Equal("https://example.com/5", wb2.Worksheet(1).Cell(2, 3).GetHyperlink().ExternalAddress.ToString());
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void RenameSheet_ValidatesAndUpdatesFormulas()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Data").Cell("A1").Value = 7;
            wb.AddWorksheet("Report").Cell("A1").FormulaA1 = "Data!A1*2";
            wb.SaveAs(path);
        }

        try
        {
            using (var doc = WorkbookDocument.Open(path))
            {
                Assert.Equal("DATA", doc.RenameSheet(0, " DATA "));  // same sheet, only case changes: allowed
                Assert.Throws<ArgumentException>(() => doc.RenameSheet(0, "report"));
                Assert.Throws<ArgumentException>(() => doc.RenameSheet(0, "a:b"));
                Assert.Equal("Sales 2026", doc.RenameSheet(0, "Sales 2026"));
                Assert.Equal(["Sales 2026", "Report"], doc.SheetNames);
                Assert.Equal("14", doc.GetCell(1, 1, 1).Text);
                doc.SaveAs(path);
            }

            using var wb2 = new XLWorkbook(path);
            Assert.Equal(14.0, wb2.Worksheet("Report").Cell("A1").Value.GetNumber());
            Assert.Contains("Sales 2026", wb2.Worksheet("Report").Cell("A1").FormulaA1);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void UndoRedo_CellsAndStructuralEdits()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").FormulaA1 = "A1*10";
            ws.Cell("A3").Value = "keep";
            wb.SaveAs(path);
        }

        try
        {
            using var doc = WorkbookDocument.Open(path);
            string A(int r) => doc.GetCell(0, r, 1).Content;
            Assert.False(doc.CanUndo);

            doc.SetCell(0, 2, 1, "plain");   // formula -> text
            doc.SetCells(0, 1, 2, [["x", "y"], ["=A1+1"]]);
            Assert.Equal("y", doc.GetCell(0, 1, 3).Text);
            Assert.Equal("2", doc.GetCell(0, 2, 2).Text);

            Assert.True(doc.Undo());         // whole paste in one step
            Assert.Equal(default, doc.GetCell(0, 1, 2));
            Assert.Equal(default, doc.GetCell(0, 2, 2));
            Assert.True(doc.Undo());
            Assert.Equal("=A1*10", A(2));    // formula is back
            Assert.False(doc.CanUndo);
            Assert.True(doc.Redo());
            Assert.Equal("plain", A(2));

            doc.MoveRowsToNewSheet(0, [1], withHeader: false); // one step for copy + delete
            Assert.Equal(["S", "Duplicates"], doc.SheetNames);
            Assert.Equal("plain", A(1));
            Assert.False(doc.CanRedo);       // a new edit clears redo
            Assert.True(doc.Undo());
            Assert.Equal(["S"], doc.SheetNames);
            Assert.Equal(["1", "plain", "keep"], new[] { A(1), A(2), A(3) });
            Assert.True(doc.Redo());
            Assert.Equal(["S", "Duplicates"], doc.SheetNames);

            Assert.Throws<ArgumentException>(() => doc.RenameSheet(0, "duplicates"));
            Assert.True(doc.Undo());         // failed rename recorded nothing, so this undoes the move
            Assert.Equal(["S"], doc.SheetNames);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void CopyCell_ShiftsRelativeReferences()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("B1").FormulaA1 = "A1*10+$A$1";
            wb.SaveAs(path);
        }

        try
        {
            using var doc = WorkbookDocument.Open(path);
            doc.CopyCell(0, 1, 2, 0, 2, 2);   // B1 -> B2
            Assert.Equal(new CellView("21", "=A2*10+$A$1", true), doc.GetCell(0, 2, 2));
            doc.CopyCell(0, 1, 1, 0, 3, 1);   // value A1 -> A3
            Assert.Equal("1", doc.GetCell(0, 3, 1).Text);
            Assert.True(doc.Undo());
            Assert.Equal(default, doc.GetCell(0, 3, 1));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void JumpEdge_MovesLikeCtrlArrow()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            foreach (var r in new[] { 1, 2, 3, 6, 7 }) ws.Cell(r, 1).Value = r;
            ws.Cell("D9").Value = "corner";
            wb.SaveAs(path);
        }

        try
        {
            using var doc = WorkbookDocument.Open(path);
            Assert.Equal((3, 1), doc.JumpEdge(0, 1, 1, 1, 0)); // end of block
            Assert.Equal((6, 1), doc.JumpEdge(0, 3, 1, 1, 0)); // next filled
            Assert.Equal((7, 1), doc.JumpEdge(0, 6, 1, 1, 0));
            Assert.Equal((9, 1), doc.JumpEdge(0, 7, 1, 1, 0)); // nothing below: edge of used range
            Assert.Equal((1, 1), doc.JumpEdge(0, 1, 1, -1, 0)); // already at top
            Assert.Equal((1, 4), doc.JumpEdge(0, 1, 1, 0, 1));  // right edge
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void ParseColumns_ListsAndRanges()
    {
        Assert.Equal([1, 3], WorkbookDocument.ParseColumns("A,C"));
        Assert.Equal([1, 2, 3, 4, 5], WorkbookDocument.ParseColumns(" a , c:e , b"));
        Assert.Equal([2, 3, 4], WorkbookDocument.ParseColumns("D:B"));
        Assert.Equal([28], WorkbookDocument.ParseColumns("AB"));
        Assert.Null(WorkbookDocument.ParseColumns(""));
        Assert.Null(WorkbookDocument.ParseColumns("A,1"));
        Assert.Null(WorkbookDocument.ParseColumns("A:B:C"));
    }

    [Fact]
    public void Find_WrapsAroundBothWays()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("B1").Value = "Apple pie";
            ws.Cell("A2").Value = "banana";
            ws.Cell("C3").Value = "apple";
            wb.SaveAs(path);
        }

        try
        {
            using var doc = WorkbookDocument.Open(path);
            Assert.Equal((1, 2), doc.Find(0, "APPLE", 1, 1));
            Assert.Equal((3, 3), doc.Find(0, "apple", 1, 2));
            Assert.Equal((1, 2), doc.Find(0, "apple", 3, 3));                 // wraps forward
            Assert.Equal((3, 3), doc.Find(0, "apple", 1, 2, backward: true)); // wraps backward
            Assert.Equal((2, 1), doc.Find(0, "nan", 2, 1));                   // only match = itself
            Assert.Null(doc.Find(0, "kiwi", 1, 1));
        }
        finally { File.Delete(path); }
    }
}
