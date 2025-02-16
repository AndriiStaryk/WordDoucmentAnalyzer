using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentAnalyzer;

static class TableGenerator
{
    public static Table CreateDailyTasksDescriptionTable(List<List<string>> rows, int minRowsCount = 28, int maxRowsCount = 28)
    {
        Table table = new Table(new TableProperties(
            new TableWidth() { Width = "100%", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Size = 4 },
                new BottomBorder() { Val = BorderValues.Single, Size = 4 },
                new LeftBorder() { Val = BorderValues.Single, Size = 4 },
                new RightBorder() { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4 }
            )
        ));

        List<string> firstHeaderRow = new List<string> { "№ з/п", "Назва робіт", "Термін виконання", "", "Примітки" };
        List<string> secondHeaderRow = new List<string> { "", "", "з", "по", "" };

        table.Append(CreateRow(firstHeaderRow, isHeader: true));
        table.Append(CreateRow(secondHeaderRow, isHeader: true));

        int rowsAdded = 2;
        foreach (var row in rows.Take(maxRowsCount))
        {
            row.Insert(0, (rowsAdded - 1).ToString());
            row.Insert(row.Count(), "");
            table.Append(CreateRow(row));
            rowsAdded++;
        }

        while (rowsAdded < minRowsCount)
        {
            table.Append(CreateRow(new List<string> { "", "", "", "", "" }));
            rowsAdded++;
        }

        return table;
    }

    private static TableRow CreateRow(List<string> columns, bool isHeader = false)
    {
        TableRow row = new TableRow();
        int[] columnWidths = { 10, 50, 15, 15, 10 };

        for (int i = 0; i < columns.Count; i++)
        {
            row.Append(CreateCell(columns[i], isHeader, columnWidths[i]));
        }

        return row;
    }

    private static TableCell CreateCell(string text, bool isHeader, int widthPercentage)
    {
        return new TableCell(
            new TableCellProperties(
                new TableCellWidth() { Width = $"{widthPercentage}%", Type = TableWidthUnitValues.Pct },
                new TableCellBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 6 },
                    new BottomBorder() { Val = BorderValues.Single, Size = 6 },
                    new LeftBorder() { Val = BorderValues.Single, Size = 6 },
                    new RightBorder() { Val = BorderValues.Single, Size = 6 }
                ),
                new Shading() { Val = ShadingPatternValues.Clear, Fill = "FFFFFF" }
            ),
            new Paragraph(
                new Run(
                    new RunProperties(
                        new Bold() { Val = isHeader ? OnOffValue.FromBoolean(true) : OnOffValue.FromBoolean(false) }
                    ),
                    new Text(text) { Space = SpaceProcessingModeValues.Preserve }
                )
            )
        );
    }

    public static Table CreateSimpleTableBasedOnLines(List<string> lines,
                                                int minRowsCount = 23, int maxRowsCount = 23)
    {
        Table table = new Table(new TableProperties(
            new TableWidth() { Width = "100%", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Size = 4 },
                new BottomBorder() { Val = BorderValues.Single, Size = 4 },
                new LeftBorder() { Val = BorderValues.Single, Size = 4 },
                new RightBorder() { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4 }
            )
        ));

        table.AppendChild(new TableGrid(new GridColumn() { Width = "5000" }));

        int rowsAdded = 0;

        foreach (string line in lines.Take(maxRowsCount))
        {
            AddRow(table, line);
            rowsAdded++;
        }

        while (rowsAdded < minRowsCount)
        {
            AddRow(table, "");
            rowsAdded++;
        }

        return table;
    }

    private static void AddRow(Table table, string text)
    {
        TableRow row = new TableRow();
        TableCell cell = new TableCell(
            new TableCellProperties(
                new TableCellWidth() { Width = "100%", Type = TableWidthUnitValues.Pct },
                new TableCellMargin(
                new TopMargin() { Width = "20", Type = TableWidthUnitValues.Dxa }, // 1pt (20 twips)
                new BottomMargin() { Width = "20", Type = TableWidthUnitValues.Dxa },
                new LeftMargin() { Width = "40", Type = TableWidthUnitValues.Dxa },
                new RightMargin() { Width = "40", Type = TableWidthUnitValues.Dxa }
            ),
                new Shading() { Val = ShadingPatternValues.Clear, Fill = "FFFFFF" }
            ),
            new Paragraph(
                new Run(
                    new RunProperties(
                        new FontSize() { Val = "22" },
                        new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
                    ),
                    new Text(text) { Space = SpaceProcessingModeValues.Preserve }
                )
            )

        );
        row.Append(cell);
        table.Append(row);
    }

}
