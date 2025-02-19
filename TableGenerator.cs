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
    public static Table CreateDailyTasksDescriptionTable(List<List<string>> rows, int minRowsCount = 27, int maxRowsCount = 27)
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
                new Shading() { Val = ShadingPatternValues.Clear, Fill = "FFFFFF" }
            ),
            new Paragraph(
                new Run(
                    new RunProperties(
                        new Bold() { Val = isHeader ? OnOffValue.FromBoolean(true) : OnOffValue.FromBoolean(false) },
                        new FontSize() { Val = "20" },
                        new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
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
                        new FontSize() { Val = "20" },
                        new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
                    ),
                    new Text(text) { Space = SpaceProcessingModeValues.Preserve }
                )
            )

        );
        row.Append(cell);
        table.Append(row);
    }

    public enum MergeDirection
    {
        Horizontal,
        Vertical
    }

    public static void MergeCells(Table table, List<(int row, int col)> cellCoords, MergeDirection direction)
    {
        if (table == null || cellCoords.Count < 2)
            throw new ArgumentException("Invalid table or not enough cells to merge.");

        cellCoords.Sort((a, b) => direction == MergeDirection.Horizontal ? a.col.CompareTo(b.col) : a.row.CompareTo(b.row));

        if (direction == MergeDirection.Horizontal)
        {
            MergeCellsHorizontally(table, cellCoords);
        }
        else
        {
            MergeCellsVertically(table, cellCoords);
        }
    }

    private static void MergeCellsHorizontally(Table table, List<(int row, int col)> cellCoords)
    {
        int rowIdx = cellCoords[0].row;
        TableRow row = table.Elements<TableRow>().ElementAt(rowIdx);

        TableCell firstCell = row.Elements<TableCell>().ElementAt(cellCoords[0].col);
        firstCell.TableCellProperties = new TableCellProperties(new GridSpan() { Val = (int)cellCoords.Count });

        // Remove merged cells except the first one
        for (int i = 1; i < cellCoords.Count; i++)
        {
            row.RemoveChild(row.Elements<TableCell>().ElementAt(cellCoords[1].col)); 
        }
    }

    private static void MergeCellsVertically(Table table, List<(int row, int col)> cellCoords)
    {
        int colIdx = cellCoords[0].col;

        for (int i = 0; i < cellCoords.Count; i++)
        {
            TableRow row = table.Elements<TableRow>().ElementAt(cellCoords[i].row);
            TableCell cell = row.Elements<TableCell>().ElementAt(colIdx);

            if (i == 0)
            {
                // First cell starts the merge
                cell.TableCellProperties = new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart });
            }
            else
            {
                // Subsequent cells continue the merge
                cell.TableCellProperties = new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Continue });
            }
        }
    }

}
