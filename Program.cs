// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

var cellRefReg = new System.Text.RegularExpressions.Regex("[A-Z]+", System.Text.RegularExpressions.RegexOptions.Compiled);

uint? getCellIndex(Cell cell)
{
    if (cell.CellReference == null) return null;

    var match = cellRefReg.Match(cell.CellReference.Value.ToUpper());
    if (match == null) return null;

    var columnName = match.Value;
    uint index = 0;
    foreach (char c in columnName)
    {
        index = index * 26 + ((uint)c - (uint)'A' + 1);
    }
    return index;
}

string? getCellString(Cell cell, WorkbookPart bookPart)
{
    var text = cell.CellValue.InnerText;
    if (cell.DataType?.Value == CellValues.SharedString)
    {
        var table = bookPart.SharedStringTablePart?.SharedStringTable;
        if (table != null)
        {
            return table.ElementAt(int.Parse(text)).InnerText;
        }
    }
    return text;
}

using (var doc = SpreadsheetDocument.Open("battle.xlsx", false))
{
    var bookPart = doc.WorkbookPart;
    if (bookPart == null) return;

    foreach (var sheet in bookPart.Workbook.Descendants<Sheet>())
    {
        if (sheet == null) continue;

        Console.WriteLine("Sheet: {0}", sheet.Name?.Value);
        var sheetPart = bookPart.GetPartById(sheet.Id?.Value ?? "") as WorksheetPart;
        if (sheetPart == null) continue;

        foreach (var row in sheetPart.Worksheet.Descendants<Row>())
        {
            if (row == null) continue;

            Console.WriteLine("\tRow: {0}", row.RowIndex?.Value);
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell == null) continue;
                Console.WriteLine("\t\t[{0}:{1},{2}] {3}", cell.CellReference?.Value, row.RowIndex?.Value, getCellIndex(cell), getCellString(cell, bookPart));
            }
        }
    }
}