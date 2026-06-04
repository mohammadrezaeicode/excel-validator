using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public static class ExcelValidator
{
    public static ValidationResult Validate(string filePath)
    {
        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);

            // Workbook
            var workbookPart = document.WorkbookPart
                ?? throw new Exception("WorkbookPart is missing.");

            var workbook = workbookPart.Workbook;

            // Sheets
            foreach (Sheet sheet in workbook.Sheets!)
            {
                var worksheetPart =
                    (WorksheetPart)workbookPart.GetPartById(sheet.Id!);

                // Force worksheet parsing
                _ = worksheetPart.Worksheet;

                // Force row/cell parsing
                foreach (var row in worksheetPart.Worksheet.Descendants<Row>())
                {
                    foreach (var cell in row.Descendants<Cell>())
                    {
                        _ = cell.CellReference?.Value;
                    }
                }
            }

            // Shared strings
            if (workbookPart.SharedStringTablePart != null)
            {
                _ = workbookPart.SharedStringTablePart.SharedStringTable;
            }

            // Styles
            if (workbookPart.WorkbookStylesPart != null)
            {
                _ = workbookPart.WorkbookStylesPart.Stylesheet;
            }

            return new ValidationResult
            {
                Valid = true,
                Message = "Workbook opened successfully."
            };
        }
        catch (Exception ex)
        {
            return new ValidationResult
            {
                Valid = false,
                Message = ex.Message
            };
        }
    }
}