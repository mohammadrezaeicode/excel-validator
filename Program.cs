using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.IO.Compression;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  validator <excel-file>");
            return 1;
        }

        string filePath = args[0];

        if (!File.Exists(filePath))
        {
            Console.WriteLine("ERROR: File not found.");
            return 1;
        }

        if (Path.GetExtension(filePath).ToLower() != ".xlsx")
        {
            Console.WriteLine("ERROR: Not an .xlsx file.");
            return 1;
        }

        // STEP 1 — ZIP validation
        try
        {
            using var zip = ZipFile.OpenRead(filePath);

            var requiredEntries = new[]
            {
                "[Content_Types].xml",
                "xl/workbook.xml"
            };

            foreach (var entry in requiredEntries)
            {
                if (!zip.Entries.Any(e => e.FullName == entry))
                {
                    Console.WriteLine($"ERROR: Missing entry: {entry}");
                    return 1;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERROR: Invalid ZIP structure.");
            Console.WriteLine(ex.Message);
            return 1;
        }

        // STEP 2 — OpenXML validation
        try
        {
            using var document =
                SpreadsheetDocument.Open(filePath, false);

            var validator = new OpenXmlValidator();

            var errors = validator.Validate(document).ToList();

            if (errors.Any())
            {
                Console.WriteLine("ERROR: OpenXML validation failed.");

                foreach (var error in errors)
                {
                    Console.WriteLine("--------------------------------");
                    Console.WriteLine($"Description: {error.Description}");
                    Console.WriteLine($"Part: {error.Part}");
                    Console.WriteLine($"Path: {error.Path.XPath}");
                }

                return 1;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERROR: Excel file is corrupted.");
            Console.WriteLine(ex.Message);
            return 1;
        }

        Console.WriteLine("SUCCESS: Excel file is valid.");

        return 0;
    }
}