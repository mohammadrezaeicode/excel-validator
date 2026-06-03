using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.IO.Compression;
namespace App.utils
{
  public static class Validator
  {
    public static string validate(MemoryStream stream){

        try
        {
            using var zip =  new ZipArchive(stream, ZipArchiveMode.Read);

            var requiredEntries = new[]
            {
                "[Content_Types].xml",
                "xl/workbook.xml"
            };

            foreach (var entry in requiredEntries)
            {
                if (!zip.Entries.Any(e => e.FullName == entry))
                {
                    return $"ERROR: Missing entry: {entry}";
                }
            }
        }
        catch (Exception ex)
        {
            return "ERROR: Invalid ZIP structure.\n"+ex.Message;
        }

        // STEP 2 — OpenXML validation
        try
        {
            using var document =
                SpreadsheetDocument.Open(stream, false);

            var validator = new OpenXmlValidator();

            var errors = validator.Validate(document).ToList();

            if (errors.Any())
            {
                string message="ERROR: OpenXML validation failed.\n";

                foreach (var error in errors)
                {
                    message+="--------------------------------\n";
                    message+=$"Description: {error.Description}\n";
                    message+=$"Part: {error.Part}\n";
                    message+=$"Path: {error.Path.XPath}\n";
                }

                return message;
            }
        }
        catch (Exception ex)
        {
            return "ERROR: Excel file is corrupted.\n"+ex.Message;
        }

        return "SUCCESS: Excel file is valid.";
    }
    public static string validate(string filePath){
         if (!File.Exists(filePath))
        {
            return "ERROR: File not found.";
        }

        if (Path.GetExtension(filePath).ToLower() != ".xlsx")
        {
            return "ERROR: Not an .xlsx file.";
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
                    return $"ERROR: Missing entry: {entry}";
                }
            }
        }
        catch (Exception ex)
        {
            return "ERROR: Invalid ZIP structure.\n"+ex.Message;
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
                string message="ERROR: OpenXML validation failed.\n";

                foreach (var error in errors)
                {
                    message+="--------------------------------\n";
                    message+=$"Description: {error.Description}\n";
                    message+=$"Part: {error.Part}\n";
                    message+=$"Path: {error.Path.XPath}\n";
                }

                return message;
            }
        }
        catch (Exception ex)
        {
            return "ERROR: Excel file is corrupted.\n"+ex.Message;
        }

        return "SUCCESS: Excel file is valid.";
    }
} 
}
