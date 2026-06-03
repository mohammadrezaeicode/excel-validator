using DocumentFormat.OpenXml.Packaging;
using App.utils;
var builder = WebApplication.CreateBuilder(args);

var app = builder.Build();

app.MapPost("/validate", async (IFormFile file) =>
{
    if (file == null || file.Length == 0)
    {
        return Results.BadRequest(new
        {
            IsValid = false,
            Message = "No file uploaded"
        });
    }

    try
    {
        using var stream = file.OpenReadStream();

        using var memory = new MemoryStream();
        await stream.CopyToAsync(memory);
        memory.Position = 0;

        // using var document =
        //     SpreadsheetDocument.Open(memory, false);
        string result = Validator.validate(memory);
        Console.WriteLine(result);
        if(result.Contains("ERROR")){
            return Results.BadRequest(new{
            IsValid = false, Message = result }) ;
        }
        return Results.Ok(new
        {
            IsValid = true,
            Message = "Excel file is valid"
        });
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new
        {
            IsValid = false,
            Message = ex.Message
        });
    }
})
.DisableAntiforgery();

app.Run();