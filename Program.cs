using System.Diagnostics;
using App.utils;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddCors(o =>
    o.AddDefaultPolicy(p => p.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod())
);

var app = builder.Build();

app.MapGet(
    "/health",
    async () =>
    {
        return Results.Ok("healthy");
    }
);

app.MapPost(
        "/validate2",
        async (IFormFile file) =>
        {
            if (file == null || file.Length == 0)
                return Results.BadRequest("No file uploaded.");

            var tempFile = Path.Combine(
                Path.GetTempPath(),
                $"{Guid.NewGuid()}{Path.GetExtension(file.FileName)}"
            );

            try
            {
                await using (var stream = File.Create(tempFile))
                {
                    await file.CopyToAsync(stream);
                }

                var result = ExcelValidator.Validate(tempFile);

                return Results.Ok(result);
            }
            finally
            {
                if (File.Exists(tempFile))
                    File.Delete(tempFile);
            }
        }
    )
    .DisableAntiforgery();

app.MapPost(
        "/validate",
        async (IFormFile file) =>
        {
            if (file == null || file.Length == 0)
            {
                return Results.BadRequest(new { IsValid = false, Message = "No file uploaded" });
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
                bool result2 = Validator.IsExcelFileValid(memory);
                Console.WriteLine(result);
                Console.WriteLine(result2);
                if (result.Contains("ERROR"))
                {
                    return Results.BadRequest(new { IsValid = false, Message = result });
                }
                return Results.Ok(new { IsValid = true, Message = "Excel file is valid" });
            }
            catch (Exception ex)
            {
                return Results.BadRequest(new { IsValid = false, Message = ex.Message });
            }
        }
    )
    .DisableAntiforgery();

app.MapPost(
        "/convert",
        async (IFormFile file, HttpContext ctx) =>
        {
            bool empty = false;
            if (ctx.Request.Query.TryGetValue("empty", out var val))
            {
                empty = val != "0" && val != "false";
            }
            // ── 1. Basic file checks ──────────────────────────────────────────────
            if (file is null || file.Length == 0)
                return Results.BadRequest(new { error = "No file uploaded." });

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext is not ".xlsx" and not ".xls")
                return Results.BadRequest(new { error = "Only .xlsx / .xls files are supported." });

            if (file.Length < 100)
                return Results.BadRequest(
                    new { error = "File is too small to be a valid Excel file." }
                );

            // ── 2. Copy to temp input file ────────────────────────────────────────
            var inputPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}{ext}");
            var outputDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(outputDir);

            await using (var stream = File.Create(inputPath))
                await file.CopyToAsync(stream);

            try
            {
                // ── 3. Validate magic bytes (xlsx = ZIP PK header) ────────────────
                var header = new byte[4];
                await using (var fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                    await fs.ReadAsync(header);

                if (
                    header[0] != 0x50
                    || header[1] != 0x4B
                    || header[2] != 0x03
                    || header[3] != 0x04
                )
                    return Results.BadRequest(
                        new
                        {
                            error = "File is corrupt or not a valid .xlsx (invalid file header).",
                        }
                    );

                // ── 4. Validate with EPPlus ───────────────────────────────────────
                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(new FileInfo(inputPath)))
                    {
                        if (package.Workbook.Worksheets.Count == 0)
                            return Results.BadRequest(
                                new { error = "Excel file has no worksheets." }
                            );

                        var ws = package.Workbook.Worksheets[0];
                        if (ws.Dimension is null && !empty)
                            return Results.BadRequest(
                                new
                                {
                                    error = $"Sheet '{ws.Name}' is completely empty — nothing to render.",
                                }
                            );
                    }
                }
                catch { }
                // ── 5. Convert via LibreOffice ────────────────────────────────────
                var userProfile = Path.Combine(Path.GetTempPath(), $"lo_{Guid.NewGuid()}");
                Directory.CreateDirectory(userProfile);

                var psi = new ProcessStartInfo
                {
                    FileName = "libreoffice",
                    Arguments =
                        $"--headless -env:UserInstallation=file://{userProfile} "
                        + $"--convert-to png --outdir \"{outputDir}\" \"{inputPath}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                };

                using var proc = Process.Start(psi)!;
                var stdout = await proc.StandardOutput.ReadToEndAsync();
                var stderr = await proc.StandardError.ReadToEndAsync();
                await proc.WaitForExitAsync();

                try
                {
                    Directory.Delete(userProfile, true);
                }
                catch { }

                if (proc.ExitCode != 0)
                    return Results.Problem(
                        $"LibreOffice conversion failed (exit {proc.ExitCode}).\n{stderr}"
                    );

                // ── 6. Validate the output image ─────────────────────────────────
                var baseName = Path.GetFileNameWithoutExtension(file.FileName);
                var outputPath = Path.Combine(outputDir, baseName + ".png");

                if (!File.Exists(outputPath))
                {
                    // LibreOffice may sanitise the filename — grab whatever .png landed
                    var found = Directory.GetFiles(outputDir, "*.png").FirstOrDefault();
                    if (found is null)
                        return Results.Problem(
                            "Conversion produced no image. The file may be too corrupted to render."
                        );
                    outputPath = found;
                }

                var imageInfo = new FileInfo(outputPath);
                if (imageInfo.Length < 1024 && !empty)
                {
                    File.Delete(outputPath);
                    return Results.BadRequest(
                        new
                        {
                            error = $"Rendered image is suspiciously small ({imageInfo.Length} B) — sheet may be blank or corrupt.",
                        }
                    );
                }

                // ── 7. Stream image back ──────────────────────────────────────────
                var imageBytes = await File.ReadAllBytesAsync(outputPath);
                return Results.File(
                    imageBytes,
                    "image/png",
                    $"{Path.GetFileNameWithoutExtension(file.FileName)}.png"
                );
            }
            finally
            {
                try
                {
                    File.Delete(inputPath);
                }
                catch { }
                try
                {
                    Directory.Delete(outputDir, true);
                }
                catch { }
            }
        }
    )
    .DisableAntiforgery();

app.Run();
