// Validates .xlsx files against the Open XML schema, using Microsoft's own
// validator. Every other reader in this suite answers "can something read
// this"; this one answers "does this conform to the format as specified",
// which is the axis Excel enforces and LibreOffice does not.
//
// Prints one line per file, then the errors. Exit code is the number of files
// that failed, capped at 100, so the calling test can distinguish "the
// validator did not run" from "the validator found problems".

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length == 0)
{
    Console.Error.WriteLine("usage: xlsx-schema-check <file.xlsx>...");
    return 2;
}

// Office2007 is the oldest schema an .xlsx can target, so it is the strictest
// bar this package can be held to: a file valid against it is valid for every
// later version too. Validating against a newer target would quietly permit
// constructs that older Excel refuses, which is the opposite of what this
// package promises.
var validator = new OpenXmlValidator(FileFormatVersions.Office2007)
{
    // 0 means no limit. A truncated list would make a run look better than it
    // was, and these files are small enough that the full list is readable.
    MaxNumberOfErrors = 0,
};

var failed = 0;

foreach (var path in args)
{
    List<ValidationErrorInfo> errors;

    try
    {
        using var document = SpreadsheetDocument.Open(path, false);
        errors = validator.Validate(document).ToList();
    }
    catch (Exception error)
    {
        // Failing to open at all is a stronger result than failing to validate,
        // not a reason to skip the file.
        failed++;
        Console.WriteLine($"FAIL {Path.GetFileName(path)} — could not be opened");
        Console.WriteLine($"  {error.GetType().Name}: {error.Message}");
        continue;
    }

    if (errors.Count == 0)
    {
        Console.WriteLine($"ok   {Path.GetFileName(path)}");
        continue;
    }

    failed++;
    Console.WriteLine($"FAIL {Path.GetFileName(path)} — {errors.Count} schema error(s)");

    foreach (var error in errors)
    {
        Console.WriteLine($"  [{error.ErrorType}] {error.Description}");

        if (error.Part is not null)
        {
            Console.WriteLine($"    part: {error.Part.Uri}");
        }

        if (error.Path?.XPath is { Length: > 0 } xpath)
        {
            Console.WriteLine($"    path: {xpath}");
        }
    }
}

Console.WriteLine(
    failed == 0
        ? $"{args.Length} file(s) conform to the Office2007 schema"
        : $"{failed} of {args.Length} file(s) failed schema validation");

return Math.Min(failed, 100);
