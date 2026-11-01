using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

var options = Options.Parse(args);

if (!File.Exists(options.TemplatePath))
{
    Console.Error.WriteLine($"Template not found: {options.TemplatePath}");
    return 1;
}

Directory.CreateDirectory(options.OutputDirectory);
var startMonth = new DateTime(options.Year, options.Month ?? DateTime.Today.Month, 1);
var months = options.Month is null
    ? [startMonth, startMonth.AddMonths(1)]
    : new[] { startMonth };
var types = options.Type == "both" ? ["payroll", "project"] : new[] { options.Type };

foreach (var month in months)
{
    foreach (var type in types)
    {
        var templatePath = options.Type != "both" && type == options.Type
            ? options.TemplatePath
            : Directory.EnumerateFiles(
                Environment.CurrentDirectory,
                type == "project" ? "Project Timesheet*.xlsx" : "Payroll Timesheet*.xlsx")
                .FirstOrDefault()
                ?? throw new ArgumentException($"No {type} .xlsx template found.");
        var outputPath = Path.Combine(
            options.OutputDirectory,
            WorkbookGenerator.GetFileName(templatePath, month, type));

        WorkbookGenerator.Generate(
            templatePath,
            outputPath,
            month,
            type,
            options.PublicHolidays,
            options.MedicalLeave,
            options.AnnualLeave);
        Console.WriteLine($"Created {outputPath}");
    }
}

return 0;

internal sealed record Options(
    string TemplatePath,
    string OutputDirectory,
    int Year,
    int? Month,
    string Type,
    IReadOnlySet<int> PublicHolidays,
    IReadOnlySet<int> MedicalLeave,
    IReadOnlySet<int> AnnualLeave)
{
    private static readonly int DefaultYear = DateTime.Today.Year;

    public static Options Parse(string[] args)
    {
        var type = (GetValue(args, "--type") ?? "both").ToLowerInvariant();
        if (type is not ("payroll" or "project" or "both"))
            throw new ArgumentException($"Invalid type: {type}. Use payroll, project, or both.");
        var pattern = type == "project" ? "Project Timesheet*.xlsx" : "Payroll Timesheet*.xlsx";
        var template = Path.GetFullPath(
            GetValue(args, "--template") ??
            Directory.EnumerateFiles(Environment.CurrentDirectory, pattern).FirstOrDefault() ??
            throw new ArgumentException($"No {type} .xlsx template found. Use --template <path>."));
        var output = Path.GetFullPath(GetValue(args, "--output") ?? "generated");
        var year = ParseInt(GetValue(args, "--year"), "year", DefaultYear, 1, 9999);
        var monthText = GetValue(args, "--month");
        int? month = monthText is null ? null : ParseInt(monthText, "month", 1, 1, 12);
        var holidays = ParsePublicHolidays(GetValue(args, "--public-holidays"));
        var medicalLeave = ParsePublicHolidays(GetValue(args, "--medical-leave"));
        var annualLeave = ParsePublicHolidays(GetValue(args, "--annual-leave"));
        return new Options(template, output, year, month, type, holidays, medicalLeave, annualLeave);
    }

    private static string? GetValue(string[] args, string name)
    {
        var index = Array.IndexOf(args, name);
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static int ParseInt(string? value, string name, int fallback, int min, int max)
    {
        if (value is null)
            return fallback;
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var result) ||
            result < min || result > max)
            throw new ArgumentException($"Invalid {name}: {value}");
        return result;
    }

    private static IReadOnlySet<int> ParsePublicHolidays(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return new HashSet<int>();

        var holidays = new HashSet<int>();
        foreach (var item in value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            holidays.Add(ParseInt(item, "public holiday day", 0, 1, 31));
        return holidays;
    }
}

internal static class WorkbookGenerator
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static string GetFileName(string templatePath, DateTime month, string type)
    {
        var stem = Path.GetFileNameWithoutExtension(templatePath);
        var start = type == "project"
            ? month.AddDays(2)
            : month;
        var end = type == "project"
            ? month.AddMonths(1).AddDays(1)
            : new DateTime(month.Year, month.Month, DateTime.DaysInMonth(month.Year, month.Month));
        var marker = stem.IndexOf(type == "project" ? " 3 " : " 1 ", StringComparison.OrdinalIgnoreCase);
        var prefix = marker >= 0 ? stem[..marker] : stem;
        return type == "project"
            ? $"{prefix} 3 {start.ToString("MMM", CultureInfo.InvariantCulture)} to {end.Day} {end.ToString("MMM", CultureInfo.InvariantCulture)}.xlsx"
            : $"{prefix} 1 {start.ToString("MMM", CultureInfo.InvariantCulture)} to {end.Day}.xlsx";
    }

    public static void Generate(
        string templatePath,
        string outputPath,
        DateTime month,
        string type,
        IReadOnlySet<int> publicHolidays,
        IReadOnlySet<int> medicalLeave,
        IReadOnlySet<int> annualLeave)
    {
        File.Copy(templatePath, outputPath, overwrite: true);
        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var entry = archive.GetEntry("xl/worksheets/sheet1.xml")
            ?? throw new InvalidDataException("The template does not contain xl/worksheets/sheet1.xml.");
        var stylesEntry = archive.GetEntry("xl/styles.xml")
            ?? throw new InvalidDataException("The template does not contain xl/styles.xml.");

        XDocument document;
        using (var stream = entry.Open())
            document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        XDocument styles;
        using (var stream = stylesEntry.Open())
            styles = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        var whiteStyles = AddFillStyles(styles, FindSolidFill(styles, "FFFFFFFF"));
        var grayStyles = AddGrayStyles(styles);
        var noBorderStyles = AddNoBorderStyles(styles);

        var sheetData = document.Root?.Element(Main + "sheetData")
            ?? throw new InvalidDataException("The worksheet does not contain sheet data.");
        var cells = sheetData.Descendants(Main + "c")
            .ToDictionary(cell => (string?)cell.Attribute("r") ?? "", StringComparer.OrdinalIgnoreCase);

        var rangeStart = type == "project" ? month.AddDays(2) : month;
        var rangeEnd = type == "project"
            ? month.AddMonths(1).AddDays(1)
            : new DateTime(month.Year, month.Month, DateTime.DaysInMonth(month.Year, month.Month));
        SetNumber(cells, "S5", ExcelSerial(rangeStart));
        SetNumber(cells, "AA5", ExcelSerial(rangeEnd));

        var days = (rangeEnd - rangeStart).Days + 1;
        if (days > 31)
            throw new InvalidDataException("The template supports a maximum of 31 daily columns.");
        if (days == 31)
        {
            CloneCell(sheetData, cells, "AF8", "AG8");
            CloneCell(sheetData, cells, "AF9", "AG9");
        }
        for (var index = 0; index < days; index++)
        {
            var date = rangeStart.AddDays(index);
            SetNumber(cells, ColumnName(index + 3) + "8", date.Day);
        }
        for (var index = days; index < 31; index++)
        {
            var column = ColumnName(index + 3);
            for (var row = 8; row <= 14; row++)
            {
                var address = column + row;
                ClearCell(cells, address);
                SetNoBorder(cells, noBorderStyles, address);
            }
        }

        var workingDays = 0;
        for (var index = 0; index < days; index++)
        {
            var date = rangeStart.AddDays(index);
            var cellAddress = ColumnName(index + 3) + "9";
            var column = ColumnName(index + 3);
            var dateAddress = column + "8";
            SetStyle(cells, whiteStyles, dateAddress);
            SetStyle(cells, whiteStyles, cellAddress);
            if (publicHolidays.Contains(date.Day))
            {
                SetText(cells, cellAddress, "PH");
                SetGray(cells, grayStyles, dateAddress);
                SetGray(cells, grayStyles, cellAddress);
            }
            else if (medicalLeave.Contains(date.Day))
            {
                SetText(cells, cellAddress, "MC");
                SetGray(cells, grayStyles, dateAddress);
                SetGray(cells, grayStyles, cellAddress);
            }
            else if (annualLeave.Contains(date.Day))
            {
                SetText(cells, cellAddress, "AL");
                SetGray(cells, grayStyles, dateAddress);
                SetGray(cells, grayStyles, cellAddress);
            }
            else if (date.DayOfWeek == DayOfWeek.Saturday)
            {
                SetText(cells, cellAddress, "SA");
                SetGray(cells, grayStyles, dateAddress);
                SetGray(cells, grayStyles, cellAddress);
            }
            else if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                SetText(cells, cellAddress, "SU");
                SetGray(cells, grayStyles, dateAddress);
                SetGray(cells, grayStyles, cellAddress);
            }
            else
            {
                SetNumber(cells, cellAddress, 8);
                workingDays++;
            }
        }

        SetNumber(cells, "E14", workingDays);
        SetNumber(cells, "Q14", workingDays);
        UpdateDimension(document, days == 31 ? "AG" : "AF");

        entry.Delete();
        var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Optimal);
        using var output = replacement.Open();
        document.Save(output);

        stylesEntry.Delete();
        var stylesReplacement = archive.CreateEntry("xl/styles.xml", CompressionLevel.Optimal);
        using var stylesOutput = stylesReplacement.Open();
        styles.Save(stylesOutput);
    }

    private static Dictionary<int, int> AddFillStyles(XDocument styles, int fillId)
    {
        var cellXfs = styles.Root?.Element(Main + "cellXfs")
            ?? throw new InvalidDataException("The template does not contain cell styles.");
        var existing = cellXfs.Elements(Main + "xf").ToList();
        var grayStyles = new Dictionary<int, int>();
        var originalCount = existing.Count;

        for (var index = 0; index < originalCount; index++)
        {
            var style = existing[index];
            if (style.Attribute("fillId")?.Value == fillId.ToString(CultureInfo.InvariantCulture))
            {
                grayStyles[index] = index;
                continue;
            }

            var grayStyle = new XElement(style);
            grayStyle.SetAttributeValue("fillId", fillId);
            grayStyle.SetAttributeValue("applyFill", "1");
            cellXfs.Add(grayStyle);
            grayStyles[index] = cellXfs.Elements(Main + "xf").Count() - 1;
        }

        cellXfs.SetAttributeValue("count", cellXfs.Elements(Main + "xf").Count());
        return grayStyles;
    }

    private static Dictionary<int, int> AddGrayStyles(XDocument styles)
    {
        return AddFillStyles(styles, FindSolidFill(styles, "FFD8D8D8"));
    }

    private static int FindSolidFill(XDocument styles, string color)
    {
        var fills = styles.Root?.Element(Main + "fills")
            ?? throw new InvalidDataException("The template does not contain fills.");
        var index = 0;
        foreach (var fill in fills.Elements(Main + "fill"))
        {
            var pattern = fill.Element(Main + "patternFill");
            var foreground = pattern?.Element(Main + "fgColor");
            if ((string?)pattern?.Attribute("patternType") == "solid" &&
                (string?)foreground?.Attribute("rgb") == color)
                return index;
            index++;
        }

        throw new InvalidDataException($"The template does not contain fill color {color}.");
    }

    private static Dictionary<int, int> AddNoBorderStyles(XDocument styles)
    {
        var cellXfs = styles.Root?.Element(Main + "cellXfs")
            ?? throw new InvalidDataException("The template does not contain cell styles.");
        var existing = cellXfs.Elements(Main + "xf").ToList();
        var noBorderStyles = new Dictionary<int, int>();

        for (var index = 0; index < existing.Count; index++)
        {
            var style = existing[index];
            if (style.Attribute("borderId")?.Value == "0")
            {
                noBorderStyles[index] = index;
                continue;
            }

            var noBorderStyle = new XElement(style);
            noBorderStyle.SetAttributeValue("borderId", "0");
            noBorderStyle.SetAttributeValue("applyBorder", "1");
            cellXfs.Add(noBorderStyle);
            noBorderStyles[index] = cellXfs.Elements(Main + "xf").Count() - 1;
        }

        cellXfs.SetAttributeValue("count", cellXfs.Elements(Main + "xf").Count());
        return noBorderStyles;
    }

    private static void SetNumber(Dictionary<string, XElement> cells, string address, double value)
    {
        if (!cells.TryGetValue(address, out var cell))
            throw new InvalidDataException($"The template is missing cell {address}.");
        cell.SetAttributeValue("t", null);
        cell.SetElementValue(Main + "v", value.ToString(CultureInfo.InvariantCulture));
        cell.Element(Main + "f")?.Remove();
    }

    private static void ClearCell(Dictionary<string, XElement> cells, string address)
    {
        if (!cells.TryGetValue(address, out var cell))
            return;
        cell.SetAttributeValue("t", null);
        cell.Element(Main + "v")?.Remove();
        cell.Element(Main + "f")?.Remove();
    }

    private static void SetText(Dictionary<string, XElement> cells, string address, string value)
    {
        if (!cells.TryGetValue(address, out var cell))
            throw new InvalidDataException($"The template is missing cell {address}.");
        cell.SetAttributeValue("t", "inlineStr");
        cell.Element(Main + "v")?.Remove();
        cell.Element(Main + "f")?.Remove();
        cell.Element(Main + "is")?.Remove();
        cell.Add(new XElement(Main + "is", new XElement(Main + "t", value)));
    }

    private static void SetGray(
        Dictionary<string, XElement> cells,
        IReadOnlyDictionary<int, int> grayStyles,
        string address)
    {
        SetStyle(cells, grayStyles, address);
    }

    private static void SetStyle(
        Dictionary<string, XElement> cells,
        IReadOnlyDictionary<int, int> styles,
        string address)
    {
        if (!cells.TryGetValue(address, out var cell))
            return;
        if (!int.TryParse(cell.Attribute("s")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var style))
            return;
        if (styles.TryGetValue(style, out var replacement))
            cell.SetAttributeValue("s", replacement);
    }

    private static void SetNoBorder(
        Dictionary<string, XElement> cells,
        IReadOnlyDictionary<int, int> noBorderStyles,
        string address)
    {
        if (!cells.TryGetValue(address, out var cell))
            return;
        if (!int.TryParse(cell.Attribute("s")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var style))
            return;
        cell.SetAttributeValue("s", noBorderStyles[style]);
    }

    private static void CloneCell(
        XElement sheetData,
        Dictionary<string, XElement> cells,
        string sourceAddress,
        string destinationAddress)
    {
        if (cells.ContainsKey(destinationAddress))
            return;
        if (!cells.TryGetValue(sourceAddress, out var source))
            throw new InvalidDataException($"The template is missing cell {sourceAddress}.");

        var clone = new XElement(source);
        clone.SetAttributeValue("r", destinationAddress);
        clone.Element(Main + "v")?.Remove();
        clone.Element(Main + "f")?.Remove();
        var rowNumber = int.Parse(destinationAddress[2..], CultureInfo.InvariantCulture);
        var row = sheetData.Elements(Main + "row").Single(r => (string?)r.Attribute("r") == rowNumber.ToString(CultureInfo.InvariantCulture));
        row.Add(clone);
        cells[destinationAddress] = clone;
    }

    private static void UpdateDimension(XDocument document, string lastColumn)
    {
        document.Root?.Element(Main + "dimension")?.SetAttributeValue("ref", $"A1:{lastColumn}26");
    }

    private static double ExcelSerial(DateTime date)
    {
        return (date.Date - new DateTime(1899, 12, 30)).TotalDays;
    }

    private static string ColumnName(int column)
    {
        var result = new StringBuilder();
        while (column > 0)
        {
            column--;
            result.Insert(0, (char)('A' + column % 26));
            column /= 26;
        }
        return result.ToString();
    }
}
