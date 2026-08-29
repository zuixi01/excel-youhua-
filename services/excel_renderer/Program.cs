using System.Security.Cryptography;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

const string RendererVersion = "0.1.0";
if (args.Length == 1 && args[0] == "--version")
{
    Console.WriteLine(JsonSerializer.Serialize(new { name = "ExcelRenderer", version = RendererVersion }, JsonOptions.Default));
    return 0;
}

var operationResults = new List<OperationResult>();
string? temporaryOutput = null;
try
{
    var arguments = Args.Parse(args);
    ValidateInvocation(arguments);
    temporaryOutput = Path.GetFullPath(arguments.Output) + ".auditor-" + Guid.NewGuid().ToString("N") + ".tmp";
    var manifest = JsonSerializer.Deserialize<RenderManifest>(File.ReadAllText(arguments.Manifest), JsonOptions.Default)
        ?? throw new InvalidDataException("manifest is empty");
    ValidateManifest(manifest);
    var actualHash = Hash(arguments.Input);
    if (!StringComparer.OrdinalIgnoreCase.Equals(actualHash, manifest.InputSha256))
        throw new InvalidDataException("input SHA-256 does not match manifest");
    var baselineValidationErrors = PackageValidationErrors(arguments.Input).ToHashSet(StringComparer.Ordinal);
    File.Copy(arguments.Input, temporaryOutput, false);
    File.SetAttributes(temporaryOutput, FileAttributes.Normal);
    var styleCache = new Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint>();
    using (var document = SpreadsheetDocument.Open(temporaryOutput, true))
    {
        var workbookPart = document.WorkbookPart ?? throw new InvalidDataException("workbook part is missing");
        var ordered = manifest.Operations.Select((operation, index) => (operation, index)).OrderBy(item => OperationPhase(item.operation.Type))
            .ThenBy(item => item.operation.Type == "insert_column" ? -ColumnNumber(item.operation.Before ?? item.operation.After ?? "A") : 0)
            .ThenBy(item => item.index);
        foreach (var (operation, index) in ordered)
        {
            try
            {
                switch (operation.Type)
                {
                    case "mark_cell": case "mark_row": ApplyMark(workbookPart, operation, styleCache); break;
                    case "set_cell": case "set_cell_after_insert": ApplySetCell(workbookPart, operation, styleCache); break;
                    case "insert_column": InsertColumn(workbookPart, operation, styleCache); break;
                    case "append_row": AppendRow(workbookPart, operation, styleCache); break;
                    case "add_or_replace_report_sheet": AddReportSheet(workbookPart, operation, Path.GetDirectoryName(Path.GetFullPath(arguments.Manifest))!); break;
                    default: throw new InvalidDataException($"unknown operation type: {operation.Type}");
                }
                operationResults.Add(new(index, operation.Type, operation.DifferenceId, arguments.DryRun ? "previewed" : "applied", null, null));
            }
            catch (Exception exception)
            {
                operationResults.Add(new(index, operation.Type, operation.DifferenceId, "failed", "OPERATION_FAILED", exception.Message));
                throw;
            }
        }
        AddMetadataSheet(workbookPart, manifest);
        (workbookPart.Workbook ?? throw new InvalidDataException("workbook is missing")).Save();
    }
    var resultContentHash = WorkbookContentHash(temporaryOutput);
    using (var document = SpreadsheetDocument.Open(temporaryOutput, true))
    {
        UpdateMetadataResultHash(document.WorkbookPart ?? throw new InvalidDataException("workbook part is missing"), resultContentHash);
    }
    var validationErrors = PackageValidationErrors(temporaryOutput)
        .Where(item => !baselineValidationErrors.Contains(item)).Take(10).ToList();
    if (validationErrors.Count > 0)
        throw new InvalidDataException("OUTPUT_VERIFICATION_FAILED: " + String.Join(" | ", validationErrors));
    var outputHash = Hash(temporaryOutput);
    if (!arguments.DryRun) File.Move(temporaryOutput, arguments.Output, true);
    Console.WriteLine(JsonSerializer.Serialize(new { success = true, dry_run = arguments.DryRun, output_sha256 = outputHash, result_content_sha256 = resultContentHash, operation_results = operationResults }, JsonOptions.Default));
    return 0;
}
catch (Exception exception)
{
    var code = exception.Message.Contains("OUTPUT_VERIFICATION_FAILED", StringComparison.Ordinal) ? "OUTPUT_VERIFICATION_FAILED"
        : exception.Message.StartsWith("UNSUPPORTED_FEATURE:", StringComparison.Ordinal) ? "UNSUPPORTED_FEATURE"
        : exception is ArgumentException or FileNotFoundException ? "ARGUMENT_INVALID"
        : exception is InvalidDataException ? "MANIFEST_OR_STRUCTURE_INVALID" : "RENDER_FAILED";
    Console.Error.WriteLine(JsonSerializer.Serialize(new { success = false, error_code = code, message = exception.Message, operation_results = operationResults }, JsonOptions.Default));
    return 1;
}
finally
{
    if (temporaryOutput is not null && File.Exists(temporaryOutput)) File.Delete(temporaryOutput);
}

static void ApplyMark(WorkbookPart workbookPart, RenderOperation operation, Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint> styleCache)
{
    var worksheetPart = Worksheet(workbookPart, operation.Sheet!);
    if (operation.Type == "mark_cell")
    {
        var cell = FindOrCreateCell(worksheetPart, operation.Cell!);
        cell.StyleIndex = FillStyle(workbookPart, cell.StyleIndex, operation.FillColor!, null, styleCache);
        if (!String.IsNullOrWhiteSpace(operation.Comment)) AddOrReplaceComment(worksheetPart, operation.Cell!, operation.Comment!);
    }
    else
    {
        var row = FindOrCreateRow(worksheetPart, operation.Row!.Value);
        var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
        var maxColumn = sheetData?.Descendants<Cell>()
            .Select(cell => ColumnNumber(CellColumn(cell.CellReference?.Value ?? "A")))
            .DefaultIfEmpty(1).Max() ?? 1;
        for (var column = 1; column <= maxColumn; column++)
        {
            var reference = ColumnName(column) + operation.Row.Value;
            var existing = row.Elements<Cell>().FirstOrDefault(cell => String.Equals(cell.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
            var sourceStyle = existing?.StyleIndex ?? sheetData?.Elements<Row>()
                .Where(candidate => candidate.RowIndex?.Value != operation.Row.Value)
                .OrderBy(candidate => Math.Abs((long)(candidate.RowIndex?.Value ?? 0u) - operation.Row.Value))
                .SelectMany(candidate => candidate.Elements<Cell>())
                .FirstOrDefault(cell => ColumnNumber(CellColumn(cell.CellReference?.Value ?? "A")) == column)?.StyleIndex;
            var cell = existing ?? FindOrCreateCell(worksheetPart, reference);
            cell.StyleIndex = FillStyle(workbookPart, sourceStyle ?? cell.StyleIndex, operation.FillColor!, null, styleCache);
        }
        if (!String.IsNullOrWhiteSpace(operation.Comment)) AddOrReplaceComment(worksheetPart, "A" + operation.Row.Value, operation.Comment!);
    }
    (worksheetPart.Worksheet ?? throw new InvalidDataException("worksheet is missing")).Save();
}

static void ApplySetCell(WorkbookPart workbookPart, RenderOperation operation, Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint> styleCache)
{
    var worksheetPart = Worksheet(workbookPart, operation.Sheet!);
    var cell = FindOrCreateCell(worksheetPart, operation.Cell!);
    WriteSafeValue(cell, operation.Value, operation.FieldType, Uses1904DateSystem(workbookPart));
    cell.StyleIndex = FillStyle(workbookPart, cell.StyleIndex, operation.FillColor!, operation.NumberFormat, styleCache);
    if (!String.IsNullOrWhiteSpace(operation.Comment)) AddOrReplaceComment(worksheetPart, operation.Cell!, operation.Comment!);
    RecalculateDimension(worksheetPart.Worksheet!);
    (worksheetPart.Worksheet ?? throw new InvalidDataException("worksheet is missing")).Save();
}

static void AppendRow(WorkbookPart workbookPart, RenderOperation operation, Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint> styleCache)
{
    var worksheetPart = Worksheet(workbookPart, operation.Sheet!);
    var date1904 = Uses1904DateSystem(workbookPart);
    var targetRow = operation.Row is null ? null : FindOrCreateRow(worksheetPart, operation.Row.Value);
    var previousRow = operation.Row is null ? null : worksheetPart.Worksheet?.GetFirstChild<SheetData>()?.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value + 1u == operation.Row.Value);
    if (targetRow is not null && previousRow is not null)
    {
        targetRow.StyleIndex = previousRow.StyleIndex;
        targetRow.CustomFormat = previousRow.CustomFormat;
        targetRow.Height = previousRow.Height;
        targetRow.CustomHeight = previousRow.CustomHeight;
    }
    foreach (var value in operation.Values ?? [])
    {
        if (String.IsNullOrWhiteSpace(value.Cell)) continue;
        var rowIndex = UInt32.Parse(new string(value.Cell.Where(Char.IsDigit).ToArray()));
        var columnName = CellColumn(value.Cell);
        var sourceCell = rowIndex > 1 ? worksheetPart.Worksheet?.GetFirstChild<SheetData>()?.Elements<Row>()
            .FirstOrDefault(item => item.RowIndex?.Value == rowIndex - 1)?.Elements<Cell>()
            .FirstOrDefault(item => CellColumn(item.CellReference?.Value ?? "") == columnName) : null;
        var cell = FindOrCreateCell(worksheetPart, value.Cell);
        if (!String.IsNullOrWhiteSpace(value.FormulaTemplate))
        {
            cell.DataType = null; cell.CellValue = null; cell.InlineString = null;
            cell.CellFormula = new CellFormula(value.FormulaTemplate.Replace("{row}", rowIndex.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal).TrimStart('='));
        }
        else WriteSafeValue(cell, value.Value, value.FieldType, date1904);
        cell.StyleIndex = FillStyle(workbookPart, sourceCell?.StyleIndex ?? cell.StyleIndex, operation.FillColor!, value.NumberFormat, styleCache);
    }
    if (!String.IsNullOrWhiteSpace(operation.Comment) && operation.Row is not null)
        AddOrReplaceComment(worksheetPart, "A" + operation.Row.Value, operation.Comment!);
    if (operation.Row is not null) ExtendStructuresForAppendedRow(worksheetPart, operation.Row.Value);
    RecalculateDimension(worksheetPart.Worksheet!);
    (worksheetPart.Worksheet ?? throw new InvalidDataException("worksheet is missing")).Save();
}

static void WriteSafeValue(Cell cell, JsonElement? value, string? fieldType = null, bool date1904 = false)
{
    cell.CellFormula = null;
    cell.CellValue = null;
    cell.InlineString = null;
    if (value is null || value.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
    {
        cell.DataType = null;
        return;
    }
    if (value.Value.ValueKind == JsonValueKind.Number)
    {
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(value.Value.GetRawText());
        return;
    }
    if (value.Value.ValueKind is JsonValueKind.True or JsonValueKind.False)
    {
        cell.DataType = CellValues.Boolean;
        cell.CellValue = new CellValue(value.Value.GetBoolean() ? "1" : "0");
        return;
    }
    var textValue = value.Value.ToString();
    if (fieldType is "integer" or "decimal" && Double.TryParse(textValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && Double.IsFinite(number))
    {
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(number.ToString("R", CultureInfo.InvariantCulture));
        return;
    }
    if (fieldType is "date" or "datetime" && DateTimeOffset.TryParse(textValue, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal, out var timestamp))
    {
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(ToExcelSerial(timestamp.DateTime, date1904).ToString("R", CultureInfo.InvariantCulture));
        return;
    }
    // Strings are always written as inline text. Values beginning with =, +, - or
    // @ therefore cannot be interpreted as formulas by Excel.
    cell.DataType = CellValues.InlineString;
    cell.InlineString = new InlineString(new Text(textValue) { Space = SpaceProcessingModeValues.Preserve });
}

static bool Uses1904DateSystem(WorkbookPart workbookPart) => workbookPart.Workbook?.WorkbookProperties?.Date1904?.Value == true;

static double ToExcelSerial(DateTime value, bool date1904)
{
    var epoch = date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
    var serial = (value - epoch).TotalDays;
    // Excel's 1900 system deliberately preserves Lotus 1-2-3's fictitious
    // 1900-02-29. Dates before that boundary use a one-day correction.
    if (!date1904 && serial > 0d && serial <= 60d) serial -= 1d;
    return serial;
}

static void InsertColumn(WorkbookPart workbookPart, RenderOperation operation, Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint> styleCache)
{
    var worksheetPart = Worksheet(workbookPart, operation.Sheet!);
    EnsureInsertCanBeMaintained(workbookPart, worksheetPart);
    var target = operation.Before is not null ? ColumnNumber(operation.Before) : ColumnNumber(operation.After!) + 1;
    var worksheet = worksheetPart.Worksheet ?? throw new InvalidDataException("worksheet is missing");
    var sheetData = worksheet.GetFirstChild<SheetData>() ?? throw new InvalidDataException("sheet data is missing");
    foreach (var row in sheetData.Elements<Row>())
    {
        foreach (var cell in row.Elements<Cell>().OrderByDescending(cell => ColumnNumber(CellColumn(cell.CellReference?.Value ?? "A"))))
        {
            var column = ColumnNumber(CellColumn(cell.CellReference?.Value ?? "A"));
            if (column >= target) cell.CellReference = ColumnName(column + 1) + (row.RowIndex?.Value ?? throw new InvalidDataException("row index is missing"));
        }
    }
    var commentsPart = worksheetPart.WorksheetCommentsPart;
    if (commentsPart?.Comments?.CommentList is not null)
    {
        foreach (var comment in commentsPart.Comments.CommentList.Elements<Comment>())
        {
            var referenceValue = comment.Reference?.Value;
            if (String.IsNullOrEmpty(referenceValue)) continue;
            var column = ColumnNumber(CellColumn(referenceValue));
            if (column >= target) comment.Reference = ColumnName(column + 1) + new string(referenceValue.Where(Char.IsDigit).ToArray());
        }
        commentsPart.Comments.Save();
        RebuildCommentVml(worksheetPart);
    }
    UpdateColumnDependentReferences(workbookPart, worksheetPart, operation, target);
    var reference = ColumnName(target) + operation.HeaderRow;
    var header = FindOrCreateCell(worksheetPart, reference);
    header.DataType = CellValues.InlineString;
    header.InlineString = new InlineString(new Text(operation.HeaderValue ?? operation.CanonicalField ?? ""));
    var adjacentHeader = sheetData.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value == operation.HeaderRow)?.Elements<Cell>()
        .Where(item => item.CellReference?.Value != reference).OrderBy(item => Math.Abs(ColumnNumber(CellColumn(item.CellReference?.Value ?? "A")) - target)).FirstOrDefault();
    header.StyleIndex = FillStyle(workbookPart, adjacentHeader?.StyleIndex ?? header.StyleIndex, operation.FillColor!, null, styleCache);
    var lastRow = sheetData.Elements<Row>().Select(item => item.RowIndex?.Value ?? 0u).DefaultIfEmpty(0u).Max();
    var firstDataRow = operation.DataStartRow ?? (operation.HeaderRow ?? 1u) + 1u;
    var formulaRows = operation.FormulaRows?.ToHashSet();
    for (var rowIndex = firstDataRow; rowIndex <= lastRow; rowIndex++)
    {
        var dataCell = FindOrCreateCell(worksheetPart, ColumnName(target) + rowIndex);
        var adjacent = FindNearestStyledCell(worksheetPart, rowIndex, target);
        dataCell.StyleIndex = FillStyle(workbookPart, adjacent?.StyleIndex ?? dataCell.StyleIndex, "", operation.NumberFormat, styleCache);
        if (!String.IsNullOrWhiteSpace(operation.FormulaTemplate) && (formulaRows is null || formulaRows.Contains(rowIndex)))
        {
            var formula = operation.FormulaTemplate.Replace("{row}", rowIndex.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal).TrimStart('=');
            dataCell.CellFormula = new CellFormula(formula.TrimStart('='));
            dataCell.DataType = null;
            dataCell.CellValue = null;
            dataCell.InlineString = null;
        }
    }
    ApplyColumnValidation(worksheet, operation, target, firstDataRow, lastRow);
    if (!String.IsNullOrWhiteSpace(operation.Comment)) AddOrReplaceComment(worksheetPart, reference, operation.Comment!);
    RecalculateDimension(worksheet);
    worksheet.Save();
}

static Cell? FindNearestStyledCell(WorksheetPart worksheetPart, uint rowIndex, int targetColumn)
{
    var row = worksheetPart.Worksheet?.GetFirstChild<SheetData>()?.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value == rowIndex);
    return row?.Elements<Cell>().Where(item => item.StyleIndex is not null)
        .OrderBy(item => Math.Abs(ColumnNumber(CellColumn(item.CellReference?.Value ?? "A")) - targetColumn)).FirstOrDefault();
}

static void ApplyColumnValidation(Worksheet worksheet, RenderOperation operation, int targetColumn, uint firstDataRow, uint lastRow)
{
    if (operation.Validation is null || lastRow < firstDataRow) return;
    var validation = operation.Validation;
    var dataValidation = new DataValidation { AllowBlank = validation.AllowBlank, SequenceOfReferences = new ListValue<StringValue> { InnerText = $"{ColumnName(targetColumn)}{firstDataRow}:{ColumnName(targetColumn)}{lastRow}" } };
    if (validation.Type == "list" && validation.Values is { Count: > 0 })
    {
        var list = String.Join(",", validation.Values.Select(value => value.Replace("\"", "\"\"")));
        if (list.Length > 250) throw new InvalidDataException("inline validation list exceeds safe Excel limit");
        dataValidation.Type = DataValidationValues.List;
        dataValidation.Formula1 = new Formula1("\"" + list + "\"");
    }
    else if (validation.Type is "integer" or "decimal")
    {
        dataValidation.Type = validation.Type == "integer" ? DataValidationValues.Whole : DataValidationValues.Decimal;
        if (validation.Min is not null && validation.Max is not null)
        {
            dataValidation.Operator = DataValidationOperatorValues.Between;
            dataValidation.Formula1 = new Formula1(validation.Min);
            dataValidation.Formula2 = new Formula2(validation.Max);
        }
        else if (validation.Min is not null)
        {
            dataValidation.Operator = DataValidationOperatorValues.GreaterThanOrEqual;
            dataValidation.Formula1 = new Formula1(validation.Min);
        }
        else if (validation.Max is not null)
        {
            dataValidation.Operator = DataValidationOperatorValues.LessThanOrEqual;
            dataValidation.Formula1 = new Formula1(validation.Max);
        }
        else return;
    }
    else return;
    var validations = worksheet.Elements<DataValidations>().FirstOrDefault();
    if (validations is null)
    {
        validations = new DataValidations();
        var following = worksheet.ChildElements.FirstOrDefault(item => item is Hyperlinks or PrintOptions or PageMargins or PageSetup or HeaderFooter or Drawing or LegacyDrawing or TableParts or ExtensionList);
        if (following is null) worksheet.Append(validations); else worksheet.InsertBefore(validations, following);
    }
    validations.Append(dataValidation);
    validations.Count = (uint)validations.ChildElements.Count;
}

static void ExtendStructuresForAppendedRow(WorksheetPart worksheetPart, uint rowIndex)
{
    foreach (var tablePart in worksheetPart.TableDefinitionParts)
    {
        var reference = tablePart.Table?.Reference?.Value;
        if (reference is null) continue;
        var match = Regex.Match(reference, @"^(?<start>\$?[A-Z]+\$?\d+):(?<endcol>\$?[A-Z]+)\$?(?<endrow>\d+)$");
        if (!match.Success || UInt32.Parse(match.Groups["endrow"].Value) + 1u != rowIndex) continue;
        var expanded = match.Groups["start"].Value + ":" + match.Groups["endcol"].Value + rowIndex;
        tablePart.Table!.Reference = expanded;
        if (tablePart.Table.AutoFilter is not null) tablePart.Table.AutoFilter.Reference = expanded;
        tablePart.Table.Save();
    }
    var worksheet = worksheetPart.Worksheet!;
    var worksheetAutoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
    if (worksheetAutoFilter?.Reference?.Value is string filter)
    {
        var match = Regex.Match(filter, @"^(?<start>\$?[A-Z]+\$?\d+):(?<endcol>\$?[A-Z]+)\$?(?<endrow>\d+)$");
        if (match.Success && UInt32.Parse(match.Groups["endrow"].Value) + 1u == rowIndex)
            worksheetAutoFilter.Reference = match.Groups["start"].Value + ":" + match.Groups["endcol"].Value + rowIndex;
    }
    foreach (var validation in worksheet.Descendants<DataValidation>())
    {
        if (validation.SequenceOfReferences?.InnerText is not string references) continue;
        validation.SequenceOfReferences = new ListValue<StringValue> { InnerText = String.Join(" ", references.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(reference => ExtendRangeRow(reference, rowIndex))) };
    }
}

static string ExtendRangeRow(string reference, uint appendedRow)
{
    var match = Regex.Match(reference, @"^(?<start>\$?[A-Z]+\$?\d+):(?<endcol>\$?[A-Z]+)\$?(?<endrow>\d+)$");
    if (match.Success && UInt32.Parse(match.Groups["endrow"].Value) + 1u == appendedRow)
        return match.Groups["start"].Value + ":" + match.Groups["endcol"].Value + appendedRow;
    var single = Regex.Match(reference, @"^(?<cell>(?<column>\$?[A-Z]+)\$?(?<row>\d+))$");
    return single.Success && UInt32.Parse(single.Groups["row"].Value) + 1u == appendedRow
        ? single.Groups["cell"].Value + ":" + single.Groups["column"].Value + appendedRow
        : reference;
}

static void RecalculateDimension(Worksheet worksheet)
{
    var cells = worksheet.GetFirstChild<SheetData>()?.Descendants<Cell>().Select(cell => cell.CellReference?.Value).Where(value => !String.IsNullOrWhiteSpace(value)).ToList() ?? [];
    if (cells.Count == 0) return;
    var minColumn = cells.Min(value => ColumnNumber(CellColumn(value!)));
    var maxColumn = cells.Max(value => ColumnNumber(CellColumn(value!)));
    var minRow = cells.Min(value => UInt32.Parse(new string(value!.Where(Char.IsDigit).ToArray())));
    var maxRow = cells.Max(value => UInt32.Parse(new string(value!.Where(Char.IsDigit).ToArray())));
    var reference = ColumnName(minColumn) + minRow + ":" + ColumnName(maxColumn) + maxRow;
    if (worksheet.SheetDimension is null) worksheet.InsertAt(new SheetDimension { Reference = reference }, 0);
    else worksheet.SheetDimension.Reference = reference;
    foreach (var row in worksheet.GetFirstChild<SheetData>()?.Elements<Row>() ?? []) row.Spans = new ListValue<StringValue> { InnerText = $"{minColumn}:{maxColumn}" };
}

static void AddReportSheet(WorkbookPart workbookPart, RenderOperation operation, string manifestDirectory)
{
    var workbook = workbookPart.Workbook ?? throw new InvalidDataException("workbook is missing");
    var sheets = workbook.Sheets ?? workbook.AppendChild(new Sheets());
    var name = operation.Name ?? "核验报告";
    var existing = sheets.Elements<Sheet>().FirstOrDefault(item => item.Name == name);
    if (existing is not null)
    {
        var oldPart = workbookPart.GetPartById(existing.Id!);
        existing.Remove();
        workbookPart.DeletePart(oldPart);
    }
    var reportPath = Path.GetFullPath(Path.Combine(manifestDirectory, operation.SourceJson ?? "report.json"));
    var relativeReportPath = Path.GetRelativePath(manifestDirectory, reportPath);
    if (Path.IsPathRooted(relativeReportPath) || relativeReportPath == ".." || relativeReportPath.StartsWith(".." + Path.DirectorySeparatorChar)) throw new InvalidDataException("report path escapes manifest directory");
    using var report = JsonDocument.Parse(File.ReadAllText(reportPath));
    var part = workbookPart.AddNewPart<WorksheetPart>();
    var data = new SheetData();
    part.Worksheet = new Worksheet(data);
    AppendTextRow(data, 1, "项目", "值");
    uint rowIndex = 2;
    if (report.RootElement.TryGetProperty("summary", out var summary))
        foreach (var property in summary.EnumerateObject()) AppendTextRow(data, rowIndex++, property.Name, property.Value.ToString());
    rowIndex++;
    AppendTextRow(data, rowIndex++, "类型", "级别", "工作表", "单元格", "字段", "业务主键", "Excel 原值", "Excel 规范值", "标准原值", "标准规范值", "规则 ID", "动作", "修复状态", "说明");
    if (report.RootElement.TryGetProperty("differences", out var differences))
        foreach (var item in differences.EnumerateArray())
            AppendTextRow(data, rowIndex++, JsonValue(item, "type"), JsonValue(item, "severity"), JsonValue(item, "sheet_name"), JsonValue(item, "cell"), JsonValue(item, "canonical_field"), JsonValue(item, "business_key"), JsonValue(item, "excel_raw_value"), JsonValue(item, "excel_normalized_value"), JsonValue(item, "standard_raw_value"), JsonValue(item, "standard_normalized_value"), JsonValue(item, "rule_id"), JsonValue(item, "render_action"), JsonValue(item, "repair_status"), JsonValue(item, "message"));
    part.Worksheet.Save();
    var nextId = sheets.Elements<Sheet>().Select(item => item.SheetId?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1u;
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(part), SheetId = nextId, Name = name });
}

static void AddMetadataSheet(WorkbookPart workbookPart, RenderManifest manifest)
{
    const string name = "__ExcelAuditorMetadata";
    var workbook = workbookPart.Workbook ?? throw new InvalidDataException("workbook is missing");
    var sheets = workbook.Sheets ?? workbook.AppendChild(new Sheets());
    var existing = sheets.Elements<Sheet>().FirstOrDefault(item => item.Name == name);
    if (existing is not null)
    {
        var oldPart = workbookPart.GetPartById(existing.Id!);
        existing.Remove();
        workbookPart.DeletePart(oldPart);
    }
    var part = workbookPart.AddNewPart<WorksheetPart>();
    var data = new SheetData();
    part.Worksheet = new Worksheet(data);
    AppendTextRow(data, 1, "key", "value");
    var entries = new (string Key, string Value)[] {
        ("job_id", manifest.JobId),
        ("schema_id", manifest.Metadata?.SchemaId ?? ""),
        ("schema_version", manifest.Metadata?.SchemaVersion ?? ""),
        ("schema_sha256", manifest.Metadata?.SchemaSha256 ?? ""),
        ("standard_snapshot_id", manifest.Metadata?.StandardSnapshotId ?? ""),
        ("standard_sha256", manifest.Metadata?.StandardSha256 ?? ""),
        ("input_sha256", manifest.InputSha256),
        ("result_content_sha256", ""),
        ("operation_count", manifest.Operations.Count.ToString(CultureInfo.InvariantCulture)),
    };
    uint row = 2;
    foreach (var entry in entries) AppendTextRow(data, row++, entry.Key, entry.Value);
    part.Worksheet.Save();
    var nextId = sheets.Elements<Sheet>().Select(item => item.SheetId?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1u;
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(part), SheetId = nextId, Name = name, State = SheetStateValues.VeryHidden });
}

static void UpdateMetadataResultHash(WorkbookPart workbookPart, string resultContentHash)
{
    var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>().FirstOrDefault(item => item.Name == "__ExcelAuditorMetadata")
        ?? throw new InvalidDataException("metadata worksheet is missing");
    var worksheetPart = workbookPart.GetPartById(sheet.Id?.Value ?? throw new InvalidDataException("metadata worksheet relationship is missing")) as WorksheetPart
        ?? throw new InvalidDataException("metadata worksheet part is invalid");
    var cell = FindOrCreateCell(worksheetPart, "B9");
    cell.DataType = CellValues.InlineString;
    cell.CellValue = null;
    cell.CellFormula = null;
    cell.InlineString = new InlineString(new Text(resultContentHash));
    worksheetPart.Worksheet?.Save();
    workbookPart.Workbook?.Save();
}

static void AppendTextRow(SheetData data, uint index, params string[] values)
{
    var row = new Row { RowIndex = index };
    for (var column = 0; column < values.Length; column++)
        row.Append(new Cell { CellReference = ColumnName(column + 1) + index, DataType = CellValues.InlineString, InlineString = new InlineString(new Text(values[column] ?? "")) });
    data.Append(row);
}

static string JsonValue(JsonElement item, string name) => item.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null ? value.ToString() : "";

static void UpdateColumnDependentReferences(WorkbookPart workbookPart, WorksheetPart worksheetPart, RenderOperation operation, int target)
{
    var worksheet = worksheetPart.Worksheet ?? throw new InvalidDataException("worksheet is missing");
    var dimension = worksheet.SheetDimension;
    if (dimension?.Reference?.Value is string dimensionReference) dimension.Reference = ShiftRange(dimensionReference, target);
    foreach (var sheetView in worksheet.Elements<SheetViews>().SelectMany(views => views.Elements<SheetView>()))
    {
        var pane = sheetView.Pane;
        if (pane?.TopLeftCell?.Value is string topLeftCell)
            pane.TopLeftCell = ShiftCellReference(topLeftCell, target);
        if (pane?.HorizontalSplit?.Value is double horizontalSplit && target <= horizontalSplit)
            pane.HorizontalSplit = horizontalSplit + 1d;
        foreach (var selection in sheetView.Elements<Selection>())
        {
            if (selection.ActiveCell?.Value is string activeCell)
                selection.ActiveCell = ShiftCellReference(activeCell, target);
            if (selection.SequenceOfReferences?.InnerText is string references && references.Length > 0)
                selection.SequenceOfReferences = new ListValue<StringValue> { InnerText = String.Join(" ", references.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(value => ShiftRange(value, target))) };
        }
    }
    foreach (var merge in worksheet.Descendants<MergeCell>())
        if (merge.Reference?.Value is string reference) merge.Reference = ShiftRange(reference, target);
    var sheetAutoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
    if (sheetAutoFilter is not null) UpdateAutoFilterForInsertedColumn(sheetAutoFilter, target);
    foreach (var validation in worksheet.Descendants<DataValidation>())
    {
        if (validation.SequenceOfReferences?.InnerText is string references && references.Length > 0)
            validation.SequenceOfReferences = new ListValue<StringValue> { InnerText = String.Join(" ", references.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(value => ShiftRange(value, target))) };
        if (validation.Formula1?.Text is string formula1) validation.Formula1.Text = ShiftFormulaForInsertedColumn(formula1, operation.Sheet ?? "", operation.Sheet ?? "", target);
        if (validation.Formula2?.Text is string formula2) validation.Formula2.Text = ShiftFormulaForInsertedColumn(formula2, operation.Sheet ?? "", operation.Sheet ?? "", target);
    }
    foreach (var conditional in worksheet.Elements<ConditionalFormatting>())
    {
        if (conditional.SequenceOfReferences?.InnerText is string references && references.Length > 0)
            conditional.SequenceOfReferences = new ListValue<StringValue> { InnerText = String.Join(" ", references.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(value => ShiftRange(value, target))) };
        foreach (var formula in conditional.Descendants<Formula>())
            if (formula.Text is string text) formula.Text = ShiftFormulaForInsertedColumn(text, operation.Sheet ?? "", operation.Sheet ?? "", target);
    }
    foreach (var hyperlink in worksheet.Descendants<Hyperlink>())
    {
        if (hyperlink.Reference?.Value is string reference) hyperlink.Reference = ShiftRange(reference, target);
        if (hyperlink.Location?.Value is string location)
            hyperlink.Location = ShiftFormulaForInsertedColumn(location, operation.Sheet ?? "", operation.Sheet ?? "", target);
    }
    foreach (var columns in worksheet.Elements<Columns>())
        foreach (var column in columns.Elements<Column>())
        {
            var min = column.Min?.Value ?? 1u;
            var max = column.Max?.Value ?? min;
            var shifted = ShiftInterval((int)min, (int)max, target);
            column.Min = (uint)shifted.Start; column.Max = (uint)shifted.End;
        }
    foreach (var tablePart in worksheetPart.TableDefinitionParts)
    {
        var table = tablePart.Table;
        if (table?.Reference?.Value is not string tableReference) continue;
        var bounds = RangeColumns(tableReference);
        table.Reference = ExpandTableRangeForInsertedColumn(tableReference, target);
        if (table.AutoFilter is not null) UpdateAutoFilterForInsertedColumn(table.AutoFilter, target);
        if (target >= bounds.Start && target <= bounds.End + 1 && table.TableColumns is not null)
        {
            var baseName = operation.HeaderValue ?? operation.CanonicalField ?? "Inserted";
            var existingNames = table.TableColumns.Elements<TableColumn>().Select(item => item.Name?.Value ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase);
            var name = baseName;
            for (var suffix = 2; existingNames.Contains(name); suffix++) name = baseName + suffix;
            var id = table.TableColumns.Elements<TableColumn>().Select(item => item.Id?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1u;
            table.TableColumns.InsertAt(new TableColumn { Id = id, Name = name }, target - bounds.Start);
            table.TableColumns.Count = (uint)table.TableColumns.ChildElements.Count;
        }
        table.Save();
    }
    var workbook = workbookPart.Workbook ?? throw new InvalidDataException("workbook is missing");
    var sheetName = operation.Sheet ?? "";
    foreach (var sheet in workbook.Sheets?.Elements<Sheet>() ?? [])
    {
        if (sheet.Id?.Value is not string relationshipId) continue;
        if (workbookPart.GetPartById(relationshipId) is not WorksheetPart formulaPart) continue;
        var currentSheetName = sheet.Name?.Value ?? "";
        foreach (var cell in formulaPart.Worksheet?.Descendants<Cell>() ?? [])
            if (cell.CellFormula?.Text is string formula)
                cell.CellFormula.Text = ShiftFormulaForInsertedColumn(formula, currentSheetName, sheetName, target);
        if (!ReferenceEquals(formulaPart, worksheetPart))
        {
            foreach (var validation in formulaPart.Worksheet?.Descendants<DataValidation>() ?? [])
            {
                if (validation.Formula1?.Text is string formula1) validation.Formula1.Text = ShiftFormulaForInsertedColumn(formula1, currentSheetName, sheetName, target);
                if (validation.Formula2?.Text is string formula2) validation.Formula2.Text = ShiftFormulaForInsertedColumn(formula2, currentSheetName, sheetName, target);
            }
            foreach (var conditional in formulaPart.Worksheet?.Elements<ConditionalFormatting>() ?? [])
                foreach (var formula in conditional.Descendants<Formula>())
                    if (formula.Text is string text) formula.Text = ShiftFormulaForInsertedColumn(text, currentSheetName, sheetName, target);
            foreach (var hyperlink in formulaPart.Worksheet?.Descendants<Hyperlink>() ?? [])
                if (hyperlink.Location?.Value is string location)
                    hyperlink.Location = ShiftFormulaForInsertedColumn(location, currentSheetName, sheetName, target);
        }
        foreach (var tablePart in formulaPart.TableDefinitionParts)
        {
            foreach (var formula in tablePart.Table?.Descendants<CalculatedColumnFormula>() ?? [])
                if (formula.Text is string text) formula.Text = ShiftFormulaForInsertedColumn(text, currentSheetName, sheetName, target);
            foreach (var formula in tablePart.Table?.Descendants<TotalsRowFormula>() ?? [])
                if (formula.Text is string text) formula.Text = ShiftFormulaForInsertedColumn(text, currentSheetName, sheetName, target);
            tablePart.Table?.Save();
        }
        formulaPart.Worksheet?.Save();
    }
    var targetSheetIndex = (workbook.Sheets?.Elements<Sheet>() ?? []).Select((item, index) => (item, index))
        .Where(item => String.Equals(item.item.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
        .Select(item => item.index).DefaultIfEmpty(-1).First();
    foreach (var definedName in workbook.DefinedNames?.Elements<DefinedName>() ?? [])
    {
        if (String.IsNullOrWhiteSpace(definedName.Text)) continue;
        definedName.Text = definedName.LocalSheetId?.Value == (uint)targetSheetIndex
            ? ShiftFormulaForInsertedColumn(definedName.Text, sheetName, sheetName, target)
            : ShiftQualifiedFormula(definedName.Text, sheetName, target);
    }
    if (workbookPart.CalculationChainPart is not null) workbookPart.DeletePart(workbookPart.CalculationChainPart);
    workbook.CalculationProperties ??= new CalculationProperties();
    workbook.CalculationProperties.ForceFullCalculation = true;
    workbook.CalculationProperties.FullCalculationOnLoad = true;
}

static void EnsureInsertCanBeMaintained(WorkbookPart workbookPart, WorksheetPart worksheetPart)
{
    if (worksheetPart.DrawingsPart is not null)
        throw new InvalidDataException("UNSUPPORTED_FEATURE: inserting a column in a worksheet with drawings or charts requires manual review");
    if (worksheetPart.PivotTableParts.Any())
        throw new InvalidDataException("UNSUPPORTED_FEATURE: inserting a column in a worksheet with a pivot table requires manual review");
    if (workbookPart.Parts.Any(item => item.OpenXmlPart.Uri.OriginalString.StartsWith("/xl/externalLinks/", StringComparison.OrdinalIgnoreCase)))
        throw new InvalidDataException("UNSUPPORTED_FEATURE: inserting a column in a workbook with external links requires manual review");
    foreach (var formula in workbookPart.WorksheetParts.SelectMany(part => part.Worksheet?.Descendants<CellFormula>() ?? []))
    {
        if (formula.FormulaType?.Value == CellFormulaValues.Array
            || formula.FormulaType?.Value == CellFormulaValues.Shared
            || HasUnsupportedFormulaReference(formula.Text ?? ""))
            throw new InvalidDataException("UNSUPPORTED_FEATURE: complex, shared, or array formulas cannot be safely rewritten during column insertion");
    }
    foreach (var worksheet in workbookPart.WorksheetParts)
    {
        foreach (var validation in worksheet.Worksheet?.Descendants<DataValidation>() ?? [])
        {
            EnsureDependentFormulaCanBeShifted(validation.Formula1?.Text, "data validation");
            EnsureDependentFormulaCanBeShifted(validation.Formula2?.Text, "data validation");
        }
        foreach (var conditional in worksheet.Worksheet?.Elements<ConditionalFormatting>() ?? [])
            foreach (var formula in conditional.Descendants<Formula>())
                EnsureDependentFormulaCanBeShifted(formula.Text, "conditional formatting");
        foreach (var hyperlink in worksheet.Worksheet?.Descendants<Hyperlink>() ?? [])
            EnsureDependentFormulaCanBeShifted(hyperlink.Location?.Value, "internal hyperlink");
    }
    foreach (var definedName in workbookPart.Workbook?.DefinedNames?.Elements<DefinedName>() ?? [])
        EnsureDependentFormulaCanBeShifted(definedName.Text, "defined name");
}

static void EnsureDependentFormulaCanBeShifted(string? formula, string source)
{
    if (!String.IsNullOrWhiteSpace(formula) && HasUnsupportedFormulaReference(formula))
        throw new InvalidDataException($"UNSUPPORTED_FEATURE: complex {source} formula cannot be safely rewritten during column insertion");
}

static bool HasUnsupportedFormulaReference(string formula) =>
    formula.Contains('[', StringComparison.Ordinal)
    || formula.Contains('#', StringComparison.Ordinal)
    || formula.Contains('{', StringComparison.Ordinal)
    || Regex.IsMatch(formula, @"(?:^|[^A-Z0-9_])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?:$|[^A-Z0-9_])", RegexOptions.IgnoreCase)
    || Regex.IsMatch(formula, @"(?:^|[^0-9])\$?\d+:\$?\d+(?:$|[^0-9])")
    || Regex.IsMatch(formula, @"[^!]+:[^!]+!")
    || Regex.IsMatch(formula, @"\b[A-Z_][A-Z0-9_.]*\[[^\]]+\]", RegexOptions.IgnoreCase);

static void UpdateAutoFilterForInsertedColumn(AutoFilter autoFilter, int target)
{
    var originalReference = autoFilter.Reference?.Value;
    if (String.IsNullOrWhiteSpace(originalReference)) return;
    var bounds = RangeColumns(originalReference);
    var insertedInside = target >= bounds.Start && target <= bounds.End + 1;
    var relativeIndex = target - bounds.Start;
    if (insertedInside && relativeIndex <= bounds.End - bounds.Start)
    {
        foreach (var filterColumn in autoFilter.Elements<FilterColumn>())
            if (filterColumn.ColumnId?.Value is uint columnId && columnId >= relativeIndex)
                filterColumn.ColumnId = columnId + 1u;
    }
    autoFilter.Reference = ExpandTableRangeForInsertedColumn(originalReference, target);
    foreach (var sortState in autoFilter.Descendants<SortState>())
    {
        if (sortState.Reference?.Value is string sortReference) sortState.Reference = ShiftRange(sortReference, target);
        foreach (var condition in sortState.Descendants<SortCondition>())
            if (condition.Reference?.Value is string conditionReference) condition.Reference = ShiftRange(conditionReference, target);
    }
}

static string ExpandTableRangeForInsertedColumn(string reference, int target)
{
    var match = Regex.Match(reference, @"^(?<startabs>\$?)(?<startcol>[A-Z]+)(?<startrow>\$?\d+):(?<endabs>\$?)(?<endcol>[A-Z]+)(?<endrow>\$?\d+)$");
    if (!match.Success) return ShiftRange(reference, target);
    var start = ColumnNumber(match.Groups["startcol"].Value);
    var end = ColumnNumber(match.Groups["endcol"].Value);
    if (target < start || target > end + 1) return ShiftRange(reference, target);
    var newStart = target == start ? start : start;
    var newEnd = end + 1;
    return match.Groups["startabs"].Value + ColumnName(newStart) + match.Groups["startrow"].Value + ":" + match.Groups["endabs"].Value + ColumnName(newEnd) + match.Groups["endrow"].Value;
}

static string ShiftFormulaForInsertedColumn(string formula, string currentSheetName, string targetSheetName, int target)
{
    const string pattern = @"(?<![A-Z0-9_\]\[])(?:(?<sheet>'(?:[^']|'')+'|[^'""\s!+\-*/^&=(),;:{}\[\]]+)!)?(?<reference>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![A-Z0-9_])";
    return ReplaceOutsideStringLiterals(formula, pattern, match =>
    {
        var qualifier = match.Groups["sheet"];
        if (!qualifier.Success)
            return String.Equals(currentSheetName, targetSheetName, StringComparison.OrdinalIgnoreCase)
                ? ShiftRange(match.Groups["reference"].Value, target)
                : match.Value;
        var qualifiedSheet = qualifier.Value;
        if (qualifiedSheet.StartsWith("'", StringComparison.Ordinal) && qualifiedSheet.EndsWith("'", StringComparison.Ordinal))
            qualifiedSheet = qualifiedSheet[1..^1].Replace("''", "'");
        return String.Equals(qualifiedSheet, targetSheetName, StringComparison.OrdinalIgnoreCase)
            ? qualifier.Value + "!" + ShiftRange(match.Groups["reference"].Value, target)
            : match.Value;
    });
}

static string ShiftQualifiedFormula(string formula, string sheetName, int target)
{
    var quoted = "'" + sheetName.Replace("'", "''") + "'!";
    var plain = sheetName + "!";
    var pattern = $@"(?<sheet>{Regex.Escape(quoted)}|{Regex.Escape(plain)})(?<reference>\$?[A-Z]{{1,3}}\$?\d+(?::\$?[A-Z]{{1,3}}\$?\d+)?)";
    return ReplaceOutsideStringLiterals(formula, pattern, match => match.Groups["sheet"].Value + ShiftRange(match.Groups["reference"].Value, target));
}

static string ReplaceOutsideStringLiterals(string formula, string pattern, MatchEvaluator evaluator)
{
    var output = new System.Text.StringBuilder(formula.Length + 8);
    var chunkStart = 0;
    var index = 0;
    while (index < formula.Length)
    {
        if (formula[index] != '"') { index++; continue; }
        output.Append(Regex.Replace(formula[chunkStart..index], pattern, evaluator));
        var literalStart = index++;
        while (index < formula.Length)
        {
            if (formula[index] != '"') { index++; continue; }
            if (index + 1 < formula.Length && formula[index + 1] == '"') { index += 2; continue; }
            index++;
            break;
        }
        output.Append(formula[literalStart..index]);
        chunkStart = index;
    }
    output.Append(Regex.Replace(formula[chunkStart..], pattern, evaluator));
    return output.ToString();
}

static string ShiftRange(string reference, int target)
{
    var parts = reference.Split(':', 2);
    var first = ShiftCellReference(parts[0], target);
    if (parts.Length == 1) return first;
    var second = ShiftCellReference(parts[1], target);
    var startColumn = ColumnNumber(CellColumn(parts[0].Replace("$", "")));
    var endColumn = ColumnNumber(CellColumn(parts[1].Replace("$", "")));
    if (startColumn < target && endColumn >= target) second = ShiftCellReference(parts[1], target, force: true);
    return first + ":" + second;
}

static string ShiftCellReference(string reference, int target, bool force = false)
{
    var match = Regex.Match(reference, @"^(?<absolute>\$?)(?<column>[A-Z]{1,3})(?<row>\$?\d+)?$");
    if (!match.Success) return reference;
    var column = ColumnNumber(match.Groups["column"].Value);
    if (!force && column < target) return reference;
    return match.Groups["absolute"].Value + ColumnName(column + 1) + match.Groups["row"].Value;
}

static (int Start, int End) RangeColumns(string reference)
{
    var parts = reference.Replace("$", "").Split(':', 2);
    var start = ColumnNumber(CellColumn(parts[0]));
    var end = parts.Length == 2 ? ColumnNumber(CellColumn(parts[1])) : start;
    return (start, end);
}

static (int Start, int End) ShiftInterval(int start, int end, int target)
{
    if (start >= target) return (start + 1, end + 1);
    if (end >= target) return (start, end + 1);
    return (start, end);
}

static void AddOrReplaceComment(WorksheetPart worksheetPart, string reference, string text)
{
    var commentsPart = worksheetPart.WorksheetCommentsPart ?? worksheetPart.AddNewPart<WorksheetCommentsPart>();
    if (commentsPart.Comments is null)
        commentsPart.Comments = new Comments(new Authors(new Author("Excel Auditor")), new CommentList());
    var authors = commentsPart.Comments.Authors ?? commentsPart.Comments.PrependChild(new Authors());
    var authorList = authors.Elements<Author>().ToList();
    var authorId = authorList.FindIndex(author => author.Text == "Excel Auditor");
    if (authorId < 0) { authors.Append(new Author("Excel Auditor")); authorId = authorList.Count; }
    var list = commentsPart.Comments.CommentList ?? commentsPart.Comments.AppendChild(new CommentList());
    list.Elements<Comment>().FirstOrDefault(item => item.Reference?.Value == reference)?.Remove();
    list.Append(new Comment
    {
        Reference = reference,
        AuthorId = (uint)authorId,
        CommentText = new CommentText(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }))
    });
    commentsPart.Comments.Save();
    RebuildCommentVml(worksheetPart);
}

static void RebuildCommentVml(WorksheetPart worksheetPart)
{
    var comments = worksheetPart.WorksheetCommentsPart?.Comments?.CommentList?.Elements<Comment>().ToList() ?? [];
    var vmlPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
    if (vmlPart is not null)
    {
        using var existingStream = vmlPart.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(existingStream);
        var existingXml = reader.ReadToEnd();
        try
        {
            XNamespace existingVml = "urn:schemas-microsoft-com:vml";
            XNamespace existingExcel = "urn:schemas-microsoft-com:office:excel";
            var document = XDocument.Parse(existingXml, LoadOptions.PreserveWhitespace);
            var unsafeShape = document.Descendants(existingVml + "shape").Any(shape =>
            {
                var clientData = shape.Descendants(existingExcel + "ClientData").FirstOrDefault();
                return clientData is null || !StringComparer.OrdinalIgnoreCase.Equals((string?)clientData.Attribute("ObjectType"), "Note");
            });
            if (unsafeShape || document.Descendants(existingExcel + "Macro").Any())
                throw new InvalidDataException("legacy VML shapes or controls cannot be safely rewritten");
        }
        catch (System.Xml.XmlException exception)
        {
            throw new InvalidDataException("legacy VML cannot be parsed safely", exception);
        }
    }
    if (vmlPart is null)
    {
        vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>();
        var worksheet = worksheetPart.Worksheet ?? throw new InvalidDataException("worksheet is missing");
        worksheet.Append(new LegacyDrawing { Id = worksheetPart.GetIdOfPart(vmlPart) });
    }
    XNamespace v = "urn:schemas-microsoft-com:vml";
    XNamespace o = "urn:schemas-microsoft-com:office:office";
    XNamespace x = "urn:schemas-microsoft-com:office:excel";
    var root = new XElement("xml",
        new XAttribute(XNamespace.Xmlns + "v", v), new XAttribute(XNamespace.Xmlns + "o", o), new XAttribute(XNamespace.Xmlns + "x", x),
        new XElement(o + "shapelayout", new XAttribute(v + "ext", "edit"), new XElement(o + "idmap", new XAttribute(v + "ext", "edit"), new XAttribute("data", "1"))),
        new XElement(v + "shapetype", new XAttribute("id", "_x0000_t202"), new XAttribute("coordsize", "21600,21600"), new XAttribute(o + "spt", "202"), new XAttribute("path", "m,l,21600r21600,l21600,xe"),
            new XElement(v + "stroke", new XAttribute("joinstyle", "miter")), new XElement(v + "path", new XAttribute("gradientshapeok", "t"), new XAttribute(o + "connecttype", "rect"))));
    var shapeId = 1025;
    foreach (var comment in comments)
    {
        var reference = comment.Reference?.Value ?? "A1";
        var row = UInt32.Parse(new string(reference.Where(Char.IsDigit).ToArray())) - 1u;
        var column = (uint)ColumnNumber(CellColumn(reference)) - 1u;
        root.Add(new XElement(v + "shape", new XAttribute("id", $"_x0000_s{shapeId++}"), new XAttribute("type", "#_x0000_t202"), new XAttribute("style", "position:absolute;margin-left:80pt;margin-top:5pt;width:108pt;height:59pt;z-index:1;visibility:hidden"), new XAttribute("fillcolor", "#ffffe1"), new XAttribute(o + "insetmode", "auto"),
            new XElement(v + "fill", new XAttribute("color2", "#ffffe1")), new XElement(v + "shadow", new XAttribute("on", "t"), new XAttribute("color", "black"), new XAttribute("obscured", "t")), new XElement(v + "path", new XAttribute(o + "connecttype", "none")),
            new XElement(v + "textbox", new XAttribute("style", "mso-direction-alt:auto"), new XElement("div", new XAttribute("style", "text-align:left"))),
            new XElement(x + "ClientData", new XAttribute("ObjectType", "Note"), new XElement(x + "MoveWithCells"), new XElement(x + "SizeWithCells"), new XElement(x + "Anchor", $"{column}, 15, {row}, 2, {column + 3}, 15, {row + 4}, 4"), new XElement(x + "AutoFill", "False"), new XElement(x + "Row", row), new XElement(x + "Column", column))));
    }
    using var stream = new MemoryStream();
    new XDocument(root).Save(stream, SaveOptions.DisableFormatting);
    stream.Position = 0;
    vmlPart.FeedData(stream);
}

static WorksheetPart Worksheet(WorkbookPart workbookPart, string name)
{
    var workbook = workbookPart.Workbook ?? throw new InvalidDataException("workbook is missing");
    var sheet = workbook.Sheets?.Elements<Sheet>().SingleOrDefault(item => item.Name == name)
        ?? throw new InvalidDataException($"sheet not found: {name}");
    return (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
}

static Row FindOrCreateRow(WorksheetPart part, uint index)
{
    var worksheet = part.Worksheet ?? throw new InvalidDataException("worksheet is missing");
    var data = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
    var row = data.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value == index);
    if (row is not null) return row;
    row = new Row { RowIndex = index };
    var following = data.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value > index);
    if (following is null) data.Append(row); else data.InsertBefore(row, following);
    return row;
}

static Cell FindOrCreateCell(WorksheetPart part, string reference)
{
    var rowIndex = UInt32.Parse(new string(reference.Where(Char.IsDigit).ToArray()));
    var row = FindOrCreateRow(part, rowIndex);
    var cell = row.Elements<Cell>().FirstOrDefault(item => item.CellReference?.Value == reference);
    if (cell is not null) return cell;
    cell = new Cell { CellReference = reference };
    var column = ColumnNumber(CellColumn(reference));
    var following = row.Elements<Cell>().FirstOrDefault(item => ColumnNumber(CellColumn(item.CellReference?.Value ?? "A")) > column);
    if (following is null) row.Append(cell); else row.InsertBefore(cell, following);
    return cell;
}

static uint FillStyle(WorkbookPart workbookPart, UInt32Value? sourceStyleIndex, string? rgb, string? numberFormat, Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint> cache)
{
    rgb ??= "";
    numberFormat ??= "";
    var sourceIndex = sourceStyleIndex?.Value ?? 0;
    var cacheKey = (sourceIndex, rgb, numberFormat);
    if (cache.TryGetValue(cacheKey, out var cached)) return cached;
    var styles = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
    styles.Stylesheet ??= new Stylesheet(new Fonts(new Font()) { Count = 1 }, new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })) { Count = 1 }, new Borders(new Border()) { Count = 1 }, new CellFormats(new CellFormat()) { Count = 1 });
    var formats = styles.Stylesheet.CellFormats ?? styles.Stylesheet.AppendChild(new CellFormats());
    var source = formats.Elements<CellFormat>().ElementAtOrDefault((int)sourceIndex) ?? new CellFormat();
    var format = (CellFormat)source.CloneNode(true);
    if (!String.IsNullOrWhiteSpace(rgb))
    {
        var fills = styles.Stylesheet.Fills ?? styles.Stylesheet.AppendChild(new Fills());
        var existingFill = fills.Elements<Fill>().Select((fill, index) => (fill, index)).FirstOrDefault(item =>
            StringComparer.OrdinalIgnoreCase.Equals(item.fill.PatternFill?.ForegroundColor?.Rgb?.Value, "FF" + rgb) && item.fill.PatternFill?.PatternType?.Value == PatternValues.Solid);
        uint fillId;
        if (existingFill.fill is not null) fillId = (uint)existingFill.index;
        else
        {
            fills.Append(new Fill(new PatternFill(new ForegroundColor { Rgb = "FF" + rgb }, new BackgroundColor { Indexed = 64 }) { PatternType = PatternValues.Solid }));
            fills.Count = (uint)fills.ChildElements.Count;
            fillId = fills.Count - 1;
        }
        format.FillId = fillId; format.ApplyFill = true;
    }
    if (!String.IsNullOrWhiteSpace(numberFormat))
    {
        var numberFormats = styles.Stylesheet.NumberingFormats ?? styles.Stylesheet.InsertAt(new NumberingFormats(), 0);
        var existing = numberFormats.Elements<NumberingFormat>().FirstOrDefault(item => item.FormatCode?.Value == numberFormat);
        uint formatId;
        if (existing is not null) formatId = existing.NumberFormatId?.Value ?? 164u;
        else
        {
            var used = numberFormats.Elements<NumberingFormat>().Select(item => item.NumberFormatId?.Value ?? 163u).ToHashSet();
            formatId = 164u; while (used.Contains(formatId)) formatId++;
            numberFormats.Append(new NumberingFormat { NumberFormatId = formatId, FormatCode = numberFormat });
            numberFormats.Count = (uint)numberFormats.ChildElements.Count;
        }
        format.NumberFormatId = formatId; format.ApplyNumberFormat = true;
    }
    formats.Append(format); formats.Count = (uint)formats.ChildElements.Count;
    styles.Stylesheet.Save();
    var created = formats.Count - 1;
    cache[cacheKey] = created;
    return created;
}

static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
static string WorkbookContentHash(string path)
{
    string metadataEntry;
    using (var package = SpreadsheetDocument.Open(path, false))
    {
        var workbookPart = package.WorkbookPart ?? throw new InvalidDataException("workbook part is missing");
        var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>().FirstOrDefault(item => item.Name == "__ExcelAuditorMetadata")
            ?? throw new InvalidDataException("metadata worksheet is missing");
        var worksheetPart = workbookPart.GetPartById(sheet.Id?.Value ?? throw new InvalidDataException("metadata worksheet relationship is missing")) as WorksheetPart
            ?? throw new InvalidDataException("metadata worksheet part is invalid");
        metadataEntry = worksheetPart.Uri.OriginalString.TrimStart('/');
    }
    using var archive = ZipFile.OpenRead(path);
    using var hasher = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
    foreach (var entry in archive.Entries.Where(item => !StringComparer.Ordinal.Equals(item.FullName, metadataEntry)).OrderBy(item => item.FullName, StringComparer.Ordinal))
    {
        hasher.AppendData(Encoding.UTF8.GetBytes(entry.FullName));
        hasher.AppendData([0]);
        using var stream = entry.Open();
        var buffer = new byte[81920];
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) hasher.AppendData(buffer, 0, read);
        hasher.AppendData([0]);
    }
    return Convert.ToHexString(hasher.GetHashAndReset()).ToLowerInvariant();
}
static IEnumerable<string> PackageValidationErrors(string path)
{
    using var package = SpreadsheetDocument.Open(path, false);
    var errors = new OpenXmlValidator().Validate(package)
        .Select(item => $"{item.Part?.Uri}:{item.Description}").ToList();
    errors.AddRange(CriticalStructureErrors(package));
    return errors;
}

static IEnumerable<string> CriticalStructureErrors(SpreadsheetDocument package)
{
    var workbookPart = package.WorkbookPart;
    if (workbookPart?.Workbook?.Sheets is null)
    {
        yield return "workbook:missing sheets collection";
        yield break;
    }
    foreach (var sheet in workbookPart.Workbook.Sheets.Elements<Sheet>())
    {
        if (sheet.Id?.Value is not string relationshipId || workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
        {
            yield return $"worksheet:{sheet.Name?.Value ?? "<unnamed>"}:invalid relationship";
            continue;
        }
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in worksheetPart.Worksheet?.GetFirstChild<SheetData>()?.Elements<Row>() ?? [])
        {
            if (row.RowIndex?.Value is not uint rowIndex || rowIndex == 0)
            {
                yield return $"worksheet:{sheet.Name?.Value ?? "<unnamed>"}:row index is missing or zero";
                continue;
            }
            foreach (var cell in row.Elements<Cell>())
            {
                var reference = cell.CellReference?.Value;
                if (String.IsNullOrWhiteSpace(reference) || !Regex.IsMatch(reference, "^[A-Z]{1,3}[1-9][0-9]*$"))
                {
                    yield return $"worksheet:{sheet.Name?.Value ?? "<unnamed>"}:invalid cell reference '{reference ?? "<null>"}'";
                    continue;
                }
                var column = ColumnNumber(CellColumn(reference));
                var referenceRow = UInt32.Parse(new string(reference.Where(Char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
                if (column is < 1 or > 16384)
                    yield return $"worksheet:{sheet.Name?.Value ?? "<unnamed>"}:cell reference exceeds Excel column limit '{reference}'";
                if (referenceRow != rowIndex)
                    yield return $"worksheet:{sheet.Name?.Value ?? "<unnamed>"}:cell '{reference}' does not belong to row {rowIndex}";
                if (!seen.Add(reference))
                    yield return $"worksheet:{sheet.Name?.Value ?? "<unnamed>"}:duplicate cell reference '{reference}'";
            }
        }
    }
}
static string CellColumn(string reference) => new(reference.TakeWhile(Char.IsLetter).ToArray());
static int ColumnNumber(string name) { var value = 0; foreach (var c in name.ToUpperInvariant()) value = value * 26 + c - 'A' + 1; return value; }
static string ColumnName(int number) { var value = ""; while (number > 0) { number--; value = (char)('A' + number % 26) + value; number /= 26; } return value; }

static int OperationPhase(string type) => type switch { "mark_cell" or "mark_row" => 0, "set_cell" => 1, "insert_column" => 2, "set_cell_after_insert" => 3, "append_row" => 4, "add_or_replace_report_sheet" => 5, _ => 99 };

static void ValidateManifest(RenderManifest manifest)
{
    if (manifest.ManifestVersion != "1.0") throw new InvalidDataException("unsupported manifest version");
    if (String.IsNullOrWhiteSpace(manifest.JobId) || String.IsNullOrWhiteSpace(manifest.InputSha256) || !Regex.IsMatch(manifest.InputSha256, "^[0-9a-fA-F]{64}$"))
        throw new InvalidDataException("manifest identity or input hash is invalid");
    if (manifest.Operations is null) throw new InvalidDataException("manifest operations are required");
    ValidateMetadata(manifest.Metadata);
    var known = new HashSet<string> { "mark_cell", "mark_row", "set_cell", "insert_column", "set_cell_after_insert", "append_row", "add_or_replace_report_sheet" };
    foreach (var operation in manifest.Operations)
    {
        if (operation is null) throw new InvalidDataException("manifest operations must not contain null items");
        if (!known.Contains(operation.Type)) throw new InvalidDataException($"unknown operation type: {operation.Type}");
        if (operation.Type is not "add_or_replace_report_sheet" && String.IsNullOrWhiteSpace(operation.Sheet)) throw new InvalidDataException($"{operation.Type} requires sheet");
        if (operation.Type == "add_or_replace_report_sheet" && operation.Sheet is not null) throw new InvalidDataException("report operation must not declare sheet");
        if (operation.Type is "mark_cell" or "set_cell" or "set_cell_after_insert")
        {
            if (!IsValidCellReference(operation.Cell)) throw new InvalidDataException($"{operation.Type} requires a valid Excel cell");
        }
        else if (operation.Cell is not null) throw new InvalidDataException($"{operation.Type} must not declare cell");
        if (operation.Type is "mark_row" or "append_row")
        {
            if (!IsValidRow(operation.Row)) throw new InvalidDataException($"{operation.Type} requires a valid Excel row");
        }
        else if (operation.Row is not null) throw new InvalidDataException($"{operation.Type} must not declare row");
        if (operation.Type == "insert_column" && (
            (String.IsNullOrWhiteSpace(operation.Before) == String.IsNullOrWhiteSpace(operation.After))
            || !IsValidRow(operation.HeaderRow)
            || String.IsNullOrWhiteSpace(operation.CanonicalField)
            || String.IsNullOrWhiteSpace(operation.HeaderValue)
            || String.IsNullOrWhiteSpace(operation.FillColor)))
            throw new InvalidDataException("insert_column requires exactly one valid before/after column, canonical_field, header_row, header_value and fill_color");
        if (operation.Type == "insert_column" && !IsValidColumnReference(operation.Before ?? operation.After)) throw new InvalidDataException("insert_column anchor must be an Excel column from A to XFD");
        if (operation.Type == "insert_column" && operation.After is not null && ColumnNumber(operation.After) == 16384) throw new InvalidDataException("insert_column cannot insert after XFD");
        if (operation.Type == "insert_column" && operation.DataStartRow is not null && operation.DataStartRow <= operation.HeaderRow) throw new InvalidDataException("insert_column data_start_row must be after header_row");
        if (operation.Type == "insert_column" && operation.FormulaRows?.Any(row => row < (operation.DataStartRow ?? (operation.HeaderRow ?? 1u) + 1u) || row > 1048576u) == true) throw new InvalidDataException("insert_column formula_rows must be valid data rows");
        if (operation.Type == "insert_column" && operation.FormulaRows is not null && (String.IsNullOrWhiteSpace(operation.FormulaTemplate) || operation.FormulaRows.Count != operation.FormulaRows.Distinct().Count())) throw new InvalidDataException("insert_column formula_rows require a formula_template and must be unique");
        if (operation.Type == "insert_column") ValidateFormulaTemplate(operation.FormulaTemplate, "insert_column formula_template");
        else if (operation.Before is not null || operation.After is not null || operation.CanonicalField is not null || operation.HeaderRow is not null || operation.HeaderValue is not null || operation.Validation is not null || operation.FormulaTemplate is not null || operation.DataStartRow is not null || operation.FormulaRows is not null)
            throw new InvalidDataException($"{operation.Type} contains insert_column-only fields");
        if (operation.Type == "insert_column" && operation.Validation is not null) ValidateValidation(operation.Validation);
        if (operation.Type == "append_row") ValidateAppendValues(operation);
        else if (operation.Values is not null) throw new InvalidDataException($"{operation.Type} must not declare values");
        if (operation.Type == "add_or_replace_report_sheet")
        {
            if (!IsValidSheetName(operation.Name) || operation.Name == "__ExcelAuditorMetadata" || String.IsNullOrWhiteSpace(operation.SourceJson)) throw new InvalidDataException("report operation requires a valid non-reserved name and source_json");
            if (operation.FillColor is not null || operation.Comment is not null) throw new InvalidDataException("report operation contains unsupported presentation fields");
        }
        else if (operation.Name is not null || operation.SourceJson is not null) throw new InvalidDataException($"{operation.Type} contains report-only fields");
        if (operation.Type is not ("set_cell" or "set_cell_after_insert" or "insert_column") && (operation.Value is not null || operation.FieldType is not null || operation.NumberFormat is not null))
            throw new InvalidDataException($"{operation.Type} contains typed-cell-only fields");
        if (operation.Type == "insert_column" && operation.Value is not null) throw new InvalidDataException("insert_column must not declare value");
        ValidateFieldType(operation.FieldType);
        ValidateNumberFormat(operation.NumberFormat);
        if (operation.Type is "set_cell" or "set_cell_after_insert")
        {
            if (operation.FieldType is null) throw new InvalidDataException($"{operation.Type} requires field_type");
            ValidateTypedValue(operation.Value, operation.FieldType);
        }
        if ((operation.Type is "mark_cell" or "mark_row" or "set_cell" or "set_cell_after_insert" or "append_row") && String.IsNullOrWhiteSpace(operation.FillColor)) throw new InvalidDataException($"{operation.Type} requires fill_color");
        if (!String.IsNullOrWhiteSpace(operation.FillColor) && !Regex.IsMatch(operation.FillColor, "^[0-9A-Fa-f]{6}$")) throw new InvalidDataException("fill color must be six hexadecimal digits");
        if (operation.Comment?.Length > 32767) throw new InvalidDataException("operation comment exceeds Excel's safe length");
    }
}

static void ValidateMetadata(RenderMetadata? metadata)
{
    if (metadata is null) return;
    foreach (var (name, value) in new[] {
        ("schema_sha256", metadata.SchemaSha256),
        ("standard_sha256", metadata.StandardSha256),
        ("result_sha256", metadata.ResultSha256),
    })
        if (value is not null && !Regex.IsMatch(value, "^[0-9a-fA-F]{64}$")) throw new InvalidDataException($"metadata {name} must be SHA-256");
}

static void ValidateAppendValues(RenderOperation operation)
{
    if (operation.Values is not { Count: > 0 }) throw new InvalidDataException("append_row requires non-empty values");
    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    foreach (var value in operation.Values)
    {
        if (value is null || !IsValidCellReference(value.Cell)) throw new InvalidDataException("append_row values require valid cells");
        var cellReference = value.Cell!;
        var row = UInt32.Parse(new string(cellReference.Where(Char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
        if (row != operation.Row) throw new InvalidDataException("append_row value cells must belong to the declared row");
        if (!seen.Add(cellReference)) throw new InvalidDataException("append_row value cells must be unique");
        ValidateFieldType(value.FieldType);
        ValidateNumberFormat(value.NumberFormat);
        ValidateFormulaTemplate(value.FormulaTemplate, "append_row formula_template");
        if (value.FormulaTemplate is null)
        {
            if (value.FieldType is null) throw new InvalidDataException("append_row values require field_type");
            ValidateTypedValue(value.Value, value.FieldType);
        }
    }
}

static void ValidateValidation(RenderValidation validation)
{
    if (validation.Type == "list")
    {
        if (validation.Values is not { Count: > 0 } || validation.Values.Any(String.IsNullOrEmpty) || validation.Min is not null || validation.Max is not null)
            throw new InvalidDataException("list validation requires non-empty values and no numeric bounds");
        var list = String.Join(",", validation.Values.Select(value => value.Replace("\"", "\"\"")));
        if (list.Length > 250) throw new InvalidDataException("inline validation list exceeds safe Excel limit");
        return;
    }
    if (validation.Type is not ("integer" or "decimal") || validation.Values is not null || (validation.Min is null && validation.Max is null))
        throw new InvalidDataException("numeric validation requires integer/decimal type, bounds, and no list values");
    Decimal? minimum = ParseValidationBound(validation.Min);
    Decimal? maximum = ParseValidationBound(validation.Max);
    if (minimum is not null && maximum is not null && minimum > maximum) throw new InvalidDataException("validation minimum must not exceed maximum");
}

static decimal? ParseValidationBound(string? value)
{
    if (value is null) return null;
    if (!Decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)) throw new InvalidDataException("validation bounds must be finite invariant decimals");
    return parsed;
}

static void ValidateFormulaTemplate(string? formula, string source)
{
    if (formula is null) return;
    var withoutRowPlaceholder = formula.Replace("{row}", "", StringComparison.Ordinal);
    if (formula.Length is < 2 or > 512 || !formula.StartsWith('=') || withoutRowPlaceholder.Contains('{') || withoutRowPlaceholder.Contains('}') || Regex.IsMatch(formula, @"\[[^\]]+\]|https?://|(?:WEBSERVICE|HYPERLINK|RTD|CALL)\s*\(", RegexOptions.IgnoreCase))
        throw new InvalidDataException($"{source} is unsafe or invalid");
}

static void ValidateFieldType(string? fieldType)
{
    if (fieldType is null) return;
    var known = new HashSet<string> { "string", "integer", "decimal", "date", "datetime", "boolean", "enum", "phone", "id_code", "postal_code", "set", "json", "fuzzy_string" };
    if (!known.Contains(fieldType)) throw new InvalidDataException($"unknown field_type: {fieldType}");
}

static void ValidateTypedValue(JsonElement? value, string fieldType)
{
    if (value is null || value.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined) return;
    var kind = value.Value.ValueKind;
    if (fieldType is "integer" or "decimal")
    {
        if (kind is not (JsonValueKind.Number or JsonValueKind.String)) throw new InvalidDataException($"{fieldType} values must be JSON numbers or numeric strings");
        var text = kind == JsonValueKind.Number ? value.Value.GetRawText() : value.Value.GetString()!;
        if (!IsExcelSafeNumericText(text, fieldType == "integer")) throw new InvalidDataException($"{fieldType} value exceeds Excel's safe numeric write precision or range");
        return;
    }
    if (fieldType is "date" or "datetime")
    {
        if (kind != JsonValueKind.String || !DateTimeOffset.TryParse(value.Value.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal, out _))
            throw new InvalidDataException($"{fieldType} values must be ISO date/time strings");
        return;
    }
    if (fieldType == "boolean")
    {
        if (kind is not (JsonValueKind.True or JsonValueKind.False)) throw new InvalidDataException("boolean values must be JSON booleans");
        return;
    }
    if (kind != JsonValueKind.String) throw new InvalidDataException($"{fieldType} values must be JSON strings");
}

static bool IsExcelSafeNumericText(string text, bool requireInteger)
{
    var match = Regex.Match(text, @"^-?(?<coefficient>(?:0|[1-9][0-9]*)(?:\.[0-9]+)?)(?:[eE][+-]?[0-9]+)?$");
    if (!match.Success) return false;
    var coefficientDigits = match.Groups["coefficient"].Value.Replace(".", "", StringComparison.Ordinal);
    var significantDigits = coefficientDigits.TrimStart('0').TrimEnd('0');
    if (significantDigits.Length > 15) return false;
    if (!Double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) || !Double.IsFinite(parsed)) return false;
    if (significantDigits.Length > 0 && parsed == 0d) return false;
    var absolute = Math.Abs(parsed);
    if (absolute != 0d && (absolute < 2.2251E-308 || absolute > 9.99999999999999E+307)) return false;
    return !requireInteger || Math.Truncate(parsed) == parsed;
}

static void ValidateNumberFormat(string? numberFormat)
{
    if (numberFormat is not null && (numberFormat.Length > 255 || numberFormat.Any(Char.IsControl))) throw new InvalidDataException("number_format exceeds Excel's safe contract");
}

static bool IsValidRow(uint? row) => row is >= 1u and <= 1048576u;
static bool IsValidColumnReference(string? reference) => !String.IsNullOrWhiteSpace(reference) && Regex.IsMatch(reference, "^[A-Z]{1,3}$") && ColumnNumber(reference) is >= 1 and <= 16384;
static bool IsValidCellReference(string? reference)
{
    if (String.IsNullOrWhiteSpace(reference)) return false;
    var match = Regex.Match(reference, "^(?<column>[A-Z]{1,3})(?<row>[1-9][0-9]*)$");
    return match.Success && IsValidColumnReference(match.Groups["column"].Value) && UInt32.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var row) && row <= 1048576u;
}

static bool IsValidSheetName(string? name) => !String.IsNullOrWhiteSpace(name) && name.Length <= 31 && !name.Any(character => "[]:*?/\\".Contains(character)) && name[0] != '\'' && name[^1] != '\'';

static void ValidateInvocation(Args arguments)
{
    var input = Path.GetFullPath(arguments.Input);
    var output = Path.GetFullPath(arguments.Output);
    var manifest = Path.GetFullPath(arguments.Manifest);
    var comparer = OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
    if (!File.Exists(input)) throw new FileNotFoundException("input workbook does not exist");
    if (!File.Exists(manifest)) throw new FileNotFoundException("manifest does not exist");
    if (comparer.Equals(input, output)) throw new ArgumentException("input and output paths must be different");
    if (comparer.Equals(manifest, output)) throw new ArgumentException("manifest and output paths must be different");
    var inputExtension = Path.GetExtension(input);
    var outputExtension = Path.GetExtension(output);
    if (!new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".xlsx", ".xlsm" }.Contains(inputExtension) || !StringComparer.OrdinalIgnoreCase.Equals(inputExtension, outputExtension))
        throw new ArgumentException("input and output must use the same supported Excel extension");
    var outputDirectory = Path.GetDirectoryName(output);
    if (String.IsNullOrWhiteSpace(outputDirectory) || !Directory.Exists(outputDirectory)) throw new ArgumentException("output directory does not exist");
}

sealed record Args(string Input, string Output, string Manifest, bool DryRun)
{
    public static Args Parse(string[] values)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        var known = new HashSet<string>(StringComparer.Ordinal) { "--input", "--output", "--manifest" };
        var dryRun = false;
        for (var index = 0; index < values.Length; index++)
        {
            if (values[index] == "--dry-run")
            {
                if (dryRun) throw new ArgumentException("duplicate argument: --dry-run");
                dryRun = true;
                continue;
            }
            if (!known.Contains(values[index])) throw new ArgumentException($"unknown argument: {values[index]}");
            if (index + 1 >= values.Length) throw new ArgumentException($"missing value for {values[index]}");
            if (!map.TryAdd(values[index], values[++index])) throw new ArgumentException($"duplicate argument: {values[index - 1]}");
        }
        foreach (var required in known)
            if (!map.ContainsKey(required)) throw new ArgumentException($"missing required argument: {required}");
        return new Args(map["--input"], map["--output"], map["--manifest"], dryRun);
    }
}

sealed record RenderManifest(string ManifestVersion, string JobId, string InputSha256, List<RenderOperation> Operations, RenderMetadata? Metadata);
sealed record RenderMetadata(string? SchemaId, string? SchemaVersion, string? SchemaSha256, string? StandardSnapshotId, string? StandardSha256, string? ResultSha256);
sealed record RenderOperation(string Type, string? Sheet, string? Cell, uint? Row, string? Before, string? After, string? CanonicalField, uint? HeaderRow, string? HeaderValue, string? FillColor, string? Comment, string? Name, string? SourceJson, JsonElement? Value, List<RenderValue>? Values, string? DifferenceId, string? FieldType, string? NumberFormat, RenderValidation? Validation, string? FormulaTemplate, uint? DataStartRow, List<uint>? FormulaRows);
sealed record RenderValue(string? Cell, JsonElement? Value, string? FieldType, string? NumberFormat, string? FormulaTemplate);
sealed record RenderValidation(string? Type, List<string>? Values, string? Min, string? Max, bool AllowBlank = true);
sealed record OperationResult(int OperationIndex, string Type, string? DifferenceId, string Status, string? ErrorCode, string? Message);
static class JsonOptions { public static readonly JsonSerializerOptions Default = new() { PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow }; }
