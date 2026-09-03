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
                    case "set_number_format": ApplyNumberFormat(workbookPart, operation, styleCache); break;
                    case "insert_column": InsertColumn(workbookPart, operation, styleCache); break;
                    case "append_row": AppendRow(workbookPart, operation, styleCache); break;
                    case "add_or_replace_product_sheets": AddProductSheets(workbookPart, operation, Path.GetDirectoryName(Path.GetFullPath(arguments.Manifest))!, styleCache); break;
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
    SynchronizeTableHeaderName(worksheetPart, operation.Cell!, operation.Value);
    cell.StyleIndex = FillStyle(workbookPart, cell.StyleIndex, operation.FillColor!, operation.NumberFormat, styleCache);
    if (!String.IsNullOrWhiteSpace(operation.Comment)) AddOrReplaceComment(worksheetPart, operation.Cell!, operation.Comment!);
    RecalculateDimension(worksheetPart.Worksheet!);
    (worksheetPart.Worksheet ?? throw new InvalidDataException("worksheet is missing")).Save();
}

static void SynchronizeTableHeaderName(WorksheetPart worksheetPart, string reference, JsonElement? value)
{
    if (value is not JsonElement element || element.ValueKind != JsonValueKind.String) return;
    var name = element.GetString();
    if (String.IsNullOrWhiteSpace(name)) return;
    var referenceRowText = new string(reference.Where(Char.IsDigit).ToArray());
    if (!UInt32.TryParse(referenceRowText, NumberStyles.None, CultureInfo.InvariantCulture, out var referenceRow)) return;
    var referenceColumn = ColumnNumber(CellColumn(reference));
    foreach (var tablePart in worksheetPart.TableDefinitionParts)
    {
        var table = tablePart.Table;
        if (table?.Reference?.Value is not string tableReference || (table.HeaderRowCount?.Value ?? 1u) == 0u) continue;
        var bounds = RangeBounds(tableReference);
        if (referenceRow != bounds.StartRow || referenceColumn < bounds.StartColumn || referenceColumn > bounds.EndColumn) continue;
        var columns = table.TableColumns?.Elements<TableColumn>().ToList();
        var index = referenceColumn - bounds.StartColumn;
        if (columns is null || index < 0 || index >= columns.Count) continue;
        columns[index].Name = name;
        table.Save();
    }
}

static void ApplyNumberFormat(WorkbookPart workbookPart, RenderOperation operation, Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint> styleCache)
{
    var worksheetPart = Worksheet(workbookPart, operation.Sheet!);
    var cell = FindOrCreateCell(worksheetPart, operation.Cell!);
    cell.StyleIndex = FillStyle(workbookPart, cell.StyleIndex, "", operation.NumberFormat, styleCache);
    if (!String.IsNullOrWhiteSpace(operation.Comment)) AddOrReplaceComment(worksheetPart, operation.Cell!, operation.Comment!);
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

static void AddProductSheets(
    WorkbookPart workbookPart,
    RenderOperation operation,
    string manifestDirectory,
    Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint> styleCache)
{
    var productPath = SafeManifestSourcePath(manifestDirectory, operation.SourceJson ?? "product-result.json");
    using var document = JsonDocument.Parse(File.ReadAllText(productPath));
    var root = document.RootElement;
    if (!root.TryGetProperty("category_sheets", out var categorySheets) || categorySheets.ValueKind != JsonValueKind.Array)
        throw new InvalidDataException("product result requires category_sheets array");
    var merchantHeaderColor = JsonValue(root, "merchant_extra_header_color");
    if (String.IsNullOrWhiteSpace(merchantHeaderColor)) merchantHeaderColor = "D9D9D9";
    if (!Regex.IsMatch(merchantHeaderColor, "^[0-9A-Fa-f]{6}$")) throw new InvalidDataException("product merchant header color is invalid");
    var issues = new Dictionary<(string Category, uint SourceRow, string Field), ProductIssue>();
    if (root.TryGetProperty("issues", out var issueArray) && issueArray.ValueKind == JsonValueKind.Array)
    {
        foreach (var issue in issueArray.EnumerateArray())
        {
            var category = JsonValue(issue, "category_id");
            var field = JsonValue(issue, "field_id");
            if (String.IsNullOrWhiteSpace(category) || String.IsNullOrWhiteSpace(field)
                || !issue.TryGetProperty("excel_row", out var sourceRowElement)
                || !sourceRowElement.TryGetUInt32(out var sourceRow)) continue;
            var color = JsonValue(issue, "color");
            var issueType = JsonValue(issue, "issue_type");
            if (Regex.IsMatch(color, "^[0-9A-Fa-f]{6}$"))
                issues[(category, sourceRow, field)] = new ProductIssue(color, issueType);
        }
    }
    var usedNames = (workbookPart.Workbook?.Sheets?.Elements<Sheet>() ?? [])
        .Select(item => item.Name?.Value ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase);
    var validationListCache = new Dictionary<string, string>(StringComparer.Ordinal);
    var validationSheetName = UniqueSheetName("__ExcelAuditorLists", "lists", usedNames);
    foreach (var category in categorySheets.EnumerateArray())
    {
        var categoryId = JsonValue(category, "category_id");
        var sheetName = JsonValue(category, "worksheet_name");
        if (!IsValidSheetName(sheetName) || sheetName == "__ExcelAuditorMetadata")
            throw new InvalidDataException("product result contains an invalid worksheet name");
        var plan = category.GetProperty("plan");
        var fields = ParseProductFields(plan.GetProperty("fields"));
        var rows = category.GetProperty("rows");
        var sourceRows = category.GetProperty("source_excel_rows");
        AddOrReplaceProductSheet(workbookPart, sheetName, categoryId, fields, rows, sourceRows, issues, merchantHeaderColor, styleCache, validationListCache, validationSheetName);
        usedNames.Add(sheetName);
        if (category.TryGetProperty("sku_rows", out var skuRows) && skuRows.ValueKind == JsonValueKind.Array && skuRows.GetArrayLength() > 0)
        {
            var skuFields = fields.Where(item => item.Source is "fixed" or "platform_specification").ToList();
            var skuName = UniqueSheetName(sheetName + "-SKU", categoryId, usedNames);
            var skuSourceRows = category.TryGetProperty("sku_source_excel_rows", out var skuSourceElement)
                ? skuSourceElement
                : sourceRows;
            AddOrReplaceProductSheet(workbookPart, skuName, categoryId, skuFields, skuRows, skuSourceRows, issues, merchantHeaderColor, styleCache, validationListCache, validationSheetName);
            usedNames.Add(skuName);
        }
    }
    AddProductReviewSheets(workbookPart, root, usedNames, styleCache);
}

static void AddProductReviewSheets(
    WorkbookPart workbookPart,
    JsonElement root,
    HashSet<string> usedNames,
    Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint> styleCache)
{
    var sourceHeaders = new List<string>();
    if (root.TryGetProperty("source_headers", out var headersElement))
    {
        if (headersElement.ValueKind != JsonValueKind.Array)
            throw new InvalidDataException("product source_headers must be an array");
        var ordinal = 1;
        foreach (var header in headersElement.EnumerateArray())
        {
            if (header.ValueKind != JsonValueKind.String)
                throw new InvalidDataException("product source_headers must contain strings");
            var value = header.GetString() ?? "";
            sourceHeaders.Add(String.IsNullOrWhiteSpace(value) ? $"原列{ordinal}" : value);
            ordinal++;
        }
    }
    if (root.TryGetProperty("unresolved_rows", out var unresolved) && unresolved.ValueKind == JsonValueKind.Array && unresolved.GetArrayLength() > 0)
    {
        var rows = new List<List<string>>();
        foreach (var item in unresolved.EnumerateArray())
        {
            var resolution = item.GetProperty("category_resolution");
            var candidates = resolution.TryGetProperty("candidates", out var candidateArray) && candidateArray.ValueKind == JsonValueKind.Array
                ? String.Join(" | ", candidateArray.EnumerateArray().Select(candidate => $"{JsonValue(candidate, "field_id")}:{JsonValue(candidate, "title")}"))
                : "";
            var row = new List<string> {
                JsonValue(item, "excel_row"),
                JsonValue(resolution, "status"),
                JsonValue(resolution, "match_type"),
                JsonValue(resolution, "raw_category_id"),
                JsonValue(resolution, "raw_category"),
                candidates,
            };
            if (item.TryGetProperty("values", out var values) && values.ValueKind == JsonValueKind.Array)
                row.AddRange(values.EnumerateArray().Select(ProductReviewText));
            rows.Add(row);
        }
        var name = UniqueSheetName("待审核商品", "review", usedNames);
        AddGeneratedProductSheet(
            workbookPart,
            name,
            ["源行", "状态", "匹配类型", "原类目ID", "原类目", "候选类目", .. sourceHeaders],
            rows,
            "D9D2E9",
            styleCache
        );
        usedNames.Add(name);
    }
    if (root.TryGetProperty("issues", out var issues) && issues.ValueKind == JsonValueKind.Array && issues.GetArrayLength() > 0)
    {
        var rows = issues.EnumerateArray().Select(item => new List<string> {
            JsonValue(item, "excel_row"),
            JsonValue(item, "category_id"),
            JsonValue(item, "field_id"),
            JsonValue(item, "issue_type"),
            item.TryGetProperty("raw_value", out var raw) ? ProductReviewText(raw) : "",
            JsonValue(item, "message"),
        }).ToList();
        var name = UniqueSheetName("问题清单", "issues", usedNames);
        AddGeneratedProductSheet(
            workbookPart,
            name,
            ["源行", "类目ID", "字段ID", "问题类型", "原值", "说明"],
            rows,
            "F9CB9C",
            styleCache
        );
        usedNames.Add(name);
    }
}

static string ProductReviewText(JsonElement value) => value.ValueKind switch {
    JsonValueKind.Null or JsonValueKind.Undefined => "",
    JsonValueKind.String => value.GetString() ?? "",
    JsonValueKind.Object or JsonValueKind.Array => value.GetRawText(),
    _ => value.ToString(),
};

static void AddGeneratedProductSheet(
    WorkbookPart workbookPart,
    string name,
    List<string> headers,
    List<List<string>> rows,
    string headerColor,
    Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint> styleCache)
{
    if (headers.Count == 0 || headers.Count > 16384 || rows.Count > 1048575)
        throw new InvalidDataException("product review sheet exceeds Excel limits");
    var part = workbookPart.AddNewPart<WorksheetPart>();
    var data = new SheetData();
    part.Worksheet = new Worksheet(data);
    var header = new Row { RowIndex = 1u };
    for (var column = 0; column < headers.Count; column++)
    {
        var cell = new Cell {
            CellReference = ColumnName(column + 1) + "1",
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text(headers[column]) { Space = SpaceProcessingModeValues.Preserve }),
            StyleIndex = FillStyle(workbookPart, 0u, headerColor, null, styleCache),
        };
        header.Append(cell);
    }
    data.Append(header);
    for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
    {
        var row = new Row { RowIndex = (uint)rowIndex + 2u };
        for (var column = 0; column < Math.Min(headers.Count, rows[rowIndex].Count); column++)
        {
            var cell = new Cell { CellReference = ColumnName(column + 1) + row.RowIndex!.Value };
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new Text(rows[rowIndex][column]) { Space = SpaceProcessingModeValues.Preserve });
            row.Append(cell);
        }
        data.Append(row);
    }
    var endColumn = ColumnName(headers.Count);
    part.Worksheet.Append(new AutoFilter { Reference = $"A1:{endColumn}{Math.Max(1, rows.Count + 1)}" });
    RecalculateDimension(part.Worksheet);
    part.Worksheet.Save();
    var workbook = workbookPart.Workbook ?? throw new InvalidDataException("workbook is missing");
    var sheets = workbook.Sheets ?? workbook.AppendChild(new Sheets());
    var nextId = sheets.Elements<Sheet>().Select(item => item.SheetId?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1u;
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(part), SheetId = nextId, Name = name });
}

static List<ProductField> ParseProductFields(JsonElement plannedFields)
{
    if (plannedFields.ValueKind != JsonValueKind.Array) throw new InvalidDataException("product plan fields must be an array");
    var fields = new List<ProductField>();
    foreach (var planned in plannedFields.EnumerateArray())
    {
        var field = planned.GetProperty("field");
        var id = JsonValue(field, "field_id");
        var title = JsonValue(field, "title");
        var source = JsonValue(field, "source");
        var fieldType = JsonValue(field, "field_type");
        var numberFormat = JsonValue(field, "number_format");
        var timezone = JsonValue(field, "timezone");
        var required = field.TryGetProperty("required", out var requiredElement) && requiredElement.ValueKind == JsonValueKind.True;
        if (field.TryGetProperty("required", out requiredElement) && requiredElement.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw new InvalidDataException("product field required must be a boolean");
        var allowBlank = !required;
        string? minimum = null;
        string? maximum = null;
        if (field.TryGetProperty("validation", out var validationElement) && validationElement.ValueKind != JsonValueKind.Null)
        {
            if (validationElement.ValueKind != JsonValueKind.Object)
                throw new InvalidDataException("product field validation must be an object");
            if (validationElement.TryGetProperty("nullable", out var nullableElement))
            {
                if (nullableElement.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
                    throw new InvalidDataException("product field validation nullable must be a boolean");
                allowBlank = !required && nullableElement.GetBoolean();
            }
            minimum = JsonValue(validationElement, "min");
            maximum = JsonValue(validationElement, "max");
        }
        var enumValues = new List<string>();
        if (field.TryGetProperty("enum_values", out var enumElement))
        {
            if (enumElement.ValueKind != JsonValueKind.Array)
                throw new InvalidDataException("product field enum_values must be an array");
            foreach (var item in enumElement.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.String || String.IsNullOrWhiteSpace(item.GetString()))
                    throw new InvalidDataException("product field enum_values must contain non-blank strings");
                enumValues.Add(item.GetString()!);
            }
            if (enumValues.Count != enumValues.Distinct(StringComparer.Ordinal).Count())
                throw new InvalidDataException("product field enum_values must be unique");
        }
        if (String.IsNullOrWhiteSpace(id) || String.IsNullOrWhiteSpace(title)
            || source is not ("fixed" or "platform_attribute" or "platform_specification" or "merchant_extra"))
            throw new InvalidDataException("product field identity or source is invalid");
        ValidateFieldType(fieldType);
        ValidateNumberFormat(String.IsNullOrWhiteSpace(numberFormat) ? null : numberFormat);
        if (!String.IsNullOrWhiteSpace(timezone) && fieldType != "datetime")
            throw new InvalidDataException("product field timezone is only valid for datetime fields");
        if (!String.IsNullOrWhiteSpace(timezone))
        {
            if (timezone.Length > 255 || timezone.Any(Char.IsControl))
                throw new InvalidDataException("product field timezone is invalid");
            ResolveTimeZone(timezone);
        }
        fields.Add(new ProductField(
            id,
            title,
            source,
            fieldType,
            String.IsNullOrWhiteSpace(numberFormat) ? null : numberFormat,
            String.IsNullOrWhiteSpace(timezone) ? null : timezone,
            enumValues,
            allowBlank,
            String.IsNullOrWhiteSpace(minimum) ? null : minimum,
            String.IsNullOrWhiteSpace(maximum) ? null : maximum
        ));
    }
    if (fields.Count == 0 || fields.Count > 16384 || fields.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count() != fields.Count)
        throw new InvalidDataException("product field count or identity is invalid");
    return fields;
}

static void AddOrReplaceProductSheet(
    WorkbookPart workbookPart,
    string name,
    string categoryId,
    List<ProductField> fields,
    JsonElement rows,
    JsonElement sourceRows,
    Dictionary<(string Category, uint SourceRow, string Field), ProductIssue> issues,
    string merchantHeaderColor,
    Dictionary<(uint SourceStyle, string Fill, string NumberFormat), uint> styleCache,
    Dictionary<string, string> validationListCache,
    string validationSheetName)
{
    if (rows.ValueKind != JsonValueKind.Array || sourceRows.ValueKind != JsonValueKind.Array || rows.GetArrayLength() != sourceRows.GetArrayLength())
        throw new InvalidDataException("product rows and source_excel_rows must be aligned arrays");
    if (rows.GetArrayLength() > 1048575) throw new InvalidDataException("product sheet exceeds Excel's row limit");
    RemoveSheetIfPresent(workbookPart, name);
    var part = workbookPart.AddNewPart<WorksheetPart>();
    var data = new SheetData();
    var view = new SheetView { WorkbookViewId = 0u };
    view.Append(new Pane {
        VerticalSplit = 1d,
        TopLeftCell = "A2",
        ActivePane = PaneValues.BottomLeft,
        State = PaneStateValues.Frozen,
    });
    var worksheet = new Worksheet(new SheetViews(view), data);
    part.Worksheet = worksheet;
    var header = new Row { RowIndex = 1u };
    for (var index = 0; index < fields.Count; index++)
    {
        var field = fields[index];
        var cell = new Cell {
            CellReference = ColumnName(index + 1) + "1",
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text(field.Title) { Space = SpaceProcessingModeValues.Preserve }),
        };
        var headerColor = field.Source switch {
            "fixed" => "DDEBF7",
            "platform_attribute" => "E2F0D9",
            "platform_specification" => "FFF2CC",
            _ => merchantHeaderColor,
        };
        cell.StyleIndex = FillStyle(workbookPart, 0u, headerColor, null, styleCache);
        header.Append(cell);
    }
    data.Append(header);
    var rowElements = rows.EnumerateArray().ToArray();
    var sourceElements = sourceRows.EnumerateArray().ToArray();
    for (var rowIndex = 0; rowIndex < rowElements.Length; rowIndex++)
    {
        if (rowElements[rowIndex].ValueKind != JsonValueKind.Object || !sourceElements[rowIndex].TryGetUInt32(out var sourceRow))
            throw new InvalidDataException("product output row or source row is invalid");
        var row = new Row { RowIndex = (uint)rowIndex + 2u };
        for (var columnIndex = 0; columnIndex < fields.Count; columnIndex++)
        {
            var field = fields[columnIndex];
            var cell = new Cell { CellReference = ColumnName(columnIndex + 1) + row.RowIndex!.Value };
            JsonElement? value = rowElements[rowIndex].TryGetProperty(field.Id, out var property) ? property : null;
            issues.TryGetValue((categoryId, sourceRow, field.Id), out var issue);
            if (field.Source == "merchant_extra")
            {
                WriteSafeValue(cell, value, field.FieldType, Uses1904DateSystem(workbookPart));
            }
            else
            {
                try
                {
                    ValidateTypedValue(value, field.FieldType, field.Timezone);
                    WriteSafeValue(cell, value, field.FieldType, Uses1904DateSystem(workbookPart));
                }
                catch (InvalidDataException) when (issue is not null)
                {
                    WriteLiteralText(cell, value);
                }
            }
            if (issue is not null || !String.IsNullOrWhiteSpace(field.NumberFormat))
                cell.StyleIndex = FillStyle(workbookPart, 0u, issue?.Color, field.NumberFormat, styleCache);
            row.Append(cell);
        }
        data.Append(row);
    }
    var endColumn = ColumnName(fields.Count);
    var endRow = Math.Max(1, rows.GetArrayLength() + 1);
    worksheet.Append(new AutoFilter { Reference = $"A1:{endColumn}{endRow}" });
    ApplyProductFieldValidations(workbookPart, worksheet, fields, validationListCache, validationSheetName);
    RecalculateDimension(worksheet);
    worksheet.Save();
    var workbook = workbookPart.Workbook ?? throw new InvalidDataException("workbook is missing");
    var sheets = workbook.Sheets ?? workbook.AppendChild(new Sheets());
    var nextId = sheets.Elements<Sheet>().Select(item => item.SheetId?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1u;
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(part), SheetId = nextId, Name = name });
}

static void ApplyProductFieldValidations(
    WorkbookPart workbookPart,
    Worksheet worksheet,
    List<ProductField> fields,
    Dictionary<string, string> validationListCache,
    string validationSheetName)
{
    for (var index = 0; index < fields.Count; index++)
    {
        var field = fields[index];
        DataValidation? validation = null;
        if (field.EnumValues.Count > 0)
        {
            var rangeName = EnsureProductValidationList(
                workbookPart,
                field.EnumValues,
                validationListCache,
                validationSheetName
            );
            validation = new DataValidation {
                Type = DataValidationValues.List,
                AllowBlank = field.AllowBlank,
                ShowErrorMessage = true,
                ErrorTitle = "Invalid value",
                Error = "Select a value from the platform catalog list.",
                Formula1 = new Formula1(rangeName),
            };
        }
        else if (field.FieldType is "integer" or "decimal" && (field.Minimum is not null || field.Maximum is not null))
        {
            validation = new DataValidation {
                Type = field.FieldType == "integer" ? DataValidationValues.Whole : DataValidationValues.Decimal,
                AllowBlank = field.AllowBlank,
                ShowErrorMessage = true,
                ErrorTitle = "Invalid number",
                Error = "Enter a value within the platform catalog bounds.",
            };
            if (field.Minimum is not null && field.Maximum is not null)
            {
                validation.Operator = DataValidationOperatorValues.Between;
                validation.Formula1 = new Formula1(field.Minimum);
                validation.Formula2 = new Formula2(field.Maximum);
            }
            else if (field.Minimum is not null)
            {
                validation.Operator = DataValidationOperatorValues.GreaterThanOrEqual;
                validation.Formula1 = new Formula1(field.Minimum);
            }
            else
            {
                validation.Operator = DataValidationOperatorValues.LessThanOrEqual;
                validation.Formula1 = new Formula1(field.Maximum!);
            }
        }
        if (validation is null) continue;
        var column = ColumnName(index + 1);
        validation.SequenceOfReferences = new ListValue<StringValue> { InnerText = $"{column}2:{column}1048576" };
        var validations = worksheet.Elements<DataValidations>().FirstOrDefault();
        if (validations is null)
        {
            validations = new DataValidations();
            var following = worksheet.ChildElements.FirstOrDefault(item => item.LocalName is
                "hyperlinks" or "printOptions" or "pageMargins" or "pageSetup" or "headerFooter" or
                "rowBreaks" or "colBreaks" or "customProperties" or "cellWatches" or "ignoredErrors" or
                "smartTags" or "drawing" or "legacyDrawing" or "legacyDrawingHF" or "picture" or
                "oleObjects" or "controls" or "webPublishItems" or "tableParts" or "extLst");
            if (following is null) worksheet.Append(validations);
            else worksheet.InsertBefore(validations, following);
        }
        validations.Append(validation);
        validations.Count = (uint)validations.ChildElements.Count;
    }
}

static string EnsureProductValidationList(
    WorkbookPart workbookPart,
    List<string> values,
    Dictionary<string, string> cache,
    string sheetName)
{
    var key = String.Join("\u001F", values);
    if (cache.TryGetValue(key, out var existingName)) return existingName;
    var workbook = workbookPart.Workbook ?? throw new InvalidDataException("workbook is missing");
    var sheets = workbook.Sheets ?? workbook.AppendChild(new Sheets());
    var sheet = sheets.Elements<Sheet>().FirstOrDefault(item => StringComparer.Ordinal.Equals(item.Name?.Value, sheetName));
    WorksheetPart part;
    if (sheet is null)
    {
        part = workbookPart.AddNewPart<WorksheetPart>();
        part.Worksheet = new Worksheet(new SheetData());
        var nextId = sheets.Elements<Sheet>().Select(item => item.SheetId?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1u;
        sheet = new Sheet {
            Id = workbookPart.GetIdOfPart(part),
            SheetId = nextId,
            Name = sheetName,
            State = SheetStateValues.VeryHidden,
        };
        sheets.Append(sheet);
    }
    else
    {
        part = workbookPart.GetPartById(sheet.Id?.Value ?? throw new InvalidDataException("validation worksheet relationship is missing")) as WorksheetPart
            ?? throw new InvalidDataException("validation worksheet part is invalid");
    }
    var data = part.Worksheet?.GetFirstChild<SheetData>() ?? throw new InvalidDataException("validation worksheet data is missing");
    var columnNumber = cache.Count + 1;
    var column = ColumnName(columnNumber);
    for (var index = 0; index < values.Count; index++)
    {
        var cell = FindOrCreateCell(part, column + (index + 1).ToString(CultureInfo.InvariantCulture));
        cell.CellFormula = null;
        cell.CellValue = null;
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Text(values[index]) { Space = SpaceProcessingModeValues.Preserve });
    }
    RecalculateDimension(part.Worksheet!);
    part.Worksheet!.Save();
    var definedNames = workbook.DefinedNames ?? workbook.AppendChild(new DefinedNames());
    var ordinal = definedNames.Elements<DefinedName>().Count(item => item.Name?.Value?.StartsWith("_ExcelAuditorList", StringComparison.Ordinal) == true) + 1;
    string name;
    do name = "_ExcelAuditorList" + ordinal++.ToString(CultureInfo.InvariantCulture);
    while (definedNames.Elements<DefinedName>().Any(item => StringComparer.OrdinalIgnoreCase.Equals(item.Name?.Value, name)));
    definedNames.Append(new DefinedName {
        Name = name,
        Text = $"'{sheetName.Replace("'", "''", StringComparison.Ordinal)}'!${column}$1:${column}${values.Count}",
    });
    cache[key] = name;
    return name;
}

static void WriteLiteralText(Cell cell, JsonElement? value)
{
    cell.CellFormula = null;
    cell.CellValue = null;
    cell.InlineString = null;
    if (value is null || value.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
    {
        cell.DataType = null;
        return;
    }
    cell.DataType = CellValues.InlineString;
    cell.InlineString = new InlineString(new Text(value.Value.ToString()) { Space = SpaceProcessingModeValues.Preserve });
}

static void RemoveSheetIfPresent(WorkbookPart workbookPart, string name)
{
    var sheets = workbookPart.Workbook?.Sheets;
    var existing = sheets?.Elements<Sheet>().FirstOrDefault(item => StringComparer.OrdinalIgnoreCase.Equals(item.Name?.Value, name));
    if (existing is null) return;
    var oldPart = workbookPart.GetPartById(existing.Id!);
    existing.Remove();
    workbookPart.DeletePart(oldPart);
}

static string UniqueSheetName(string requested, string categoryId, HashSet<string> used)
{
    var sanitized = Regex.Replace(requested, @"[\[\]:*?/\\]", "-");
    if (sanitized.Length > 31) sanitized = sanitized[..31];
    if (!used.Contains(sanitized) && IsValidSheetName(sanitized)) return sanitized;
    var suffix = "-" + categoryId;
    if (suffix.Length > 12) suffix = suffix[..12];
    var candidate = sanitized[..Math.Min(sanitized.Length, 31 - suffix.Length)] + suffix;
    var counter = 2;
    while (used.Contains(candidate))
    {
        var marker = "-" + counter++;
        candidate = sanitized[..Math.Min(sanitized.Length, 31 - suffix.Length - marker.Length)] + suffix + marker;
    }
    return candidate;
}

static string SafeManifestSourcePath(string manifestDirectory, string source)
{
    var path = Path.GetFullPath(Path.Combine(manifestDirectory, source));
    var relative = Path.GetRelativePath(manifestDirectory, path);
    if (Path.IsPathRooted(relative) || relative == ".." || relative.StartsWith(".." + Path.DirectorySeparatorChar))
        throw new InvalidDataException("source path escapes manifest directory");
    return path;
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
    const string pattern = @"(?i)(?<![A-Z0-9_\]\[])(?:(?<sheet>'(?:[^']|'')+'|[^'""\s!+\-*/^&=(),;:{}\[\]]+)!)?(?<reference>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![A-Z0-9_]|\s*\()";
    return ReplaceOutsideStringLiterals(formula, pattern, match =>
    {
        if (!IsValidFormulaReference(match.Groups["reference"].Value)) return match.Value;
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
    var pattern = $@"(?i)(?<sheet>{Regex.Escape(quoted)}|{Regex.Escape(plain)})(?<reference>\$?[A-Z]{{1,3}}\$?\d+(?::\$?[A-Z]{{1,3}}\$?\d+)?)(?![A-Z0-9_]|\s*\()";
    return ReplaceOutsideStringLiterals(formula, pattern, match => IsValidFormulaReference(match.Groups["reference"].Value)
        ? match.Groups["sheet"].Value + ShiftRange(match.Groups["reference"].Value, target)
        : match.Value);
}

static bool IsValidFormulaReference(string reference) => reference.Split(':', 2).All(item =>
{
    var match = Regex.Match(item.Replace("$", "", StringComparison.Ordinal), @"^(?<column>[A-Z]{1,3})(?<row>[1-9][0-9]*)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    return match.Success
        && ColumnNumber(match.Groups["column"].Value) <= 16384
        && UInt32.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var row)
        && row <= 1048576u;
});

static string ReplaceOutsideStringLiterals(string formula, string pattern, MatchEvaluator evaluator)
{
    var output = new System.Text.StringBuilder(formula.Length + 8);
    var chunkStart = 0;
    var index = 0;
    while (index < formula.Length)
    {
        if (formula[index] != '"') { index++; continue; }
        output.Append(Regex.Replace(formula[chunkStart..index], pattern, evaluator, RegexOptions.CultureInvariant));
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
    output.Append(Regex.Replace(formula[chunkStart..], pattern, evaluator, RegexOptions.CultureInvariant));
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
    var match = Regex.Match(reference, @"^(?<absolute>\$?)(?<column>[A-Z]{1,3})(?<row>\$?\d+)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    if (!match.Success) return reference;
    var column = ColumnNumber(match.Groups["column"].Value);
    if (!force && column < target) return reference;
    if (column >= 16384)
        throw new InvalidDataException("UNSUPPORTED_FEATURE: column insertion would shift a cell reference beyond XFD");
    return match.Groups["absolute"].Value + ColumnName(column + 1) + match.Groups["row"].Value;
}

static (int Start, int End) RangeColumns(string reference)
{
    var parts = reference.Replace("$", "").Split(':', 2);
    var start = ColumnNumber(CellColumn(parts[0]));
    var end = parts.Length == 2 ? ColumnNumber(CellColumn(parts[1])) : start;
    return (start, end);
}

static (int StartColumn, int EndColumn, uint StartRow, uint EndRow) RangeBounds(string reference)
{
    var parts = reference.Replace("$", "").Split(':', 2);
    var startRowText = new string(parts[0].Where(Char.IsDigit).ToArray());
    var endCell = parts.Length == 2 ? parts[1] : parts[0];
    var endRowText = new string(endCell.Where(Char.IsDigit).ToArray());
    if (!UInt32.TryParse(startRowText, NumberStyles.None, CultureInfo.InvariantCulture, out var startRow)
        || !UInt32.TryParse(endRowText, NumberStyles.None, CultureInfo.InvariantCulture, out var endRow))
        throw new InvalidDataException($"invalid range reference: {reference}");
    return (ColumnNumber(CellColumn(parts[0])), ColumnNumber(CellColumn(endCell)), startRow, endRow);
}

static (int Start, int End) ShiftInterval(int start, int end, int target)
{
    if (start >= target) return (start + 1, end + 1);
    if (end >= target) return (start, end + 1);
    return (start, end);
}

static void AddOrReplaceComment(WorksheetPart worksheetPart, string reference, string text)
{
    const string auditMarker = "\n\n[Excel Auditor]\n";
    var commentsPart = worksheetPart.WorksheetCommentsPart ?? worksheetPart.AddNewPart<WorksheetCommentsPart>();
    if (commentsPart.Comments is null)
        commentsPart.Comments = new Comments(new Authors(new Author("Excel Auditor")), new CommentList());
    var authors = commentsPart.Comments.Authors ?? commentsPart.Comments.PrependChild(new Authors());
    var authorList = authors.Elements<Author>().ToList();
    var authorId = authorList.FindIndex(author => author.Text == "Excel Auditor");
    if (authorId < 0) { authors.Append(new Author("Excel Auditor")); authorId = authorList.Count; }
    var list = commentsPart.Comments.CommentList ?? commentsPart.Comments.AppendChild(new CommentList());
    var existing = list.Elements<Comment>().FirstOrDefault(item => item.Reference?.Value == reference);
    var outputText = text;
    var outputAuthorId = (uint)authorId;
    if (existing is not null)
    {
        var existingText = existing.CommentText?.InnerText ?? "";
        var existingAuthorId = existing.AuthorId?.Value;
        var existingAuthor = existingAuthorId is not null && existingAuthorId.Value < (uint)authors.ChildElements.Count
            ? authors.Elements<Author>().ElementAt((int)existingAuthorId.Value).Text
            : null;
        if (!StringComparer.Ordinal.Equals(existingAuthor, "Excel Auditor"))
        {
            var markerIndex = existingText.IndexOf(auditMarker, StringComparison.Ordinal);
            var originalText = markerIndex >= 0 ? existingText[..markerIndex] : existingText;
            var previousAuditText = markerIndex >= 0 ? existingText[(markerIndex + auditMarker.Length)..] : "";
            var auditText = MergeAuditCommentText(previousAuditText, text);
            outputText = originalText + auditMarker + auditText;
            if (existingAuthorId is not null) outputAuthorId = existingAuthorId.Value;
        }
        else
        {
            outputText = MergeAuditCommentText(existingText, text);
        }
    }
    if (outputText.Length > 32767)
        throw new InvalidDataException("UNSUPPORTED_FEATURE: existing comment leaves insufficient room for an audit comment without data loss");
    var outputComment = existing ?? new Comment();
    outputComment.Reference = reference;
    outputComment.AuthorId = outputAuthorId;
    outputComment.CommentText = new CommentText(new Run(new Text(outputText) { Space = SpaceProcessingModeValues.Preserve }));
    if (existing is null) list.Append(outputComment);
    commentsPart.Comments.Save();
    RebuildCommentVml(worksheetPart);
}

static string MergeAuditCommentText(string existing, string next)
{
    if (String.IsNullOrEmpty(existing)) return next;
    if (StringComparer.Ordinal.Equals(existing, next)) return existing;
    var entries = existing.Split(" | ", StringSplitOptions.None);
    return entries.Contains(next, StringComparer.Ordinal) ? existing : existing + " | " + next;
}

static void RebuildCommentVml(WorksheetPart worksheetPart)
{
    var comments = worksheetPart.WorksheetCommentsPart?.Comments?.CommentList?.Elements<Comment>().ToList() ?? [];
    var vmlPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
    XNamespace v = "urn:schemas-microsoft-com:vml";
    XNamespace o = "urn:schemas-microsoft-com:office:office";
    XNamespace x = "urn:schemas-microsoft-com:office:excel";
    XDocument document;
    if (vmlPart is not null)
    {
        using var existingStream = vmlPart.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(existingStream);
        var existingXml = reader.ReadToEnd();
        try
        {
            document = XDocument.Parse(existingXml, LoadOptions.PreserveWhitespace);
            var unsafeShape = document.Descendants(v + "shape").Any(shape =>
            {
                var clientData = shape.Descendants(x + "ClientData").FirstOrDefault();
                return clientData is null || !StringComparer.OrdinalIgnoreCase.Equals((string?)clientData.Attribute("ObjectType"), "Note");
            });
            if (unsafeShape || document.Descendants(x + "Macro").Any())
                throw new InvalidDataException("UNSUPPORTED_FEATURE: legacy VML shapes or controls cannot be safely rewritten");
        }
        catch (System.Xml.XmlException exception)
        {
            throw new InvalidDataException("UNSUPPORTED_FEATURE: legacy VML cannot be parsed safely", exception);
        }
    }
    else
    {
        vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>();
        var worksheet = worksheetPart.Worksheet ?? throw new InvalidDataException("worksheet is missing");
        var legacyDrawing = new LegacyDrawing { Id = worksheetPart.GetIdOfPart(vmlPart) };
        var following = worksheet.ChildElements.FirstOrDefault(item => item.LocalName is
            "legacyDrawingHF" or "picture" or "oleObjects" or "controls" or "webPublishItems" or "tableParts" or "extLst");
        if (following is null) worksheet.Append(legacyDrawing);
        else worksheet.InsertBefore(legacyDrawing, following);
        document = CreateCommentVmlDocument(v, o, x);
    }

    var root = document.Root ?? throw new InvalidDataException("UNSUPPORTED_FEATURE: legacy VML has no document element");
    var shapes = root.Descendants(v + "shape").ToList();
    if (shapes.Count > comments.Count)
        throw new InvalidDataException("UNSUPPORTED_FEATURE: comment VML contains more Note shapes than comments");
    for (var index = 0; index < shapes.Count; index++)
        MoveExistingCommentShape(shapes[index], comments[index], x);

    if (comments.Count > shapes.Count)
    {
        EnsureCommentShapeType(root, v, o);
        var nextShapeId = Math.Max(1025, shapes.Select(shape => ParseCommentShapeId((string?)shape.Attribute("id"))).DefaultIfEmpty(1024).Max() + 1);
        foreach (var comment in comments.Skip(shapes.Count))
            root.Add(CreateDefaultCommentShape(comment, nextShapeId++, v, o, x));
    }

    using var stream = new MemoryStream();
    document.Save(stream, SaveOptions.DisableFormatting);
    stream.Position = 0;
    vmlPart.FeedData(stream);
}

static XDocument CreateCommentVmlDocument(XNamespace v, XNamespace o, XNamespace x)
{
    return new XDocument(new XElement("xml",
        new XAttribute(XNamespace.Xmlns + "v", v), new XAttribute(XNamespace.Xmlns + "o", o), new XAttribute(XNamespace.Xmlns + "x", x),
        new XElement(o + "shapelayout", new XAttribute(v + "ext", "edit"), new XElement(o + "idmap", new XAttribute(v + "ext", "edit"), new XAttribute("data", "1"))),
        new XElement(v + "shapetype", new XAttribute("id", "_x0000_t202"), new XAttribute("coordsize", "21600,21600"), new XAttribute(o + "spt", "202"), new XAttribute("path", "m,l,21600r21600,l21600,xe"),
            new XElement(v + "stroke", new XAttribute("joinstyle", "miter")), new XElement(v + "path", new XAttribute("gradientshapeok", "t"), new XAttribute(o + "connecttype", "rect")))));
}

static void EnsureCommentShapeType(XElement root, XNamespace v, XNamespace o)
{
    if (root.Elements(v + "shapetype").Any(item => StringComparer.Ordinal.Equals((string?)item.Attribute("id"), "_x0000_t202"))) return;
    var shapeType = new XElement(v + "shapetype", new XAttribute("id", "_x0000_t202"), new XAttribute("coordsize", "21600,21600"), new XAttribute(o + "spt", "202"), new XAttribute("path", "m,l,21600r21600,l21600,xe"),
        new XElement(v + "stroke", new XAttribute("joinstyle", "miter")), new XElement(v + "path", new XAttribute("gradientshapeok", "t"), new XAttribute(o + "connecttype", "rect")));
    var firstShape = root.Elements(v + "shape").FirstOrDefault();
    if (firstShape is null) root.Add(shapeType);
    else firstShape.AddBeforeSelf(shapeType);
}

static int ParseCommentShapeId(string? identifier)
{
    var match = Regex.Match(identifier ?? "", @"^_x0000_s(?<id>\d+)$", RegexOptions.CultureInvariant);
    return match.Success && Int32.TryParse(match.Groups["id"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value) ? value : 0;
}

static void MoveExistingCommentShape(XElement shape, Comment comment, XNamespace x)
{
    var clientData = shape.Descendants(x + "ClientData").SingleOrDefault()
        ?? throw new InvalidDataException("UNSUPPORTED_FEATURE: Note shape is missing ClientData");
    var rowElement = clientData.Elements(x + "Row").SingleOrDefault();
    var columnElement = clientData.Elements(x + "Column").SingleOrDefault();
    if (rowElement is null || columnElement is null
        || !Int32.TryParse(rowElement.Value.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out var previousRow)
        || !Int32.TryParse(columnElement.Value.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out var previousColumn))
        throw new InvalidDataException("UNSUPPORTED_FEATURE: Note shape coordinates cannot be safely interpreted");

    var (row, column) = ZeroBasedCommentCoordinates(comment);
    var anchor = clientData.Elements(x + "Anchor").SingleOrDefault();
    if (anchor is null)
    {
        anchor = new XElement(x + "Anchor", DefaultCommentAnchor(row, column));
        columnElement.AddBeforeSelf(anchor);
    }
    else
    {
        var values = anchor.Value.Split(',').Select(item => item.Trim()).ToArray();
        if (values.Length != 8 || values.Any(item => !Int32.TryParse(item, NumberStyles.Integer, CultureInfo.InvariantCulture, out _)))
            throw new InvalidDataException("UNSUPPORTED_FEATURE: Note shape anchor cannot be safely interpreted");
        var positions = values.Select(item => Int32.Parse(item, CultureInfo.InvariantCulture)).ToArray();
        var columnDelta = checked(column - previousColumn);
        var rowDelta = checked(row - previousRow);
        var shiftedStartColumn = (long)positions[0] + columnDelta;
        var shiftedEndColumn = (long)positions[4] + columnDelta;
        var shiftedStartRow = (long)positions[2] + rowDelta;
        var shiftedEndRow = (long)positions[6] + rowDelta;
        if (shiftedStartColumn < 0 || shiftedStartColumn > Int32.MaxValue || shiftedEndColumn < 0 || shiftedEndColumn > Int32.MaxValue
            || shiftedStartRow < 0 || shiftedStartRow > Int32.MaxValue || shiftedEndRow < 0 || shiftedEndRow > Int32.MaxValue)
            throw new InvalidDataException("UNSUPPORTED_FEATURE: Note shape anchor would move outside the worksheet");
        positions[0] = (int)shiftedStartColumn;
        positions[4] = (int)shiftedEndColumn;
        positions[2] = (int)shiftedStartRow;
        positions[6] = (int)shiftedEndRow;
        anchor.Value = String.Join(", ", positions.Select(item => item.ToString(CultureInfo.InvariantCulture)));
    }
    rowElement.Value = row.ToString(CultureInfo.InvariantCulture);
    columnElement.Value = column.ToString(CultureInfo.InvariantCulture);
}

static XElement CreateDefaultCommentShape(Comment comment, int shapeId, XNamespace v, XNamespace o, XNamespace x)
{
    var (row, column) = ZeroBasedCommentCoordinates(comment);
    return new XElement(v + "shape", new XAttribute("id", $"_x0000_s{shapeId}"), new XAttribute("type", "#_x0000_t202"), new XAttribute("style", "position:absolute;margin-left:80pt;margin-top:5pt;width:108pt;height:59pt;z-index:1;visibility:hidden"), new XAttribute("fillcolor", "#ffffe1"), new XAttribute(o + "insetmode", "auto"),
        new XElement(v + "fill", new XAttribute("color2", "#ffffe1")), new XElement(v + "shadow", new XAttribute("on", "t"), new XAttribute("color", "black"), new XAttribute("obscured", "t")), new XElement(v + "path", new XAttribute(o + "connecttype", "none")),
        new XElement(v + "textbox", new XAttribute("style", "mso-direction-alt:auto"), new XElement("div", new XAttribute("style", "text-align:left"))),
        new XElement(x + "ClientData", new XAttribute("ObjectType", "Note"), new XElement(x + "MoveWithCells"), new XElement(x + "SizeWithCells"), new XElement(x + "Anchor", DefaultCommentAnchor(row, column)), new XElement(x + "AutoFill", "False"), new XElement(x + "Row", row), new XElement(x + "Column", column)));
}

static (int Row, int Column) ZeroBasedCommentCoordinates(Comment comment)
{
    var reference = comment.Reference?.Value ?? throw new InvalidDataException("comment reference is missing");
    return (checked((int)UInt32.Parse(new string(reference.Where(Char.IsDigit).ToArray()), CultureInfo.InvariantCulture) - 1), ColumnNumber(CellColumn(reference)) - 1);
}

static string DefaultCommentAnchor(int row, int column)
{
    return $"{column}, 15, {row}, 2, {column + 3}, 15, {row + 4}, 4";
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
    var existingFormat = formats.Elements<CellFormat>()
        .Select((item, index) => (item, index))
        .FirstOrDefault(item => item.item.OuterXml == format.OuterXml);
    if (existingFormat.item is not null)
    {
        var existingIndex = (uint)existingFormat.index;
        cache[cacheKey] = existingIndex;
        return existingIndex;
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

static int OperationPhase(string type) => type switch { "mark_cell" or "mark_row" => 0, "set_number_format" or "set_cell" => 1, "insert_column" => 2, "set_cell_after_insert" => 3, "append_row" => 4, "add_or_replace_product_sheets" => 5, "add_or_replace_report_sheet" => 6, _ => 99 };

static void ValidateManifest(RenderManifest manifest)
{
    if (manifest.ManifestVersion != "1.0") throw new InvalidDataException("unsupported manifest version");
    if (String.IsNullOrWhiteSpace(manifest.JobId) || String.IsNullOrWhiteSpace(manifest.InputSha256) || !Regex.IsMatch(manifest.InputSha256, "^[0-9a-fA-F]{64}$"))
        throw new InvalidDataException("manifest identity or input hash is invalid");
    if (manifest.Operations is null) throw new InvalidDataException("manifest operations are required");
    ValidateMetadata(manifest.Metadata);
    var known = new HashSet<string> { "mark_cell", "mark_row", "set_cell", "set_number_format", "insert_column", "set_cell_after_insert", "append_row", "add_or_replace_product_sheets", "add_or_replace_report_sheet" };
    foreach (var operation in manifest.Operations)
    {
        if (operation is null) throw new InvalidDataException("manifest operations must not contain null items");
        if (!known.Contains(operation.Type)) throw new InvalidDataException($"unknown operation type: {operation.Type}");
        var globalOperation = operation.Type is "add_or_replace_report_sheet" or "add_or_replace_product_sheets";
        if (!globalOperation && String.IsNullOrWhiteSpace(operation.Sheet)) throw new InvalidDataException($"{operation.Type} requires sheet");
        if (globalOperation && operation.Sheet is not null) throw new InvalidDataException($"{operation.Type} must not declare sheet");
        if (operation.Type is "mark_cell" or "set_cell" or "set_number_format" or "set_cell_after_insert")
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
        else if (operation.Type == "add_or_replace_product_sheets")
        {
            if (operation.Name is not null || String.IsNullOrWhiteSpace(operation.SourceJson)) throw new InvalidDataException("product sheet operation requires source_json and no name");
            if (operation.FillColor is not null || operation.Comment is not null) throw new InvalidDataException("product sheet operation contains unsupported presentation fields");
        }
        else if (operation.Name is not null || operation.SourceJson is not null) throw new InvalidDataException($"{operation.Type} contains report-only fields");
        if (operation.Type is not ("set_cell" or "set_cell_after_insert" or "set_number_format" or "insert_column") && (operation.Value is not null || operation.FieldType is not null || operation.NumberFormat is not null))
            throw new InvalidDataException($"{operation.Type} contains typed-cell-only fields");
        if (operation.Type is not ("set_cell" or "set_cell_after_insert") && operation.Timezone is not null)
            throw new InvalidDataException($"{operation.Type} contains typed-value-only timezone");
        if (operation.Type == "insert_column" && operation.Value is not null) throw new InvalidDataException("insert_column must not declare value");
        if (operation.Type == "set_number_format" && (operation.Value is not null || operation.FieldType is not null || String.IsNullOrWhiteSpace(operation.NumberFormat)))
            throw new InvalidDataException("set_number_format requires number_format and must not declare value or field_type");
        if (operation.Type == "set_number_format" && operation.FillColor is not null)
            throw new InvalidDataException("set_number_format must not declare fill_color");
        ValidateFieldType(operation.FieldType);
        ValidateNumberFormat(operation.NumberFormat);
        if (operation.Type is "set_cell" or "set_cell_after_insert")
        {
            if (operation.FieldType is null) throw new InvalidDataException($"{operation.Type} requires field_type");
            ValidateTypedValue(operation.Value, operation.FieldType, operation.Timezone);
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
        if (value.FormulaTemplate is not null && value.Timezone is not null)
            throw new InvalidDataException("append_row formula values must not declare timezone");
        if (value.FormulaTemplate is null)
        {
            if (value.FieldType is null) throw new InvalidDataException("append_row values require field_type");
            ValidateTypedValue(value.Value, value.FieldType, value.Timezone);
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
    var formulaCode = FormulaCodeOutsideStringLiterals(formula);
    if (formula.Length is < 2 or > 512 || !formula.StartsWith('=') || withoutRowPlaceholder.Contains('{') || withoutRowPlaceholder.Contains('}') || formulaCode is null
        || Regex.IsMatch(formulaCode, @"\[|https?://|\||(?<![A-Z0-9_.])(?:(?:_XLFN|_XLWS|_XLL)\.)?(?:WEBSERVICE|HYPERLINK|RTD|CALL|DDE|EXEC|REGISTER(?:\.ID)?|RUN|IMAGE|STOCKHISTORY)\s*\(", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        throw new InvalidDataException($"{source} is unsafe or invalid");
}

static string? FormulaCodeOutsideStringLiterals(string formula)
{
    var output = new StringBuilder(formula.Length);
    var index = 0;
    while (index < formula.Length)
    {
        if (formula[index] != '"')
        {
            output.Append(formula[index++]);
            continue;
        }
        output.Append(' ');
        index++;
        var closed = false;
        while (index < formula.Length)
        {
            output.Append(' ');
            if (formula[index] != '"') { index++; continue; }
            if (index + 1 < formula.Length && formula[index + 1] == '"')
            {
                output.Append(' ');
                index += 2;
                continue;
            }
            index++;
            closed = true;
            break;
        }
        if (!closed) return null;
    }
    return output.ToString();
}

static void ValidateFieldType(string? fieldType)
{
    if (fieldType is null) return;
    var known = new HashSet<string> { "string", "integer", "decimal", "date", "datetime", "boolean", "enum", "phone", "id_code", "postal_code", "set", "json", "fuzzy_string" };
    if (!known.Contains(fieldType)) throw new InvalidDataException($"unknown field_type: {fieldType}");
}

static void ValidateTypedValue(JsonElement? value, string fieldType, string? timezoneName = null)
{
    if (timezoneName is not null && fieldType != "datetime") throw new InvalidDataException("timezone is only valid for datetime values");
    if (timezoneName is not null && (String.IsNullOrWhiteSpace(timezoneName) || timezoneName.Length > 255 || timezoneName.Any(Char.IsControl)))
        throw new InvalidDataException("datetime timezone is invalid");
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
        if (kind != JsonValueKind.String || !DateTimeOffset.TryParse(value.Value.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal, out var timestamp))
            throw new InvalidDataException($"{fieldType} values must be ISO date/time strings");
        if (fieldType == "datetime" && timezoneName is not null) ValidateDatetimeTimezone(timestamp, timezoneName);
        return;
    }
    if (fieldType == "boolean")
    {
        if (kind is not (JsonValueKind.True or JsonValueKind.False)) throw new InvalidDataException("boolean values must be JSON booleans");
        return;
    }
    if (kind != JsonValueKind.String) throw new InvalidDataException($"{fieldType} values must be JSON strings");
}

static void ValidateDatetimeTimezone(DateTimeOffset timestamp, string timezoneName)
{
    var target = ResolveTimeZone(timezoneName);
    var wallTime = DateTime.SpecifyKind(timestamp.DateTime, DateTimeKind.Unspecified);
    if (target.IsInvalidTime(wallTime)) throw new InvalidDataException("datetime is nonexistent in the declared timezone");
    if (target.IsAmbiguousTime(wallTime)) throw new InvalidDataException("datetime is ambiguous in the declared timezone and cannot be preserved by Excel");
    if (target.GetUtcOffset(wallTime) != timestamp.Offset) throw new InvalidDataException("datetime offset does not match the declared timezone");
}

static TimeZoneInfo ResolveTimeZone(string timezoneName)
{
    try
    {
        return TimeZoneInfo.FindSystemTimeZoneById(timezoneName);
    }
    catch (Exception exception) when (exception is TimeZoneNotFoundException or InvalidTimeZoneException)
    {
        string? platformId = null;
        var converted = OperatingSystem.IsWindows()
            ? TimeZoneInfo.TryConvertIanaIdToWindowsId(timezoneName, out platformId)
            : TimeZoneInfo.TryConvertWindowsIdToIanaId(timezoneName, out platformId);
        if (converted && platformId is not null)
        {
            try
            {
                return TimeZoneInfo.FindSystemTimeZoneById(platformId);
            }
            catch (Exception convertedException) when (convertedException is TimeZoneNotFoundException or InvalidTimeZoneException)
            {
                // Fall through to the stable manifest error below.
            }
        }
    }
    throw new InvalidDataException("datetime timezone is unknown");
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
sealed record RenderOperation(string Type, string? Sheet, string? Cell, uint? Row, string? Before, string? After, string? CanonicalField, uint? HeaderRow, string? HeaderValue, string? FillColor, string? Comment, string? Name, string? SourceJson, JsonElement? Value, List<RenderValue>? Values, string? DifferenceId, string? FieldType, string? NumberFormat, RenderValidation? Validation, string? FormulaTemplate, uint? DataStartRow, List<uint>? FormulaRows, string? Timezone);
sealed record ProductField(
    string Id,
    string Title,
    string Source,
    string FieldType,
    string? NumberFormat,
    string? Timezone,
    List<string> EnumValues,
    bool AllowBlank,
    string? Minimum,
    string? Maximum
);
sealed record ProductIssue(string Color, string Type);
sealed record RenderValue(string? Cell, JsonElement? Value, string? FieldType, string? NumberFormat, string? FormulaTemplate, string? Timezone);
sealed record RenderValidation(string? Type, List<string>? Values, string? Min, string? Max, bool AllowBlank = true);
sealed record OperationResult(int OperationIndex, string Type, string? DifferenceId, string Status, string? ErrorCode, string? Message);
static class JsonOptions { public static readonly JsonSerializerOptions Default = new() { PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow }; }
