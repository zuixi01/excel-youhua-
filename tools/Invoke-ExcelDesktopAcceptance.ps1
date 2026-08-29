[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]] $WorkbookPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $EvidencePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-FileSha256 {
    param([Parameter(Mandatory = $true)][string] $Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Release-ComObject {
    param([object] $Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Get-CriticalPackageParts {
    param([Parameter(Mandatory = $true)][string] $Path)
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    $parts = [Collections.Generic.List[object]]::new()
    try {
        $entries = $archive.Entries | Where-Object {
            $_.FullName -match '^(xl/vbaProject(?:Signature)?\.bin|xl/activeX/|xl/ctrlProps/|xl/embeddings/|xl/drawings/.*\.vml$|xl/drawings/_rels/.*\.vml\.rels$)'
        } | Sort-Object FullName
        foreach ($entry in $entries) {
            $stream = $entry.Open()
            $sha = [Security.Cryptography.SHA256]::Create()
            try {
                $digest = $sha.ComputeHash($stream)
                $hex = ([BitConverter]::ToString($digest)).Replace("-", "").ToLowerInvariant()
                $parts.Add([ordered]@{ name = $entry.FullName; length = $entry.Length; sha256 = $hex })
            }
            finally {
                $sha.Dispose()
                $stream.Dispose()
            }
        }
    }
    finally {
        $archive.Dispose()
    }
    return @($parts)
}

function Get-WorksheetNames {
    param([Parameter(Mandatory = $true)][object] $Workbook)
    $names = [Collections.Generic.List[string]]::new()
    $worksheets = $Workbook.Worksheets
    try {
        for ($index = 1; $index -le $worksheets.Count; $index++) {
            $sheet = $worksheets.Item($index)
            try {
                $names.Add([string]$sheet.Name)
            }
            finally {
                Release-ComObject $sheet
            }
        }
    }
    finally {
        Release-ComObject $worksheets
    }
    return @($names)
}

$resolved = [Collections.Generic.List[string]]::new()
$baseNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($candidate in $WorkbookPath) {
    $item = Get-Item -LiteralPath $candidate
    if (-not $item.PSIsContainer -and $item.Extension.ToLowerInvariant() -in @(".xlsx", ".xlsm")) {
        if (-not $baseNames.Add($item.Name)) {
            throw "Workbook base names must be unique in acceptance evidence."
        }
        $resolved.Add($item.FullName)
    }
    else {
        throw "Every acceptance input must be an .xlsx or .xlsm file."
    }
}
if (($resolved | Where-Object { [IO.Path]::GetExtension($_).ToLowerInvariant() -eq ".xlsx" }).Count -lt 1 -or
    ($resolved | Where-Object { [IO.Path]::GetExtension($_).ToLowerInvariant() -eq ".xlsm" }).Count -lt 1) {
    throw "Acceptance requires at least one .xlsx and one .xlsm workbook."
}

$evidenceFullPath = [IO.Path]::GetFullPath($EvidencePath)
if (Test-Path -LiteralPath $evidenceFullPath) {
    throw "Evidence output already exists; choose a new path to preserve prior evidence."
}
$evidenceDirectory = [IO.Path]::GetDirectoryName($evidenceFullPath)
if (-not [IO.Directory]::Exists($evidenceDirectory)) {
    throw "Evidence output directory does not exist."
}

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ("excel-auditor-acceptance-" + [Guid]::NewGuid().ToString("N"))
[void][IO.Directory]::CreateDirectory($temporaryRoot)
$excel = $null
$workbooks = $null
$results = [Collections.Generic.List[object]]::new()
$excelVersion = ""
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
    $excelVersion = [string]$excel.Version
    $workbooks = $excel.Workbooks

    foreach ($source in $resolved) {
        $extension = [IO.Path]::GetExtension($source).ToLowerInvariant()
        $roundtrip = Join-Path $temporaryRoot ([Guid]::NewGuid().ToString("N") + $extension)
        $opened = $false
        $savedCopy = $false
        $roundtripOpened = $false
        $criticalPartsEqual = $false
        $criticalPartNames = @()
        $sourceCriticalParts = @()
        $worksheetNames = @()
        $roundtripSha256 = ""
        $workbook = $null
        $roundtripWorkbook = $null
        try {
            $sourceCriticalParts = @(Get-CriticalPackageParts $source)
            $criticalPartNames = @($sourceCriticalParts | ForEach-Object { [string]$_.name })
            $workbook = $workbooks.Open($source, 0, $true)
            $opened = $true
            $worksheetNames = Get-WorksheetNames $workbook
            $workbook.SaveCopyAs($roundtrip)
            $savedCopy = Test-Path -LiteralPath $roundtrip
        }
        catch {
            $opened = $false
        }
        finally {
            if ($null -ne $workbook) {
                try { $workbook.Close($false) } catch { }
                Release-ComObject $workbook
            }
        }
        if ($savedCopy) {
            try {
                $roundtripWorkbook = $workbooks.Open($roundtrip, 0, $true)
                $roundtripOpened = $true
                $roundtripNames = Get-WorksheetNames $roundtripWorkbook
                $roundtripOpened = (@($roundtripNames) -join "`0") -eq (@($worksheetNames) -join "`0")
            }
            catch {
                $roundtripOpened = $false
            }
            finally {
                if ($null -ne $roundtripWorkbook) {
                    try { $roundtripWorkbook.Close($false) } catch { }
                    Release-ComObject $roundtripWorkbook
                }
            }
            if ($roundtripOpened) {
                $roundtripSha256 = Get-FileSha256 $roundtrip
                $beforeParts = @($sourceCriticalParts) | ConvertTo-Json -Compress -Depth 5
                $afterParts = @(Get-CriticalPackageParts $roundtrip) | ConvertTo-Json -Compress -Depth 5
                $criticalPartsEqual = $beforeParts -ceq $afterParts
            }
        }
        $results.Add([ordered]@{
            file_name = [IO.Path]::GetFileName($source)
            extension = $extension
            input_sha256 = Get-FileSha256 $source
            roundtrip_sha256 = $roundtripSha256
            opened = $opened
            roundtrip_opened = $roundtripOpened
            saved_copy = $savedCopy
            worksheet_names = @($worksheetNames)
            critical_part_count = $criticalPartNames.Count
            critical_part_names = @($criticalPartNames)
            critical_parts_equal = $criticalPartsEqual
        })
    }

    $xlsxCount = @($results | Where-Object { $_.extension -eq ".xlsx" }).Count
    $xlsmCount = @($results | Where-Object { $_.extension -eq ".xlsm" }).Count
    $allChecksPassed = @($results | Where-Object {
        -not ($_.opened -and $_.roundtrip_opened -and $_.saved_copy -and $_.critical_parts_equal)
    }).Count -eq 0
    $evidence = [ordered]@{
        schema_version = "1.0"
        generated_at = [DateTimeOffset]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
        excel_version = $excelVersion
        macro_execution = "force_disabled"
        files = @($results)
        summary = [ordered]@{
            total = $results.Count
            xlsx = $xlsxCount
            xlsm = $xlsmCount
            all_checks_passed = $allChecksPassed
        }
    }
    [IO.File]::WriteAllText($evidenceFullPath, ($evidence | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    if (-not $allChecksPassed) {
        [Console]::Error.WriteLine("One or more Excel desktop acceptance checks failed; evidence was preserved.")
        exit 2
    }
    Write-Output "Excel desktop automated evidence created: $evidenceFullPath"
}
finally {
    Release-ComObject $workbooks
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        Release-ComObject $excel
    }
    if ([IO.Directory]::Exists($temporaryRoot)) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
