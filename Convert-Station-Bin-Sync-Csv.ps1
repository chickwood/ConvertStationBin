# Convert-StationCsv.ps1
# Converts ekidata.jp station CSV to the binary format used by nearest_station_locator.
# Mirrors the logic in station_manager.dart (parseCsv + _toBinary).
# Then follows the logic to generate a readable CSV.
#
# Usage:
#   .\Convert-StationCsv.ps1 -CsvPath "station20260409free.csv" -CsvPriority "station_priority.csv"
#
# Output: station_data.bin and station_readable.csv in the current directory.

param(
    [Parameter(Mandatory=$false)][string]$CsvPath = "station20260409free.csv",
	[Parameter(Mandatory=$false)][string]$CsvPriority = "station_priority.csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── 1. Timestamp from filename (8 digits, e.g. 20260206) ──────────
# ── 路径健壮性处理 ──────────────────────────────────────────────
$finalPath = [System.IO.Path]::GetFullPath($CsvPath)
# 如果解析出的绝对路径不存在，则尝试在当前目录下查找
if (-not (Test-Path $finalPath)) {
    $finalPath = Join-Path (Get-Location) $CsvPath
}
if (-not (Test-Path $finalPath)) {
    Write-Error "找不到 CSV 文件: $CsvPath (尝试路径: $finalPath)"
    Write-Error "退出"
    exit
}
# ────────────────────────────────────────────────────────────────
$fileName = [System.IO.Path]::GetFileName($finalPath)
$tsMatch = [regex]::Match($fileName, '\d{8}')
$date = if ($tsMatch.Success) { [uint32]$tsMatch.Value } else { [uint32]0 }
Write-Host "Date      : $date"

# ── 2. Parse Main CSV (First Independent Read) ──────────────────
# 加载 .NET 文本解析库
Add-Type -AssemblyName "Microsoft.VisualBasic"

$allRows = New-Object System.Collections.Generic.List[PSObject]

$parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($finalPath, [System.Text.Encoding]::UTF8)
$parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
$parser.SetDelimiters(",")
$parser.HasFieldsEnclosedInQuotes = $true

$headerRow = $parser.ReadFields()
$header = $headerRow | ForEach-Object { $_.Trim().ToLower() }

function Col([string]$name) {
    $i = [array]::IndexOf($header, $name)
    if ($i -lt 0) { throw "Required column '$name' not found." }
    return $i
}

$iCd = Col 'station_cd'
$iGcd = Col 'station_g_cd'
$iName = Col 'station_name'
$iLat = Col 'lat'
$iLon = Col 'lon'
$iStatus = Col 'e_status'

while (-not $parser.EndOfData) {
    $row = $parser.ReadFields()

    # 仅当 e_status 字段完全为空时才跳过
    if ($iStatus -ge 0 -and $row[$iStatus].Trim() -eq '') { continue }

    $allRows.Add([PSCustomObject]@{
        station_cd = $row[$iCd].Trim()
        station_g_cd = $row[$iGcd].Trim()
        station_name = $row[$iName].Trim()
        lon = $row[$iLon].Trim()
        lat = $row[$iLat].Trim()
    })
}
$parser.Close()

# 打印读取完成后的集合长度
Write-Host "Raw Rows  : $($allRows.Count)"

# ── 3. Priority CSV (Second Independent Read) ───────────────
$priorityMap = @{}

# ── 路径健壮性处理 ──────────────────────────────────────────────
$finalPriority = [System.IO.Path]::GetFullPath($CsvPriority)
# 如果解析出的绝对路径不存在，则尝试在当前目录下查找
if (-not (Test-Path $finalPriority)) {
    $finalPriority = Join-Path (Get-Location) $CsvPriority
}
if (-not (Test-Path $finalPriority)) {
    Write-Error "无优先 CSV 文件: $CsvPriority (尝试路径: $finalPriority)"
    Write-Error "忽略优先 CSV 文件继续处理"
} else {
    $pParser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($finalPriority, [System.Text.Encoding]::UTF8)
    $pParser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $pParser.SetDelimiters(",")
    $pParser.HasFieldsEnclosedInQuotes = $true

    $pHeaderRow = $pParser.ReadFields() | ForEach-Object { $_.Trim().ToLower() }
    $pHeader = $pHeaderRow | ForEach-Object { $_.Trim().ToLower() }
	
    function PCol([string]$name) {
        $i = [array]::IndexOf($pHeader, $name)
        if ($i -lt 0) { throw "Required column '$name' not found." }
        return $i
    }

    $pIdxCd = PCol "station_cd"
    $pIdxGcd = PCol "station_g_cd"
    $pIdxName = PCol "station_name"
    $pIdxLon = PCol "lon"
    $pIdxLat = PCol "lat"

    while (-not $pParser.EndOfData) {
        $pFields = $pParser.ReadFields()
        $pCd = $pFields[$pIdxCd].Trim()
        if ($pCd) {
            $priorityMap[$pCd] = [PSCustomObject]@{
                station_g_cd = if ($pIdxGcd -ge 0) { $pFields[$pIdxGcd].Trim() } else { "" }
                station_name = if ($pIdxName -ge 0) { $pFields[$pIdxName].Trim() } else { "" }
                lon = if ($pIdxLon -ge 0) { $pFields[$pIdxLon].Trim() } else { "" }
                lat = if ($pIdxLat -ge 0) { $pFields[$pIdxLat].Trim() } else { "" }
            }
        }
    }
    $pParser.Close()
}

# 打印读取完成后的集合长度
Write-Host "Priority  : $($priorityMap.Count)"

# ── 4. Apply Overrides (After BOTH are loaded) ───────────────────
foreach ($rowObj in $allRows) {
    $cd = $rowObj.station_cd
    if ($priorityMap.ContainsKey($cd)) {
        $patch = $priorityMap[$cd]
        if (-not [string]::IsNullOrWhiteSpace($patch.station_g_cd)) { $rowObj.station_g_cd = $patch.station_g_cd }
        if (-not [string]::IsNullOrWhiteSpace($patch.station_name)) { $rowObj.station_name = $patch.station_name }
        if (-not [string]::IsNullOrWhiteSpace($patch.lon)) { $rowObj.lon = $patch.lon }
        if (-not [string]::IsNullOrWhiteSpace($patch.lat)) { $rowObj.lat = $patch.lat }
    }
}

# ── 5. Grouping and Merging ───────────────────────────────────────
$groups = $allRows | Group-Object -Property station_g_cd

$lats = New-Object System.Collections.Generic.List[float]
$lons = New-Object System.Collections.Generic.List[float]
$nameStrings = New-Object System.Collections.Generic.List[string]

function Get-Norm ([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return "" }
    return $name.Trim().
        Replace("(", "（").Replace(")", "）").
        Replace("〈", "（").Replace("〉", "）")
}

foreach ($g in $groups) {
    $mainEntry = $g.Group | Where-Object { $_.station_cd -eq $g.Name } | Select-Object -First 1
    if (-not $mainEntry) { $mainEntry = $g.Group[0] }

    $normSet = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($item in $g.Group) {
        [void]$normSet.Add((Get-Norm $item.station_name))
    }

    $lats.Add([float]$mainEntry.lat)
    $lons.Add([float]$mainEntry.lon)
    $nameStrings.Add("$($g.Name)," + ([string]::Join("/", $normSet)))
}

$count = $lats.Count
Write-Host "Stations  : $count"

# ── 6. Serialize to binary (_toBinary) then serialize to CSV (mirrors _toBinary layout) ───────────────
# Layout:
#   [float32 lat  * count]   lat  区  (count * 4 bytes, offset 0)
#   [float32 lon  * count]   lon  区  (count * 4 bytes, offset count*4)
#   [float32 cos  * count]   cos  区  (count * 4 bytes, offset count*8)
#   [uint32  off  * count]   name offset 区 (count * 4 bytes, offset count*12)
#   [uint8   ...         ]   name data 区  (variable, offset count*16)
#   [uint32  count       ]   4 bytes
#   [uint32  date        ]   4 bytes

# Precompute name offsets
$nameBytes = [byte[][]]::new($count)
$offsets = [uint32[]]::new($count)
$currentNamePos = [uint32]0
for ($i = 0; $i -lt $count; $i++) {
    $nameBytes[$i] = [System.Text.Encoding]::UTF8.GetBytes($nameStrings[$i])
    $offsets[$i] = $currentNamePos
    $currentNamePos += $nameBytes[$i].Length
}
$totalNameBytes = $currentNamePos

# bin
$totalSize = ($count * 16) + $totalNameBytes + 8
$buffer = [byte[]]::new($totalSize)
$ms = [System.IO.MemoryStream]::new($buffer)
$bw = [System.IO.BinaryWriter]::new($ms)
# csv
$sb = [System.Text.StringBuilder]::new()

# Header for the readable CSV format
[void]$sb.AppendLine("station_g_cd,station_name,lat,lon,coslatrad")

# lat区
for ($i = 0; $i -lt $count; $i++) { $bw.Write([float]$lats[$i]) }
# lon区
for ($i = 0; $i -lt $count; $i++) { $bw.Write([float]$lons[$i]) }
# coslatrad区 (cos(lat * pi/180))
$DEG_TO_RAD = [math]::PI / 180.0
for ($i = 0; $i -lt $count; $i++) {
    $cosLatRad = [float]([math]::Cos([double]$lats[$i] * $DEG_TO_RAD))
    $bw.Write($cosLatRad)

    # Mirroring the sequence and logic of the binary serialization loop
    $line = "{0},{1},{2},{3}" -f $nameStrings[$i], $lats[$i], $lons[$i], $cosLatRad
    [void]$sb.AppendLine($line)
}
# name offset区
for ($i = 0; $i -lt $count; $i++) { $bw.Write([uint32]$offsets[$i]) }
# name data区
for ($i = 0; $i -lt $count; $i++) { $bw.Write($nameBytes[$i]) }
# 尾部
$bw.Write([uint32]$count)
$bw.Write([uint32]$date)

# ── 4. Write output ──────────────────────────────────────────────
$bw.Flush()
$outPath = Join-Path (Get-Location) "station_data.bin"
[System.IO.File]::WriteAllBytes($outPath, $buffer)
$outCsvPath = Join-Path (Get-Location) "station_readable.csv"
[System.IO.File]::WriteAllText($outCsvPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
$kb1 = [math]::Round((New-Object System.IO.FileInfo($outPath)).Length / 1024, 1)
$kb2 = [math]::Round((New-Object System.IO.FileInfo($outCsvPath)).Length / 1024, 1)
Write-Host "Written   : $outPath ($kb1 KB) $outCsvPath ($kb2 KB)"
