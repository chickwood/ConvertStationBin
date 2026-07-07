# Convert-StationCsv.ps1
# Converts ekidata.jp station CSV to the binary format used by nearest_station_locator.
# Mirrors the logic in station_manager.dart (parseCsv + _toBinary).
# Then follows the logic to generate a readable CSV.
#
# Usage:
#   .\Convert-StationCsv.ps1 -CsvPath "station20260528free.csv" -CsvPriority "station_priority.csv"
#
# Output: station_data_xyz.bin and station_readable_xyz.csv in the current directory.

param(
    [Parameter(Mandatory=$false)][string]$CsvPath = "station20260528free.csv",
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
    Write-Error "找不到 CSV 文件: $CsvPath (尝试路径: $finalPath)`n使用 -CsvPath 参数指定文件"
    exit
}
# ────────────────────────────────────────────────────────────────
$fileName = [System.IO.Path]::GetFileName($finalPath)
$tsMatch = [regex]::Match($fileName, '\d{8}')
$date = if ($tsMatch.Success) { [uint32]$tsMatch.Value } else { [uint32]0 }
Write-Host "Date      : $date"

# 加载 .NET 文本解析库
Add-Type -AssemblyName "Microsoft.VisualBasic"

# ── 2. Parse Main CSV (First Independent Read) ──────────────────
$allMap = @{}

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
$iLon = Col 'lon'
$iLat = Col 'lat'
$iStatus = Col 'e_status'

while (-not $parser.EndOfData) {
    $row = $parser.ReadFields()

    # 仅当 e_status 字段完全为空时才跳过
    if ($iStatus -ge 0 -and $row[$iStatus].Trim() -eq '') { continue }

    $cd = $row[$iCd].Trim()
    # 以 station_cd 作为 Key 存入哈希表
    $allMap[$cd] = [PSCustomObject]@{
        station_cd = $cd
        station_g_cd = $row[$iGcd].Trim()
        station_name = $row[$iName].Trim()
        lon = $row[$iLon].Trim()
        lat = $row[$iLat].Trim()
        status = $row[$iStatus].Trim()
    }
}
$parser.Close()

# 打印读取完成后的集合长度
Write-Host "Raw Rows  : $($allMap.Count)"

# ── 3. Priority CSV (Second Independent Read) ───────────────
$priorityMap = @{}
$STATUS_DEL = 9

# ── 路径健壮性处理 ──────────────────────────────────────────────
$finalPriority = [System.IO.Path]::GetFullPath($CsvPriority)
# 如果解析出的绝对路径不存在，则尝试在当前目录下查找
if (-not (Test-Path $finalPriority)) {
    $finalPriority = Join-Path (Get-Location) $CsvPriority
}
if (-not (Test-Path $finalPriority)) {
    Write-Warning "无优先 CSV 文件: $CsvPriority (尝试路径: $finalPriority)`n忽略优先 CSV 文件继续处理或使用 -CsvPriority 参数指定文件"
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
    $pIdxStatus = PCol "e_status"

    while (-not $pParser.EndOfData) {
        $pFields = $pParser.ReadFields()
        $pCd = $pFields[$pIdxCd].Trim()
        if ($pCd) {
            $priorityMap[$pCd] = [PSCustomObject]@{
                station_cd = $pCd
                station_g_cd = if ($pIdxGcd -ge 0) { $pFields[$pIdxGcd].Trim() } else { "" }
                station_name = if ($pIdxName -ge 0) { $pFields[$pIdxName].Trim() } else { "" }
                lon = if ($pIdxLon -ge 0) { $pFields[$pIdxLon].Trim() } else { "" }
                lat = if ($pIdxLat -ge 0) { $pFields[$pIdxLat].Trim() } else { "" }
                status = if ($pIdxStatus -ge 0) { $pFields[$pIdxStatus].Trim() } else { "" }
            }
        }
    }
    $pParser.Close()
}

# 打印读取完成后的集合长度
Write-Host "Priority  : $($priorityMap.Count)"

# ── 4. Apply Overrides (After BOTH are loaded) ───────────────────
foreach ($cd in $priorityMap.Keys) {
    $patch = $priorityMap[$cd]
    if ($allMap.ContainsKey($cd)) {
        # 如果 allMap 已存在该 Key，执行字段覆盖
        $rowObj = $allMap[$cd]
        if (-not [string]::IsNullOrWhiteSpace($patch.station_g_cd)) { $rowObj.station_g_cd = $patch.station_g_cd }
        if (-not [string]::IsNullOrWhiteSpace($patch.station_name)) { $rowObj.station_name = $patch.station_name }
        if (-not [string]::IsNullOrWhiteSpace($patch.lon)) { $rowObj.lon = $patch.lon }
        if (-not [string]::IsNullOrWhiteSpace($patch.lat)) { $rowObj.lat = $patch.lat }
        if (-not [string]::IsNullOrWhiteSpace($patch.status)) { $rowObj.status = $patch.status }
    } else {
        $allMap[$cd] = $patch
    }
}

# ── 5. Grouping and Merging ───────────────────────────────────────
$groups = $allMap.Values | Where-Object { $_.status -ne '9' } | Group-Object -Property station_g_cd

$nodeList = New-Object System.Collections.Generic.List[PSObject]
$DEG_TO_RAD = [math]::PI / 180.0

function Get-Norm ([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return "" }
    return $name.Trim().
        Replace("(", "（").Replace(")", "）").
        Replace("〈", "（").Replace("〉", "）")
}

foreach ($g in $groups) {
    $mainEntry = $g.Group | Where-Object { $_.station_cd -eq $g.Name } | Select-Object -First 1
    if (-not $mainEntry) { $mainEntry = $g.Group[0] }

    $normList = New-Object System.Collections.Generic.List[string]
    $normName = Get-Norm $mainEntry.station_name
    $normList.Add($normName)
    foreach ($item in $g.Group) {
        $normName = Get-Norm $item.station_name
        if (-not $normList.Contains($normName)) {
            $normList.Add($normName)
        }
    }
    # 处理名称后缀逻辑
    $joinedNames = [string]::Join("/", $normList)
    if ($g.Group.status -notcontains '0') {
        $joinedNames += "*"
    }

    $latVal = [double]$mainEntry.lat
    $lonVal = [double]$mainEntry.lon
    $latRad = $latVal * $DEG_TO_RAD
    $lonRad = $lonVal * $DEG_TO_RAD

    $nodeList.Add([PSCustomObject]@{
        station_g_cd = $g.Name
        names = $joinedNames
        lat = [double]$latVal
        lon = [double]$lonVal
        vx = [double]([math]::Cos($latRad) * [math]::Cos($lonRad))
        vy = [double]([math]::Cos($latRad) * [math]::Sin($lonRad))
        vz = [double]([math]::Sin($latRad))
        left = $null
        right = $null
        newId = -1
        isLeaf = $false
    })
}

$count = $nodeList.Count
Write-Host "Stations  : $count"

# ── 6. Build KD-Tree (Iterative, Stack-based) ────────────────────
# left_count(n): 完全二叉树中，n个节点的左子树节点数
# 使用此函数代替 n//2，确保树满足完全二叉树约束（最多1个单侧非叶节点）
function left_count([int]$n) {
    if ($n -le 1) { return 0 }
    $h = [int][math]::Floor([math]::Log($n, 2))
    $lastLevel = $n - ([math]::Pow(2, $h) - 1)
    $leftLast = [math]::Min($lastLevel, [math]::Pow(2, $h - 1))
    return [int](([math]::Pow(2, $h - 1) - 1) + $leftLast)
}
# 使用4个并行Stack代替PSCustomObject栈帧，避免属性存取时类型信息丢失
$stkNodes = New-Object System.Collections.Generic.Stack[object] # ArrayList per frame
$stkDepth = New-Object System.Collections.Generic.Stack[int]
$stkParent = New-Object System.Collections.Generic.Stack[object] # PSCustomObject or $null
$stkSide = New-Object System.Collections.Generic.Stack[string]

$initList = New-Object System.Collections.ArrayList
foreach ($nd in $nodeList) { [void]$initList.Add($nd) }
$stkNodes.Push($initList)
$stkDepth.Push(0)
$stkParent.Push($null)
$stkSide.Push('')

$treeRoot = $null

while ($stkNodes.Count -gt 0) {
    $currList = [System.Collections.ArrayList]$stkNodes.Pop()
    $depth = [int]$stkDepth.Pop()
    $parent = $stkParent.Pop()
    $side = [string]$stkSide.Pop()

    $n = $currList.Count
    if ($n -eq 0) { continue }

    $prop = ("vx","vy","vz")[$depth % 3]
    # Sort返回单元素时是裸对象，用ArrayList接收保证Count可用
    $sortedList = New-Object System.Collections.ArrayList
    foreach ($nd in ($currList | Sort-Object -Property $prop)) { [void]$sortedList.Add($nd) }
    $mid = left_count $n
    $node = $sortedList[$mid]

    if ($null -eq $parent) { $treeRoot = $node }
    elseif ($side -eq 'L') { $parent.left = $node }
    else { $parent.right = $node }

    $leftList = New-Object System.Collections.ArrayList
    $rightList = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $mid; $i++) { [void]$leftList.Add($sortedList[$i]) }
    for ($i = $mid + 1; $i -lt $n; $i++) { [void]$rightList.Add($sortedList[$i]) }

    $node.isLeaf = ($leftList.Count -eq 0 -and $rightList.Count -eq 0)

    if ($rightList.Count -gt 0) {
        $stkNodes.Push($rightList); $stkDepth.Push($depth + 1); $stkParent.Push($node); $stkSide.Push('R')
    }
    if ($leftList.Count -gt 0) {
        $stkNodes.Push($leftList); $stkDepth.Push($depth + 1); $stkParent.Push($node); $stkSide.Push('L')
    }
}

# ── 7. Reindex: Non-leaf [0, mCount), Leaf [mCount, count) ────────
[array]$internalNodes = @($nodeList | Where-Object { $_.isLeaf -eq $false })
[array]$leafNodes = @($nodeList | Where-Object { $_.isLeaf -eq $true })

$mCount = $internalNodes.Count
$lCount = $leafNodes.Count
$nCount = $mCount + $lCount
Write-Host "M Count   : $mCount"
Write-Host "L Count   : $lCount"
Write-Host "M+L Count : $nCount"

for ($i = 0; $i -lt $mCount; $i++) { $internalNodes[$i].newId = $i }
for ($i = 0; $i -lt $leafNodes.Count; $i++) { $leafNodes[$i].newId = $i + $mCount }

$finalNodes = New-Object System.Collections.Generic.List[PSObject]
foreach ($n in $internalNodes) { $finalNodes.Add($n) }
foreach ($n in $leafNodes) { $finalNodes.Add($n) }

# Validate mCount constraint: at most one non-leaf may have a single child
$NULL_IDX = [uint32]4294967295 # [uint32]::MaxValue
$singleChildCount = 0
for ($i = 0; $i -lt $mCount; $i++) {
    $nd = $finalNodes[$i]
    $hasLeft = ($null -ne $nd.left)
    $hasRight = ($null -ne $nd.right)
    if (-not $hasLeft -or -not $hasRight) {
        $singleChildCount++
        if ($singleChildCount -eq 1) {
            Write-Host "[INFO] Non-leaf node #$($nd.newId) ($($nd.station_g_cd)) is the first single-child non-leaf."
        } elseif ($singleChildCount -gt 1) {
            Write-Host "[ERROR] Non-leaf node #$($nd.newId) ($($nd.station_g_cd)) is the $singleChildCount-th single-child non-leaf (max 1 allowed)!"
        }
    }
}

$root = [uint32]$treeRoot.newId
Write-Host "Root      : $root ($($treeRoot.station_g_cd))"

# ── 8. Serialize to binary (_toBinary) then serialize to CSV (mirrors _toBinary layout) ───────────────
# Layout:
#   [float64 lat  * count]   lat  区
#   [float64 lon  * count]   lon  区
#   [float64 vx   * count]   vx   区
#   [float64 vy   * count]   vy   区
#   [float64 vz   * count]   vz   区
#   [uint32  left * mCount]  leftChildIndex  区 (non-leaf only)
#   [uint32  right* mCount]  rightChildIndex 区 (non-leaf only)
#   [uint32  off  * count]   name offset 区
#   [uint8   ...         ]   name data 区  (variable)
#   [uint32  root   ]   4 bytes
#   [uint32  mCount      ]   4 bytes
#   [uint32  count       ]   4 bytes
#   [uint32  date        ]   4 bytes

# Precompute name offsets
$nameBytes = [byte[][]]::new($count)
$offsets = [uint32[]]::new($count)
$currentNamePos = [uint32]0
for ($i = 0; $i -lt $count; $i++) {
    $nameBytes[$i] = [System.Text.Encoding]::UTF8.GetBytes("$($finalNodes[$i].station_g_cd),$($finalNodes[$i].names)")
    $offsets[$i] = $currentNamePos
    $currentNamePos += $nameBytes[$i].Length
}
$totalNameBytes = $currentNamePos

# bin
$totalSize = (5 * 8 * $count) + (2 * 4 * $mCount) + (4 * $count) + $totalNameBytes + 16
$buffer = [byte[]]::new($totalSize)
$ms = [System.IO.MemoryStream]::new($buffer)
$bw = [System.IO.BinaryWriter]::new($ms)
# csv
$sb = [System.Text.StringBuilder]::new()
# Header for the readable CSV format
[void]$sb.AppendLine("station_g_cd,station_name,lat,lon,vx,vy,vz,left,right,isLeaf")

Write-Host "lat,lon区"
# lat区
for ($i = 0; $i -lt $count; $i++) { $bw.Write([double]$finalNodes[$i].lat) }
# lon区
for ($i = 0; $i -lt $count; $i++) { $bw.Write([double]$finalNodes[$i].lon) }
Write-Host "lat,lon区完了"
Write-Host "x,y,z区"
# vx区
for ($i = 0; $i -lt $count; $i++) { $bw.Write([double]$finalNodes[$i].vx) }
# vy区
for ($i = 0; $i -lt $count; $i++) { $bw.Write([double]$finalNodes[$i].vy) }
# vz区
for ($i = 0; $i -lt $count; $i++) { $bw.Write([double]$finalNodes[$i].vz) }
Write-Host "x,y,z区完了"
Write-Host "left区"
# leftChildIndex区 (non-leaf only)
for ($i = 0; $i -lt $mCount; $i++) {
    $idx = if ($null -ne $finalNodes[$i].left) { [uint32]$finalNodes[$i].left.newId } else { $NULL_IDX }
    $bw.Write($idx)
}
Write-Host "left区完了"
Write-Host "right区"
# rightChildIndex区 (non-leaf only)
for ($i = 0; $i -lt $mCount; $i++) {
    $idx = if ($null -ne $finalNodes[$i].right) { [uint32]$finalNodes[$i].right.newId } else { $NULL_IDX }
    $bw.Write($idx)
}
Write-Host "right区完了"
Write-Host "offset区"
# name offset区
for ($i = 0; $i -lt $count; $i++) { $bw.Write([uint32]$offsets[$i]) }
Write-Host "offset区完了"
Write-Host "name区"
# name data区
for ($i = 0; $i -lt $count; $i++) { $bw.Write($nameBytes[$i]) }
Write-Host "name区完了"
Write-Host "校验区"
# 尾部
$bw.Write([uint32]$root)
$bw.Write([uint32]$mCount)
$bw.Write([uint32]$count)
$bw.Write([uint32]$date)
Write-Host "校验区完了"

# ── 9. Write output ──────────────────────────────────────────────
Write-Host "Write output"
$bw.Flush()
$outPath = Join-Path (Get-Location) "station_data_xyz.bin"
[System.IO.File]::WriteAllBytes($outPath, $buffer)
Write-Host "Write output完了"

# ── 10. Readable buffer ──────────────────────────────────────────────
# Mirroring the sequence and logic of the binary serialization loop
for ($i = 0; $i -lt $count; $i++) {
    $nd = $finalNodes[$i]
    if ($nd.isLeaf) {
        $leftStr = ""
        $rightStr = ""
    } else {
        $leftStr = if ($null -ne $nd.left) { [string]$nd.left.newId } else { [string]$NULL_IDX }
        $rightStr = if ($null -ne $nd.right) { [string]$nd.right.newId } else { [string]$NULL_IDX }
    }
    $line = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}" -f `
        $nd.station_g_cd, $nd.names, $nd.lat, $nd.lon, $nd.vx, $nd.vy, $nd.vz, $leftStr, $rightStr, $nd.isLeaf
    [void]$sb.AppendLine($line)
}

# ── 11. Write readable ──────────────────────────────────────────────
Write-Host "Write readable"
$outCsvPath = Join-Path (Get-Location) "station_readable_xyz.csv"
[System.IO.File]::WriteAllText($outCsvPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
Write-Host "Write readable完了"

# ── 12. Console write result ──────────────────────────────────────────────
$kb1 = [math]::Round((New-Object System.IO.FileInfo($outPath)).Length / 1024, 1)
$kb2 = [math]::Round((New-Object System.IO.FileInfo($outCsvPath)).Length / 1024, 1)
Write-Host "Written   : $outPath ($kb1 KB) $outCsvPath ($kb2 KB)"
