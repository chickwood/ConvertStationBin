param (
    [Parameter(Mandatory=$false)][string]$CsvPath = "station20260430free.csv",
    [Parameter(Mandatory=$false)][string]$OutputFile = "DivergenceReport.txt",
    [Parameter(Mandatory=$false)][string]$ExtraFile = "ExtraEntries.txt"
)

$CsvPriority = "station_priority.csv"

# ── 路径健壮性处理 ──────────────────────────────────────────────
$finalPath = [System.IO.Path]::GetFullPath($CsvPath)

# 如果解析出的绝对路径不存在，则尝试在当前目录下查找
if (-not (Test-Path $finalPath)) {
    $finalPath = Join-Path (Get-Location) $CsvPath
}

if (-not (Test-Path $finalPath)) {
    Write-Error "找不到CSV文件: $CsvPath (尝试路径: $finalPath)"
    exit
}
# ────────────────────────────────────────────────────────────────

# ── 路径健壮性处理 ──────────────────────────────────────────────
$finalPriority = [System.IO.Path]::GetFullPath($CsvPriority)

# 如果解析出的绝对路径不存在，则尝试在当前目录下查找
if (-not (Test-Path $finalPriority)) {
    $finalPriority = Join-Path (Get-Location) $CsvPriority
}

if (-not (Test-Path $finalPriority)) {
    Write-Error "无优先CSV文件: $CsvPriority (尝试路径: $finalPriority)"
}
# ────────────────────────────────────────────────────────────────

# 辅助函数：将任何形式的括号统一替换为全角〈〉，用于逻辑比对
function Get-NormalizedName ($name) {
    if (-not $name) { return "" }
    # 将全角（）和半角() 统一替换为 〈〉
    $n = $name.Trim() -replace "\(", "〈" -replace "\)", "〉" -replace "（", "〈" -replace "）", "〉"
    return $n
}

# 1. 导入数据
$data = Import-Csv -Path $finalPath -Encoding UTF8

# ── 数据修正与覆盖逻辑 (数据B 覆盖 数据A) ────────────────────────
if (Test-Path $finalPriority) {
    $priorityData = Import-Csv -Path $finalPriority -Encoding UTF8
    $priorityMap = @{}
    foreach ($row in $priorityData) {
        if ($row.station_cd) { $priorityMap[$row.station_cd.Trim()] = $row }
    }

    foreach ($item in $data) {
        $cd = $item.station_cd.Trim()
        if ($priorityMap.ContainsKey($cd)) {
            $patch = $priorityMap[$cd]
            if (-not [string]::IsNullOrWhiteSpace($patch.station_g_cd)) { $item.station_g_cd = $patch.station_g_cd.Trim() }
            if (-not [string]::IsNullOrWhiteSpace($patch.station_name)) { $item.station_name = $patch.station_name.Trim() }
            if (-not [string]::IsNullOrWhiteSpace($patch.lon)) { $item.lon = $patch.lon.Trim() }
            if (-not [string]::IsNullOrWhiteSpace($patch.lat)) { $item.lat = $patch.lat.Trim() }
        }
    }
}
# ────────────────────────────────────────────────────────────────

# 2. 按 station_g_cd 分组
$groups = $data | Group-Object -Property station_g_cd

$report = New-Object System.Collections.Generic.List[string]
$extraReport = New-Object System.Collections.Generic.List[string]

$report.Add("GCD Divergence Report - $(Get-Date)")
$report.Add("Normalization: All brackets (),（） treated as 〈〉 for comparison.")
$report.Add("-" * 60)

$conflictCount = 0

foreach ($g in $groups) {
    $gcd = $g.Name

    # 寻找基准站
    $mainEntry = $g.Group | Where-Object { $_.station_cd -eq $gcd } | Select-Object -First 1
    if (-not $mainEntry) {
        Write-Error "组CD不存在对应车站CD可能是被删除的新干线车站 $gcd "
        $mainEntry = $g.Group[0]
    }

    # 基准站的原始信息
    $n0_raw = $mainEntry.station_name.Trim()
    $c0_raw = $mainEntry.station_cd.Trim()
    $p0 = "$($mainEntry.lon.Trim()),$($mainEntry.lat.Trim())"

    # 用于逻辑比对的归一化名字
    $n0_norm = Get-NormalizedName $n0_raw

    $aliases = New-Object System.Collections.Generic.List[string]
    $independentEntries = [Ordered]@{}

    foreach ($item in $g.Group) {
        # 跳过基准站自身
        if ($item.station_cd -eq $mainEntry.station_cd) { continue }

        $currNameRaw = $item.station_name.Trim()
        #$currCD = $item.station_cd.Trim()
        $currNameNorm = Get-NormalizedName $currNameRaw
        $currPos = "$($item.lon.Trim()),$($item.lat.Trim())"

        # 1. 名字相同判断（使用归一化后的名字比对）
        if ($currNameNorm -eq $n0_norm) { continue }

        # 2. 名字不同，坐标相同：归为别名
        if ($currPos -eq $p0) {
            # 存入别名时带上 CD
            $aliasEntry = "$currNameRaw"
            if (-not $aliases.Contains($aliasEntry)) { $aliases.Add($aliasEntry) }
        }
        # 3. 名字不同，坐标不同：归为独立项
        else {
            if (-not $independentEntries.Contains($currPos)) {
                $independentEntries[$currPos] = New-Object System.Collections.Generic.List[string]
            }
            $list = $independentEntries[$currPos]
            $indepEntry = "$currNameRaw"
            if (-not $list.Contains($indepEntry)) {
                $list.Add($indepEntry)
            }
        }
    }

    # 判定输出条件
    if ($aliases.Count -gt 0 -or $independentEntries.Count -gt 0) {
        $conflictCount++

        # 格式化主条目输出 (包含基准 CD)
        $mainDisplayName = "[$c0_raw] $n0_raw"
        if ($aliases.Count -gt 0) { $mainDisplayName += "/" + ($aliases -join "/") }

        $report.Add("Group CD: $gcd")
        $report.Add("  -> $mainDisplayName | $p0")

        foreach ($pos in $independentEntries.Keys) {
            $otherNames = $independentEntries[$pos]
            $otherDisplayName = if ($otherNames.Count -gt 1) {
                $first = $otherNames[0]
                $rest = ($otherNames | Select-Object -Skip 1) -join "/"
                "$first($rest)"
            } else {
                $otherNames[0]
            }

            # 记录到主报告
            $line = "  -> $otherDisplayName | $pos"
            $report.Add($line)
            # 同时记录到附加文件报告
            $extraReport.Add($line)
        }
        $report.Add("-" * 60)
    }
}

$summary = "扫描完成！共发现 $conflictCount 组符合条件的车站组。"
$report.Add($summary)

$report
$extraReport

# 输出主报告
$report | Out-File -FilePath $OutputFile -Encoding UTF8

# 输出附加文件 (非主条目行)
if ($extraReport.Count -gt 0) {
    $extraReport | Out-File -FilePath $ExtraFile -Encoding UTF8
}

Write-Host $summary -ForegroundColor Green
Write-Host "Done. Main report: $OutputFile, Extra items: $ExtraFile" -ForegroundColor Green
