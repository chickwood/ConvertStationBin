# Read-StationBin.ps1
# 读取 station_data_xyz.bin 文件中指定索引的车站信息。
#
# Usage:
#   .\Read-StationBin.ps1 -Index 4
#   .\Read-StationBin.ps1 -BinPath "my_station_data_xyz.bin" -Index 10

param(
    [Parameter(Mandatory=$false)][string]$BinPath = "station_data_xyz.bin",
    [Parameter(Mandatory=$true)][int]$Index
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── 路径健壮性处理 ──────────────────────────────────────────────
# 1. 尝试获取绝对路径
$finalPath = [System.IO.Path]::GetFullPath($BinPath)

# 2. 如果绝对路径不存在，尝试在当前工作目录拼接
if (-not (Test-Path $finalPath)) {
    $finalPath = Join-Path (Get-Location) $BinPath
}

# 3. 最终检查
if (-not (Test-Path $finalPath)) {
    Write-Error "找不到二进制文件: $BinPath (尝试路径: $finalPath)"
    exit
}
# ────────────────────────────────────────────────────────────────

# 将整个文件读入内存
$bytes = [System.IO.File]::ReadAllBytes($finalPath)
$totalLen = $bytes.Length

# ── 1. 解析尾部信息获取总数 ──────────────────────────────────
# 倒数第 16-13 字节是 root (uint32)
$root = [BitConverter]::ToUInt32($bytes, $totalLen - 16)
# 倒数第 12-9 字节是 mCount (uint32)
$mCount = [BitConverter]::ToUInt32($bytes, $totalLen - 12)
# 倒数第 8-5 字节是 Count (uint32)
$count = [BitConverter]::ToUInt32($bytes, $totalLen - 8)
# 倒数第 4-1 字节是 Date (uint32)
$date  = [BitConverter]::ToUInt32($bytes, $totalLen - 4)

if ($Index -lt 0 -or $Index -ge $count) {
    Write-Error "索引超出范围。有效范围: 0 到 $($count - 1)"
    exit
} elseif ($Index -ge $mCount) {
    Write-Host "索引超出非叶节点范围。非叶节点范围: 0 到 $($mCount - 1)"
}

Write-Host "File Path : $finalPath" -ForegroundColor Gray
Write-Host "File Date : $date"
Write-Host "Total     : $count stations"
Write-Host "Non Leaf  : $mCount stations"
Write-Host "Root      : $root"
Write-Host "----------------------------"

# ── 2. 计算各区偏移量 (与生成脚本 layout 严格对应) ──────────────
$latBase    = 0
$lonBase    = $count * 8
$xBase      = $count * 16
$yBase      = $count * 24
$zBase      = $count * 32
$lBase      = $count * 40
$rBase      = $count * 40 + $mCount * 4
$offsetBase = $count * 40 + $mCount * 8
$nameDataBase = $count * 44 + $mCount * 8

# ── 3. 提取指定索引的数据 ────────────────────────────────────
$fieldOffset = $Index * 8

# 读取 Lat, Lon, Cos (float64)
$lat = [BitConverter]::ToDouble($bytes, $latBase + $fieldOffset)
$lon = [BitConverter]::ToDouble($bytes, $lonBase + $fieldOffset)
$x   = [BitConverter]::ToDouble($bytes, $xBase   + $fieldOffset)
$y   = [BitConverter]::ToDouble($bytes, $yBase   + $fieldOffset)
$z   = [BitConverter]::ToDouble($bytes, $zBase   + $fieldOffset)
$l   = ($Index -ge $mCount) ? "" : [BitConverter]::ToUInt32($bytes, $lBase   + $Index * 4)
$r   = ($Index -ge $mCount) ? "" : [BitConverter]::ToUInt32($bytes, $rBase   + $Index * 4)
$isLeaf = ($Index -ge $mCount) ? "True" : "False"

# 读取 Name 字符串
$nameStartRel = [BitConverter]::ToUInt32($bytes, $offsetBase + $Index * 4)

# 计算长度：取下一个 offset 或到数据区末尾
if ($Index -lt $count - 1) {
    $nameEndRel = [BitConverter]::ToUInt32($bytes, $offsetBase + $Index * 4 + 4)
} else {
    # 减 8 是因为文件末尾有 8 字节的 count 和 date
    $nameEndRel = $totalLen - $nameDataBase - 8
}

$nameLen = $nameEndRel - $nameStartRel
$nameStr = [System.Text.Encoding]::UTF8.GetString($bytes, $nameDataBase + $nameStartRel, $nameLen)

# ── 4. 输出结果 ──────────────────────────────────────────────
[PSCustomObject]@{
    Index    = $Index
    GCD_Name = $nameStr
    Lat      = $lat
    Lon      = $lon
    X        = $x  
	Y        = $y  
	Z        = $z  
	Left     = $l  
	Right    = $r  
	IsLeaf   = $isLeaf
} | Format-List
