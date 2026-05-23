<#
.SYNOPSIS
    从 FTP 服务器下载当天日期的早读 PPT 到桌面，并自动清理过期副本
.DESCRIPTION
    1. 删除桌面上所有带时间戳且日期超过1天的 PPT 文件
    2. 从配置的 FTP 路径获取当天日期 PPT；若未找到，则从备用路径查找
    3. 下载所有当天日期的 PPT 到桌面，添加日期戳，并用默认程序打开
    4. 打开 PPT 前自动最小化所有窗口（仅在有PPT打开时执行，且不包含本程序窗口）
.COPYRIGHT
    Copyright (c) 嵇子扬
.NOTES
    保存时请选择 "UTF-8 with BOM" 编码，否则中文注释可能导致解析错误
    配置文件 config.json 必须与脚本放在同一目录下
#>

[Console]::OutputEncoding = [System.Text.Encoding]::Default

# 显示版权信息
Write-Host "========================================" -ForegroundColor DarkBlue
Write-Host "  早读 PPT 自动下载工具" -ForegroundColor White
Write-Host "  Copyright (c) 嵇子扬" -ForegroundColor White
Write-Host "========================================" -ForegroundColor DarkBlue
Write-Host ""

# ================= 读取配置文件 =================
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path $scriptDir "config.json"

if (-not (Test-Path $configPath)) {
    Write-Host "[错误] 未找到配置文件 config.json" -ForegroundColor Red
    Write-Host "请将 config.json.example 复制为 config.json 并填写正确的 FTP 信息。" -ForegroundColor Yellow
    Start-Sleep -Seconds 10
    exit 1
}

try {
    $config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
} catch {
    Write-Host "[错误] 配置文件解析失败，请检查 JSON 格式是否正确" -ForegroundColor Red
    Start-Sleep -Seconds 10
    exit 1
}

# 必要字段检查
$required = @('ftp_server', 'ftp_username', 'ftp_password', 'primary_remote_path')
foreach ($field in $required) {
    if (-not $config.$field) {
        Write-Host "[错误] 配置文件缺少必要字段: $field" -ForegroundColor Red
        Start-Sleep -Seconds 10
        exit 1
    }
}

$ftpServer   = $config.ftp_server
$ftpUser     = $config.ftp_username
$ftpPass     = $config.ftp_password
$remotePath1 = $config.primary_remote_path
if ($config.secondary_remote_path) {
    $remotePath2 = $config.secondary_remote_path
} else {
    $remotePath2 = $null
}

# 确保远程路径以斜杠结尾
if (-not $remotePath1.EndsWith('/')) { $remotePath1 += '/' }
if ($remotePath2 -and -not $remotePath2.EndsWith('/')) { $remotePath2 += '/' }

$desktop = [Environment]::GetFolderPath("Desktop")
$cred    = New-Object System.Net.NetworkCredential($ftpUser, $ftpPass)
$today   = (Get-Date).Date
# =================================================

# --- 第一步：清理桌面上过期 PPT ---
Write-Host "正在清理过期文件..." -ForegroundColor DarkBlue
$thresholdDate = $today.AddDays(-1)
$datePattern   = '_(?<date>\d{8})\.ppt[x]?$'

Get-ChildItem -Path $desktop -Filter "*.ppt*" | ForEach-Object {
    if ($_.Name -match $datePattern) {
        $dateStr = $Matches['date']
        try {
            $fileDate = [datetime]::ParseExact($dateStr, 'yyyyMMdd', $null)
            if ($fileDate -lt $thresholdDate) {
                Write-Host "  删除过期文件: $($_.Name)  (日期: $dateStr)"
                Remove-Item $_.FullName -Force -ErrorAction Continue
            }
        } catch { }
    }
}
Write-Host "清理完成。" -ForegroundColor White
Write-Host ""

# --- 辅助函数：获取指定远程目录下所有 PPT，并筛选当天日期的文件 ---
function Get-TodayPptFromFtp {
    param([string]$RemoteDir)
    $baseUri = "ftp://$ftpServer$RemoteDir"

    # 列出目录
    $req = [System.Net.FtpWebRequest]::Create($baseUri)
    $req.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectory
    $req.Credentials = $cred
    $req.UsePassive = $true
    $res = $req.GetResponse()
    $stream = $res.GetResponseStream()
    $rdr = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::Default)
    $allNames = ($rdr.ReadToEnd() -split "`r`n") | Where-Object { $_ -ne "" }
    $rdr.Close(); $res.Close()

    $pptFiles = @($allNames | Where-Object { $_ -match '\.ppt[x]?$' })
    if ($pptFiles.Count -eq 0) { return @() }

    # 正则：支持个位数的月/日，前面的0可选
    $dateRegex = '(?<!\d)(0?[1-9]|1[0-2])\.(0?[1-9]|[12]\d|3[01])(?!\d)'

    $todayFiles = @()
    foreach ($f in $pptFiles) {
        # 净化文件名，防止路径注入
        $safeFileName = [System.IO.Path]::GetFileName($f)

        $extractedDateStr = $null
        $extractedDate = $null
        $matches = [regex]::Matches($f, $dateRegex)
        foreach ($m in $matches) {
            $month = [int]$m.Groups[1].Value
            $day   = [int]$m.Groups[2].Value
            try {
                $extractedDate = Get-Date -Year $today.Year -Month $month -Day $day
                $extractedDateStr = $m.Value
                break
            } catch { }
        }

        # 获取 FTP 真实修改时间（使用净化后的文件名）
        $ftpDate = $null
        $furi = $baseUri + $safeFileName
        $dreq = [System.Net.FtpWebRequest]::Create($furi)
        $dreq.Method = [System.Net.WebRequestMethods+Ftp]::GetDateTimestamp
        $dreq.Credentials = $cred
        $dreq.UsePassive = $true
        try {
            $dres = $dreq.GetResponse()
            $ftpDate = $dres.LastModified
            $dres.Close()
        } catch { }

        # 判断是否为当天文件
        $isToday = $false
        if ($extractedDate -and $extractedDate.Date -eq $today) {
            $isToday = $true
        } elseif ($ftpDate -and $ftpDate.Date -eq $today) {
            $isToday = $true
        }

        $displayDateStr = if ($ftpDate) {
            $ftpDate.ToString('yyyy/MM/dd HH:mm:ss')
        } elseif ($extractedDate) {
            $extractedDate.ToString('yyyy/MM/dd HH:mm:ss')
        } else {
            '未知'
        }

        if ($extractedDateStr) {
            $displayLine = "$f -> 提取日期：$extractedDateStr；系统日期：$displayDateStr"
        } else {
            $displayLine = "$f -> $displayDateStr"
        }

        if ($isToday) {
            $todayFiles += [PSCustomObject]@{ 
                Name = $safeFileName
                SortKey = if ($ftpDate) { $ftpDate } elseif ($extractedDate) { $extractedDate } else { [DateTime]::MinValue }
                RemoteDir = $RemoteDir
            }
            Write-Host "  [当天] " -NoNewline -ForegroundColor DarkMagenta
            Write-Host $displayLine
        } else {
            Write-Host "  [忽略] " -NoNewline -ForegroundColor DarkYellow
            Write-Host $displayLine
        }
    }
    return $todayFiles
}

# --- 第二步：在默认路径查找当天 PPT ---
Write-Host "正在默认路径查找当天 PPT: $remotePath1" -ForegroundColor DarkBlue
$todayPptList = Get-TodayPptFromFtp -RemoteDir $remotePath1

# --- 第三步：若未找到且有备用路径，则在备用路径查找 ---
if ($todayPptList.Count -eq 0 -and $remotePath2) {
    Write-Host "默认路径未找到当天 PPT，尝试备用路径: $remotePath2" -ForegroundColor DarkBlue
    $todayPptList = Get-TodayPptFromFtp -RemoteDir $remotePath2
}

# --- 第四步：若无当天 PPT，直接退出 ---
if ($todayPptList.Count -eq 0) {
    Write-Host "[信息] 未找到任何当天日期的 PPT 文件。" -ForegroundColor DarkBlue
    Start-Sleep -Seconds 5
    exit 0
}

# --- 第五步：下载所有当天 PPT 到桌面 ---
Write-Host "`n开始下载..." -ForegroundColor DarkBlue
$timestamp = Get-Date -Format "yyyyMMdd"
$downloadedFiles = @()

foreach ($item in $todayPptList) {
    # 二次净化，防御性编程
    $cleanName = [System.IO.Path]::GetFileName($item.Name)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($cleanName)
    $ext      = [System.IO.Path]::GetExtension($cleanName)
    $localName = "${baseName}_${timestamp}${ext}"
    $localFile = Join-Path $desktop $localName

    $sourceUri = "ftp://$ftpServer$($item.RemoteDir)$cleanName"
    Write-Host "下载: " -NoNewline -ForegroundColor DarkBlue
    Write-Host "$($item.Name) -> $localName"
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Credentials = $cred
        $wc.DownloadFile($sourceUri, $localFile)
        $wc.Dispose()
        $downloadedFiles += $localFile
        Write-Host "  完成。" -ForegroundColor White
    } catch {
        Write-Host "  [错误] 下载失败: $_" -ForegroundColor Red
    }
}

# --- 第六步：打开下载的 PPT 并最小化其他窗口 ---
if ($downloadedFiles.Count -gt 0) {
    Write-Host "`n正在最小化其他窗口..." -ForegroundColor DarkBlue
    
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class WindowHelper {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
"@
    $currentWindow = [WindowHelper]::GetForegroundWindow()
    
    $shell = New-Object -ComObject Shell.Application
    [void]$shell.MinimizeAll()
    Start-Sleep -Milliseconds 500
    Write-Host "窗口已最小化。" -ForegroundColor White

    [void][WindowHelper]::ShowWindow($currentWindow, 9)
    
    Write-Host "正在打开所有当天 PPT..." -ForegroundColor White
    foreach ($file in $downloadedFiles) {
        Start-Process $file
        Start-Sleep -Milliseconds 500
    }
} else {
    Write-Host "[警告] 所有文件下载失败。" -ForegroundColor DarkRed
}

Write-Host "操作完成。" -ForegroundColor White
Start-Sleep -Seconds 5