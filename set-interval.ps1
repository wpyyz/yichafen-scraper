param()
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $here 'interval.txt'

Write-Host ''
Write-Host '============================================'
Write-Host '  设置查询间隔'
Write-Host '============================================'
Write-Host ''
Write-Host '易查分有反爬限制，查询太密会被临时封禁。'
Write-Host ''
Write-Host '   6  秒  快，但接近触发点，不推荐'
Write-Host '  12  秒  默认，一个班 70 人约 14 分钟'
Write-Host '  20  秒  稳，约 25 分钟'
Write-Host '  30  秒  最保险，约 37 分钟'
Write-Host ''
Write-Host ("当前设置：" + $(if (Test-Path $file) { (Get-Content $file -Raw).Trim() } else { '12' }) + ' 秒')
Write-Host ''

$v = Read-Host '请输入新的间隔秒数（直接回车 = 12）'
if ([string]::IsNullOrWhiteSpace($v)) { $v = '12' }
$v = $v.Trim()

if ($v -notmatch '^\d+$') {
    Write-Host ''
    Write-Host '输入的不是纯数字，用默认 12 秒。' -ForegroundColor Yellow
    $v = '12'
}

$n = [int]$v
if ($n -lt 3) {
    Write-Host '小于 3 秒太激进，已按 3 秒处理。' -ForegroundColor Yellow
    $n = 3
}

Set-Content -Path $file -Value $n -Encoding ASCII
Write-Host ''
Write-Host ("已保存：" + $n + " 秒") -ForegroundColor Green
Write-Host '下次双击「一键查分.bat」就按这个间隔跑。'
Write-Host ''
