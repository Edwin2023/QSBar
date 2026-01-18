# QSBar 一键发布脚本 (Gitee 版)
# 用法: .\Publish.ps1 -Version "1.0.0.1" -Log "修复了图标显示并优化了更新流程"

param (
    [string]$Version,
    [string]$Log
)

$RootDir = Get-Location
$DllSource = "$RootDir\QSBar\bin\Release\QSBar.dll"
$PublishDir = "$RootDir\publish"
$VersionJson = "$RootDir\version.json"

# 1. 检查参数
if (-not $Version -or -not $Log) {
    Write-Host "错误: 请提供版本号和更新日志。" -ForegroundColor Red
    Write-Host "用法示例: .\Publish.ps1 -Version '1.0.0.1' -Log '更新说明内容'" -ForegroundColor Yellow
    exit
}

Write-Host "--- 开始发布流程 v$Version ---" -ForegroundColor Cyan

# 2. 编译 Release 版本 (可选，如果已手动编译可注释)
# Write-Host "正在编译 Release 版本..."
# & "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" /p:Configuration=Release /t:Rebuild

# 3. 检查编译结果
if (-not (Test-Path $DllSource)) {
    Write-Host "错误: 找不到编译后的 DLL: $DllSource" -ForegroundColor Red
    exit
}

# 4. 更新 publish 文件夹
if (-not (Test-Path $PublishDir)) { New-Item -ItemType Directory -Path $PublishDir }
Copy-Item $DllSource -Destination "$PublishDir\QSBar.dll" -Force
Write-Host "已复制最新 DLL 到发布目录。" -ForegroundColor Green

# 5. 修改 version.json
$json = Get-Content $VersionJson | ConvertFrom-Json
$json.version = $Version
$json.changeLog = $Log
$json | ConvertTo-Json | Set-Content $VersionJson -Encoding UTF8
Write-Host "已更新 version.json 版本号和日志。" -ForegroundColor Green

# 6. 提交并推送到 Gitee
Write-Host "正在推送至 Gitee..." -ForegroundColor Cyan
git add .
git commit -m "Release v$Version : $Log"
git push origin master

Write-Host "--- 发布成功！ ---" -ForegroundColor Green
Write-Host "用户现在将收到 v$Version 的更新提示。" -ForegroundColor Cyan
