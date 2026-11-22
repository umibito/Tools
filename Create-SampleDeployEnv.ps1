# =======================================================
# Deploy Tool Test Environment
# =======================================================
<#
C:\DeployTest
│
├─ Package
│   ├─ Scripts
│   │   ├─ SampleJob1.ps1
│   │   └─ SampleJob2.ps1
│   │
│   ├─ Config
│   │   ├─ appsettings.json
│   │   └─ setting.xml
│   │
│   └─ Readme
│       └─ version.txt
│
├─ Package.zip
│
├─ Dest
│   ├─ Scripts
│   │   └─ SampleJob1.ps1   (旧版)
│   │
│   └─ Config
│       └─ appsettings.json (旧版)
│
├─ Backup
│
└─ Temp
#>

$Root = "C:\Users\nzmfi\OneDrive\Myライブラリ\PowerShell\tmp"

Write-Host "Create Test Environment..."

# 初期化
if (Test-Path $Root) {
    Remove-Item $Root -Recurse -Force
}


# フォルダ作成

try {
    @(
        "$Root\Package\Scripts",
        "$Root\Package\Config",
        "$Root\Package\Readme",
        "$Root\Dest\Scripts",
        "$Root\Dest\Config",
        "$Root\Backup",
        "$Root\Temp"
    ) | ForEach-Object {

        New-Item `
            -Path $_ `
            -ItemType Directory `
            -Force | Out-Null
    }
}
catch {
    Write-Host "Error: $_"
    exit 1
}


# -------------------------------------------------------
# ZIP用ファイル
# -------------------------------------------------------

@'
Write-Host "SampleJob1 Start"
Get-Date
Write-Host "SampleJob1 End"
'@ | Set-Content `
    "$Root\Package\Scripts\SampleJob1.ps1" `
    -Encoding UTF8

@'
Write-Host "Backup Process"
'@ | Set-Content `
    "$Root\Package\Scripts\SampleJob2.ps1" `
    -Encoding UTF8

@'
{
  "ApplicationName": "SampleApp",
  "Version": "1.0.0",
  "LogLevel": "Information"
}
'@ | Set-Content `
    "$Root\Package\Config\appsettings.json" `
    -Encoding UTF8

@'
<?xml version="1.0" encoding="utf-8"?>
<Settings>
  <ServerName>TEST-SRV01</ServerName>
  <Port>8080</Port>
</Settings>
'@ | Set-Content `
    "$Root\Package\Config\setting.xml" `
    -Encoding UTF8

@'
Release Version : 1.0.0
Deploy Date     : 2026-11-18
'@ | Set-Content `
    "$Root\Package\Readme\version.txt" `
    -Encoding UTF8


# -------------------------------------------------------
# 展開先の旧版ファイル
# -------------------------------------------------------

@'
Write-Host "OLD VERSION"
'@ | Set-Content `
    "$Root\Dest\Scripts\SampleJob1.ps1" `
    -Encoding UTF8

@'
{
  "ApplicationName": "OldApp",
  "Version": "0.9.0"
}
'@ | Set-Content `
    "$Root\Dest\Config\appsettings.json" `
    -Encoding UTF8

# -------------------------------------------------------
# ZIP作成
# -------------------------------------------------------

try {
    $ZipFile = "$Root\Package.zip"

    if (Test-Path $ZipFile) {
        Remove-Item $ZipFile -Force
    }

    Compress-Archive `
        -Path "$Root\Package\*" `
        -DestinationPath $ZipFile `
        -Force

}
catch {
    Write-Host "Error: $_"
    exit 1
}

# -------------------------------------------------------
# マッピング設定
# -------------------------------------------------------

@'
{
    "Scripts": "C:\Users\nzmfi\OneDrive\Myライブラリ\PowerShell\tmp\DeployTest\\Dest\\Scripts",
    "Config":  "C:\Users\nzmfi\OneDrive\Myライブラリ\PowerShell\tmp\DeployTest\\Dest\\Config"
}
'@ | Set-Content `
    "$Root\Mapping.json" `
    -Encoding UTF8

Write-Host ""
Write-Host "Environment Created"
Write-Host "Root : $Root"
Write-Host "ZIP  : $ZipFile"
