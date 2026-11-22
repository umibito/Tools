$ZipFile = "C:\Users\nzmfi\OneDrive\work\PowerShell\tmpPackage.zip"
$TempDir = "C:\Users\nzmfi\OneDrive\work\PowerShell\Temp"
$MapFile = "C:\Users\nzmfi\OneDrive\work\PowerShell\Mapping.json"
$LogDirectory = "C:\Users\nzmfi\OneDrive\work\PowerShell\Logs"
$LogFile = Join-Path $LogDirectory ("DeployTool_{0}.log" -f (Get-Date -Format 'yyyyMMdd'))
$RunTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$BackupDirectory = Join-Path (Split-Path -Parent $ZipFile) (Join-Path 'Backup' $RunTimestamp)
$DiffDirectory = Join-Path (Split-Path -Parent $ZipFile) 'Diff'

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $LogLine = '[{0}] [{1}] {2}' -f $Timestamp, $Level, $Message
    $Color = switch ($Level) {
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Gray' }
    }

    Write-Host $LogLine -ForegroundColor $Color

    if (-not (Test-Path -LiteralPath $LogDirectory -PathType Container)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    [System.IO.File]::AppendAllText(
        $LogFile,
        $LogLine + [Environment]::NewLine,
        [System.Text.UTF8Encoding]::new($false)
    )
}

$ErrorActionPreference = 'Stop'

function Write-DiffReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$DestinationRoot,

        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$ReportPath
    )

    $ReportLines = [System.Collections.Generic.List[string]]::new()
    $ReportLines.Add("Deploy差分レポート: $Category")
    $ReportLines.Add("作成日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $ReportLines.Add("パッケージ: $SourceRoot")
    $ReportLines.Add("コピー先: $DestinationRoot")
    $ReportLines.Add('')

    $SourceFiles = @(Get-ChildItem -LiteralPath $SourceRoot -File -Recurse -ErrorAction Stop)
    $SourceRelativePaths = @(
        $SourceFiles | ForEach-Object {
            $_.FullName.Substring($SourceRoot.TrimEnd('\').Length).TrimStart('\')
        }
    )
    if ($SourceFiles.Count -eq 0) {
        $ReportLines.Add('対象ファイルはありません。')
    }

    foreach ($SourceFile in $SourceFiles) {
        $RelativePath = $SourceFile.FullName.Substring($SourceRoot.TrimEnd('\').Length).TrimStart('\')
        $DestinationFile = Join-Path $DestinationRoot $RelativePath
        $ReportLines.Add("=== $RelativePath ===")

        if (-not (Test-Path -LiteralPath $DestinationFile -PathType Leaf)) {
            $ReportLines.Add('[追加] コピー先に同名ファイルがありません。')
            $ReportLines.Add('')
            continue
        }

        $IsTextFile = $SourceFile.Extension -match '^\.(ps1|psm1|psd1|json|xml|txt|ini|config|cfg|ya?ml|properties|bat|cmd|sh)$'
        if ($IsTextFile) {
            $Before = @(Get-Content -LiteralPath $DestinationFile -ErrorAction Stop)
            $After = @(Get-Content -LiteralPath $SourceFile.FullName -ErrorAction Stop)
            $Changes = @(Compare-Object -ReferenceObject $Before -DifferenceObject $After)

            if ($Changes.Count -eq 0) {
                $ReportLines.Add('[変更なし]')
            }
            else {
                $ReportLines.Add('[変更あり] 差分行（< コピー先のみ、> パッケージのみ）:')
                foreach ($Change in $Changes) {
                    $ReportLines.Add(("{0} {1}" -f $Change.SideIndicator, $Change.InputObject))
                }
            }
        }
        else {
            $BeforeHash = (Get-FileHash -LiteralPath $DestinationFile -Algorithm SHA256).Hash
            $AfterHash = (Get-FileHash -LiteralPath $SourceFile.FullName -Algorithm SHA256).Hash
            if ($BeforeHash -eq $AfterHash) {
                $ReportLines.Add('[変更なし]')
            }
            else {
                $ReportLines.Add('[変更あり] バイナリファイルのためSHA256で比較しました。')
                $ReportLines.Add("< コピー先: $BeforeHash")
                $ReportLines.Add("> パッケージ: $AfterHash")
            }
        }

        $ReportLines.Add('')
    }

    if (Test-Path -LiteralPath $DestinationRoot -PathType Container) {
        $DestinationFiles = @(Get-ChildItem -LiteralPath $DestinationRoot -File -Recurse -ErrorAction Stop)
        foreach ($DestinationFile in $DestinationFiles) {
            $RelativePath = $DestinationFile.FullName.Substring($DestinationRoot.TrimEnd('\').Length).TrimStart('\')
            if ($SourceRelativePaths -notcontains $RelativePath) {
                $ReportLines.Add("=== $RelativePath ===")
                $ReportLines.Add('[コピー先のみ] パッケージに同じ相対パスのファイルがありません。')
                $ReportLines.Add('')
            }
        }
    }

    $ReportParent = Split-Path -Parent $ReportPath
    if (-not (Test-Path -LiteralPath $ReportParent -PathType Container)) {
        New-Item -Path $ReportParent -ItemType Directory -Force | Out-Null
    }
    [System.IO.File]::AppendAllLines($ReportPath, $ReportLines, [System.Text.UTF8Encoding]::new($false))
    Write-Log -Message "差分レポートを出力しました: $ReportPath"
}

# 以下main処理
function Main {
    .{
        Write-Log -Message "Main処理を開始します。"

        try {
            Write-Log -Message "Deploy started. Package: $ZipFile"
        
            # 一時フォルダ作成
            if (Test-Path -LiteralPath $TempDir) {
                Write-Log -Message "既存の一時フォルダを削除します: $TempDir"
                Remove-Item -LiteralPath $TempDir -Recurse -Force
            }

            New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
            Write-Log -Message "一時フォルダを作成しました: $TempDir"

            # 解凍
            Write-Log -Message "パッケージを解凍します: $ZipFile"
            Expand-Archive `
                -Path $ZipFile `
                -DestinationPath $TempDir `
                -Force

            # マッピング読込
            Write-Log -Message "マッピング設定を読み込みます: $MapFile"
            $Mappings = Get-Content -LiteralPath $MapFile -Raw | ConvertFrom-Json

            # コピー前の差分レポート出力（スクリプトと設定ファイルを分ける）
            $ScriptReport = Join-Path $DiffDirectory "DeployDiff_Scripts_$RunTimestamp.txt"
            $ConfigReport = Join-Path $DiffDirectory "DeployDiff_Config_$RunTimestamp.txt"
            $ScriptRoots = @($Mappings.PSObject.Properties | Where-Object { $_.Name -match 'script' })
            $ConfigRoots = @($Mappings.PSObject.Properties | Where-Object { $_.Name -match 'config' })

            if ($ScriptRoots.Count -gt 0) {
                foreach ($Mapping in $ScriptRoots) {
                    Write-DiffReport `
                        -SourceRoot (Join-Path $TempDir $Mapping.Name) `
                        -DestinationRoot $Mapping.Value `
                        -Category 'スクリプト' `
                        -ReportPath $ScriptReport
                }
            }
            else {
                Write-Log -Message 'スクリプト用マッピングがないため、スクリプト差分レポートを出力しません。' -Level WARN
            }

            if ($ConfigRoots.Count -gt 0) {
                foreach ($Mapping in $ConfigRoots) {
                    Write-DiffReport `
                        -SourceRoot (Join-Path $TempDir $Mapping.Name) `
                        -DestinationRoot $Mapping.Value `
                        -Category '設定ファイル' `
                        -ReportPath $ConfigReport
                }
            }
            else {
                Write-Log -Message '設定ファイル用マッピングがないため、設定ファイル差分レポートを出力しません。' -Level WARN
            }

            # コピー
            $Mappings.PSObject.Properties | ForEach-Object {
                $Source = Join-Path $TempDir $_.Name
                $Destination = $_.Value

                Write-Log -Message "コピーします: $Source -> $Destination"

                # 上書き対象の既存ファイルを、マッピング名と相対パスを保って退避する。
                $SourceFiles = @(Get-ChildItem -LiteralPath $Source -File -Recurse -ErrorAction Stop)
                foreach ($SourceFile in $SourceFiles) {
                    $RelativePath = $SourceFile.FullName.Substring($Source.TrimEnd('\').Length).TrimStart('\')
                    $DestinationFile = Join-Path $Destination $RelativePath
                    if (Test-Path -LiteralPath $DestinationFile -PathType Leaf) {
                        $BackupPath = Join-Path (Join-Path $BackupDirectory $_.Name) $RelativePath
                        $BackupParent = Split-Path -Parent $BackupPath
                        if (-not (Test-Path -LiteralPath $BackupParent -PathType Container)) {
                            New-Item -Path $BackupParent -ItemType Directory -Force | Out-Null
                        }
                        Copy-Item -LiteralPath $DestinationFile -Destination $BackupPath -Force
                        Write-Log -Message "コピー先をバックアップしました: $DestinationFile -> $BackupPath"
                    }
                }

                Copy-Item `
                    -Path "$Source\*" `
                    -Destination $Destination `
                    -Recurse `
                    -Force
            }

            Write-Log -Message 'Deploy completed successfully.'
        }
        catch {
            Write-Log -Message ("Deploy failed: {0}" -f $_.Exception.Message) -Level ERROR
            throw
        }
        finally {
            # 後片付け
            if (Test-Path -LiteralPath $TempDir) {
                try {
                    Remove-Item -LiteralPath $TempDir -Recurse -Force
                    Write-Log -Message "一時フォルダを削除しました: $TempDir"
                }
                catch {
                    Write-Log -Message ("一時フォルダの削除に失敗しました: {0}" -f $_.Exception.Message) -Level WARN
                }
            }
        }

    }
}
