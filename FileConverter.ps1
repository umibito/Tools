[CmdletBinding()]
param()

# エラーを停止して、変換中の問題を GUI に表示できるようにする。
$ErrorActionPreference = 'Stop'

# Windows Forms で画面を作るためのアセンブリを読み込む。
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# CSV/TXT に出力する値を、必要に応じてダブルクォートで囲む。
function ConvertTo-QuotedField {
    param(
        [AllowNull()][object]$Value,
        [bool]$Quote
    )

    $text = if ($null -eq $Value) { '' } else { [string]$Value }

    if ($Quote) {
        # 値に含まれる " は "" に置き換える（CSV のエスケープ規則）。
        return '"' + $text.Replace('"', '""') + '"'
    }

    return $text
}

# 区切り文字付きテキストを読み込み、引用符や改行を考慮して行と列に分解する。
function ConvertFrom-DelimitedText {
    param(
        [string]$Text,
        [char]$Delimiter
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $row = [System.Collections.Generic.List[string]]::new()
    $field = [System.Text.StringBuilder]::new()
    $quoted = $false

    # 文字を1つずつ処理し、引用符内の区切り文字や改行を値として保持する。
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]

        if ($quoted) {
            if ($ch -eq '"') {
                if (($i + 1) -lt $Text.Length -and $Text[$i + 1] -eq '"') {
                    [void]$field.Append('"')
                    $i++
                }
                else {
                    $quoted = $false
                }
            }
            else {
                [void]$field.Append($ch)
            }
        }
        elseif ($ch -eq '"' -and $field.Length -eq 0) {
            $quoted = $true
        }
        elseif ($ch -eq $Delimiter) {
            $row.Add($field.ToString())
            [void]$field.Clear()
        }
        elseif ($ch -eq "`r" -or $ch -eq "`n") {
            # CRLF は1つの改行として扱い、完成した行を記録する。
            $row.Add($field.ToString())
            [void]$field.Clear()
            $rows.Add($row.ToArray())
            $row.Clear()

            if ($ch -eq "`r" -and ($i + 1) -lt $Text.Length -and $Text[$i + 1] -eq "`n") {
                $i++
            }
        }
        else {
            [void]$field.Append($ch)
        }
    }

    if ($quoted) {
        throw 'テキストのダブルクォートが閉じられていません。'
    }

    # 最後の行（末尾改行がない場合や、最後が区切り文字の場合）を追加する。
    if ($field.Length -gt 0 -or $row.Count -gt 0 -or
        ($Text.Length -gt 0 -and $Text[$Text.Length - 1] -eq $Delimiter)) {
        $row.Add($field.ToString())
        $rows.Add($row.ToArray())
    }

    return ,$rows.ToArray()
}

# GUI で選ばれた文字コード名から .NET のエンコーディング設定を取得する。
function Get-TextEncoding {
    param([string]$Name)

    switch ($Name) {
        'UTF-8 (BOMあり)' { return [System.Text.UTF8Encoding]::new($true) }
        'UTF-8 (BOMなし)' { return [System.Text.UTF8Encoding]::new($false) }
        'Shift-JIS' { return [System.Text.Encoding]::GetEncoding(932) }
        'UTF-16 LE' { return [System.Text.Encoding]::Unicode }
    }
}

# 選択された方向に応じて、Excel ブックとテキストファイルを相互変換する。
function Invoke-Conversion {
    param(
        $Mode,
        $Source,
        $Destination,
        $Quote,
        $EncodingName
    )

    $excel = $null
    $workbook = $null

    try {
        if ($Mode -eq 'Excel → テキスト') {
            # Excel の先頭シートの使用範囲を取得し、CSV またはタブ区切りで保存する。
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Open($Source, 0, $true)

            $sheet = $workbook.Worksheets.Item(1)
            $range = $sheet.UsedRange
            $values = $range.Value2
            $rows = [System.Collections.Generic.List[string]]::new()
            $rowCount = [int]$range.Rows.Count
            $colCount = [int]$range.Columns.Count
            $delimiter = if ([IO.Path]::GetExtension($Destination) -ieq '.csv') { ',' } else { "`t" }

            # セルの値を文字列化し、選択された引用符設定を適用する。
            for ($r = 1; $r -le $rowCount; $r++) {
                $fields = [System.Collections.Generic.List[string]]::new()

                for ($c = 1; $c -le $colCount; $c++) {
                    if ($rowCount -eq 1 -and $colCount -eq 1) {
                        $value = $values
                    }
                    else {
                        $value = $values[$r, $c]
                    }

                    $fields.Add((ConvertTo-QuotedField $value $Quote))
                }

                $rows.Add(($fields -join $delimiter))
            }

            # GUI で選択した文字コードでテキストファイルを書き出す。
            [IO.File]::WriteAllLines($Destination, $rows, (Get-TextEncoding $EncodingName))
        }
        else {
            # CSV/TXT を選択した文字コードで読み込み、拡張子から区切り文字を決める。
            $text = [IO.File]::ReadAllText($Source, (Get-TextEncoding $EncodingName))
            $delimiter = if ([IO.Path]::GetExtension($Source) -ieq '.csv') { [char]',' } else { [char]"`t" }
            $records = @(ConvertFrom-DelimitedText $text $delimiter)

            # Excel の新規ブックを作り、各セルを文字列形式にして値を設定する。
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Add()
            $sheet = $workbook.Worksheets.Item(1)

            for ($r = 0; $r -lt $records.Count; $r++) {
                for ($c = 0; $c -lt $records[$r].Count; $c++) {
                    $cell = $sheet.Cells.Item($r + 1, $c + 1)
                    $cell.NumberFormat = '@'
                    $cell.Value2 = [string]$records[$r][$c]
                }
            }

            # 出力拡張子に合わせて XLSX または XLS 形式で保存する。
            $format = if ([IO.Path]::GetExtension($Destination) -ieq '.xls') { 56 } else { 51 }
            $workbook.SaveAs($Destination, $format)
        }
    }
    finally {
        # 変換の成否にかかわらず Excel を閉じ、COM オブジェクトを解放する。
        if ($workbook) {
            $workbook.Close($false) | Out-Null
        }
        if ($excel) {
            $excel.Quit()
        }

        foreach ($comObject in @($range, $sheet, $workbook, $excel)) {
            if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
            }
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# メイン画面の基本設定を行う。
$form = [Windows.Forms.Form]::new()
$form.Text = 'ファイル変換ツール'
$form.Size = [Drawing.Size]::new(620, 390)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Font = [Drawing.Font]::new('Meiryo UI', 9)

# 指定位置とサイズでラベルを追加する共通処理。
function Add-Label($text, $x, $y, $w = 130) {
    $label = [Windows.Forms.Label]::new()
    $label.Text = $text
    $label.Location = [Drawing.Point]::new($x, $y)
    $label.Size = [Drawing.Size]::new($w, 25)
    $form.Controls.Add($label)
}

# 変換方向を選ぶドロップダウンを配置する。
Add-Label '変換の種類' 18 20
$mode = [Windows.Forms.ComboBox]::new()
$mode.DropDownStyle = 'DropDownList'
$mode.Items.AddRange(@('Excel → テキスト', 'テキスト → Excel'))
$mode.SelectedIndex = 0
$mode.Location = [Drawing.Point]::new(155, 18)
$mode.Size = [Drawing.Size]::new(420, 25)
$form.Controls.Add($mode)

# 変換元ファイルのパス入力欄と参照ボタンを配置する。
Add-Label '変換元ファイル' 18 62
$source = [Windows.Forms.TextBox]::new()
$source.Location = [Drawing.Point]::new(155, 60)
$source.Size = [Drawing.Size]::new(330, 25)
$form.Controls.Add($source)

$sourceButton = [Windows.Forms.Button]::new()
$sourceButton.Text = '参照...'
$sourceButton.Location = [Drawing.Point]::new(495, 58)
$sourceButton.Size = [Drawing.Size]::new(80, 28)
$form.Controls.Add($sourceButton)

# 出力先のパス入力欄と保存先選択ボタンを配置する。
Add-Label '出力先ファイル' 18 102
$destination = [Windows.Forms.TextBox]::new()
$destination.Location = [Drawing.Point]::new(155, 100)
$destination.Size = [Drawing.Size]::new(330, 25)
$form.Controls.Add($destination)

$destinationButton = [Windows.Forms.Button]::new()
$destinationButton.Text = '参照...'
$destinationButton.Location = [Drawing.Point]::new(495, 98)
$destinationButton.Size = [Drawing.Size]::new(80, 28)
$form.Controls.Add($destinationButton)

# テキスト出力時の引用符設定と文字コード選択欄を配置する。
Add-Label 'テキスト出力設定' 18 150
$quote = [Windows.Forms.CheckBox]::new()
$quote.Text = 'ヘッダ・値をダブルクォートで囲む'
$quote.Location = [Drawing.Point]::new(155, 148)
$quote.Size = [Drawing.Size]::new(350, 25)
$quote.Checked = $true
$form.Controls.Add($quote)

Add-Label 'エンコード' 18 188
$encoding = [Windows.Forms.ComboBox]::new()
$encoding.DropDownStyle = 'DropDownList'
$encoding.Items.AddRange(@('UTF-8 (BOMあり)', 'UTF-8 (BOMなし)', 'Shift-JIS', 'UTF-16 LE'))
$encoding.SelectedIndex = 0
$encoding.Location = [Drawing.Point]::new(155, 186)
$encoding.Size = [Drawing.Size]::new(250, 25)
$form.Controls.Add($encoding)

# 実行ボタンと閉じるボタンを配置する。
$run = [Windows.Forms.Button]::new()
$run.Text = '実行'
$run.Location = [Drawing.Point]::new(390, 285)
$run.Size = [Drawing.Size]::new(90, 35)
$form.Controls.Add($run)

$close = [Windows.Forms.Button]::new()
$close.Text = '閉じる'
$close.Location = [Drawing.Point]::new(485, 285)
$close.Size = [Drawing.Size]::new(90, 35)
$form.Controls.Add($close)

# 変換元の参照ボタンで形式に合うファイルを選び、出力先の初期値も設定する。
$sourceButton.Add_Click({
    $dialog = [Windows.Forms.OpenFileDialog]::new()
    $dialog.Filter = if ($mode.SelectedIndex -eq 0) {
        'Excel (*.xlsx;*.xls)|*.xlsx;*.xls'
    }
    else {
        'テキスト (*.csv;*.txt)|*.csv;*.txt'
    }

    if ($dialog.ShowDialog() -eq 'OK') {
        $source.Text = $dialog.FileName
        $ext = if ($mode.SelectedIndex -eq 0) { '.csv' } else { '.xlsx' }
        $destination.Text = [IO.Path]::ChangeExtension($dialog.FileName, $ext)
    }
})

# 出力先の参照ボタンで変換方向に応じた保存形式を選ぶ。
$destinationButton.Add_Click({
    $dialog = [Windows.Forms.SaveFileDialog]::new()
    $dialog.OverwritePrompt = $true
    $dialog.Filter = if ($mode.SelectedIndex -eq 0) {
        'CSV (*.csv)|*.csv|テキスト (*.txt)|*.txt'
    }
    else {
        'Excel (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls'
    }

    if ($dialog.ShowDialog() -eq 'OK') {
        $destination.Text = $dialog.FileName
    }
})

# テキスト出力以外では引用符の設定を使わないため、選択に応じて有効/無効を切り替える。
$mode.Add_SelectedIndexChanged({
    $quote.Enabled = ($mode.SelectedIndex -eq 0)
})

# 実行ボタンで入力値を検証してから変換を開始し、結果またはエラーを画面に表示する。
$run.Add_Click({
    try {
        if (-not (Test-Path -LiteralPath $source.Text -PathType Leaf)) {
            throw '変換元ファイルを選択してください。'
        }

        if ([string]::IsNullOrWhiteSpace($destination.Text)) {
            throw '出力先ファイルを選択してください。'
        }

        $expected = if ($mode.SelectedIndex -eq 0) { @('.csv', '.txt') } else { @('.xlsx', '.xls') }
        if ([IO.Path]::GetExtension($destination.Text).ToLowerInvariant() -notin $expected) {
            throw "出力先の拡張子は $($expected -join ', ') から選択してください。"
        }

        if ([IO.Path]::GetFullPath($source.Text) -eq [IO.Path]::GetFullPath($destination.Text)) {
            throw '変換元と出力先に同じファイルは指定できません。'
        }

        # 二重実行を防ぎ、変換中であることをカーソルで知らせる。
        $run.Enabled = $false
        $form.UseWaitCursor = $true
        $form.Refresh()

        Invoke-Conversion $mode.Text $source.Text $destination.Text $quote.Checked $encoding.Text
        [Windows.Forms.MessageBox]::Show('変換が完了しました。', '完了', 'OK', 'Information') | Out-Null
    }
    catch {
        [Windows.Forms.MessageBox]::Show($_.Exception.Message, '変換エラー', 'OK', 'Error') | Out-Null
    }
    finally {
        # 成否にかかわらず画面操作を再び有効にする。
        $run.Enabled = $true
        $form.UseWaitCursor = $false
    }
})

# 閉じるボタンを押したら画面を終了する。
$close.Add_Click({
    $form.Close()
})

# GUI を表示し、ユーザーの操作を待つ。
[void]$form.ShowDialog()
