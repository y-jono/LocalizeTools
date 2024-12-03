# 実行の開始時刻を記録
$startTime = Get-Date

# ベースパス（必要に応じて変更してください）
$basePath = "C:\path\to\LocalizedHtmlGenerator"

# DLLのパス
$dllPath = Join-Path $basePath "LocalizedHtmlGenerator.dll"
$closedXmlPath = Join-Path $basePath "ClosedXML.dll"
$openXmlPath = Join-Path $basePath "DocumentFormat.OpenXml.dll"

# DLLをロード
Add-Type -Path $openXmlPath
Add-Type -Path $closedXmlPath
Add-Type -Path $dllPath

# パス設定
$templatePath = Join-Path $basePath "template.html"
$excelPath = Join-Path $basePath "resources.xlsm"
$outputBaseName = "OutputBase"

# 出力先ディレクトリを指定（ユーザーが設定可能）
$outputDirectory = Join-Path $basePath "OutputHtml"  # 例として "OutputHtml" フォルダに出力

# 出力先ディレクトリが存在しない場合は作成
if (-Not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

# ファイルの存在を確認
if (-Not (Test-Path $excelPath)) {
    Write-Error "The Excel file '$excelPath' could not be found. Please check the file path and try again."
    exit
}

# 各言語ごとにHTMLを生成
$cultures = @("English", "Japanese", "French") # CSVファイルのヘッダーに合わせる
foreach ($culture in $cultures) {
    try {
        $htmlGenerator = New-Object LocalizedHtmlGenerator
        $htmlPath = $htmlGenerator.GenerateHtml($outputBaseName, $culture, $templatePath, $excelPath, $outputDirectory)
        Write-Output "HTML for $culture has been generated and saved to $htmlPath"
    } catch {
        Write-Error ("Error generating HTML for culture {0}: {1}" -f $culture, $_.Exception.Message)
    }
}

# 実行の終了時刻を記録
$endTime = Get-Date

# 実行時間の計算と出力
$totalTime = $endTime - $startTime
Write-Output "Total execution time: $($totalTime.TotalSeconds) seconds"
