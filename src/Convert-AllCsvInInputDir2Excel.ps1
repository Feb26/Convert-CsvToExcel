Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path $PSScriptRoot -parent
Set-Location $projectRoot

#########
# import
#########
. .\src\Convert-Csv2Excel.ps1

###########
# Settings
###########
$settings = Get-Content -Path ".\settings.json" -Raw -Encoding Default | ConvertFrom-Json

#######
# Main 
#######
$inpuCsvList = Get-ChildItem -Path $settings.inputDir -File -Recurse -Filter "*.$($settings.extension)" |
    Select-Object -Property Directory, Name, Length, LastWriteTime |
    Out-GridView -PassThru  -Title "Excelに変換するファイルを選択してください" |
    ForEach-Object { Join-Path $_.Directory $_.Name }

foreach ($inputCsv in $inpuCsvList) {
    # 入力ファイルのパスを出力ファイルのパスに変換する
    # e.g. ".\input\dir\file.csv" -> ".\output\dir\file.csv.xlsx"
    $inputCsvDir = Split-Path -Path $inputCsv -Parent
    Push-Location $settings.inputDir
        $relativeInputCsv =  Resolve-Path $inputCsv -Relative
    Pop-Location
    Push-Location $settings.outputDir
        $relativeOutputCsv = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($relativeInputCsv)
        $outputXlsx = "$relativeOutputCsv.xlsx"
    Pop-Location
    $xlsxParentDir = Split-Path $outputXlsx -Parent
    if (-not (Test-Path $xlsxParentDir)) {
        New-Item $xlsxParentDir -ItemType Directory | Out-Null
    }

    ConvertCsvToExcel -inputCsv $inputCsv -outputXlsx $outputXlsx -delimiter $settings.delimiter -adjustColumnWidth $settings.adjustColumnWidth
}
