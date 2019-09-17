Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"
$VerbosePreference = "Continue"
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
$inpuCsvList = Get-ChildItem -Path $settings.path.inputDir -File -Recurse -Filter "*.$($settings.inputFile.extension)" |
    Select-Object -Property Directory, Name, Length, LastWriteTime |
    Out-GridView -PassThru  -Title "Excelに変換するファイルを選択してください" |
    ForEach-Object { Join-Path $_.Directory $_.Name }

foreach ($inputCsv in $inpuCsvList) {
    # 入力ファイルのパスを出力ファイルのパスに変換する
    # e.g. ".\input\dir\file.csv" -> ".\output\dir\file.csv.xlsx"
    $inputCsvDir = Split-Path -Path $inputCsv -Parent
    Push-Location $settings.path.inputDir
        $relativeInputCsv =  Resolve-Path $inputCsv -Relative
    Pop-Location
    Push-Location $settings.path.outputDir
        $relativeOutputCsv = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($relativeInputCsv)
        $outputXlsx = "$relativeOutputCsv.xlsx"
    Pop-Location
    $xlsxParentDir = Split-Path $outputXlsx -Parent
    if (-not (Test-Path $xlsxParentDir)) {
        New-Item $xlsxParentDir -ItemType Directory | Out-Null
    }

    ConvertCsvToExcel -inputCsv $inputCsv `
            -outputXlsx $outputXlsx `
            -delimiter $settings.inputFile.delimiter `
            -textQualifier $settings.inputFile.textQualifier `
            -font $settings.outputFile.font `
            -adjustColumnWidth $settings.outputFile.adjustColumnWidth
}
