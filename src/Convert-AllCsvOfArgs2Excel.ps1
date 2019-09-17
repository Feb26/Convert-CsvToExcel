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

###########
# function
###########
function convertCsvFile([string] $inputCsv) {
    $inputCsvFileName = [System.IO.Path]::GetFileName($inputCsv)
    $relativePath = Join-Path $settings.path.outputDir "$inputCsvFileName.xlsx"
    $outputXlsx = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($relativePath)
    
    ConvertCsvToExcel -inputCsv $inputCsv `
            -outputXlsx $outputXlsx `
            -delimiter $settings.inputFile.delimiter `
            -textQualifier $settings.inputFile.textQualifier `
            -font $settings.outputFile.font `
            -adjustColumnWidth $settings.outputFile.adjustColumnWidth
}

#######
# Main 
#######
$inpuFileList = $Args # バッチファイルにD&Dされたファイルのパスを入力ファイルとする

foreach ($inputFile in $inpuFileList) {
    if (Test-Path -Path $inputFile -PathType Leaf) {
        # ファイルの場合はそのまま変換
        convertCsvFile $inputFile
    } elseif (Test-Path -Path $inputFile -PathType Container) {
        # ディレクトリの場合はファイルを選択してから変換
        $inputCsvList = Get-ChildItem -Path $inputFile -File -Recurse -Filter "*.$($settings.inputFile.extension)" |
            Select-Object -Property Directory, Name, Length, LastWriteTime |
            Out-GridView -PassThru -Title "Excelに変換するファイルを選択してください" |
            ForEach-Object { Join-Path $_.Directory $_.Name }
        foreach ($inputCsv in $inputCsvList) {
            convertCsvFile $inputCsv
        }
    } else {
        Write-Warning "`"$inputFile`"が存在しません。スキップします。"
        continue
    }
}
