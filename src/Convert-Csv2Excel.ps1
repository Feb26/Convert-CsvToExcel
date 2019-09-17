Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ConvertCsvToExcel {
    Param(
        # 入力ファイルのパス
        [Parameter(mandatory)]
        [string] $inputCsv,

        # 出力ファイルのパス
        [Parameter(mandatory)]
        [string] $outputXlsx,

        # 入力ファイルのデリミタ（"," or "\t" or ";"）
        [string] $delimiter = ",",

        # 入力ファイルの文字列の括り
        [string] $textQualifier = "`"",

        # 出力ファイルのフォント
        [string] $font = "ＭＳ Ｐゴシック",

        # 出力ファイルの列幅の自動調整をするか
        [bool] $adjustColumnWidth = $true
    )

    Begin {
        $dllPath = Resolve-Path ".\lib\EPPlus.dll"
        [void][Reflection.Assembly]::LoadFrom($dllPath)
        function Get-EOL($inputCsv, [Text.Encoding] $encoding) {
            $byteContent = Get-Content -LiteralPath $inputCsv -Encoding Byte -TotalCount 2048
            $stringContent = $encoding.GetString($byteContent)
            if ($stringContent.IndexOf("`r`n") -ge 0) {
                return "`r`n"
            } else {
                return "`n"
            }
        }

        function Get-DateTime() {
            return Get-Date -Format "yyyy-mm-dd hh:MM:ss"
        }
    }

    Process {
        $inputCsvfileName = [IO.Path]::GetFileName($inputCsv)
        Write-Verbose ("[$(Get-DateTime)] Begin: $inputCsvfileName")

        # 文字エンコーディングの判別
        [Text.Encoding] $inputEncoding = . .\src\Resolve-Encoding -LiteralPath $inputCsv
        if ($inputEncoding -eq $null) {
            throw "$inputCsv の文字エンコーディングの判別に失敗しました。"
        }
        Write-Verbose ("[$(Get-DateTime)] Encoding: $($inputEncoding.EncodingName)")
        
        # 改行コードの判別
        $EOL = Get-EOL -inputCsv $inputCsv -encoding $inputEncoding
        $eolString = if ($EOL -eq "`r`n") { "CRLF" } else { "LF" }
        Write-Verbose ("[$(Get-DateTime)] EOL: $eolString")

        # 出力先が存在する場合は事前に削除する
        if (Test-Path $outputXlsx) {
            Remove-Item $outputXlsx
        }

        try {
            $excelPackage = [OfficeOpenXml.ExcelPackage]::new($outputXlsx)
            [OfficeOpenXml.ExcelWorkbook] $book = $excelPackage.Workbook
            [OfficeOpenXml.ExcelWorksheet] $sheet = $book.Worksheets.Add("Sheet1")
            
            # シートの設定
            try {
                $sheetName = [System.IO.Path]::GetFileName($inputCsv)
                $sheet.Name = $sheetName
            } catch {
                Write-Warning "`"${sheetName}`"をシート名に設定できませんでした。シート名に使用できない文字がファイル名に含まれているか、ファイル名が32文字以上の可能性があります。"
            }

            # ブックの設定
            $book.Styles.Fonts[0].Name = $font

            # CSVのインポート処理の設定
            $textFormat = [OfficeOpenXml.ExcelTextFormat]::new()
            $textFormat.DataTypes = ,[OfficeOpenXml.eDataTypes]::String * 1000
            $textFormat.Delimiter = $delimiter
            $textFormat.Encoding = $inputEncoding
            $textFormat.EOL = $EOL
            $textFormat.TextQualifier = $textQualifier
            $loadedRange = $sheet.Cells["A1"].LoadFromText((Get-Item $inputCsv), $textFormat)

            # 列幅調整
            if ($adjustColumnWidth) {
                $sheet.Cells.AutoFitColumns(3)
            }
            
            $excelPackage.Save()
            Write-Verbose ("[$(Get-DateTime)] Finish: $inputCsvfileName")
        } catch {
            Write-Error $_.Exception

            # re-throwして終了
            break
        } finally {
            if ($excelPackage -ne $null) {
                $excelPackage.Dispose()
            }
        }
    }

    End {
    }

}
