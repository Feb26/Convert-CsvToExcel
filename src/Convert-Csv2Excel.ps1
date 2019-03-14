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

        # デリミタ（"," or "\t" or ";"）
        [string] $delimiter = ",",

        # 列幅の自動調整をするか
        [bool] $adjustColumnWidth = $true
    )

    Begin {
        $excel = New-Object -ComObject excel.application
        $excel.Visible = $false
        $workbooks = $excel.Workbooks

        function Clear-Resource() {
            # Excelのゾンビプロセスが発生しないように、WorkbooksオブジェクトとExcelオブジェクトを適切に破棄する
            # 多くの場合、WorkbooksオブジェクトとExcelオブジェクトだけを処理すればプロセスは正常に解放される模様
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbooks)
            $workbooks = $null

            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
            $excel = $null

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }

    Process {
        $inputEncoding = . .\src\Resolve-Encoding -LiteralPath $inputCsv
        if ($inputEncoding -eq $null) {
            throw "$inputCsv のエンコーディングの判定に失敗しました。"
        }
        try {
            try {
                # Workbookの新規作成(Worksheetも同時に作成される)
                $workbook = $workbooks.Add(1)
                try {
                    $worksheet = $workbook.Worksheets.Item(1)
                    try {
                        $worksheet.Name = [System.IO.Path]::GetFileName($inputCsv)
                    } catch {
                        Write-Warning "`"${inputCsv}`"のファイル名をシート名に設定できませんでした。シート名に使用できない文字がファイル名に含まれているか、ファイル名が32文字以上の可能性があります。"
                    }

                    try {
                        # Excelで データ » テキストファイル を選択した際のテキストファイルのインポートウィザードを実行する
                        $absInputCsvPath = Convert-Path -Path $inputCsv
                        $txtConnector = ("TEXT;$absInputCsvPath") # 絶対パスでの指定が必須
                        $r = $worksheet.Range("A1")
                        try {
                            $connector = $worksheet.QueryTables.Add($txtConnector, $r)
                            try {
                                $query = $worksheet.QueryTables.item($connector.name)

                                # クエリの設定
                                $query.TextFileOtherDelimiter = $delimiter
                                $query.TextFilePlatform = $inputEncoding.CodePage
                                $query.TextFileParseType  = [Microsoft.Office.Interop.Excel.XlTextParsingType]::xlDelimited
                                $query.TextFileColumnDataTypes = ,[Microsoft.Office.Interop.Excel.XlColumnDataType]::xlTextFormat * $worksheet.Cells.Columns.Count # すべてテキストとしてインポート
                                $query.AdjustColumnWidth = $adjustColumnWidth

                                # クエリの実行
                                [void] $query.Refresh()
                                $query.Delete()
                            } finally {
                                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($query)
                                $query = $null
                            }
                        } finally {
                            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($r)
                            $r = $null
                        }
                    } finally {
                        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($connector)
                        $connector = $null
                    }
                } finally {
                    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet)
                    $worksheet = $null
                }
            } finally {
            if ($workbook -ne $null) {
                # ワークブックの保存
                $workbook.SaveAs($outputXlsx, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
                $workbook.Close($false)
            }
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
            $workbook = $null
            }
        } catch {
            Write-Error $_.Exception
            Clear-Resource

            # re-throwして終了
            break
        }
    }

    End {
        Clear-Resource
    }

}
