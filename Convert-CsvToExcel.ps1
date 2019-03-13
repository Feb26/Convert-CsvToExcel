Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

enum TextFilePlatform {
    Shift_JIS = 932
    UTF_8 = 65001
}

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

        # エンコーディング
        [ValidateSet("Shift-JIS" , "UTF-8")]
        [string] $encoding = "Shift-JIS",

        # 列幅の自動調整をするか
        [bool] $adjustColumnWidth = $true
    )

    Begin {
        $excel = New-Object -ComObject excel.application
        $excel.Visible = $false
        $workbooks = $excel.Workbooks

        # ウィンドウハンドルからプロセスIDを取得するWin32APIを使用
        # 同一セッションで2回以上実行する（ISEから実行した場合など）とエラーとなるため事前にチェックする
        if ( -not ("Win32API" -as [type])) {
            Add-Type '
                using System;
                using System.Runtime.InteropServices;
                public class Win32API {
                    [DllImport("user32.dll")]
                    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
                }
            '
        }

        function Clear-Resource() {
            # Excelを終了する前にプロセスIDを取得しておく
            $excelHwnd = $excel.Hwnd
            $processId = 0
            [void][Win32API]::GetWindowThreadProcessId($excelHwnd, [ref]$processId)

            # Excelのゾンビプロセスが発生しないように、WorkbooksオブジェクトとExcelオブジェクトを適切に破棄する
            # 多くの場合、WorkbooksオブジェクトとExcelオブジェクトだけを処理すればプロセスは正常に解放される模様
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbooks)
            $workbooks = $null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
            $excel = $null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            # どうしても生き残ってしまった場合はプロセスをKillする
            Get-Process -Name "EXCEL" |
                    Where-Object { $_.Id -eq $processId } |
                    ForEach-Object { $_.Kill() }
        }
    }

    Process {
        try {
            # Workbookの新規作成(Worksheetも同時に作成される)
            $workbook = $workbooks.Add(1)
            $worksheet = $workbook.Worksheets.Item(1)
            try {
                $worksheet.Name = [System.IO.Path]::GetFileName($inputCsv)
            } catch {
                Write-Warning "`"${inputCsv}`"のファイル名をシート名に設定できませんでした。シート名に使用できない文字がファイル名に含まれているか、ファイル名が32文字以上の可能性があります。"
            }

            # Excelで データ » テキストファイル を選択した際のテキストファイルのインポートウィザードを実行する
            $absInputCsvPath = Convert-Path -Path $inputCsv
            $txtConnector = ("TEXT;$absInputCsvPath") # 絶対パスでの指定が必須
            $connector = $worksheet.QueryTables.Add($txtConnector, $worksheet.Range("A1"))
            $query = $worksheet.QueryTables.item($connector.name)

            # クエリの設定
            $query.TextFileOtherDelimiter = $delimiter
            $textFilePlatform = if ($encoding -eq "Shift-JIS") {
                [TextFilePlatform]::Shift_JIS
            } elseif ($encoding -eq "UTF-8") {
                [TextFilePlatform]::UTF_8
            }
            $query.TextFilePlatform = $textFilePlatform
            $query.TextFileParseType  = [Microsoft.Office.Interop.Excel.XlTextParsingType]::xlDelimited
            $query.TextFileColumnDataTypes = ,[Microsoft.Office.Interop.Excel.XlColumnDataType]::xlTextFormat * $worksheet.Cells.Columns.Count # すべてテキストとしてインポート
            $query.AdjustColumnWidth = $adjustColumnWidth

            # クエリの実行
            [void] $query.Refresh()
            $query.Delete()

            # ワークブックの保存
            $workbook.SaveAs($outputXlsx, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
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
