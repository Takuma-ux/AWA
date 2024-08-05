# PowerShellでExcelを操作するためのスクリプト

function Get-CellBackgroundColor {
    param (
        [string]$ExcelFilePath,
        [string]$SheetName,
        [string]$CellAddress
    )
    
    # Excelアプリケーションの起動
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    # Excelファイルを開く
    $workbook = $excel.Workbooks.Open($ExcelFilePath)
    $worksheet = $workbook.Sheets.Item($SheetName)
    
    # セルの背景色を取得
    $cell = $worksheet.Range($CellAddress)
    $color = $cell.Interior.Color

    # 色をRGB形式で取得
    $r = $color -band 0xFF
    $g = ($color -shr 8) -band 0xFF
    $b = ($color -shr 16) -band 0xFF

    # 結果を出力
    Write-Output "Cell: $CellAddress"
    Write-Output "Red: $r"
    Write-Output "Green: $g"
    Write-Output "Blue: $b"

    # Excelファイルを閉じる
    $workbook.Close($false)
    $excel.Quit()

    # COMオブジェクトの解放
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($cell) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# スクリプト実行部
$ExcelFilePath = "C:\\Users\\takum\\Documents\\cheerjob_auto\\test_0802.docx"  # Excelファイルのパス
$SheetName = "Sheet1"  # シート名
$CellAddress = "A1"    # セルのアドレス

Get-CellBackgroundColor -ExcelFilePath $ExcelFilePath -SheetName $SheetName -CellAddress $CellAddress
