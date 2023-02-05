#
# 関数定義
#

# ExcelファイルをCSVに変換する。(LFは削除する)
function excelToCsv($xlsxFile) {
  $workFile = $xlsxFile + ".work.csv"
  $csvFile = $xlsxFile + ".new.csv"
  
  # ExcelアプリケーションのCOMオブジェクトでCSVに変換する
  $Excel = New-Object -ComObject Excel.Application
  $Workbook = $Excel.Workbooks.Open($xlsxFile)
  $Workbook.SaveAs($workFile, 6)
  $Excel.Quit()
  
  (Get-Content $workFile -raw).Replace("`r`n","__CRLF__").Replace("`n","").Replace("__CRLF__","`r`n") | Set-Content $csvFile
  Remove-Item -Force "*.work.csv"
  
}

# csvファイルを転置して、結果を追記する関数
function transposeCsv($csvFile,$appendFile) {
  # csvファイルの内容を取得
  $csv = @(Get-Content $csvFile )
  # 各行をカンマで分割
  for ($i = 0; $i -lt $csv.Count; $i++) {
    $csv[$i] = $csv[$i].Split(",")
  }

  # 結果を格納する配列を作成
  $array = New-Object System.Collections.ArrayList

  # csvの転置処理
  for ($i = 0; $i -lt $csv[0].Count; $i++) {
    $row = New-Object System.Collections.ArrayList
    for ($y = 0; $y -lt $csv.Count; $y++) {
      $row.Add($csv[$y][$i]) > $null
    }
    $array.Add($row) > $null
  }

  # 結果を出力ファイルに追記
  foreach($a in $array){
    $str = ""
    foreach($b in $a){
      $str += $b + ","
    }
    $str = $str.Substring(0,$str.Length-1)
    $str | Out-File -Append -Encoding default $appendFile
  }
}


#
# メイン処理
#

# 出力ファイルのパスを定義
$OUTPUT_FILE = "output.csv"
# 入力フォルダのパスを定義
$INPUT_FOLDER = ".\input_data"

# 出力ファイルを削除
Remove-Item -Force ($INPUT_FOLDER + "\*.csv")
Remove-Item -Force $OUTPUT_FILE

# 入力フォルダから取得したExcelファイルそれぞれを処理
$xlsxFiles = Get-ChildItem $INPUT_FOLDER -Filter '*.xlsx' -Recurse
foreach ($file in $xlsxFiles) {
  excelToCsv -xlsxFile $file.FullName
}

# 入力フォルダから取得したcsvファイルそれぞれを処理
$csvFiles = Get-ChildItem $INPUT_FOLDER -Filter '*.new.csv' -Recurse
foreach ($file in $csvFiles) {
  transposeCsv -csvFile $file.FullName -appendFile $OUTPUT_FILE
}