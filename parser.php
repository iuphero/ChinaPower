<?php
#解析含有数据的Excel表格, 保存为json文件

require('reader/php-excel-reader/excel_reader2.php');
$data = new Spreadsheet_Excel_Reader("ChinaData.xls");

$result = [];
$yearsMapper = [
  2 => 1995,
  3 => 2000,
  4 => 2005,
  5 => 2010,
  6 => 2012,
  7 => 2013
];

for ($i=6; $i < 42; $i++) {
  for ($j=2; $j < 8; $j++) {
    $name = trim($data->val($i, 1));
    if (!empty($name)) {
      $name = str_replace(" ", "", $name); #去除文字中的空格
      $year = $yearsMapper[$j];
      $result[$year][$name] = $data->val($i, $j);
    }
  }
}

$json = json_encode($result, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);

$fp = fopen('ChinaData.json', "w");
fwrite($fp, $json);
fclose($fp);

echo '解析并保存成功';
