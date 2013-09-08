<?php

ini_set('memory_limit', '1000M');

require_once 'Classes/PHPExcel/IOFactory.php';

$excel = PHPExcel_IOFactory::load("files/2lVetPop11_POS_National.xlsx");
$writer = PHPExcel_IOFactory::createWriter($excel, 'CSV');
$writer->setDelimiter(",");
$writer->setEnclosure("");
$writer->setLineEnding("\r\n");
$writer->setSheetIndex(0);
$writer->save("files/2lVetPop11_POS_National.csv");
echo "done!";

?>
