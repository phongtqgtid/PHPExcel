<?php
require_once 'Classes/PHPExcel.php';

//$filename = 'example.xlsx';
//$inputFileType = PHPExcel_IOFactory::identify($filename);
//$objReader = PHPExcel_IOFactory::createReader($inputFileType);
//
//$objReader->setReadDataOnly(true);
//
///**  Load $inputFileName to a PHPExcel Object  **/
//$objPHPExcel = $objReader->load("$filename");
//
//$total_sheets=$objPHPExcel->getSheetCount();
//
//$allSheetName=$objPHPExcel->getSheetNames();
//$objWorksheet  = $objPHPExcel->setActiveSheetIndex(0);
//$highestRow    = $objWorksheet->getHighestRow();
//$highestColumn = $objWorksheet->getHighestColumn();
//$highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);
//$arraydata = array();
//for ($row = 2; $row <= $highestRow;++$row)
//{
//    for ($col = 0; $col <$highestColumnIndex;++$col)
//    {
//        $value=$objWorksheet->getCellByColumnAndRow($col, $row)->getValue();
//        $arraydata[$row-2][$col]=$value;
//    }
//}
//
//echo '<pre>';
//print_r($arraydata);
//echo '</pre>';


$connect = mysqli_connect("localhost", "root", "", "test");
include ("Classes/PHPExcel/IOFactory.php");
$html="<table border='1'>";
$objPHPExcel = PHPExcel_IOFactory::load('example.xlsx');
foreach ($objPHPExcel->getWorksheetIterator() as $worksheet)
{
    $highestRow = $worksheet->getHighestRow();
    for ($row=2; $row<=$highestRow; $row++)
    {
        $html.="<tr>";
        $name = mysqli_real_escape_string($connect, $worksheet->getCellByColumnAndRow(0, $row)->getValue());
        $email = mysqli_real_escape_string($connect, $worksheet->getCellByColumnAndRow(1, $row)->getValue());
        $sql = "INSERT INTO test(Name, Email) VALUES ('".$name."', '".$email."')";
        mysqli_query($connect, $sql);
        $html.= '<td>'.$name.'</td>';
        $html .= '<td>'.$email.'</td>';
        $html .= "</tr>";
    }
}
$html .= '</table>';
echo $html;
echo '<br />Đây là dữ liệu được import từ excel ';
?>