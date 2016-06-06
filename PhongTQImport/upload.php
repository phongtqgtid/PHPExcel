<?php
/**
 * Created by PhpStorm.
 * User: quatt_000
 * Date: 6/6/16
 * Time: 12:33 PM
 */

if(isset($_FILES['file']))
{
    $file = $_FILES['file'];
    //file properties
    $filename = $file['name'];
    $file_tmp= $file['tmp_name'];
    $file_size = $file['size'];
    $file_erro = $file['error'];

    $file_ext = explode('.',$filename);
    $file_ext=  strtolower(end($file_ext));
    $allow = array('txt', 'jpg','xlsx','xls','csv');
    if(in_array($file_ext,$allow))
    {
        if($file_erro === 0)
        {
            if($file_size <= 2097152)
            {
                $filename_new = uniqid('',true). '.' .$file_ext;
                $file_destination = 'Uploads/' . $filename_new;

                if(move_uploaded_file($file_tmp,$file_destination))
                {
//                    echo $file_destination;
                    $connect = mysqli_connect("localhost", "root", "", "test");
                    include ("Classes/PHPExcel/IOFactory.php");
                    $html="<table border='1'>";
                    $objPHPExcel = PHPExcel_IOFactory::load( $file_destination);
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
                }
            }
        }

    }
}