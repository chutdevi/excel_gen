<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$data_col = array();
//var_dump($list_act_report); exit;

//exit;
$ind = 0;
foreach ($title as $inTil => $til) {
         $objPHPExcel->createSheet();
         $objPHPExcel->setActiveSheetIndex($ind);
         $objPHPExcel->getActiveSheet()->setTitle("$til");

//echo $til; exit;
$sheetIndex =  strtolower(str_replace(' ', '_', $title[$ind])); 
//echo $sheetIndex . " " . count($list_act_report[$sheetIndex][0]); exit;
if (count($list_act_report[$sheetIndex]) > 0) {      
$objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 28 );
$objPHPExcel->getActiveSheet()
    ->getStyle('2')
    ->getAlignment()
    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);       
 $style =   array( 
                    'font'    => array( 'size' => 11, 
                                        'bold' => true,
                                        'color' => array('rgb' => 'FFFFFF')), 
                    'borders' => array(                                 
                                        'allborders' => array(
                                                               'style' => PHPExcel_Style_Border::BORDER_THICK,
                                                               'color' => array('rgb' => 'FFFFFF')
                                                             )
                                      )
                );                        
// echo count($list_act_report[$sheetIndex]); exit;    


$col_name = array();
$i = 0;
foreach ( range('A', 'Z') as $cm ) {
    array_push($col_name, $cm);
}
//echo count($list_act_report[$sheetIndex][0]); exit;
//var_dump($col_name); exit;

$Date_today = date('Y-m-d');
$objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('F'."1", '40 = PART , 30 = FW , 20 = RM , 10 = FG'); 
$objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('L'."1", 'DATE : '.$Date_today);  

foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) {
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."2", strtoupper($key));       
}

$objPHPExcel->getActiveSheet()->getStyle($col_name[0]."2:".$col_name[count($list_act_report[$sheetIndex][0])-1]."2")->applyFromArray($style);

$objPHPExcel->getActiveSheet()
    ->getStyle($col_name[0]."2:".$col_name[count($list_act_report[$sheetIndex][0])-1]."2")
    ->applyFromArray(
        array(
            'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => $colhead)
                           )
             )
    );


foreach(range('A',$col_name[count($list_act_report[$sheetIndex][0])]) as $columnID) {
    $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)
        ->setAutoSize(true);
}

$row = 3;
foreach ($list_act_report[$sheetIndex] as $key => $value) {
    
    $col = 0;
    foreach ($value as $body => $val) {
            if ($val == 'UNREAD') {
                $UNREADFillStyle = array( 'fill' => Style_Fill('FF6666') );   
                $UNREADFontStyle = array( 'font' => Style_Font(12,'FF0000',true) );          
                $objPHPExcel->getActiveSheet()->getStyle($col_name[0].$row.":".$col_name[count($list_act_report[$sheetIndex][0])-2].$row)->applyFromArray($UNREADFillStyle);
                $objPHPExcel->getActiveSheet()->getStyle($col_name[count($list_act_report[$sheetIndex][0])-1].$row)->applyFromArray($UNREADFontStyle);
            }

            elseif($val == 'READ') {

                $objPHPExcel->getActiveSheet()->getStyle($col_name[count($list_act_report[$sheetIndex][0])-1].$row)->applyFromArray(array( 'font' => Style_Font(11,'4C9900',true)));
            }
            elseif ($val == 'UNCOMPLETE') {
                $UNREADFillStyle = array( 'fill' => Style_Fill('FFE900') );   
                $UNREADFontStyle = array( 'font' => Style_Font(12,'FF0000',true) );          
                $objPHPExcel->getActiveSheet()->getStyle($col_name[0].$row.":".$col_name[count($list_act_report[$sheetIndex][0])-2].$row)->applyFromArray($UNREADFillStyle);
                $objPHPExcel->getActiveSheet()->getStyle($col_name[count($list_act_report[$sheetIndex][0])-1].$row)->applyFromArray($UNREADFontStyle);
                $objPHPExcel->getActiveSheet()->getStyle('K'.$row)->applyFromArray(array('font' => Style_Font(11,'FF0000',true)));
            }
            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);
           // var_dump($val); 
    }
   // var_dump($value); 
  //  exit;
    $row++;
    $objPHPExcel->getActiveSheet()->setAutoFilter($col_name[0]."2:".$col_name[count($list_act_report[$sheetIndex][0])-1]."2");
    $objPHPExcel->getActiveSheet()->freezePane('A3');
}
} else {

            $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A1', "No data ".$til.".");
            $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true)));
            //echo "Non data."; exit;
}
// $objPHPExcel->getActiveSheet()->setTitle($title);
$ind++;


}

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);
 $objPHPExcel->getActiveSheet()->getStyle('ZZ1');
$objPHPExcel->removeSheetByIndex(count($title));
$today = date("My");
//Redirect output to a client’s web browser (Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
$con = 'Content-Disposition: attachment;filename='.$filename.date('d').'.xlsx';
//echo $con; exit;
header($con);
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0


$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
// $con = 'ship_remain'.date('d').'.xlsx';
// $objWriter->save('D:/'.$con);
exit;

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


function Style_Fill($color=null) {

    return array( 'type'  => PHPExcel_Style_Fill::FILL_SOLID,                           
                  'color' => array('rgb' => $color)                                    
                );                                   
}

function Style_Font($size=11, $color='FFFFFF', $bol=false) {

    return  array(
                    'size' => $size,
                    'bold' => $bol,
                    'color' => array('rgb' => $color)
                 );                               
}


?>

 
