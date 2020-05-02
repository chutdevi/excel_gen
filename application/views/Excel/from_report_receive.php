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
$freez = 'A4';
$start_col = 3; 
//var_dump($list_act_report); exit;

//exit;
$ind = 0;
foreach ($title as $inTil => $til) {
         $objPHPExcel->createSheet();
         $objPHPExcel->setActiveSheetIndex($ind);
         $objPHPExcel->getActiveSheet()->setTitle("$til");

//echo $til; exit;
$sheetIndex =  strtolower(str_replace(' ', '_', $title[$ind])); 
$end_row = count($list_act_report[$sheetIndex])+$start_col;
//echo $sheetIndex . " " . count($list_act_report[$sheetIndex][0]); exit;
if (count($list_act_report[$sheetIndex]) > 0) {  
$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(70);    
$objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 35 );
$objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 35 );
$objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 35 );

$objPHPExcel->getActiveSheet()
    ->getStyle($start_col)
    ->getAlignment()
    ->setWrapText(true)
    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);   








     $style =   array(  
                        'font'    => array( 'size' => 11, 
                                            'bold' => true,
                                            'color' => array('rgb' => 'FFFFFF')), 
                        'borders' => array(                                 
                                            'allborders' => array(
                                                                   'style' => PHPExcel_Style_Border::BORDER_THIN,
                                                                   'color' => array('rgb' => 'FFFFFF')
                                                                 )
                                          )
                    );                         


    $col_name = array();
    $i = 0;
    foreach ( range('A', 'Z') as $cm ) {
        array_push($col_name, $cm);
    }
    foreach(range('A',$col_name[count($list_act_report[$sheetIndex][0])-6]) as $columnID) {
            $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)
                                          ->setAutoSize(true);
    }
    foreach (range('I', 'O') as $c)
         $objPHPExcel->getActiveSheet()->getColumnDimension($c)->setWidth('18');
//echo count($list_act_report[$sheetIndex][0]); exit;
//var_dump($col_name); exit;
  // row column
//echo date('Y-M-d', strtotime(date('Y')."-".date('M')."-".(date('d')+1))); exit;
    $objPHPExcel->getActiveSheet()->getStyle('I2:O'.$end_row)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('A1', 'DAILY RM AND PART RECEIVE');
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('I1', date('Y-M-d', strtotime(date('Y')."-".date('m')."-".(date('d')-1)) ));
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('L1', date('Y-M-d', strtotime(date('Y')."-".date('m')."-".date('d') )));
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('N1', date('Y-M-d', strtotime(date('Y')."-".date('m')."-".(date('d')+1)) ));
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('O1', date('Y-M-d', strtotime(date('Y')."-".date('m')."-".(date('d')+2)) ));


    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('H1', 'DATE    >>> ');
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('H2', 'TOTAL  >>> ');

    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('I2', '=SUBTOTAL(9,I4:I'. $end_row . ')');
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('J2', '=SUBTOTAL(9,J4:J'. $end_row . ')');
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('K2', '=SUBTOTAL(9,K4:K'. $end_row . ')');
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('L2', 'TO DAY');
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('M2', '=SUBTOTAL(9,M4:M'. $end_row . ')');
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('N2', '=SUBTOTAL(9,N4:N'. $end_row . ')');
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('O2', '=SUBTOTAL(9,O4:O'. $end_row . ')');

    $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array( 'font' => Style_Font(48, '000000',true)));
    $objPHPExcel->getActiveSheet()->getStyle('I1:O1')->applyFromArray(array( 'font' => Style_Font(10, '963635',true)));//330000
    $objPHPExcel->getActiveSheet()->getStyle('I2:O2')->applyFromArray(array( 'font' => Style_Font(14, '0000FF', TRUE)));
    $objPHPExcel->getActiveSheet()->getStyle('H1:H2')->applyFromArray(array( 'font' => Style_Font(16, 'FFFFFF', TRUE)));
    $objPHPExcel->getActiveSheet()->getStyle('L2')->applyFromArray(array( 'font' => Style_Font(11, 'FFFFFF', TRUE)));
                                    
    $objPHPExcel->getActiveSheet()->getStyle('H1:O2')->applyFromArray(array(
                                                                            'borders' => array( 'allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                                                                               
                                $objPHPExcel->getActiveSheet()
                                            ->getStyle('A1:O2')
                                            ->getAlignment()
                                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);                                                                                                                                  
    $objPHPExcel->getActiveSheet()->getStyle('H1:H2')->applyFromArray(array('fill' => Style_Fill('9999FF')));
    $objPHPExcel->getActiveSheet()->getStyle('I1')->applyFromArray(array('fill'  => Style_Fill('FFCC99'))); 
    $objPHPExcel->getActiveSheet()->getStyle('L1')->applyFromArray(array('fill'  => Style_Fill('FFFF99'))); 
    $objPHPExcel->getActiveSheet()->getStyle('N1')->applyFromArray(array('fill'  => Style_Fill('CCFF99'))); 
    $objPHPExcel->getActiveSheet()->getStyle('O1')->applyFromArray(array('fill'  => Style_Fill('99FFCC'))); 

    $objPHPExcel->getActiveSheet()->getStyle('I2:O2')->applyFromArray(array('fill'  => Style_Fill('FFCCCC')));
    $objPHPExcel->getActiveSheet()->getStyle('L2')->applyFromArray(array('fill'  => Style_Fill($colhead)));
    $objPHPExcel->getActiveSheet()->mergeCells('A1:'.'G2');
    $objPHPExcel->getActiveSheet()->mergeCells('I1:'.'K1');
    $objPHPExcel->getActiveSheet()->mergeCells('L1:'.'M1');
    foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) {
        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++].$start_col, strtoupper(str_replace('_', ' ', $key)));       
    }

    $objPHPExcel->getActiveSheet()->getStyle($col_name[0].$start_col.":".$col_name[count($list_act_report[$sheetIndex][0])-1].$start_col)->applyFromArray($style);



    $objPHPExcel->getActiveSheet()
        ->getStyle($col_name[0].$start_col.":".$col_name[count($list_act_report[$sheetIndex][0])-1].$start_col)
        ->applyFromArray(
            array(
                'fill' => array(
                                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                'color' => array('rgb' => $colhead)
                               )
                 )
        );




    $row = $start_col+1;
    foreach ($list_act_report[$sheetIndex] as $key => $value) {
        
        $col = 0;
        foreach ($value as $body => $val) {
                //var_dump($body); exit;
                if ($body == 'stock' && $val < 1) {
                    $UNREADFontStyle = array( 'font' => Style_Font(11,'FF0000',TRUE) );          
                    $objPHPExcel->getActiveSheet()->getStyle('L'.$row)->applyFromArray($UNREADFontStyle);
                }
                elseif ($body == 'diff' && $val < 0) {
                $UNREADFillStyle = array( 'fill' => Style_Fill('FFCCFF') );   
                $UNREADFontStyle = array( 'font' => Style_Font(11,'FF0000',true) );          
                $objPHPExcel->getActiveSheet()->getStyle($col_name[0].$row.":".$col_name[count($list_act_report[$sheetIndex][0])-1].$row)->applyFromArray($UNREADFillStyle);
                $objPHPExcel->getActiveSheet()->getStyle($col_name[count($list_act_report[$sheetIndex][0])-5].$row)->applyFromArray($UNREADFontStyle);
                }
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);
               // var_dump($val); 
        }
       // var_dump($value); 
      //  exit;
        $row++;
        $objPHPExcel->getActiveSheet()->setAutoFilter($col_name[0].$start_col.":".$col_name[count($list_act_report[$sheetIndex][0])-1].$start_col);
        $objPHPExcel->getActiveSheet()->freezePane($freez);
    }
} else {

            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "No data ".$til.".");
            $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true)));
            //echo "Non data."; exit;
}
// $objPHPExcel->getActiveSheet()->setTitle($title);
$ind++;


}

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->getStyle('ZZ1')->getNumberFormat()->setFormatCode('_-* #,##0_-;-* #,##0_-;_-* "-"??_-;_-@_-');
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

function Style_border($line='BORDER_THICK', $color='000000')
{
    return array( 'style' => $line, 'color' => array('rgb' => $color)) ;
}


function holiday($dat, $hol)
{

    foreach ($hol as $ld) 
        if ( substr( $ld['d_t'], 8,2 ) == substr( $dat, 0,2 ) ) 
            return true;
}

function set_autosize($colA = 'A' ,$colB = 'Z',  $objPHPExcel = nul, $index = 0)
{
        $objPHPExcel->setActiveSheetIndex($index); 
    foreach(range($colA, $colB) as $columnID) 
    {        
        $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);       
    }                     
}


?>

 
