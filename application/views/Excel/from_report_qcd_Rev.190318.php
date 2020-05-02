<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

$path = "D:/AppServ/www/report/script/qcd_task/log/";
$fordate1 = 'n1date1.txt'; $getfordate1 = rundata($path.$fordate1);
$suby1 = substr($getfordate1,0,4);
$subm1 = substr($getfordate1,5,2);
$fordate2 = 'n2date2.txt'; $getfordate2 = rundata($path.$fordate2);
$suby2 = substr($getfordate2,0,4);
$subm2 = substr($getfordate2,5,2);
$fordate3 = 'n3date2.txt'; $getfordate3 = rundata($path.$fordate3);
$suby3 = substr($getfordate3,0,4);
$subm3 = substr($getfordate3,5,2);
$fordate4 = 'n4date2.txt'; $getfordate4 = rundata($path.$fordate4);
$suby4 = substr($getfordate4,0,4);
$subm4 = substr($getfordate4,5,2);
$fordate = 'n5date2.txt'; $getfordate = rundata($path.$fordate);
$suby5 = substr($getfordate,0,4);
$subm5 = substr($getfordate,5,2);
$subyear = substr($getfordate,0,4);
//echo $getfordate; exit;
$ct = array( array(), array(), array(), array(), array(), array(), array(), array() ); 
$cld = array( 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ','AK', 'AL', 'AM'); 

if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

$dayA   = date('d');
$monthA = date('m');
$monthB = date('M');
$yearA  = date('Y');
$lastmonth = substr(date('Y/m/t',strtotime($yearA."/".$monthA."/".$dayA)),8, 2);
//$lastmonth = 31;
$curdate = $dayA."-".$monthB."-".$yearA;

if($monthA == "01"){$monthfull = "JANUARY";}
else if($monthA == "02"){$monthfull = "FEBRUARY";}
else if($monthA == "03"){$monthfull = "MARCH";}
else if($monthA == "04"){$monthfull = "APRIL";}
else if($monthA == "05"){$monthfull = "MAY";}
else if($monthA == "06"){$monthfull = "JUNE";}
else if($monthA == "07"){$monthfull = "JULY";}
else if($monthA == "08"){$monthfull = "AUGUST";}
else if($monthA == "09"){$monthfull = "SEPTEMBER";}
else if($monthA == "10"){$monthfull = "OCTOBER";}
else if($monthA == "11"){$monthfull = "NOVEMBER";}
else if($monthA == "12"){$monthfull = "DECEMBER";}

if($subm5 == "01"){$monthfull2 = "JANUARY"; $m5 = "Jan";}
else if($subm5 == "02"){$monthfull2 = "FEBRUARY"; $m5 = "Feb";}
else if($subm5 == "03"){$monthfull2 = "MARCH"; $m5 = "Mar";}
else if($subm5 == "04"){$monthfull2 = "APRIL"; $m5 = "Apr";}
else if($subm5 == "05"){$monthfull2 = "MAY"; $m5 = "May";}
else if($subm5 == "06"){$monthfull2 = "JUNE"; $m5 = "Jun";}
else if($subm5 == "07"){$monthfull2 = "JULY"; $m5 = "Jul";}
else if($subm5 == "08"){$monthfull2 = "AUGUST"; $m5 = "Aug";}
else if($subm5 == "09"){$monthfull2 = "SEPTEMBER"; $m5 = "Sep";}
else if($subm5 == "10"){$monthfull2 = "OCTOBER"; $m5 = "Oct";}
else if($subm5 == "11"){$monthfull2 = "NOVEMBER"; $m5 = "Nov";}
else if($subm5 == "12"){$monthfull2 = "DECEMBER"; $m5 = "Dec";}

if($subm1 == "01"){$m1 = "Jan";}
else if($subm1 == "02"){$m1 = "Feb";}
else if($subm1 == "03"){$m1 = "Mar";}
else if($subm1 == "04"){$m1 = "Apr";}
else if($subm1 == "05"){$m1 = "May";}
else if($subm1 == "06"){$m1 = "Jun";}
else if($subm1 == "07"){$m1 = "Jul";}
else if($subm1 == "08"){$m1 = "Aug";}
else if($subm1 == "09"){$m1 = "Sep";}
else if($subm1 == "10"){$m1 = "Oct";}
else if($subm1 == "11"){$m1 = "Nov";}
else if($subm1 == "12"){$m1 = "Dec";}

if($subm2 == "01"){$m2 = "Jan";}
else if($subm2 == "02"){$m2 = "Feb";}
else if($subm2 == "03"){$m2 = "Mar";}
else if($subm2 == "04"){$m2 = "Apr";}
else if($subm2 == "05"){$m2 = "May";}
else if($subm2 == "06"){$m2 = "Jun";}
else if($subm2 == "07"){$m2 = "Jul";}
else if($subm2 == "08"){$m2 = "Aug";}
else if($subm2 == "09"){$m2 = "Sep";}
else if($subm2 == "10"){$m2 = "Oct";}
else if($subm2 == "11"){$m2 = "Nov";}
else if($subm2 == "12"){$m2 = "Dec";}

if($subm3 == "01"){$m3 = "Jan";}
else if($subm3 == "02"){$m3 = "Feb";}
else if($subm3 == "03"){$m3 = "Mar";}
else if($subm3 == "04"){$m3 = "Apr";}
else if($subm3 == "05"){$m3 = "May";}
else if($subm3 == "06"){$m3 = "Jun";}
else if($subm3 == "07"){$m3 = "Jul";}
else if($subm3 == "08"){$m3 = "Aug";}
else if($subm3 == "09"){$m3 = "Sep";}
else if($subm3 == "10"){$m3 = "Oct";}
else if($subm3 == "11"){$m3 = "Nov";}
else if($subm3 == "12"){$m3 = "Dec";}

if($subm4 == "01"){$m4 = "Jan";}
else if($subm4 == "02"){$m4 = "Feb";}
else if($subm4 == "03"){$m4 = "Mar";}
else if($subm4 == "04"){$m4 = "Apr";}
else if($subm4 == "05"){$m4 = "May";}
else if($subm4 == "06"){$m4 = "Jun";}
else if($subm4 == "07"){$m4 = "Jul";}
else if($subm4 == "08"){$m4 = "Aug";}
else if($subm4 == "09"){$m4 = "Sep";}
else if($subm4 == "10"){$m4 = "Oct";}
else if($subm4 == "11"){$m4 = "Nov";}
else if($subm4 == "12"){$m4 = "Dec";}

$showfdate = $monthfull." ".$yearA."  -  ".$monthfull2." ".$subyear;
$showf = $monthB."-".$yearA;
$showf1 = $m1."-".$suby1;
$showf2 = $m2."-".$suby2;
$showf3 = $m3."-".$suby3;
$showf4 = $m4."-".$suby4;
$showf5 = $m5."-".$suby5;

$curmonth = $monthfull." ".$yearA;
//echo $dayA; echo $monthA; echo $yearA; echo $lastmonth; exit;
$month1 = "Month R1";
$month2 = "Month R2";
$month3 = "Month R3";
$daily = "Daily";

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$data_col = array();
//var_dump($list_act_report); exit;

$col_name = array();

foreach ( range('A', 'Z') as $cm ) { array_push($col_name, $cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "A".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "B".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "C".$cm); }
$i   = 0;   
$ind = 0;
// var_dump($title);
// exit();
foreach ($title as $inTil => $til) 
{
             $objPHPExcel->createSheet();
             $objPHPExcel->setActiveSheetIndex($ind);
             //$objPHPExcel->getActiveSheet()->setTitle( "$til ( ". date('Y-m-d') . " )" );
             $objPHPExcel->getActiveSheet()->setTitle( "$til" );
             //$objPHPExcel->getActiveSheet()->setShowGridlines(False);

    $sheetIndex  =  strtolower(str_replace(' ', '_', $title[$ind])); 
    $count_index = 0;
    $count_data  =  count($list_act_report[$sheetIndex]) + 5;
    $cat = ($count_data+6);

    if ($count_data - 4  > 0) 
    {   
        if ($til == 'PCL QCD Daily Report') { 
                //$count_index =  count($list_act_report[$sheetIndex][0]) - 6 ;
                $count_index =  count($list_act_report[$sheetIndex][0]) - 3;
                $count_data  =  count($list_act_report[$sheetIndex]) + 5;
                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 35 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 35 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 30 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 30 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 5 )->setRowHeight( 12 );
                $objPHPExcel->getActiveSheet()
                    ->getStyle('1:5')
                    ->getAlignment()
                    //->setWrapText(true)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);  

                foreach (range(7, 30) as $c)
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$c])->setWidth('10');
                $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('20');
                $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('24');
                $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('19');
                $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('23');
                $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('13');
                $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('35');
                $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('15');
                $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('16'); 
                $objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setWidth('10');     
                $objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setWidth('13');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AV')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AX')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AZ')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BB')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BD')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BF')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BH')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BJ')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BL')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BN')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BP')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BR')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BT')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BV')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('BX')->setWidth('12');
                                       
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('A5:'.$col_name[$count_index].'5');
                $objPHPExcel->getActiveSheet()->freezePane('A6');

                $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[$count_index].'4')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                $objPHPExcel->getActiveSheet()->getStyle('D6:D'.$count_data)->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                $objPHPExcel->getActiveSheet()->getStyle('E6:BY'.$count_data)->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'d9d9d9'))));
                $objPHPExcel->getActiveSheet()->getStyle('E6:BY'.$count_data)->applyFromArray(array('fill' => Style_Fill('FFFFFF')));
                $objPHPExcel->getActiveSheet()->getStyle('A23:BY23')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A49:BY49')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A60:BY60')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A62:BY62')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A64:BY64')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A65:BY65')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A66:BY66')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A67:BY67')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A73:BY73')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A77:BY77')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A94:BY94')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A95:BY95')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A96:BY96')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));

                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {
                    //echo $key; 
                    if($key != 'NO' && $key != 'PRODUCT_NO'){
                        if($key == 'CUSTOMER_NAME' || $key == 'CUSTOMER_NAME2' || $key == 'PRODUCT_GROUP' || $key == 'PRODUCT_GROUP2' || $key == 'MODEL' || $key == 'MODEL_REF'){
                            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."3", str_replace("_", " ", strtoupper($key)));       
                        }else{
                            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."4", str_replace("_", " ", strtoupper($key)));  
                        }
                    }
                }            
                $row = 6;
                $st=$row;
                $gt=$row;
                $ccp = $list_act_report[$sheetIndex][0]['CUSTOMER_NAME'];
                $gcp = $list_act_report[$sheetIndex][0]['PRODUCT_GROUP2'];
                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {
                    $col = 0;
                    //echo $body; exit;
                        if( $value['PRODUCT_NO'] == 1 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row ); //array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 2 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 3 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 4 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 5 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 6 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 7 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 8 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );

                    foreach ($value as $body => $val) 
                    {
                        
                        if($body != 'NO' && $body != 'PRODUCT_NO'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);                
                        }

                        if ($body == 'CUSTOMER_NAME'){
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(14,'FFFFFF',true,false,'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->applyFromArray(array('fill' => Style_Fill('1f497d')));

                            if($val != $ccp)
                            {
                                if( ($row - $st) > 0 )
                                $objPHPExcel->getActiveSheet()->getStyle('A' . $st . ':' . 'A' . ($row-1))->applyFromArray(array('fill' => Style_Fill('1f497d')));  
                                $objPHPExcel->getActiveSheet()->mergeCells( 'A' . $st . ':' . 'A' . ($row-1) );
                                //$objPHPExcel->getActiveSheet()->mergeCells( 'C' . $st . ':' . 'C' . ($row-1) );
                                $objPHPExcel->getActiveSheet()->getStyle( 'A'. ($row-1) )
                                                              ->applyFromArray(array(
                                                                'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                                $ccp = $val;
                                $st  = $row;
                            }
                        }   

                        if($body == 'PRODUCT_GROUP2')
                        {
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->applyFromArray(array('fill' => Style_Fill('daeef3')));

                            if($val != $gcp)
                            {
                                if( ($row - $gt) > 0 )
                                    $objPHPExcel->getActiveSheet()->mergeCells( 'D' . $gt . ':' . 'D' . ($row-1) );
                                	// $objPHPExcel->getActiveSheet()->getStyle( 'D'. ($row-1) )
                                 //                              ->applyFromArray(array(
                                 //                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                                    $gcp = $val;
                                    $gt  = $row;
                            }
                        }

                        if ($body == 'MODEL'){
                            if ($value['MODEL'] == '3E00') 
                            {
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->getNumberFormat()->setFormatCode('###"E00"');
                                              $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), $value['MODEL']);
                            }
                        }   

                        if (substr($body,0,1) == 'D') {
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].$row)
                                                ->applyFromArray(array('font' => Style_Font(11,'ff0000',false,false,'Calibri')));
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                        } 

                        if (substr($body,0,3) == 'MTH') {
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',false,false,'Calibri')));
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                        } 

                        if (substr($body,0,2) == 'FG') {
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        } 

                        if (substr($body,0,5) == 'STOCK') {
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0.0_-;[RED](#,##0.0)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        } 

                        if (substr($body,0,3) == 'ODR') {
                            if (substr($body,5,2) == $dayA) {
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].$row)
                                                ->applyFromArray(array('font' => Style_Font(11,'002060',true,false,'Calibri')));
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->applyFromArray(array('fill' => Style_Fill('ebf1de'))); 
                            }else{
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].$row)
                                                ->applyFromArray(array('font' => Style_Font(11,'000000',false,false,'Calibri')));
                            }
                        } 

                        if($body == 'STOCK_LEV'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=IF('.'AQ'.($row).'=0,0,('.'G'.($row).'/'.'AQ'.($row).'))' );   
                        }

                        if($body == 'ODR_ACCUM'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=SUM(I'.($row).':AM'.($row).')' );   
                        }

                        if($body == 'ODR_PROG'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=IF('.'AP'.($row).'=0,0,('.'AN'.($row).'/'.'AP'.($row).'))' );  
                            $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0%_-;[RED](#,##0%)_-;_-* "-"??_-;_-@_-'); 
                        }
                    }
                    $row++; 
                }

                //====================================SUMMARY CUSTOMER DEMAND BY PRODUCTS CATEGORY====================================//
                foreach(array('G','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY') as $cel )
                    put_data($objPHPExcel, $ct, $cel, ($count_data+4));
                // foreach(array('AN') as $cel )
                //     put_data($objPHPExcel, $ct, $cel, ($count_data+4));

                for($i = ($count_data+4); $i < ($count_data+12); $i++){
                	$temp = 'A'.$i.':BY'.$i;
                	$objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                    $objPHPExcel->getActiveSheet()->getStyle( $temp )
                                                              ->applyFromArray(array(
                                                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                      
                }

                for($i = ($count_data+4); $i < ($count_data+12); $i++){
                	$temp = 'AO'.$i;
                	$objPHPExcel->setActiveSheetIndex()->setCellValue($temp, '=IF('.'AP'.($i).'=0,0,('.'AN'.($i).'/'.'AP'.($i).'))' );
                	$objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0%_-;[RED](#,##0%)_-;_-* "-"??_-;_-@_-'); 
                }

                for($i = ($count_data+4); $i < ($count_data+12); $i++){
                	$temp = 'H'.$i;											 
                	$objPHPExcel->setActiveSheetIndex()->setCellValue($temp, '=IF('.'AQ'.($i).'=0,0,('.'G'.($i).'/'.'AQ'.($i).'))' );
                	$objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0.0_-;[RED](#,##0.0)_-;_-* "-"??_-;_-@_-'); 
                }

                for($i = ($count_data+4); $i < ($count_data+15); $i++){
                	$temp = 'A'.$i;
                	if($i < ($count_data+12)){
                		$objPHPExcel->setActiveSheetIndex()->setCellValue($temp, "c" );
                		$objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Wingdings 3')));
                		Style_Alignment($temp, 3, false, $objPHPExcel); 
                	}else if($i == ($count_data+14)){
                		$objPHPExcel->setActiveSheetIndex()->setCellValue($temp, "ISSUED BY PC SYSTEM ON ".$dayA."-".$monthB."-".$yearA);
                		$objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(12,'000000',true,true,'Calibri')));
                		Style_Alignment($temp, 9, false, $objPHPExcel); 
                	}
                }

                for($i = ($count_data+4); $i < ($count_data+12); $i++){
                	$temp = 'B'.$i;
                	$temp2 = 'B'.$i.':F'.$i;
                	if($i == ($count_data+4)){
                		$type = "WATER PUMP";
                	}else if($i == ($count_data+5)){
                		$type = "OIL PUMP";
                	}else if($i == ($count_data+6)){
                		$type = "WHEEL CYT";
                	}else if($i == ($count_data+7)){
                		$type = "FORK SHIFT";
                	}else if($i == ($count_data+8)){
                		$type = "BRAKE";
                	}else if($i == ($count_data+9)){
                		$type = "GEAR";
                	}else if($i == ($count_data+10)){
                		$type = "BEARING";
                	}else if($i == ($count_data+11)){
                		$type = "OTHER";
                	}
                	$objPHPExcel->setActiveSheetIndex()->setCellValue($temp, $type);
                	$objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(14,'000000',true,true,'Calibri')));
                	$objPHPExcel->getActiveSheet()->mergeCells($temp2);
                }
                //var_dump($count_data+4); exit;`

                $objPHPExcel->getActiveSheet()->getStyle('A'.($count_data+3))->applyFromArray(array('font' => Style_Font(14,'FFFFFF',true,true,'Calibri'))); 
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A'.($count_data+3), "SUMMARY CUSTOMER DEMAND BY PRODUCTS CATEGORY");
                $objPHPExcel->getActiveSheet()->mergeCells('A'.($count_data+3).':'.'BY'.($count_data+3));
                $objPHPExcel->getActiveSheet()->mergeCells('A'.($count_data+12).':'.'BY'.($count_data+12));
                $objPHPExcel->getActiveSheet()->getStyle('A'.($count_data+3))->applyFromArray(array('fill' => Style_Fill('974706')));
                $objPHPExcel->getActiveSheet()->getStyle('A'.($count_data+12))->applyFromArray(array('fill' => Style_Fill('974706')));
                $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('fill' => Style_Fill('1f497d'))); //BLUE COLOR
                $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(24,'FFFFFF',true,true,'Calibri Light')));  
                $objPHPExcel->getActiveSheet()->getStyle('A3:H4')->applyFromArray(array('fill' => Style_Fill('daeef3'))); //LIGHT BLUE COLOR
                $objPHPExcel->getActiveSheet()->getStyle('A3:H4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('I1:AL4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('I1:AO2')->applyFromArray(array('fill' => Style_Fill('669900'))); //GREEN COLOR
                $objPHPExcel->getActiveSheet()->getStyle('I3:AO4')->applyFromArray(array('fill' => Style_Fill('cce199'))); //LIGHT GREEN COLOR
                $objPHPExcel->getActiveSheet()->getStyle('I1:AO2')->applyFromArray(array('font' => Style_Font(16,'FFFFFF',true,true,'Calibri Light'))); 
                $objPHPExcel->getActiveSheet()->getStyle('I3:AO3')->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('I4:AO4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AP1:BY2')->applyFromArray(array('fill' => Style_Fill('ffcc00'))); //ORANGE COLOR
                $objPHPExcel->getActiveSheet()->getStyle('AP3:BY4')->applyFromArray(array('fill' => Style_Fill('ffff99'))); //LIGHT ORANGE COLOR
                $objPHPExcel->getActiveSheet()->getStyle('AP1:BY2')->applyFromArray(array('font' => Style_Font(16,'000000',true,true,'Calibri Light'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AP3:BY4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AP4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AQ4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AR4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AS4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AT4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AU4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AV4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AW4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AX4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AY4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AZ4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BA4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BB4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BC4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BD4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BE4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BF4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BG4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BH4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BI4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BJ4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BK4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BL4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BM4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BN4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BO4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BP4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BQ4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BR4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BS4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BT4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BU4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri')));
                $objPHPExcel->getActiveSheet()->getStyle('BV4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BW4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri')));  
                $objPHPExcel->getActiveSheet()->getStyle('BX4')->applyFromArray(array('font' => Style_Font(12,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('BY4')->applyFromArray(array('font' => Style_Font(12,'ff0000',true,false,'Calibri'))); 
                // $objPHPExcel->getActiveSheet()->getStyle('K1')->applyFromArray(array('fill' => Style_Fill('ddddff')));
                //$objPHPExcel->getActiveSheet()->getStyle('L1')->applyFromArray(array('fill' => Style_Fill('d9ffd9')));
                // $objPHPExcel->getActiveSheet()->getStyle('L1')->applyFromArray(array('fill' => Style_Fill('fce4d6')));
                //$objPHPExcel->getActiveSheet()->getStyle('M5:M'.$count_data)->applyFromArray(array('fill' => Style_Fill('F8F9F9')));
                //$objPHPExcel->setActiveSheetIndex($ind)->setCellValue('L3', "ACTUAL INCOME \r\n (THB)");

                //==============================================TITLE====================================================//
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "PC&L QCD DAILY REPORT"." ( ".$curdate." ) ");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A3', "CUSTOMERS");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('B3', "CUSTOMERS FILTER");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('C3', "GROUP FILTER");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('D3', "GROUP PART");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('E3', "MODEL");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('F3', "REF");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('G3', "EXP/JA");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('G4', "STOCK (QTY)");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('H3', "STOCK LEVEL");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('H4', "[ DAY ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('I1', "DAILY DELIVERY PLAN");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('I2', $curmonth);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AN3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AO3', "[ % ]");

                // foreach(range('A',$col_name[$count_index]) as $columnID) 
                //     $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('I3:'.'AJ3', "[ PCS. ]");

                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AP1', "CUSTOMER DEMAND FORECAST");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AP2', $showfdate);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('I3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('I4', "01st");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('J3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('J4', "02nd");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('K3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('K4', "03nd");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('L3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('L4', "04th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('M3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('M4', "05th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('N3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('N4', "06th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('O3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('O4', "07th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('P3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('P4', "08th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Q3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Q4', "09th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('R3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('R4', "10th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('S3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('S4', "11th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('T3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('T4', "12th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('U3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('U4', "13th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('V3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('V4', "14th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('W3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('W4', "15th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('X3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('X4', "16th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Y3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Y4', "17th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Z3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Z4', "18th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AA3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AA4', "19th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AB3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AB4', "20th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AC3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AC4', "21th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AD3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AD4', "22th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AE3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AE4', "23th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AF3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AF4', "24th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AG3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AG4', "25th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AH3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AH4', "26th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AI3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AI4', "27th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AJ3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AJ4', "28th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AK3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AK4', "29th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AL3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AL4', "30th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AM3', "[ PCS. ]");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AM4', "31th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AN4', "ACCUM");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AO4', "PROGRESS");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AP3', $showf);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AR3', $showf);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AT3', $showf);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AV3', $showf1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AX3', $showf1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AZ3', $showf1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BB3', $showf2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BD3', $showf2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BF3', $showf2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BH3', $showf3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BJ3', $showf3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BL3', $showf3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BN3', $showf4);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BP3', $showf4);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BR3', $showf4);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BT3', $showf5);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BV3', $showf5);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BX3', $showf5);
                //echo $showf."".$showf1."".$showf2; exit;
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AP4', $month1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AQ4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AR4', $month2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AS4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AT4', $month3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AU4', $daily);

                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AV4', $month1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AW4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AX4', $month2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AY4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AZ4', $month3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BA4', $daily);

                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BB4', $month1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BC4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BD4', $month2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BE4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BF4', $month3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BG4', $daily);

                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BH4', $month1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BI4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BJ4', $month2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BK4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BL4', $month3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BM4', $daily);

                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BN4', $month1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BO4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BP4', $month2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BQ4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BR4', $month3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BS4', $daily);

                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BT4', $month1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BU4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BV4', $month2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BW4', $daily);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BX4', $month3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('BY4', $daily);
                //==============================================TITLE====================================================//

                $objPHPExcel->getActiveSheet()->mergeCells('A1:'.'H2');
                $objPHPExcel->getActiveSheet()->mergeCells('I1:'.'AO1');
                $objPHPExcel->getActiveSheet()->mergeCells('I2:'.'AO2');
                $objPHPExcel->getActiveSheet()->mergeCells('AP1:'.'BY1');
                $objPHPExcel->getActiveSheet()->mergeCells('AP2:'.'BY2');
                $objPHPExcel->getActiveSheet()->mergeCells('AP3:'.'AQ3');
                $objPHPExcel->getActiveSheet()->mergeCells('AR3:'.'AS3');
                $objPHPExcel->getActiveSheet()->mergeCells('AT3:'.'AU3');
                $objPHPExcel->getActiveSheet()->mergeCells('AV3:'.'AW3');
                $objPHPExcel->getActiveSheet()->mergeCells('AX3:'.'AY3');
                $objPHPExcel->getActiveSheet()->mergeCells('AZ3:'.'BA3');
                $objPHPExcel->getActiveSheet()->mergeCells('BB3:'.'BC3');
                $objPHPExcel->getActiveSheet()->mergeCells('BD3:'.'BE3');
                $objPHPExcel->getActiveSheet()->mergeCells('BF3:'.'BG3');
                $objPHPExcel->getActiveSheet()->mergeCells('BH3:'.'BI3');
                $objPHPExcel->getActiveSheet()->mergeCells('BJ3:'.'BK3');
                $objPHPExcel->getActiveSheet()->mergeCells('BL3:'.'BM3');
                $objPHPExcel->getActiveSheet()->mergeCells('BN3:'.'BO3');
                $objPHPExcel->getActiveSheet()->mergeCells('BP3:'.'BQ3');
                $objPHPExcel->getActiveSheet()->mergeCells('BR3:'.'BS3');
                $objPHPExcel->getActiveSheet()->mergeCells('BT3:'.'BU3');
                $objPHPExcel->getActiveSheet()->mergeCells('BV3:'.'BW3');
                $objPHPExcel->getActiveSheet()->mergeCells('BX3:'.'BY3');
                //
                $objPHPExcel->getActiveSheet()->mergeCells('A3:'.'A4');
                $objPHPExcel->getActiveSheet()->mergeCells('B3:'.'B4');
                $objPHPExcel->getActiveSheet()->mergeCells('C3:'.'C4');
                $objPHPExcel->getActiveSheet()->mergeCells('D3:'.'D4');
                $objPHPExcel->getActiveSheet()->mergeCells('E3:'.'E4');
                $objPHPExcel->getActiveSheet()->mergeCells('F3:'.'F4');

             //    if($lastmonth == "28"){
	            //     $objPHPExcel->getActiveSheet()->getColumnDimension('AK')->setVisible(false);
	            //     $objPHPExcel->getActiveSheet()->getColumnDimension('AL')->setVisible(false);
	            //     $objPHPExcel->getActiveSheet()->getColumnDimension('AM')->setVisible(false);
	            // }else if($lastmonth == "29"){
	            //     $objPHPExcel->getActiveSheet()->getColumnDimension('AL')->setVisible(false);
	            //     $objPHPExcel->getActiveSheet()->getColumnDimension('AM')->setVisible(false);
	            // }else if($lastmonth == "30"){
	            //     $objPHPExcel->getActiveSheet()->getColumnDimension('AM')->setVisible(false);
	            // }
//$dayA = 30;	
                $strcol = $dayA+0;
				$hidecol = 30;
				//echo $strcol. $hidecol; exit;
				for($i = $strcol; $i <= $hidecol; $i++){
					//echo $cld[$i]; 
	            	$objPHPExcel->getActiveSheet()->getColumnDimension($cld[$i])->setVisible(false);
	            }
//exit;
	            $x = 8;
                $num = $x + $dayA;
                for ($x = 8; $x < $num-1; $x++) {
                    Style_group_Col($col_name, $x, $objPHPExcel, 1);
                }

	            if($dayA <= 10){
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AS')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AU')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AX')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AY')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AZ')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BA')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BD')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BE')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BF')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BG')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BJ')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BK')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BL')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BM')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BP')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BQ')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BR')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BS')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BV')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BW')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BX')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BY')->setVisible(false);
	            }else if($dayA <= 20){
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AU')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AV')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AW')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AZ')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BA')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BB')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BC')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BF')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BG')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BH')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BI')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BL')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BM')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BN')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BO')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BR')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BS')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BT')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BU')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BX')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BY')->setVisible(false);
	            }else if($dayA <= 31){
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AS')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AV')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AW')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AX')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('AY')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BB')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BC')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BD')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BE')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BH')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BI')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BJ')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BK')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BN')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BO')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BP')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BQ')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BT')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BU')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BV')->setVisible(false);
	                $objPHPExcel->getActiveSheet()->getColumnDimension('BW')->setVisible(false);
	            }

                Style_group_Col($col_name, 1, $objPHPExcel, 1);
                Style_group_Col($col_name, 2, $objPHPExcel, 1);
                // $y = 8;
                // $num1 = $y + $dayA;
                // $ydiff = $lastmonth - $num1;
                // $ylast = $lastmonth + $ydiff;
                // //echo $num1."/".$ylast; exit;
                // for ($y1 = $num1; $y1 < $ylast; $y1++) {
                //     Style_group_Col($col_name, $y1, $objPHPExcel, 1);
                // }

                Style_Alignment('A6:A'.$count_data, 3, false, $objPHPExcel);
                Style_Alignment('D6:D'.$count_data, 9, false, $objPHPExcel);
                Style_Alignment('E6:E'.$count_data, 9, false, $objPHPExcel);
            //}         

        }
    //==========================================NO DATA CASE=============================================//
    } else {

                    $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "No data ".$til.".");
                    $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true,'Franklin Gothic Book')));
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

function Style_Fill($color=null) {

    return array( 'type'  => PHPExcel_Style_Fill::FILL_SOLID,                           
                  'color' => array('rgb' => $color)                                    
                );                                   
}

function Style_Font($size=11, $color='FFFFFF', $bol=false, $ita=false, $fname='Calibri Light') {

    return  array(
                    'name' => $fname,
                    'size' => $size,
                    'bold' => $bol,
                    'italic'=> $ita,
                    'color' => array('rgb' => $color)
                 );                               
}

function Style_border($line='BORDER_THICK', $color='000000')
{
    return array( 'style' => $line, 'color' => array('rgb' => $color)) ;
}

function rundata($strFileName){
    //$strFileName = "D:\DATA\uq.txt";
    $objFopen = fopen($strFileName, 'r');
    $tx="";
    if ($objFopen) {
        while (!feof($objFopen)) {
            $file = fgets($objFopen, 5000);
            if ($file <> ""){
            $tx .= $file;
            }
            
        }
    fclose($objFopen);
    }
    return $tx;
}

function Style_group_Col($cell=null, $index=0, $objPHPExcel=null, $level=1, $vi=false, $co=true)
{
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setOutlineLevel($level);
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setVisible($vi);
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setCollapsed($co); 
}
function Style_group_Row($index=0, $objPHPExcel=null, $vi=false, $co=true)
{
    $objPHPExcel->getActiveSheet()->getRowDimension ($index)->setOutlineLevel(1);
    $objPHPExcel->getActiveSheet()->getRowDimension ($index)->setVisible($vi);
    $objPHPExcel->getActiveSheet()->getRowDimension ($index)->setCollapsed($co); 
}

function Style_Alignment($cell='A1', $sty=1, $swt=false, $objPHPExcel= null)
{
    switch ($sty) {
        case 1: #bottom->center
                $objPHPExcel->getActiveSheet()
                    ->getStyle($cell)
                    ->getAlignment()
                    ->setWrapText($swt)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_BOTTOM)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            break;
        case 2: #top->center
                $objPHPExcel->getActiveSheet()
                    ->getStyle($cell)
                    ->getAlignment()
                    ->setWrapText($swt)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            break;
        case 3: #center->center
                $objPHPExcel->getActiveSheet()
                    ->getStyle($cell)
                    ->getAlignment()
                    ->setWrapText($swt)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            break;
        case 4: #bottom->right
                $objPHPExcel->getActiveSheet()
                    ->getStyle($cell)
                    ->getAlignment()
                    ->setWrapText($swt)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_BOTTOM)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                   // echo $sty; exit;
            break;
        case 5: #top->right
                $objPHPExcel->getActiveSheet()
                    ->getStyle($cell)
                    ->getAlignment()
                    ->setWrapText($swt)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            break;
        case 6: #center->right
                $objPHPExcel->getActiveSheet()
                    ->getStyle($cell)
                    ->getAlignment()
                    ->setWrapText($swt)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            break;
        case 7: #bottom->left
                $objPHPExcel->getActiveSheet()
                    ->getStyle($cell)
                    ->getAlignment()
                    ->setWrapText($swt)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_BOTTOM)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                   // echo $cell; exit;
            break;
        case 8: #top->left
                $objPHPExcel->getActiveSheet()
                    ->getStyle($cell)
                    ->getAlignment()
                    ->setWrapText($swt)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            break;
        case 9: #center->left
                $objPHPExcel->getActiveSheet()
                    ->getStyle($cell)
                    ->getAlignment()
                    ->setWrapText($swt)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            break;                                                                                           
        default:
            echo "No Style_Alignment type!!"."<hr>"; exit;
            break;
    }
}

function put_data($objPHPExcel, $dat, $cell, $row)
{

  $str = "=SUBTOTAL(109,";
  foreach ( $dat as $key => $value ) 
  {
    $str = "=SUBTOTAL(109,";
      foreach( $value as $ro => $val)
      {
        $str .= $cell.$val.",";
      }
             
    $objPHPExcel->getActiveSheet()->setCellValue($cell.($row+$key), substr($str, 0, strlen($str)-1) . ")" );       
  }
}


?>

 
