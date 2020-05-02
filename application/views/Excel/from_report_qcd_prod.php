<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
//date_default_timezone_set('Europe/London');
date_default_timezone_set("Asia/Bangkok");

$ct = array( array(), array(), array(), array(), array(), array(), array(), array(), array() ); 
$ct2 = array( array(), array());  
$cld = array( 'H' ,'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ','AK', 'AL'); 

if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

$dayA   = date('d');
$monthA = date('m');
$monthB = date('M');
$yearA  = date('Y');
$curdate = $dayA."-".$monthB."-".$yearA;

if($monthA == "12"){
	$nextmonth = 1;
	$nextyear = $yearA + 1;
}else{
	$nextmonth = $monthA + 1;
	$nextyear = $yearA;
}

if($nextmonth != "12" || $nextmonth != "11" || $nextmonth != "12"){
	$nextmonth = "0".$nextmonth;
}

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

if($nextmonth == "01"){$monthfull2 = "JANUARY"; $m = "Jan";}
else if($nextmonth == "02"){$monthfull2 = "FEBRUARY"; $m = "Feb";}
else if($nextmonth == "03"){$monthfull2 = "MARCH"; $m = "Mar";}
else if($nextmonth == "04"){$monthfull2 = "APRIL"; $m = "Apr";}
else if($nextmonth == "05"){$monthfull2 = "MAY"; $m = "May";}
else if($nextmonth == "06"){$monthfull2 = "JUNE"; $m = "Jun";}
else if($nextmonth == "07"){$monthfull2 = "JULY"; $m = "Jul";}
else if($nextmonth == "08"){$monthfull2 = "AUGUST"; $m = "Aug";}
else if($nextmonth == "09"){$monthfull2 = "SEPTEMBER"; $m = "Sep";}
else if($nextmonth == "10"){$monthfull2 = "OCTOBER"; $m = "Oct";}
else if($nextmonth == "11"){$monthfull2 = "NOVEMBER"; $m = "Nov";}
else if($nextmonth == "12"){$monthfull2 = "DECEMBER"; $m = "Dec";}

$showfdate = $monthfull." ".$yearA."  -  ".$monthfull2." ".$nextyear;
$curmonth = $monthfull." ".$yearA;
$curnextmonth = $monthfull2." ".$nextyear;
$text1 = "PROD. PLAN";
$text2 = "DELIVERY ORDER";
$text3 = "Diff ( PCS.)";
$text4 = "Diff ( % )";

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
//var_dump($title);
//exit();
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
        if ($til == 'QCD Production Report') { 
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
                    ->setWrapText(true)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);  

                foreach (range(7, 38) as $c)
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$c])->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('20');
                $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('24');
                $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('19');
                $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('23');
                $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('13');
                $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('35');
                $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('15');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setWidth('20');     
                $objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setWidth('20');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setWidth('16');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setWidth('16');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setWidth('20');     
                $objPHPExcel->getActiveSheet()->getColumnDimension('AS')->setWidth('20');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setWidth('16');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AU')->setWidth('16');
                                       
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('A5:'.$col_name[$count_index].'5');
                $objPHPExcel->getActiveSheet()->freezePane('A6');

                $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[$count_index].'4')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                $objPHPExcel->getActiveSheet()->getStyle('D6:D'.$count_data)->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                $objPHPExcel->getActiveSheet()->getStyle('E6:AU'.$count_data)->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'d9d9d9'))));
                $objPHPExcel->getActiveSheet()->getStyle('E6:AU'.$count_data)->applyFromArray(array('fill' => Style_Fill('FFFFFF')));
                $objPHPExcel->getActiveSheet()->getStyle('A23:AU23')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A50:AU50')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A62:AU62')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A64:AU64')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A66:AU66')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A69:AU69')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A70:AU70')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A71:AU71')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A77:AU77')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A81:AU81')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A102:AU102')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A103:AU103')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A107:AU107')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A109:AU109')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A110:AU110')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));

                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {
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
                    //echo $gcp; exit();
                        if( $value['PRODUCT_NO'] == 1 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row ); //array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 2 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 3 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 4 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 5 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 6 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 7 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 8 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 9 ) array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );

                    foreach ($value as $body => $val) 
                    {
                        if($body != 'NO' && $body != 'PRODUCT_NO'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);                
                        }

                        if ($body == 'CUSTOMER_NAME'){
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(14,'FFFFFF',true,false,'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->applyFromArray(array('fill' => Style_Fill('c0504d')));

                            if($val != $ccp)
                            {
                                if( ($row - $st) > 0 )
                                $objPHPExcel->getActiveSheet()->getStyle('A' . $st . ':' . 'A' . ($row-1))->applyFromArray(array('fill' => Style_Fill('c0504d')));  
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
                            $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->applyFromArray(array('fill' => Style_Fill('f2f2f2')));

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
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        }   

                        if (substr($body,0,1) == 'D') {
                            if (substr($body,1,2) == $dayA) {
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].$row)
                                                ->applyFromArray(array('font' => Style_Font(11,'002060',true,false,'Calibri')));
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->applyFromArray(array('fill' => Style_Fill('ffffcc'))); 
                            }else{
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].$row)
                                                //->applyFromArray(array('font' => Style_Font(11,'ff0000',false,false,'Calibri')));
                                                ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            }
                        } 

                        if ($body == 'STOCK' || $body == 'FC_PROD_1' || $body == 'FC_PROD_2' || $body == 'FC_ORD_1' || $body == 'FC_ORD_2' || $body == 'MODEL_REF') {
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                            if ($body == 'FC_PROD_1' || $body == 'FC_PROD_2') {
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->applyFromArray(array('fill' => Style_Fill('d3ffff'))); 
                            } 
                            if ($body == 'FC_ORD_1' || $body == 'FC_ORD_2') {
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->applyFromArray(array('fill' => Style_Fill('e1eaf3'))); 
                            } 
                        } 

                        if($body == 'ACCUM'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=SUM(H'.($row).':AL'.($row).')' );
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        }

                        if($body == 'DIFF_PCS1'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=IFERROR(AN'.($row).'- AO'.($row).', 0 )' );
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        }

                        if($body == 'DIFF_PER1'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=IFERROR(AP'.($row).'/AN'.($row).', 0 )' );
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0.0%_-;[RED](#,##0.0%)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'166403',true,false,'Calibri')));
                        }

                        if($body == 'DIFF_PCS2'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=IFERROR(AR'.($row).'- AS'.($row).', 0 )' );
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        }

                        if($body == 'DIFF_PER2'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=IFERROR(AT'.($row).'/AR'.($row).', 0 )' );
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0.0%_-;[RED](#,##0.0%)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'166403',true,false,'Calibri')));
                        }

                    }
                    $row++; 
                }

                //====================================SUMMARY CUSTOMER DEMAND BY PRODUCTS CATEGORY====================================//
                foreach(array('H','G','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU' ) as $cel )
                    put_data($objPHPExcel, $ct, $cel, ($count_data+4));
                // foreach(array('AN') as $cel )
                //     put_data($objPHPExcel, $ct, $cel, ($count_data+4));

                for($i = ($count_data+4); $i < ($count_data+13); $i++){
                    $temp = 'AP'.$i;
                    $objPHPExcel->getActiveSheet()->setCellValue($temp, '=IFERROR(AN'.($i).'- AO'.($i).', 0 )' );
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));                
                }

                for($i = ($count_data+4); $i < ($count_data+13); $i++){
                    $temp = 'AQ'.$i;
                    $objPHPExcel->getActiveSheet()->setCellValue($temp, '=IFERROR(AP'.($i).'/AN'.($i).', 0 )' );
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0.0%_-;[RED](#,##0.0%)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'166403',true,false,'Calibri'))); 
                    $objPHPExcel->getActiveSheet()->getStyle( $temp )
                                                              ->applyFromArray(array(
                                                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                
                }

                for($i = ($count_data+4); $i < ($count_data+13); $i++){
                    $temp = 'AT'.$i;
                    $objPHPExcel->getActiveSheet()->setCellValue($temp, '=IFERROR(AR'.($i).'- AS'.($i).', 0 )' );
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));                
                }

                for($i = ($count_data+4); $i < ($count_data+13); $i++){
                    $temp = 'AU'.$i;
                    $objPHPExcel->getActiveSheet()->setCellValue($temp, '=IFERROR(AT'.($i).'/AR'.($i).', 0 )' );
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0.0%_-;[RED](#,##0.0%)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'166403',true,false,'Calibri')));
                    $objPHPExcel->getActiveSheet()->getStyle( $temp )
                                                              ->applyFromArray(array(
                                                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                 
                }

                for($i = ($count_data+4); $i < ($count_data+13); $i++){
                    $temp = 'A'.$i.':AP'.$i;
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                    $objPHPExcel->getActiveSheet()->getStyle( $temp )
                                                              ->applyFromArray(array(
                                                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                      
                }

                for($i = ($count_data+4); $i < ($count_data+13); $i++){
                    $temp = 'AR'.$i.':AT'.$i;
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                    $objPHPExcel->getActiveSheet()->getStyle( $temp )
                                                              ->applyFromArray(array(
                                                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                      
                }

                for($i = ($count_data+4); $i < ($count_data+16); $i++){
                    $temp = 'A'.$i;
                    if($i < ($count_data+13)){
                        $objPHPExcel->setActiveSheetIndex()->setCellValue($temp, "c" );
                        $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Wingdings 3')));
                        Style_Alignment($temp, 3, false, $objPHPExcel); 
                    }else if($i == ($count_data+15)){
                        $objPHPExcel->setActiveSheetIndex()->setCellValue($temp, "ISSUED BY PC SYSTEM ON ".$dayA."-".$monthB."-".$yearA);
                        $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(12,'000000',true,true,'Calibri')));
                        Style_Alignment($temp, 9, false, $objPHPExcel); 
                    }
                }

                for($i = ($count_data+4); $i < ($count_data+13); $i++){
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
                    }else if($i == ($count_data+12)){
                        $type = "GKN";
                    }
                    $objPHPExcel->setActiveSheetIndex()->setCellValue($temp, $type);
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(14,'000000',true,true,'Calibri')));
                    $objPHPExcel->getActiveSheet()->mergeCells($temp2);
                }

                $objPHPExcel->getActiveSheet()->getStyle('A'.($count_data+3))->applyFromArray(array('font' => Style_Font(14,'FFFFFF',true,true,'Calibri'))); 
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A'.($count_data+3), "SUMMARY CUSTOMER DEMAND BY PRODUCTS CATEGORY");
                $objPHPExcel->getActiveSheet()->mergeCells('A'.($count_data+3).':'.'AU'.($count_data+3));
                $objPHPExcel->getActiveSheet()->mergeCells('A'.($count_data+13).':'.'AU'.($count_data+13));
                $objPHPExcel->getActiveSheet()->getStyle('A'.($count_data+3))->applyFromArray(array('fill' => Style_Fill('974706')));
                $objPHPExcel->getActiveSheet()->getStyle('A'.($count_data+13))->applyFromArray(array('fill' => Style_Fill('974706')));
                $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('fill' => Style_Fill('c0504d'))); //PINK COLOR
                $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(22,'FFFFFF',true,true,'Calibri Light')));  
                $objPHPExcel->getActiveSheet()->getStyle('A3:G4')->applyFromArray(array('fill' => Style_Fill('f2f2f2'))); //GRAY COLOR
                $objPHPExcel->getActiveSheet()->getStyle('A3:G4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('H1:AM4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('H1:AM2')->applyFromArray(array('fill' => Style_Fill('e26b0a'))); //ORANGE COLOR
                $objPHPExcel->getActiveSheet()->getStyle('H3:AM4')->applyFromArray(array('fill' => Style_Fill('fde9d9'))); //LIGHT ORANGE COLOR
                $objPHPExcel->getActiveSheet()->getStyle('H1:AM2')->applyFromArray(array('font' => Style_Font(16,'FFFFFF',true,true,'Calibri Light'))); 
                $objPHPExcel->getActiveSheet()->getStyle('H3:AM3')->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('H4:AM4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AN1:AU2')->applyFromArray(array('fill' => Style_Fill('4f81bd'))); //BLUE PURPLE COLOR
                $objPHPExcel->getActiveSheet()->getStyle('AN3:AU3')->applyFromArray(array('fill' => Style_Fill('4bacc6'))); //BLUE PURPLE LIGHT COLOR
                $objPHPExcel->getActiveSheet()->getStyle('AN4:AU4')->applyFromArray(array('fill' => Style_Fill('daeef3'))); //BLUE LIGHT COLOR
                $objPHPExcel->getActiveSheet()->getStyle('AN1:AU2')->applyFromArray(array('font' => Style_Font(16,'FFFFFF',true,true,'Calibri Light'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AN3:AU3')->applyFromArray(array('font' => Style_Font(14,'FFFFFF',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AN4:AU4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 

                //==============================================TITLE====================================================//
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "QCD PRODUCTION DAILY REPORT"." ( ".$curdate." ) ");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A3', "CUSTOMERS");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('B3', "CUSTOMERS FILTER");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('C3', "GROUP FILTER");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('D3', "GROUP PART");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('E3', "MODEL");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('F3', "REF");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('G3', "EXP/JA");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('G4', "STOCK (QTY)");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('H1', "DAILY PRODUCTION PLAN");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('H2', $curmonth);

                foreach(array('H3','I3', 'J3' ,'K3','L3','M3','N3','O3','P3','Q3','R3','S3','T3','U3','V3','W3','X3','Y3','Z3','AA3','AB3','AC3','AD3','AE3','AF3','AG3','AH3','AI3','AJ3','AK3','AL3','AM3') as $cel ) 
                    $objPHPExcel->setActiveSheetIndex($ind)->setCellValue( $cel , "[ PCS. ]");

                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AN1', "DEMAND FORECAST");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AN2', $showfdate);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('H4', "01st");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('I4', "02nd");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('J4', "03nd");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('K4', "04th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('L4', "05th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('M4', "06th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('N4', "07th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('O4', "08th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('P4', "09th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Q4', "10th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('R4', "11th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('S4', "12th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('T4', "13th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('U4', "14th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('V4', "15th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('W4', "16th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('X4', "17th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Y4', "18th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Z4', "19th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AA4', "20th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AB4', "21th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AC4', "22th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AD4', "23th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AE4', "24th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AF4', "25th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AG4', "26th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AH4', "27th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AI4', "28th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AJ4', "29th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AK4', "30th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AL4', "31th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AM4', "ACCUM");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AN3', $curmonth);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AR3', $curnextmonth);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AN4', $text1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AO4', $text2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AP4', $text3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AQ4', $text4);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AR4', $text1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AS4', $text2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AT4', $text3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AU4', $text4);

                //==============================================TITLE====================================================//

                $objPHPExcel->getActiveSheet()->mergeCells('A1:'.'G2');
                $objPHPExcel->getActiveSheet()->mergeCells('H1:'.'AM1');
                $objPHPExcel->getActiveSheet()->mergeCells('H2:'.'AM2');
                $objPHPExcel->getActiveSheet()->mergeCells('AN1:'.'AU1');
                $objPHPExcel->getActiveSheet()->mergeCells('AN2:'.'AU2');
                $objPHPExcel->getActiveSheet()->mergeCells('AN3:'.'AQ3');
                $objPHPExcel->getActiveSheet()->mergeCells('AR3:'.'AU3');
                //
                $objPHPExcel->getActiveSheet()->mergeCells('A3:'.'A4');
                $objPHPExcel->getActiveSheet()->mergeCells('B3:'.'B4');
                $objPHPExcel->getActiveSheet()->mergeCells('C3:'.'C4');
                $objPHPExcel->getActiveSheet()->mergeCells('D3:'.'D4');
                $objPHPExcel->getActiveSheet()->mergeCells('E3:'.'E4');
                $objPHPExcel->getActiveSheet()->mergeCells('F3:'.'F4');

                $strcol = $dayA+0;
                $hidecol = 30;
                //echo $strcol. $hidecol; exit;
                for($i = $strcol; $i <= $hidecol; $i++){
                    //echo $cld[$i]; 
                    $objPHPExcel->getActiveSheet()->getColumnDimension($cld[$i])->setVisible(false);
                }

                $x = 7;
                $num = $x + $dayA;
                for ($x = 7; $x < $num-1; $x++) {
                    Style_group_Col($col_name, $x, $objPHPExcel, 1);
                }

                Style_group_Col($col_name, 1, $objPHPExcel, 1);
                Style_group_Col($col_name, 2, $objPHPExcel, 1);

                Style_Alignment('A6:A'.$count_data, 3, false, $objPHPExcel);
                Style_Alignment('D6:D'.$count_data, 9, false, $objPHPExcel);
                Style_Alignment('E6:E'.$count_data, 9, false, $objPHPExcel);
            //}         

        } elseif ($till = "PD6 QCD Production Report") {
            $objPHPExcel->setActiveSheetIndex(1);
            $i   = 0;

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
                    ->setWrapText(true)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);  

            foreach (range(7, 38) as $c)
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$c])->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('20');
                $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('24');
                $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('24');
                $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('24');
                $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('34');
                $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('26');
                $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('15');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setWidth('20');     
                $objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setWidth('20');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setWidth('16');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setWidth('16');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setWidth('20');     
                $objPHPExcel->getActiveSheet()->getColumnDimension('AS')->setWidth('20');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setWidth('16');
                $objPHPExcel->getActiveSheet()->getColumnDimension('AU')->setWidth('16');
                                       
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('A5:'.'AU5');
                $objPHPExcel->getActiveSheet()->freezePane('A6');

                $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[$count_index].'4')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                $objPHPExcel->getActiveSheet()->getStyle('D6:D'.$count_data)->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                $objPHPExcel->getActiveSheet()->getStyle('E6:AU'.$count_data)->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'d9d9d9'))));
                $objPHPExcel->getActiveSheet()->getStyle('E6:AU'.$count_data)->applyFromArray(array('fill' => Style_Fill('FFFFFF')));
                $objPHPExcel->getActiveSheet()->getStyle('A22:AU22')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));
                $objPHPExcel->getActiveSheet()->getStyle('A55:AU55')->applyFromArray(array('borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'808080'))));

                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {
                    if($key != 'NO' && $key != 'PRODUCT_NO' && $key != 'CUSTOMER_NAME'){
                        if( $key == 'CUSTOMER_NAME2' || $key == 'PRODUCT_GROUP' || $key == 'PRODUCT_GROUP2' || $key == 'ITEM_CD' || $key == 'ITEM_NAME' || $key == 'MODEL'){
                            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."3", str_replace("_", " ", strtoupper($key)));       
                        }else{
                            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."4", str_replace("_", " ", strtoupper($key)));  
                        }
                    }
                }  

                $row = 6;
                // $st=$row;
                // $gt=$row;
                // $ccp = $list_act_report[$sheetIndex][0]['CUSTOMER_NAME2'];
                // $gcp = $list_act_report[$sheetIndex][0]['PRODUCT_GROUP2'];
                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {
                    $col = 0;
                    //echo $body; exit;
                    //echo $gcp; exit();
                        if( $value['PRODUCT_NO'] == 1 ) array_push( $ct2[ ($value['PRODUCT_NO']-1) ], $row ); //array_push( $ct[ ($value['PRODUCT_NO']-1) ], $row );
                        elseif( $value['PRODUCT_NO'] == 2 ) array_push( $ct2[ ($value['PRODUCT_NO']-1) ], $row );

                    foreach ($value as $body => $val) 
                    {
                        if($body != 'NO' && $body != 'PRODUCT_NO' && $body != 'CUSTOMER_NAME'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);                
                        }

                        // if ($body == 'CUSTOMER_NAME2'){
                        //     $objPHPExcel->getActiveSheet()
                        //                     ->getStyle($col_name[$col-1].$row)
                        //                     ->applyFromArray(array('font' => Style_Font(14,'FFFFFF',true,false,'Calibri')));
                        //     $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->applyFromArray(array('fill' => Style_Fill('c0504d')));

                        //     if($val != $ccp)
                        //     {
                        //         if( ($row - $st) > 0 )
                        //         $objPHPExcel->getActiveSheet()->getStyle('A' . $st . ':' . 'A' . ($row-1))->applyFromArray(array('fill' => Style_Fill('c0504d')));  
                        //         $objPHPExcel->getActiveSheet()->mergeCells( 'A' . $st . ':' . 'A' . ($row-1) );
                        //         //$objPHPExcel->getActiveSheet()->mergeCells( 'C' . $st . ':' . 'C' . ($row-1) );
                        //         $objPHPExcel->getActiveSheet()->getStyle( 'A'. ($row-1) )
                        //                                       ->applyFromArray(array(
                        //                                         'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                        //         $ccp = $val;
                        //         $st  = $row;
                        //     }
                        // }   

                        // if($body == 'PRODUCT_GROUP2')
                        // {
                        //     $objPHPExcel->getActiveSheet()
                        //                     ->getStyle($col_name[$col-1].$row)
                        //                     ->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri')));
                        //     $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->applyFromArray(array('fill' => Style_Fill('f2f2f2')));

                        //     if($val != $gcp)
                        //     {
                        //         if( ($row - $gt) > 0 )
                        //             $objPHPExcel->getActiveSheet()->mergeCells( 'C' . $gt . ':' . 'C' . ($row-1) );
                        //             // $objPHPExcel->getActiveSheet()->getStyle( 'D'. ($row-1) )
                        //          //                              ->applyFromArray(array(
                        //          //                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                        //             $gcp = $val;
                        //             $gt  = $row;
                        //     }
                        // }

                        if($body == 'PRODUCT_GROUP2')
                        {
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->applyFromArray(array('fill' => Style_Fill('f2f2f2')));
                        }

                        if ($body == 'ITEM_CD'){
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        } 

                        if ($body == 'ITEM_NAME'){
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        } 

                        if ($body == 'MODEL'){
                            if ($value['MODEL'] == '3E00') 
                            {
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->getNumberFormat()->setFormatCode('###"E00"');
                                              $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), $value['MODEL']);
                            }
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        }   

                        if (substr($body,0,1) == 'D') {
                            if (substr($body,1,2) == $dayA) {
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].$row)
                                                ->applyFromArray(array('font' => Style_Font(11,'002060',true,false,'Calibri')));
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->applyFromArray(array('fill' => Style_Fill('ffffcc'))); 
                            }else{
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].$row)
                                                //->applyFromArray(array('font' => Style_Font(11,'ff0000',false,false,'Calibri')));
                                                ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                                $objPHPExcel->getActiveSheet()
                                                ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            }
                        } 

                        if ($body == 'STOCK' || $body == 'FC_PROD_1' || $body == 'FC_PROD_2' || $body == 'FC_ORD_1' || $body == 'FC_ORD_2') {
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                            if ($body == 'FC_PROD_1' || $body == 'FC_PROD_2') {
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->applyFromArray(array('fill' => Style_Fill('d3ffff'))); 
                            } 
                            if ($body == 'FC_ORD_1' || $body == 'FC_ORD_2') {
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->applyFromArray(array('fill' => Style_Fill('e1eaf3'))); 
                            } 
                        } 

                        if($body == 'ACCUM'){
                            // $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=SUM(H'.($row).':AL'.($row).')' );
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=SUM(H'.($row).':AL'.($row).')' );
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        }

                        if($body == 'DIFF_PCS1'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=IFERROR(AN'.($row).'- AO'.($row).', 0 )' );
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        }

                        if($body == 'DIFF_PER1'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=IFERROR(AP'.($row).'/AN'.($row).', 0 )' );
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0.0%_-;[RED](#,##0.0%)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'166403',true,false,'Calibri')));
                        }

                        if($body == 'DIFF_PCS2'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=IFERROR(AR'.($row).'- AS'.($row).', 0 )' );
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                        }

                        if($body == 'DIFF_PER2'){
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), '=IFERROR(AT'.($row).'/AR'.($row).', 0 )' );
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0.0%_-;[RED](#,##0.0%)_-;_-* "-"??_-;_-@_-');
                            $objPHPExcel->getActiveSheet()
                                            ->getStyle($col_name[$col-1].$row)
                                            ->applyFromArray(array('font' => Style_Font(11,'166403',true,false,'Calibri')));
                        }

                    }
                    $row++; 
                }

            //====================================SUMMARY CUSTOMER DEMAND BY PRODUCTS CATEGORY====================================//
                foreach(array('H','G','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU' ) as $cel )
                    put_data($objPHPExcel, $ct2, $cel, ($count_data+4));
                // foreach(array('AN') as $cel )
                //     put_data($objPHPExcel, $ct, $cel, ($count_data+4));

                $p1 = ($count_data+4);
                $p2 = ($count_data+5);

                for($i = ($count_data+6); $i < ($count_data+7); $i++){
                    foreach(array('G','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AP','AQ','AR','AS','AT','AU') as $cel )
                    $objPHPExcel->getActiveSheet()->setCellValue($cel.$i, '=SUM('.$cel.$p1.':'.$cel.$p2.')' );                 
                }

                for($i = ($count_data+6); $i < ($count_data+7); $i++){
                    $objPHPExcel->getActiveSheet()->setCellValue('H'.$i, '=SUM('.'H'.$p1.':'.'H'.$p2.')' );  
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle('H'.$i)->getNumberFormat()->setFormatCode('_-* #,##0.0_-;[RED](#,##0.0)_-;_-* "-"??_-;_-@_-');                
                }

                for($i = ($count_data+6); $i < ($count_data+7); $i++){
                    $objPHPExcel->getActiveSheet()->setCellValue('AO'.$i, '=SUM('.'AO'.$p1.':'.'AO'.$p2.')' );  
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle('AO'.$i)->getNumberFormat()->setFormatCode('_-* #,##0.0%_-;[RED](#,##0.0%)_-;_-* "-"??_-;_-@_-');                
                }

                for($i = ($count_data+4); $i < ($count_data+7); $i++){
                    $temp = 'AP'.$i;
                    $objPHPExcel->getActiveSheet()->setCellValue($temp, '=IFERROR(AN'.($i).'- AO'.($i).', 0 )' );
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));                
                }

                for($i = ($count_data+4); $i < ($count_data+7); $i++){
                    $temp = 'AQ'.$i;
                    $objPHPExcel->getActiveSheet()->setCellValue($temp, '=IFERROR(AP'.($i).'/AN'.($i).', 0 )' );
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0.0%_-;[RED](#,##0.0%)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'166403',true,false,'Calibri'))); 
                    $objPHPExcel->getActiveSheet()->getStyle( $temp )
                                                              ->applyFromArray(array(
                                                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                
                }

                for($i = ($count_data+4); $i < ($count_data+7); $i++){
                    $temp = 'AT'.$i;
                    $objPHPExcel->getActiveSheet()->setCellValue($temp, '=IFERROR(AR'.($i).'- AS'.($i).', 0 )' );
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));                
                }

                for($i = ($count_data+4); $i < ($count_data+7); $i++){
                    $temp = 'AU'.$i;
                    $objPHPExcel->getActiveSheet()->setCellValue($temp, '=IFERROR(AT'.($i).'/AR'.($i).', 0 )' );
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0.0%_-;[RED](#,##0.0%)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'166403',true,false,'Calibri')));
                    $objPHPExcel->getActiveSheet()->getStyle( $temp )
                                                              ->applyFromArray(array(
                                                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                 
                }

                for($i = ($count_data+4); $i < ($count_data+7); $i++){
                    $temp = 'A'.$i.':AP'.$i;
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                    $objPHPExcel->getActiveSheet()->getStyle( $temp )
                                                              ->applyFromArray(array(
                                                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                      
                }

                for($i = ($count_data+4); $i < ($count_data+7); $i++){
                    $temp = 'AR'.$i.':AT'.$i;
                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($temp)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); 
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri')));
                    $objPHPExcel->getActiveSheet()->getStyle( $temp )
                                                              ->applyFromArray(array(
                                                               'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                      
                }

                for($i = ($count_data+4); $i < ($count_data+10); $i++){
                    $temp = 'A'.$i;
                    if($i < ($count_data+7)){
                        $objPHPExcel->getActiveSheet()->setCellValue($temp, "c" );
                        $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Wingdings 3')));
                        Style_Alignment($temp, 3, false, $objPHPExcel); 
                    }else if($i == ($count_data+9)){
                        $objPHPExcel->getActiveSheet()->setCellValue($temp, "ISSUED BY PC SYSTEM ON ".$dayA."-".$monthB."-".$yearA);
                        $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(12,'000000',true,true,'Calibri')));
                        Style_Alignment($temp, 9, false, $objPHPExcel); 
                    }
                }

                //$objPHPExcel->getActiveSheet()->setCellValue('A'.($count_data+10), "c" );
                for($i = ($count_data+4); $i < ($count_data+7); $i++){
                    $temp = 'B'.$i;
                    $temp2 = 'B'.$i.':F'.$i;
                    if($i == ($count_data+4)){
                        $type = "AIR TYPE";
                    }else if($i == ($count_data+5)){
                        $type = "WATER TYPE";
                    }else if($i == ($count_data+6)){
                        $type = "TOTAL";
                    }
                    $objPHPExcel->getActiveSheet()->setCellValue($temp, $type);
                    $objPHPExcel->getActiveSheet()->getStyle($temp)->applyFromArray(array('font' => Style_Font(14,'000000',true,true,'Calibri')));
                    $objPHPExcel->getActiveSheet()->mergeCells($temp2);
                }

                $objPHPExcel->getActiveSheet()->getStyle('A'.($count_data+3))->applyFromArray(array('font' => Style_Font(14,'FFFFFF',true,true,'Calibri')));
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A'.($count_data+3), "SUMMARY CUSTOMER DEMAND BY PRODUCTS CATEGORY");
                $objPHPExcel->getActiveSheet()->mergeCells('A'.($count_data+3).':'.'AU'.($count_data+3));
                $objPHPExcel->getActiveSheet()->mergeCells('A'.($count_data+7).':'.'AU'.($count_data+7));
                $objPHPExcel->getActiveSheet()->getStyle('A'.($count_data+3))->applyFromArray(array('fill' => Style_Fill('974706')));
                $objPHPExcel->getActiveSheet()->getStyle('A'.($count_data+7))->applyFromArray(array('fill' => Style_Fill('974706')));
                $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('fill' => Style_Fill('76933c'))); //Green COLOR
                $objPHPExcel->getActiveSheet()->getStyle('A6')->applyFromArray(array('fill' => Style_Fill('76933c'))); //Green COLOR
                $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(22,'FFFFFF',true,true,'Calibri Light')));  
                $objPHPExcel->getActiveSheet()->getStyle('A6')->applyFromArray(array('font' => Style_Font(18,'FFFFFF',true,false,'Calibri Light'))); 
                $objPHPExcel->getActiveSheet()->getStyle('A3:G4')->applyFromArray(array('fill' => Style_Fill('f2f2f2'))); //GRAY COLOR
                $objPHPExcel->getActiveSheet()->getStyle('A3:G4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('H1:AM4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('H1:AM2')->applyFromArray(array('fill' => Style_Fill('e26b0a'))); //ORANGE COLOR
                $objPHPExcel->getActiveSheet()->getStyle('H3:AM4')->applyFromArray(array('fill' => Style_Fill('fde9d9'))); //LIGHT ORANGE COLOR
                $objPHPExcel->getActiveSheet()->getStyle('H1:AM2')->applyFromArray(array('font' => Style_Font(16,'FFFFFF',true,true,'Calibri Light'))); 
                $objPHPExcel->getActiveSheet()->getStyle('H3:AM3')->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('H4:AM4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AN1:AU2')->applyFromArray(array('fill' => Style_Fill('4f81bd'))); //BLUE PURPLE COLOR
                $objPHPExcel->getActiveSheet()->getStyle('AN3:AU3')->applyFromArray(array('fill' => Style_Fill('4bacc6'))); //BLUE PURPLE LIGHT COLOR
                $objPHPExcel->getActiveSheet()->getStyle('AN4:AU4')->applyFromArray(array('fill' => Style_Fill('daeef3'))); //BLUE LIGHT COLOR
                $objPHPExcel->getActiveSheet()->getStyle('AN1:AU2')->applyFromArray(array('font' => Style_Font(16,'FFFFFF',true,true,'Calibri Light'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AN3:AU3')->applyFromArray(array('font' => Style_Font(14,'FFFFFF',true,false,'Calibri'))); 
                $objPHPExcel->getActiveSheet()->getStyle('AN4:AU4')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Calibri'))); 

                //==============================================TITLE====================================================//
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "PD6 QCD PRODUCTION DAILY REPORT"." ( ".$curdate." ) ");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A3', "CUSTOMERS");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A6', "MTA");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('B3', "GROUP FILTER");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('C3', "GROUP PART");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('D3', "ITEM CODE");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('E3', "ITEM NAME");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('F3', "MODEL");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('G3', "EXP/JA");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('G4', "STOCK (QTY)");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('H1', "DAILY PRODUCTION PLAN");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('H2', $curmonth);

                foreach(array('H3','I3', 'J3' ,'K3','L3','M3','N3','O3','P3','Q3','R3','S3','T3','U3','V3','W3','X3','Y3','Z3','AA3','AB3','AC3','AD3','AE3','AF3','AG3','AH3','AI3','AJ3','AK3','AL3','AM3') as $cel ) 
                    $objPHPExcel->setActiveSheetIndex($ind)->setCellValue( $cel , "[ PCS. ]");

                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AN1', "DEMAND FORECAST");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AN2', $showfdate);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('H4', "01st");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('I4', "02nd");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('J4', "03nd");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('K4', "04th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('L4', "05th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('M4', "06th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('N4', "07th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('O4', "08th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('P4', "09th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Q4', "10th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('R4', "11th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('S4', "12th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('T4', "13th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('U4', "14th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('V4', "15th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('W4', "16th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('X4', "17th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Y4', "18th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Z4', "19th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AA4', "20th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AB4', "21th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AC4', "22th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AD4', "23th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AE4', "24th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AF4', "25th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AG4', "26th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AH4', "27th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AI4', "28th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AJ4', "29th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AK4', "30th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AL4', "31th");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AM4', "ACCUM");
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AN3', $curmonth);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AR3', $curnextmonth);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AN4', $text1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AO4', $text2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AP4', $text3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AQ4', $text4);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AR4', $text1);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AS4', $text2);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AT4', $text3);
                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('AU4', $text4);

                //==============================================TITLE====================================================//

                $objPHPExcel->getActiveSheet()->mergeCells('A1:'.'G2');
                $objPHPExcel->getActiveSheet()->mergeCells('A6:'.'A55');
                $objPHPExcel->getActiveSheet()->mergeCells('C6:'.'C22');
                $objPHPExcel->getActiveSheet()->mergeCells('C23:'.'C55');
                $objPHPExcel->getActiveSheet()->mergeCells('H1:'.'AM1');
                $objPHPExcel->getActiveSheet()->mergeCells('H2:'.'AM2');
                $objPHPExcel->getActiveSheet()->mergeCells('AN1:'.'AU1');
                $objPHPExcel->getActiveSheet()->mergeCells('AN2:'.'AU2');
                $objPHPExcel->getActiveSheet()->mergeCells('AN3:'.'AQ3');
                $objPHPExcel->getActiveSheet()->mergeCells('AR3:'.'AU3');
                //
                $objPHPExcel->getActiveSheet()->mergeCells('A3:'.'A4');
                $objPHPExcel->getActiveSheet()->mergeCells('B3:'.'B4');
                $objPHPExcel->getActiveSheet()->mergeCells('C3:'.'C4');
                $objPHPExcel->getActiveSheet()->mergeCells('D3:'.'D4');
                $objPHPExcel->getActiveSheet()->mergeCells('E3:'.'E4');
                $objPHPExcel->getActiveSheet()->mergeCells('F3:'.'F4');

                $strcol = $dayA+0;
                $hidecol = 30;
                //echo $strcol. $hidecol; exit;
                for($i = $strcol; $i <= $hidecol; $i++){
                    //echo $cld[$i]; 
                    $objPHPExcel->getActiveSheet()->getColumnDimension($cld[$i])->setVisible(false);
                }

                $x = 7;
                $num = $x + $dayA;
                for ($x = 7; $x < $num-1; $x++) {
                    Style_group_Col($col_name, $x, $objPHPExcel, 1);
                }

                Style_group_Col($col_name, 1, $objPHPExcel, 1);
                //Style_group_Col($col_name, 2, $objPHPExcel, 1);

                Style_Alignment('A6:A'.$count_data, 3, false, $objPHPExcel);
                Style_Alignment('C6:C'.$count_data, 9, false, $objPHPExcel);
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

$objPHPExcel->setActiveSheetIndex(1)->insertNewColumnBefore('H', 1);
$objPHPExcel->setActiveSheetIndex(1)->getColumnDimension('H')->setVisible(false);

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex()->insertNewColumnBefore('H', 1);
$objPHPExcel->setActiveSheetIndex()->getColumnDimension('H')->setVisible(false);
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

 
