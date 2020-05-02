<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');
ini_set('max_execution_time', 300); 
ini_set('memory_limit','2048M');
if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/PHPExcel-1.8.1/Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();
$data_col = array();
//var_dump($list_act_report); exit;
$col_name = array();
foreach ( range('B', 'Z') as $cm ) { array_push($col_name, $cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "A".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "B".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "C".$cm); }

$i   = 0;   
$ind = 0;
$T_lastM = ((date('m')-1) > 12 ) ? date('My', strtotime( (date('Y')-1). "-" ."12". "-" . '01' ) ) : date('My', strtotime( (date('Y')+0). "-" .(date('m')-1). "-" . '01' ) ) ;// exit;
$H_lastM = ((date('m')-1) > 12 ) ? date('F Y', strtotime( (date('Y')-1). "-" ."12". "-" . '01' ) ) : date('F Y', strtotime( (date('Y')+0). "-" .(date('m')-1). "-" . '01' ) ) ;// exit;
//$T_lastM = date('My',  strtotime( date('Y'). "-" .(date('m')-1). "-" . 1 ) ) ;// exit;
//$H_lastM = date('F Y', strtotime( date('Y'). "-" .(date('m')-1). "-" . 1 ) ) ;// exit;

$ex_usd = $rate[0]['CURRENCY_RATE'];
$ex_eur = $rate[1]['CURRENCY_RATE'];
$ex_jpy = $rate[2]['CURRENCY_RATE'];
// echo $ex_usd; 
// echo "<hr>";
// echo $ex_eur; 
// echo "<hr>";
// echo $ex_jpy; 
// echo "<hr>";

// exit;
foreach ($title as $inTil => $til) 
{
             $objPHPExcel->createSheet();
             $objPHPExcel->setActiveSheetIndex($ind);
             

            $sheetIndex  =  strtolower(str_replace(' ', '_', $title[$ind])); 
            $count_index = 0;
            $count_data  =  count($list_act_report[$sheetIndex]) + 5;
    if ($count_data > 0) 
    {      
#========================================================================================================================  Put field ====================================================================================        
            if( $sheetIndex == 'receive_monthly' ) 
            {
                $objPHPExcel->getActiveSheet()->setTitle( "$til of ". $T_lastM  );
                $objPHPExcel->getActiveSheet()->setShowGridlines(False);
                $st_col = 9;
                $st_dat = 11;
                $count_index =  count($list_act_report[$sheetIndex][0]) - 1 ;
                $row = $st_dat;
                $look_data = 0;
                $count_data  =  count($list_act_report[$sheetIndex]) + $row-1;
                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 10 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 10 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 8 )->setRowHeight( 10 );                
                $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 20 );
                foreach (range(4,7) as $id) {  $objPHPExcel->getActiveSheet()->getRowDimension( $id )->setRowHeight( 30 ); }
                
                $objPHPExcel->getActiveSheet()->getRowDimension( 10 )->setRowHeight( 10 );  

                $objPHPExcel->getActiveSheet()->freezePane('A'.$row);   
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(91);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('B'.($st_col+1).':'.'L'.($st_col+1));                


                $objPHPExcel->getActiveSheet()->getStyle('B2:'.$col_name[$count_index].($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000023')))); 

                //echo $row; exit;                   
                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++].$st_col, str_replace("_", " ", $key));
                } // exit;     
#========================================================================================================================  Put data ====================================================================================                

                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {               
                   $col = 0;
                    foreach ($value as $body => $val) 
                    {
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);

                                if($val == 3 && $body == 'MODEL')  $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('###"E00"');
                                if($val == "" && $body == 'PD' && $look_data === 0)
                                    $look_data = $row;
                            if ( ($body == 'PRICE_AMOUNT' && $value[$body] > 0) || ( $body == 'PLAN_AMOUNT' && $value[$body] > 0) ) 
                            {
                                    if ($value['SOURCE_CD'] == 'T00100' || $value['SOURCE_CD'] == 'M50050' || $value['SOURCE_CD'] == 'T10400' ) {

                                       // $amt = $val * $rate[2]['CURRENCY_RATE'];
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)
                                                                                                 ->applyFromArray(array('font' => Style_Font(11, '009900', false, 'Consolas')));                                                   
                                        $objPHPExcel->setActiveSheetIndex($ind)->setCellValue( $col_name[$col-1].($row), "=".$val ."*".'$C'.($count_data+6));
                                    }

                                    elseif ($value['SOURCE_CD'] == 'M50030' || $value['SOURCE_CD'] == 'M50040' || $value['SOURCE_CD'] == 'M5004A' || $value['SOURCE_CD'] == 'M5004B' || $value['SOURCE_CD'] == 'M5004C' || $value['SOURCE_CD'] == 'T10200' || $value['SOURCE_CD'] == 'T10300'){
                                        //$amt = $val * $rate[0]['CURRENCY_RATE'];
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)
                                                                                                 ->applyFromArray(array('font' => Style_Font(11, '006600', false, 'Consolas')));                                                   
                                        $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), "=".$val ."*".'$C'.($count_data+4));

                                    }elseif ($value['SOURCE_CD'] == 'M50020'){
                                        //$amt = $val * $rate[1]['CURRENCY_RATE'];
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)
                                                                                                 ->applyFromArray(array('font' => Style_Font(11, '0000CC', false, 'Consolas')));                                                   
                                        $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), "=".$val ."*".'$C'.($count_data+5));
                                    }                                                              
                            }                                
                    }
                    $row++;               
                }

                $objPHPExcel->getActiveSheet()->setCellValue('C3', 'WEEKLY FORECAST RECEIVING REPORT');
                $objPHPExcel->getActiveSheet()->setCellValue('C5',  date('F Y') . " to " . date('F Y', strtotime("+ 1 month", strtotime(date('Y-m-01') ) ) ) );
                $objPHPExcel->getActiveSheet()->setCellValue('J3',  'SUMMARY TOTAL');
                $objPHPExcel->getActiveSheet()->setCellValue('J4',  'PLAN');
                $objPHPExcel->getActiveSheet()->setCellValue('J5',  'PRICE AMOUNT');
                $objPHPExcel->getActiveSheet()->setCellValue('J6',  'PLAN NEXT MONTH');
                $objPHPExcel->getActiveSheet()->setCellValue('J7',  'PRICE NEXT MONTH');

                $objPHPExcel->getActiveSheet()->setCellValue('L4',  "=SUBTOTAL(9,J". $st_dat .":J".$count_data.")");
                $objPHPExcel->getActiveSheet()->setCellValue('L5',  "=SUBTOTAL(9,K". $st_dat .":K".$count_data.")");
                $objPHPExcel->getActiveSheet()->setCellValue('L6',  "=SUBTOTAL(9,L". $st_dat .":L".$count_data.")");
                $objPHPExcel->getActiveSheet()->setCellValue('L7',  "=SUBTOTAL(9,M". $st_dat .":M".$count_data.")");
//                 $objPHPExcel->getActiveSheet()->setCellValue('L8',  "=SUBTOTAL(9,N". $st_dat .":N".$count_data.")");
//                 $objPHPExcel->getActiveSheet()->setCellValue('L9',  "=SUBTOTAL(9,O". $st_dat .":O".$count_data.")");

                $objPHPExcel->getActiveSheet()->setCellValue('M4',  "Pcs.");
                $objPHPExcel->getActiveSheet()->setCellValue('M5',  "Thb.");
                $objPHPExcel->getActiveSheet()->setCellValue('M6',  "Pcs.");
                $objPHPExcel->getActiveSheet()->setCellValue('M7',  "Thb.");

                $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+3),  'Exchange rate');
                $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+4),  'USD');
                $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+5),  'EUR');
                $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+6),  'JPY');
                $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+4),  $ex_usd);
                $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+5),  $ex_eur);
                $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+6),  $ex_jpy);


                $objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray(array('font' => Style_Font(30,"000000",true,true)));
                $objPHPExcel->getActiveSheet()->getStyle('C5')->applyFromArray(array('font' => Style_Font(21,"000000",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('J3')->applyFromArray(array('font' => Style_Font(14,"ebf1de",true,true)));
                $objPHPExcel->getActiveSheet()->getStyle('J4:M9')->applyFromArray(array('font' => Style_Font(14,"974706",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('B'.$st_col.':'.'M'.$st_col)->applyFromArray(array('font' => Style_Font(10,"ebf1de",true,true)));
                $objPHPExcel->getActiveSheet()->getStyle('B'.$st_dat.':'.'M'.$count_data)->applyFromArray(array('font' => Style_Font(10,"000005",false,false)));

                $objPHPExcel->getActiveSheet()->getStyle('B'.($count_data+3) )->applyFromArray(array('font' => Style_Font(10,"000000",false,true)));
                $objPHPExcel->getActiveSheet()->getStyle( 'B'.($count_data+4).':'.'D'.($count_data+6) )->applyFromArray(array('font' => Style_Font(9,"000000",true,true)));


                $objPHPExcel->getActiveSheet()->getStyle('A1:M10')->applyFromArray(array('fill' => Style_Fill('FFFFFF')));
                //$objPHPExcel->getActiveSheet()->insertNewRowBefore(3,1);

                $objPHPExcel->getActiveSheet()->getStyle('J3'.':'.$col_name[$count_index]."3")->applyFromArray(array('fill' => Style_Fill('004700')));
                $objPHPExcel->getActiveSheet()->getStyle('J4'.':'.$col_name[$count_index]."7")->applyFromArray(array('fill' => Style_Fill('c6e0b4')));

                $objPHPExcel->getActiveSheet()->getStyle('B'.$st_col.':'.$col_name[$count_index].$st_col)->applyFromArray(array('fill' => Style_Fill('004700')));


                $objPHPExcel->getActiveSheet()->getStyle('C3:H4')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'00000E'))));

                $objPHPExcel->getActiveSheet()->getStyle('C5:H6')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E'))));

                $objPHPExcel->getActiveSheet()->getStyle('J4:'.$col_name[$count_index].'7')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));
                $objPHPExcel->getActiveSheet()->getStyle('J8:'.$col_name[$count_index].'8')
                                              ->applyFromArray(array(
                                                'borders' => array('top' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

                $objPHPExcel->getActiveSheet()->getStyle('B'.$st_col.':'.$col_name[$count_index].$st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));

                $objPHPExcel->getActiveSheet()->getStyle('B'.$st_dat.':'.$col_name[$count_index].$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));
                $objPHPExcel->getActiveSheet()->getStyle('B'.$count_data.':'.$col_name[$count_index].$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));

                $objPHPExcel->getActiveSheet()->getStyle('B'.$look_data.':'.'M'.$look_data)
                                              ->applyFromArray(array(
                                                'borders' => array('top'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E'))));

                $objPHPExcel->getActiveSheet()->getStyle( 'B' . ($count_data+3) .':'. 'D' .($count_data+3) )
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E'))));                                                              
// //echo $look_data; exit;

$objPHPExcel->getActiveSheet()->getStyle('J'.$st_dat.':'.$col_name[$count_index].$count_data)
                              ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');

$objPHPExcel->getActiveSheet()->getStyle('K'.$st_dat.':'.'K'.$count_data)
                              ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
$objPHPExcel->getActiveSheet()->getStyle('M'.$st_dat.':'.'M'.$count_data)
                              ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');

$objPHPExcel->getActiveSheet()->getStyle('L'.'4'.':'.'L'.'7')
                              ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
$objPHPExcel->getActiveSheet()->getStyle('L'.'5')
                              ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');                              
$objPHPExcel->getActiveSheet()->getStyle('L'.'7')
                              ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-'); 

                $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('2');              #A
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[0])->setWidth('5');     #B no
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[1])->setWidth('7');     #D plnt
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[1])->setWidth('8');     #C pd                
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[3])->setWidth('11');    #E so_no
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[4])->setWidth('19');    #F so_nm
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[5])->setWidth('17');    #G it_no
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[6])->setWidth('30');    #H it_nm
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[7])->setWidth('22');    #I model
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[8])->setWidth('21');    #J
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[9])->setWidth('21');    #K
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[10])->setWidth('21');   #L
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[11])->setWidth('21');#M
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[12])->setWidth('12');   #N
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[13])->setWidth('14.29');#M

                Style_Alignment('C2:C5',3, false, $objPHPExcel);
                Style_Alignment('J3',3, false, $objPHPExcel);
                Style_Alignment(('B'.$st_col.':'.'M'.$st_col), 3, false, $objPHPExcel);
                Style_Alignment(('B'.$st_dat.':'.'I'.$count_data), 9, false, $objPHPExcel);
                $objPHPExcel->getActiveSheet()->mergeCells('C3:'.'H4');
                $objPHPExcel->getActiveSheet()->mergeCells('C5:'.'H6');
                $objPHPExcel->getActiveSheet()->mergeCells('J3:'.'M3');

                 foreach (range( ($count_data+4) , ($count_data+6) ) as $index) Style_group_lv1_Row($index, $objPHPExcel);

//                 //echo ($count_data+4); exit;

                foreach(range(4, 7) as $r)
                {
                    $objPHPExcel->getActiveSheet()->mergeCells('J'.$r.':'.'K'.$r);
                    //$objPHPExcel->getActiveSheet()->mergeCells('J'.$r.':'.'K'.$r);                    
                }

                $objPHPExcel->getActiveSheet()->mergeCells('B' . ($count_data+3) .':'. 'D' .($count_data+3));

//                 foreach(range(($count_data+4), ($count_data+6)) as $r)
//                 {
//                     $objPHPExcel->getActiveSheet()->mergeCells('C'.$r.':'.'D'.$r);                 
//                 }               

#========================================================================================================================  Put field ==================================================================================== 
            }
            elseif( $sheetIndex == 'receive_history' ) 
            {
                $objPHPExcel->getActiveSheet()->setTitle( "History Receive last 12 month" );                
                $objPHPExcel->getActiveSheet()->setShowGridlines(False);
                //$objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('A1', 'TEST');
                //$objPHPExcel->getActiveSheet()->getStyle('A10')->getAlignment()->setTextRotation(45);
                $st_col = 18;
                $st_dat = 20;
                $count_index =  count($list_act_report[$sheetIndex][0]) - 1 ;
                $row = $st_dat;
                $i=0;
                $look_data = 0;
                $count_data  =  count($list_act_report[$sheetIndex]) + $row-1;
                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 10 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 10 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 16 )->setRowHeight( 10 );                
                $objPHPExcel->getActiveSheet()->getRowDimension( '3:16' )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 19 )->setRowHeight( 12 );  

                $objPHPExcel->getActiveSheet()->freezePane('A'.$row);   
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('B'.($st_col+1).':'.$col_name[$count_index].($st_col+1)); 
                                
                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++].($st_col-1), str_replace("_", " ", $key));
                } // exit;     
#========================================================================================================================  Put data ====================================================================================                

                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {               
                   $col = 0;
                    foreach ($value as $body => $val) 
                    {
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);
                            if($val == 3 && $body == 'MODEL')  $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('###"E00"');
                            if($val == "" && $body == 'PD' && $look_data === 0) $look_data = $row;

                            //echo substr($body, 0,1) . "<hr>";
                            if ( (substr($body, 0,1) == 'P' && $value[$body] > 0 && strlen($body) < 5 ) ) 
                            {
                                    if ($value['SOURCE_CD'] == 'T00100' || $value['SOURCE_CD'] == 'M50050' || $value['SOURCE_CD'] == 'T10400' ) {

                                       // $amt = $val * $rate[2]['CURRENCY_RATE'];
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)
                                                                                                 ->applyFromArray(array('font' => Style_Font(11, '009900', false, 'Consolas')));                                                   
                                        $objPHPExcel->setActiveSheetIndex($ind)->setCellValue( $col_name[$col-1].($row), "=".$val ."*".'$C'.($count_data+6));
                                    }

                                    elseif ($value['SOURCE_CD'] == 'M50030' || $value['SOURCE_CD'] == 'M50040' || $value['SOURCE_CD'] == 'M5004A' || $value['SOURCE_CD'] == 'M5004B' || $value['SOURCE_CD'] == 'M5004C' || $value['SOURCE_CD'] == 'T10200' || $value['SOURCE_CD'] == 'T10300'){
                                        //$amt = $val * $rate[0]['CURRENCY_RATE'];
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)
                                                                                                 ->applyFromArray(array('font' => Style_Font(11, '006600', false, 'Consolas')));                                                   
                                        $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), "=".$val ."*".'$C'.($count_data+4));

                                    }elseif ($value['SOURCE_CD'] == 'M50020'){
                                        //$amt = $val * $rate[1]['CURRENCY_RATE'];
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');
                                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)
                                                                                                 ->applyFromArray(array('font' => Style_Font(11, '0000CC', false, 'Consolas')));                                                   
                                        $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), "=".$val ."*".'$C'.($count_data+5));
                                    }                                                              
                            }                                                             
                    }//exit;
                    $row++;               
                }
                
                $objPHPExcel->getActiveSheet()->setCellValue('B5', 'MONTHLY RECEIVING HISTORY REPORT');
$st = ((date('m')-12) < 1 )  ? date('F-Y', strtotime( (date('Y')-1) . "-" . (12+(date('m')-12)). "-" . '01' ) ) : date('F-Y (ERROR)') ; 

$en = ((date('m')-1)  < 1 )  ? date('F-Y', strtotime( (date('Y')-1). "-" ."12". "-" . '01' ) ) : date('F-Y', strtotime( (date('Y')+0). "-" .(date('m')-1). "-" . '01' ) ) ;//
                $objPHPExcel->getActiveSheet()->setCellValue('B7', 'PERIOD TIME :  '. $st . ' To '. $en);
                $objPHPExcel->getActiveSheet()->setCellValue('H3', 'Summary Actual (Pcs.)');
                $objPHPExcel->getActiveSheet()->setCellValue('I3', 'Summary Price (Thb.)' );

                $objPHPExcel->getActiveSheet()->setCellValue('AH2',  'p' );                
                $objPHPExcel->getActiveSheet()->setCellValue('AH6',  'Click button to unhide' );
                $objPHPExcel->getActiveSheet()->setCellValue('AH20', 'DATA HISTORY MONTHLY RECEIVE' );

                $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+3),  'Exchange rate');
                $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+4),  'USD');
                $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+5),  'EUR');
                $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+6),  'JPY');
                $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+4),  $ex_usd);
                $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+5),  $ex_eur);
                $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+6),  $ex_jpy);

                $re_mon = 12;
                foreach(range(4, 15) as $mon)
                {
$his_month = ((date('m')-($re_mon)) < 1 )  ? date('F-Y', strtotime( (date('Y')-1) . "-" . (12+(date('m')-($re_mon--))). "-" . '01' ) ) : date('F-Y', strtotime( (date('Y')+0). "-" .(date('m')-($re_mon--)). "-" . '01' ) ) ;

                    $objPHPExcel->getActiveSheet()->setCellValue('G'.$mon , $his_month);
                }
                $sum_rA = 15;
                $sum_rP = 15;
                $switch_col = 0;
                foreach(range(8, 31) as $de_col)
                {
                    $detail = ($de_col % 2 == 0) ? "Actual (Pcs.)" : "Price (Thb.)" ;
                        $objPHPExcel->getActiveSheet()->setCellValue($col_name[$de_col].$st_col, $detail);

                    
                    if($de_col % 2 == 0)
                    {
$his_month = ((date('m')-(++$re_mon)) < 1 )  ? date('F-Y', strtotime( (date('Y')-1) . "-" . (12+(date('m')-($re_mon))). "-" . '01' ) ) : date('F-Y', strtotime( (date('Y')+0). "-" .(date('m')-($re_mon)). "-" . '01' ) ) ;
                        $objPHPExcel->getActiveSheet()->setCellValue($col_name[$de_col].($st_col-1), $his_month);

                        if($switch_col == 0)
                        {
                            $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col-1) )->applyFromArray(array('fill' => Style_Fill('002900')));
                            $switch_col = 1;
                        }
                        else
                        {
                            $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col-1) )->applyFromArray(array('fill' => Style_Fill('333300')));
                            $switch_col = 0;
                        }


                        $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) )->applyFromArray(array('fill' => Style_Fill('76933c')));
                        $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) .":".$col_name[$de_col].($count_data) )->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
                        $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) .":".$col_name[$de_col].($count_data) )->applyFromArray(array('font' => Style_Font(10,"000005",false,true)));
                              
                        $objPHPExcel->getActiveSheet()->setCellValue($col_name[6].($sum_rA--), '=SUBTOTAL(9,'.$col_name[$de_col].$st_dat.":".$col_name[$de_col].$count_data.')');
                    }
                    else
                    {
                       /// $objPHPExcel->getActiveSheet()->setCellValue($col_name[$de_col].($st_col-1), '');
                         $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col-1) )->applyFromArray(array('fill' => Style_Fill('333300')));
                         $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) )->applyFromArray(array('fill' => Style_Fill('4f6228')));
                         $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_dat) .":".$col_name[$de_col].($count_data) )->applyFromArray(array('fill' => Style_Fill('ebf1de')));
                         $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) .":".$col_name[$de_col].($count_data) )->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
                         $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) .":".$col_name[$de_col].($count_data) )->applyFromArray(array('font' => Style_Font(10,"eb2613",false,true)));                               

                        $objPHPExcel->getActiveSheet()->setCellValue($col_name[7].($sum_rP--), '=SUBTOTAL(9,'.$col_name[$de_col].$st_dat.":".$col_name[$de_col].$count_data.')');
                    }
                }

                $objPHPExcel->getActiveSheet()->getStyle('B5')->applyFromArray(array('font' => Style_Font(18,"000000",true,true)));
                $objPHPExcel->getActiveSheet()->getStyle('B7')->applyFromArray(array('font' => Style_Font(12,"000000",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('G3:I15')->applyFromArray(array('font' => Style_Font(11,"ebf1de",true,true)));
                $objPHPExcel->getActiveSheet()->getStyle('H4:I15')->applyFromArray(array('font'  => Style_Font(12,"974706",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('B'.($st_col-1).':'.'I' .($st_col-1))->applyFromArray(array('font' => Style_Font(10,"ebf1de",true,true)));
                $objPHPExcel->getActiveSheet()->getStyle('J'.($st_col-1).':'.'AG'.($st_col-1))->applyFromArray(array('font' => Style_Font(11,"ebf1de",true,true)));
                $objPHPExcel->getActiveSheet()->getStyle('J'.($st_col).  ':'.'AG'.($st_col))->applyFromArray(array('font' => Style_Font(10,"ebf1de",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('B'.($st_dat).':'.'I'.$count_data)->applyFromArray(array('font' => Style_Font(10,"000005",false,true)));

                $objPHPExcel->getActiveSheet()->getStyle('B'.($count_data+3) )->applyFromArray(array('font' => Style_Font(11,"000000",false,true)));
                $objPHPExcel->getActiveSheet()->getStyle( 'B'.($count_data+4).':'.'D'.($count_data+6) )->applyFromArray(array('font' => Style_Font(10,"000000",false,true)));   

                $objPHPExcel->getActiveSheet()->getStyle('AH2')->applyFromArray(array('font' => Style_Font(36,"00b0f0",true,false,'Wingdings 3')));
                $objPHPExcel->getActiveSheet()->getStyle('AH6')->applyFromArray(array('font' => Style_Font(14,"00b0f0",true,true,'Arial Unicode MS')));
                $objPHPExcel->getActiveSheet()->getStyle('AH20')->applyFromArray(array('font' => Style_Font(26,"00b0f0",true,true)));
                
                //$objPHPExcel->getActiveSheet()->getStyle('AH2')->getAlignment()->setTextRotation(90);
                $objPHPExcel->getActiveSheet()->getStyle('AH6')->getAlignment()->setTextRotation(-90);
                $objPHPExcel->getActiveSheet()->getStyle('AH20')->getAlignment()->setTextRotation(-90);

                $objPHPExcel->getActiveSheet()->getStyle('A1:M9')->applyFromArray(array('fill' => Style_Fill('FFFFFF')));
                //$objPHPExcel->getActiveSheet()->insertNewRowBefore(3,1);
                $objPHPExcel->getActiveSheet()->getStyle('H3'.':'.'I3')->applyFromArray(array('fill' => Style_Fill('002900')));
                $objPHPExcel->getActiveSheet()->getStyle('G4'.':'.'G15')->applyFromArray(array('fill' => Style_Fill('002900')));
                $objPHPExcel->getActiveSheet()->getStyle('H4'.':'.'I15')->applyFromArray(array('fill' => Style_Fill('c6e0b4')));



                $objPHPExcel->getActiveSheet()->getStyle('B'.($st_col-1).':'.$col_name[7].($st_col-1))->applyFromArray(array('fill' => Style_Fill('002900')));
$objPHPExcel->getActiveSheet()->getStyle( 'H4' .":".'H15' )->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
$objPHPExcel->getActiveSheet()->getStyle( 'I4' .":".'I15' )->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');

                $objPHPExcel->getActiveSheet()->getStyle('B5:F6')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'00000E'))));

                $objPHPExcel->getActiveSheet()->getStyle('B7:F8')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E'))));

                $objPHPExcel->getActiveSheet()->getStyle('H3:I15')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));
                $objPHPExcel->getActiveSheet()->getStyle('G4:G15')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                                              
                $objPHPExcel->getActiveSheet()->getStyle('H16:I16')
                                              ->applyFromArray(array(
                                                'borders' => array('top' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

                $objPHPExcel->getActiveSheet()->getStyle('B'.($st_col-1).':'.$col_name[$count_index].$st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));

                $objPHPExcel->getActiveSheet()->getStyle('B'.$st_dat.':'.$col_name[$count_index].$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));
                $objPHPExcel->getActiveSheet()->getStyle('B'.$count_data.':'.$col_name[$count_index].$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));
                $objPHPExcel->getActiveSheet()->getStyle('B'.$look_data.':'.$col_name[$count_index].$look_data)
                                              ->applyFromArray(array(
                                                'borders' => array('top'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E'))));

                $objPHPExcel->getActiveSheet()->getStyle('B2:'.$col_name[7].($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000023')))); 

                $objPHPExcel->getActiveSheet()->getStyle('J16:'.$col_name[31].($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000023'))));   

                $objPHPExcel->getActiveSheet()->getStyle( 'B' . ($count_data+3) .':'. 'D' .($count_data+3) )
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E')))); 

                foreach (range(8, 31) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel);
                foreach (range( ($count_data+4) , ($count_data+6) ) as $index) Style_group_lv1_Row($index, $objPHPExcel);

                $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('2');              #A
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[0])->setWidth('5');     #B no
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[1])->setWidth('8');     #D plnt
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[1])->setWidth('8');     #C pd                
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[3])->setWidth('19');    #E so_no
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[4])->setWidth('19');    #F so_nm
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[5])->setWidth('19');    #G it_no
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[6])->setWidth('30');    #H it_nm
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[7])->setWidth('30');    #I model    
                foreach(range(8, 31) as $key)
                    $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$key])->setWidth('14.71');
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[32])->setWidth('15.71');    #I model

                Style_Alignment('B2:B5',7, false, $objPHPExcel);
                Style_Alignment('H3:I3',3, false, $objPHPExcel);
                Style_Alignment('AH2',3, false, $objPHPExcel);
                Style_Alignment('AH6',2, false, $objPHPExcel);
                Style_Alignment('AH20',2, false, $objPHPExcel);
                Style_Alignment(('B'.($st_col-1).':'.$col_name[$count_index].$st_col), 3, false, $objPHPExcel);
                Style_Alignment(('B'.$st_dat.':'.'I'.$count_data), 9, false, $objPHPExcel);

                foreach(range(0, 7)  as $key) $objPHPExcel->getActiveSheet()->mergeCells($col_name[$key].($st_col-1).':'.$col_name[$key].$st_col);
                foreach(range(8, 31) as $key) 
                    if($key % 2 == 0)
                        $objPHPExcel->getActiveSheet()->mergeCells($col_name[$key].($st_col-1).':'.$col_name[($key+1)].($st_col-1));   
                $objPHPExcel->getActiveSheet()->mergeCells('B5'.':'.'F6');
                $objPHPExcel->getActiveSheet()->mergeCells('B7'.':'.'F8');  
                $objPHPExcel->getActiveSheet()->mergeCells('AH2'. ':'.'AH5');
                $objPHPExcel->getActiveSheet()->mergeCells('AH6'. ':'.'AH16');
                $objPHPExcel->getActiveSheet()->mergeCells('AH20'.':'.'AH'.($count_data+1));

                $objPHPExcel->getActiveSheet()->mergeCells('B' . ($count_data+3) .':'. 'D' .($count_data+3));

                foreach(range(($count_data+4), ($count_data+6)) as $r)
                {
                    $objPHPExcel->getActiveSheet()->mergeCells('C'.$r.':'.'D'.$r);                 
                }                                   
            }
#========================================================================================================================  Put data ====================================================================================         
    } else {
                    $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "No data ".$til.".");
                    $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true,false,'Franklin Gothic Book')));
    }
$ind++;

}

$objPHPExcel->setActiveSheetIndex(0);

$objPHPExcel->removeSheetByIndex(count($title));                             
                           
$today = date("My");
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
$con = 'Content-Disposition: attachment;filename='.$filename.$today.'.xlsx';
header($con);
header('Cache-Control: max-age=0');
header('Cache-Control: max-age=1');
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function Style_Fill($color=null) {

    return array( 'type'  => PHPExcel_Style_Fill::FILL_SOLID,                           
                  'color' => array('rgb' => $color)                                    
                );                                   
}

function Style_Font($size=11, $color='FFFFFF', $bol=false, $ita=false, $fname='Calibri') {

    return  array(
                    'name'  => $fname,
                    'size'  => $size,
                    'bold'  => $bol,
                    'italic'=> $ita,
                    'color' => array('rgb' => $color)
                 );                               
}
function Style_border($line='BORDER_THICK', $color='000000')
{
    return array( 'style' => $line, 'color' => array('rgb' => $color)) ;
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

function Style_group_lv1_Col($cell=null, $index=0, $objPHPExcel=null, $vi=false, $co=true)
{
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setOutlineLevel(1);
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setVisible($vi);
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setCollapsed($co); 
}
function Style_group_lv1_Row($index=0, $objPHPExcel=null, $vi=false, $co=true)
{
    $objPHPExcel->getActiveSheet()->getRowDimension ($index)->setOutlineLevel(1);
    $objPHPExcel->getActiveSheet()->getRowDimension ($index)->setVisible($vi);
    $objPHPExcel->getActiveSheet()->getRowDimension ($index)->setCollapsed($co); 
}
?>