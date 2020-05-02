<?php
//error_reporting(E_ALL);
error_reporting(E_ALL ^ E_NOTICE);
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
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "D".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "E".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "F".$cm); }

$i   = 0;   
$ind = 0;
//$T_lastM = ((date('m')-1) > 12 ) ? date('My', strtotime( (date('Y')-1). "-" ."12". "-" . '01' ) ) : date('My', strtotime( (date('Y')+0). "-" .(date('m')-1). "-" . '01' ) ) ;// exit;
//$H_lastM = ((date('m')-1) > 12 ) ? date('F Y', strtotime( (date('Y')-1). "-" ."12". "-" . '01' ) ) : date('F Y', strtotime( (date('Y')+0). "-" .(date('m')-1). "-" . '01' ) ) ;// exit;
//$T_lastM = date('My',  strtotime( date('Y'). "-" .(date('m')-1). "-" . 1 ) ) ;// exit;
//$H_lastM = date('F Y', strtotime( date('Y'). "-" .(date('m')-1). "-" . 1 ) ) ;// exit;

// $ex_usd = $rate[0]['CURRENCY_RATE'];
// $ex_eur = $rate[1]['CURRENCY_RATE'];
// $ex_jpy = $rate[2]['CURRENCY_RATE'];
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
            $count_data  =  count($list_act_report[$sheetIndex]);
    if ($count_data > 0) 
    {      
#========================================================================================================================  Put field ====================================================================================        
            if( $sheetIndex == 'defect' ) 
            {
                $objPHPExcel->getActiveSheet()->setTitle( "$til Weekly"  );
                $objPHPExcel->getActiveSheet()->setShowGridlines(False);
                $st_col = 23;
                $st_dat = 25;
                $count_index =  count($list_act_report[$sheetIndex][0]) - 2 ;
                $row = $st_dat;
                $look_data = 0;
                $count_data  =  count($list_act_report[$sheetIndex]) + $row-1;

                $gdImage = dirname(__FILE__) . '/img/NEW-TBKK-LOGO_0.png';
                // Add a drawing to the worksheetecho date('H:i:s') . " Add a drawing to the worksheet\n";
                $objDrawing = new PHPExcel_Worksheet_Drawing();
                $objDrawing->setName('Sample image');
                $objDrawing->setDescription('Sample image');
                $objDrawing->setPath($gdImage);
                //$objDrawing->setRenderingFunction(PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG);
                //$objDrawing->setMimeType(PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT);
                $objDrawing->setOffsetX(27); 
                $objDrawing->setOffsetY(40);  
                $objDrawing->setHeight(120);
                $objDrawing->setWidth(105); 
                $objDrawing->setCoordinates('B3');
                $objDrawing->setWorksheet($objPHPExcel->getActiveSheet()); 



                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 7 );

                foreach (range(2, 13) as $id ) 
                $objPHPExcel->getActiveSheet()->getRowDimension( $id )->setRowHeight( 20 );                
                $objPHPExcel->getActiveSheet()->getRowDimension( 14 )->setRowHeight( 13 );
                foreach (range(15, 18) as $id ) 
                $objPHPExcel->getActiveSheet()->getRowDimension( $id )->setRowHeight( 14 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 19 )->setRowHeight( 28 ); 
                $objPHPExcel->getActiveSheet()->getRowDimension( 20 )->setRowHeight( 4.5 ); 
                $objPHPExcel->getActiveSheet()->getRowDimension( 21 )->setRowHeight( 28 );  
                $objPHPExcel->getActiveSheet()->getRowDimension( 22 )->setRowHeight( 4.5 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 23 )->setRowHeight( 50 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 24 )->setRowHeight( 10 );

                $color_border['head'] = 'ffccff';
                $color_border['targ'] = 'ffccff';
                $color_border['deta'] = 'ffccff';
                $color_border['cost'] = 'ffff99';
                $color_border['summ'] = 'ffccff';
                $color_border['rm']   = 'b3ff99';
                $color_border['ma']   = 'ff9999';
                $color_border['as']   = 'ffcc99';
                $color_border['di']   = 'ffe699';
                $color_border['pe']   = '99ffcc';
                $color_border['oh']   = 'cc99ff';

                $style_layout['head'] = array('B' ,'J' );
                $style_layout['targ'] = array('S' ,'AG' );
                $style_layout['cost'] = array('K' ,'Q' );
                $style_layout['summ'] = array('S' ,'AG');
                $style_layout['rm']   = array('AI','AU');
                $style_layout['ma']   = array('AW','CG');
                $style_layout['as']   = array('CI','CV');
                $style_layout['di']   = array('CX','DQ');
                $style_layout['pe']   = array('DS','DX');
                $style_layout['oh']   = array('DZ','EK');

                $objPHPExcel->getActiveSheet()->freezePane('R'.$row);   
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(59);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('B'.($st_col+1).':'.'AC'.($st_col+1));            
            #====================================================================== เส้นตารางข้อมูล =============================================================================# 
                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['head'][0] . $st_dat . ':' . $style_layout['cost'][1].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'3e4241'))));

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['summ'][0] . $st_dat . ':' . 'AC'.$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'3e4241'),
                                                				   'right'    => Style_border(PHPExcel_Style_Border::BORDER_THIN ,$color_border['summ'])
                                            					  )));

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['rm'][0] . $st_dat . ':' . $style_layout['rm'][1].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'3e4241'))));

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['ma'][0] . $st_dat . ':' . $style_layout['ma'][1].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'3e4241'))));

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['as'][0] . $st_dat . ':' . $style_layout['as'][1].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'3e4241'))));   

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['di'][0] . $st_dat . ':' . $style_layout['di'][1].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'3e4241'))));   

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['pe'][0] . $st_dat . ':' . $style_layout['pe'][1].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'3e4241')))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['oh'][0] . $st_dat . ':' . $style_layout['oh'][1].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'3e4241'))));   






            #======================================================================= เส้น และ สี ของ ตางราง PPM ด้านบน ==============================================================================#   

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['head'][0] . '2' . ':' . $style_layout['head'][1] . '13' )
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['head'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['targ'][0] . '2' . ':' . $style_layout['targ'][1] . '13' )
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['targ']))));  

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['targ'][0] . '3' . ':' . $style_layout['targ'][1] . '4'  ) 
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['targ']),
                                                                   'top'      => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF"),
                                                                   'bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['targ']))));  



                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['targ'][0] . '5' . ':' . $style_layout['targ'][1] . '13' )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['targ']))));                                                

                $objPHPExcel->getActiveSheet()->getStyle( "T5:T13" )
                                              ->applyFromArray(array(
                                                'borders' => array('right'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['targ']))));        

                $objPHPExcel->getActiveSheet()->getStyle( "w5:w13" )
                                              ->applyFromArray(array(
                                                'borders' => array('right'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['targ']))));        

                $objPHPExcel->getActiveSheet()->getStyle( "Z5:Z13" )
                                              ->applyFromArray(array(
                                                'borders' => array('right'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['targ']))));    

                                                

            #======================================================================== เส้น ตารางหัว คอลัมป์ ============================================================#   


                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['head'][0] . $st_col . ':'. $style_layout['head'][1] . $st_col )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['deta'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['cost'][0] . $st_col . ':'. $style_layout['cost'][1] . $st_col )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['cost'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['summ'][0] . $st_col . ':'. $style_layout['summ'][1] . $st_col )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['summ'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['rm'][0] . $st_col . ':'. $style_layout['rm'][1] . $st_col )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['rm'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['ma'][0] . $st_col . ':'. $style_layout['ma'][1] . $st_col )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['ma'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['as'][0] . $st_col . ':'. $style_layout['as'][1] . $st_col )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['as'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['di'][0] . $st_col . ':'. $style_layout['di'][1] . $st_col )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['di'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['pe'][0] . $st_col . ':'. $style_layout['pe'][1] . $st_col )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['pe'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['oh'][0] . $st_col . ':'. $style_layout['oh'][1] . $st_col )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['oh'])))); 

                                                                                                                                                                                                                                                                                                                                            
            #======================================================================== เส้นแบ่งขอบเขตข้อมูล =======================================================================#                                  

                // $objPHPExcel->getActiveSheet()->getStyle('B14:'.$col_name[$count_index+1].($count_data+2))
                //                               ->applyFromArray(array(
                //                                 'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['deta'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['head'][0] . '15' . ':'. $style_layout['head'][1] . ($count_data+1) )
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['deta'])))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['head'][0] . '21' . ':' . $style_layout['head'][1] . '21')
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['deta'])))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['cost'][0] . '15' . ':'. $style_layout['cost'][1] . ($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['cost'])))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['summ'][0] . '15' . ':'. $style_layout['summ'][1] . ($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['summ'])))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['rm'][0] . '15' . ':'. $style_layout['rm'][1] . ($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['rm'])))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['ma'][0] . '15' . ':'. $style_layout['ma'][1] . ($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['ma'])))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['as'][0] . '15' . ':'. $style_layout['as'][1] . ($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['as'])))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['di'][0] . '15' . ':'. $style_layout['di'][1] . ($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['di'])))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['pe'][0] . '15' . ':'. $style_layout['pe'][1] . ($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['pe'])))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['oh'][0] . '15' . ':'. $style_layout['oh'][1] . ($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['oh'])))); 


            #======================================================================== เล้นแบ่ง ตางรางแถวที่ 19 SUM ข้อมูล PPM =======================================================================#       

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['cost'][0] . '19' . ':'. $style_layout['cost'][1] . '19')
                                              ->applyFromArray(array(
                                                'borders' => array('top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['summ'][0] . '19' . ':'. $style_layout['summ'][1] . '19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['summ']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['rm'][0] . '19' . ':'. $style_layout['rm'][1] . '19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['rm']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['ma'][0] . '19' . ':'. $style_layout['ma'][1] . '19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['ma']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['as'][0] . '19' . ':'. $style_layout['as'][1] . '19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['as']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['di'][0] . '19' . ':'. $style_layout['di'][1] . '19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['di']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['pe'][0] . '19' . ':'. $style_layout['pe'][1] . '19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['pe']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['oh'][0] . '19' . ':'. $style_layout['oh'][1] . '19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['oh']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") )));  

            #======================================================================== เล้นแบ่ง ตางรางแถวที่ 21 SUM ข้อมูล Pcs ========================================================#       

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['cost'][0] . '21' . ':'. $style_layout['cost'][1] . '21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['cost']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['cost']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['summ'][0] . '21' . ':'. $style_layout['summ'][1] . '21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['summ']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['summ']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['rm'][0] . '21' . ':'. $style_layout['rm'][1] . '21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['rm']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['rm']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['ma'][0] . '21' . ':'. $style_layout['ma'][1] . '21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['ma']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['ma']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['as'][0] . '21' . ':'. $style_layout['as'][1] . '21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['as']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['as']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['di'][0] . '21' . ':'. $style_layout['di'][1] . '21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['di']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['di']) )));  

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['pe'][0] . '21' . ':'. $style_layout['pe'][1] . '21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['pe']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['pe']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['oh'][0] . '21' . ':'. $style_layout['oh'][1] . '21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['oh']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['oh']) ))); 


            #======================================================================== กำหนดสี fill ========================================================# 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['head'][0] . '2' . ':'. $style_layout['head'][1] . '13')->applyFromArray(array('fill' => Style_Fill($color_border['head'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['targ'][0] . '2' . ':'. $style_layout['targ'][1] . '2')->applyFromArray(array('fill' => Style_Fill($color_border['targ'])));                          

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['head'][0] . '15' . ':'. $style_layout['head'][1] . '19')->applyFromArray(array('fill' => Style_Fill($color_border['deta'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['head'][0] . '21' . ':'. $style_layout['head'][1] . '21')->applyFromArray(array('fill' => Style_Fill($color_border['deta'])));



                $objPHPExcel->getActiveSheet()->getStyle($style_layout['rm'][0]   . '15' . ':'. $style_layout['rm'][1]   . '18')->applyFromArray(array('fill' => Style_Fill($color_border['rm'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['ma'][0]   . '15' . ':'. $style_layout['ma'][1]   . '18')->applyFromArray(array('fill' => Style_Fill($color_border['ma'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['summ'][0] . '15' . ':'. $style_layout['summ'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['summ'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['cost'][0] . '15' . ':'. $style_layout['cost'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['cost'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['cost'][0] . '19' . ':'. $style_layout['cost'][1] . '19')->applyFromArray(array('fill' => Style_Fill($color_border['cost'])));

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['as'][0] . '15' . ':' . $style_layout['as'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['as'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['di'][0] . '15' . ':' . $style_layout['di'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['di'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['pe'][0] . '15' . ':' . $style_layout['pe'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['pe'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['oh'][0] . '15' . ':' . $style_layout['oh'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['oh'])));

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['head'][0]  . ($st_col+1) .':'. $style_layout['head'][1]  . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['deta'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['cost'][0]  . ($st_col+1) .':'. $style_layout['cost'][1]  . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['cost']  )));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['summ'][0]  . ($st_col+1) .':'. $style_layout['summ'][1]  . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['summ']  )));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['rm'][0]    . ($st_col+1) .':'. $style_layout['rm'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['rm']  )));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['ma'][0]    . ($st_col+1) .':'. $style_layout['ma'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['ma']  )));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['as'][0]    . ($st_col+1) .':'. $style_layout['as'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['as']  )));                
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['di'][0]    . ($st_col+1) .':'. $style_layout['di'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['di']  )));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['pe'][0]    . ($st_col+1) .':'. $style_layout['pe'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['pe']  )));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['oh'][0]    . ($st_col+1) .':'. $style_layout['oh'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['oh']  )));



            #======================================================================== กำหนดขนาด คอลัมป์ ========================================================# 

                 $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('2.71');              
                 $objPHPExcel->getActiveSheet()->getColumnDimension('EM')->setWidth('2.71'); 

                 $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('23.71');    

                    foreach (range('C', 'D') as $id )
                     $objPHPExcel->getActiveSheet()->getColumnDimension($id)->setWidth('9.71');

                     $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('21.71');                  
                     $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('10.71');    
                     $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('59.71');    
                     $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('21.71');    
                     $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('32.71');    
                     $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('29.71'); 
                    foreach (range('K', 'Q') as $id )
                     $objPHPExcel->getActiveSheet()->getColumnDimension($id)->setWidth('16.71');   
     
                    foreach (range(17, 31) as $id )
                     $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$id])->setWidth('16.29');

                     // $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth('12.71');   
                     // $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth('12.71');
                     // $objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth('12.71');   
                     // $objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth('12.71');
                     // $objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth('12.71');
                     // $objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth('12.71');
                     // $objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setWidth('12.71');  
                                    
                    foreach (range(33, 131) as $id )
                     $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$id])->setWidth('13.29');  

                    foreach (array('R', 'AH', 'EL') as $ind_col ) 
                     $objPHPExcel->getActiveSheet()->getColumnDimension($ind_col)->setWidth('2.71'); 


            #======================================================================== การ input ========================================================================================    
                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {
                    if ($key != 'IND') 

                    {
                        
                        if ( substr($key, 0,3) == 'CD_') 

                        	 $objPHPExcel->getActiveSheet()->setCellValue($col_name[$i++].$st_col,  intval( substr($key , 3 ,3) ));

                        elseif ( $key == 'BK' || $key == 'RMM' || $key == 'MAA' || $key == 'ASD' || $key == 'DIC' || $key == 'PEE' || $key == 'OHR' || $key == 'V1' || $key == 'V2' || $key == 'V3' || $key == 'V4'  ) 

                        	 $objPHPExcel->getActiveSheet()->setCellValue($col_name[$i++].$st_col,  '' ) ;

                        else $objPHPExcel->getActiveSheet()->setCellValue($col_name[$i++].$st_col, str_replace("_", " ", $key));


                    }
                } // exit;     
       

                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {               
                   $col = 0;
                    foreach ($value as $body => $val) 
                    {

                        if ($body != 'IND') 

                        {


                                if ( $body == 'TOTAL' )

                                {

                                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col++].($row), "=U$row + V$row");

                                }

                                elseif ($body == 'TOTAL_DEFECT') 

                                {
                                
                                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col++].($row), "=SUM( W$row:AB$row )");
                                
                                }

                                else
                                {

                                     $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col++].($row), $val);
                                
                                }
                                                                   

                                if($val == 3 && $body == 'MODEL')  $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('###"E00"');


                        }

                    }
                    $row++;               
                }

            #======================================================================== จัดตำแหน่ง ข้อมูล ====================================================================================     


                //echo date('t F Y', strtotime( "$mnt month", strtotime( date('01-m-Y') ) ) ) ; exit;


                 $objPHPExcel->getActiveSheet()->setCellValue('C'  . '3', 'DEFECT  WEEKLY REPORT');
                 $objPHPExcel->getActiveSheet()->setCellValue('C'  . '6', '( Accumulate Ng ' . date('01 F') . ' to ' . date('d F', strtotime("- 1 day") ) . " )" );

                 $objPHPExcel->getActiveSheet()->setCellValue('H'  . '12', 'TBKK (Thailand) Co., Ltd.' );
                 $objPHPExcel->getActiveSheet()->setCellValue('H'  . '13', 'vol. 1.21  :  Issue by Pc System ' . date('d-m-Y') );


                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '2', 'Control PPM in TBKK process.' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '3', 'Group' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '3', 'RM ( PPM. )' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '3', 'MA ( PPM. )' );
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '3', 'PE ( PPM. )' );

                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '4', 'Target' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '4', 'Target' );
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '4', 'Target' );
                 $objPHPExcel->getActiveSheet()->setCellValue('X'  . '3', 'Ng + Act.' );
                 $objPHPExcel->getActiveSheet()->setCellValue('Y'  . '3', 'Receive' );

                 $objPHPExcel->getActiveSheet()->setCellValue('Z'    . '3', 'NG Pcs.' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AD'   . '3', 'NG PPM.' );
                 
                 $objPHPExcel->getActiveSheet()->setCellValue('X'    . '4', '(Pcs.)' );
                 $objPHPExcel->getActiveSheet()->setCellValue('Y'    . '4', '(Pcs.)' );
                 $objPHPExcel->getActiveSheet()->setCellValue('Z'    . '4', 'RM' );                 
                 $objPHPExcel->getActiveSheet()->setCellValue('AA'   . '4', 'MA' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AB'   . '4', 'PE' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AC'   . '4', 'Other' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AD'   . '4', 'RM' );                 
                 $objPHPExcel->getActiveSheet()->setCellValue('AE'   . '4', 'MA' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AF'   . '4', 'PE' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AG'   . '4', 'Other' );

                 $objPHPExcel->getActiveSheet()->setCellValue('AE'  . '19',  'PPM.' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AD'  . '19',  '×' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AG'  . '19',  'Ø' ); 
                 $objPHPExcel->getActiveSheet()->setCellValue('AE'  . '21',  'SUM (pcs.)' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AD'  . '21',  '×' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AG'  . '21',  'Ø' );                                                                           


                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '5',  'PD1 ASSY' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '6',  'PD2 ENGINE PUMP' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '7',  'PD3 BRAKE&FCD' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '8',  'PD3 KUBOTA' );                 
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '9',  'PD4 ADC 1' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '10', 'PD5 GEAR' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '11', 'PD5 GKN' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '12', 'PD6 BH' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '13', 'None' );

                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '5',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '6',  '44' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '7',  '69' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '8',  '22' );                 
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '9',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '10', '365' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '11', '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '12', '99' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '13', '0' );

                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '5',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '6',  '632' );
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '7',  '74' );
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '8',  '486' );                 
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '9',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '10',  '348' );
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '11', '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '12', '498' );
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '13', '0' );                 

                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '5',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '6',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '7',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '8',  '0' );                 
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '9',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '10', '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '11', '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '12', '20131' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '13', '0' );  

                 $objPHPExcel->getActiveSheet()->setCellValue('B'  . '15', 'DEFECT OF ' . strtoupper( date('F-Y') ) );
                 $objPHPExcel->getActiveSheet()->setCellValue('K'  . '15', 'COST');
                 $objPHPExcel->getActiveSheet()->setCellValue('K'  . '19', 'TOTAL Cost( Baht )');
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '15', 'SUMMARY NG');
                 $objPHPExcel->getActiveSheet()->setCellValue($style_layout['rm'][0] . '15', 'NG CODE RM');
                 $objPHPExcel->getActiveSheet()->setCellValue($style_layout['ma'][0] . '15', 'NG CODE MA');
                 $objPHPExcel->getActiveSheet()->setCellValue($style_layout['as'][0] . '15', 'NG CODE AS');
                 $objPHPExcel->getActiveSheet()->setCellValue($style_layout['di'][0] . '15', 'NG CODE PD4');                 
                 $objPHPExcel->getActiveSheet()->setCellValue($style_layout['pe'][0] . '15', 'NG CODE PE');
                 $objPHPExcel->getActiveSheet()->setCellValue($style_layout['oh'][0] . '15', 'NG CODE OTHER');

                 $objPHPExcel->getActiveSheet()->setCellValue('B' . $st_col, 'GROUP');

                 $objPHPExcel->getActiveSheet()->setCellValue('K' . $st_col,  'UC'."\n"."( Baht )");
                 $objPHPExcel->getActiveSheet()->setCellValue('L' . $st_col,  'RM'."\n"."( Baht )");
                 $objPHPExcel->getActiveSheet()->setCellValue('M' . $st_col,  'MA'."\n"."( Baht )");
                 $objPHPExcel->getActiveSheet()->setCellValue('N' . $st_col,  'AS'."\n"."( Baht )");
                 $objPHPExcel->getActiveSheet()->setCellValue('O' . $st_col,  'PD4'."\n"."( Baht )");
                 $objPHPExcel->getActiveSheet()->setCellValue('P' . $st_col,  'PE'."\n"."( Baht )");
                 $objPHPExcel->getActiveSheet()->setCellValue('Q' . $st_col,  'OTHER'."\n"."( Baht )");


                 $objPHPExcel->getActiveSheet()->setCellValue('S' . $st_col, 'RECEIVE'."\n"."(Pcs.)");
                 $objPHPExcel->getActiveSheet()->setCellValue('T' . $st_col, 'ACT + NG'."\n"."(Pcs.)");
                 $objPHPExcel->getActiveSheet()->setCellValue('U' . $st_col, 'ACTUAL'."\n"."(Pcs.)");
                 $objPHPExcel->getActiveSheet()->setCellValue('V' . $st_col, 'NG'."\n"."(Pcs.)");

                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . $st_col, 'RM'."\n"."(Pcs.)");
                 $objPHPExcel->getActiveSheet()->setCellValue('X'  . $st_col, 'MA'."\n"."(Pcs.)");
                 $objPHPExcel->getActiveSheet()->setCellValue('Y'  . $st_col, 'AS'."\n"."(Pcs.)");
                 $objPHPExcel->getActiveSheet()->setCellValue('Z'  . $st_col, 'PD4'."\n"."(Pcs.)");
                 $objPHPExcel->getActiveSheet()->setCellValue('AA' . $st_col, 'PE'."\n"."(Pcs.)");
                 $objPHPExcel->getActiveSheet()->setCellValue('AB' . $st_col, 'OTHER'."\n"."(Pcs.)");
                 $objPHPExcel->getActiveSheet()->setCellValue('AC' . $st_col, 'MA + PE'."\n"."(Pcs.)");

                    foreach ( range(10, 15) as $sum )   $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '21',  "=SUBTOTAL(9,". $col_name[$sum] . $st_dat .":". $col_name[$sum] .$count_data.")");
                    foreach ( range(17, 27) as $sum )   $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '21',  "=SUBTOTAL(9,". $col_name[$sum] . $st_dat .":". $col_name[$sum] .$count_data.")");
                    foreach ( range(33, 45) as $sum )   $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '21',  "=SUBTOTAL(9,". $col_name[$sum] . $st_dat .":". $col_name[$sum] .$count_data.")");
                    foreach ( range(47, 83) as $sum )   $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '21',  "=SUBTOTAL(9,". $col_name[$sum] . $st_dat .":". $col_name[$sum] .$count_data.")");
                    foreach ( range(85, 98) as $sum )   $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '21',  "=SUBTOTAL(9,". $col_name[$sum] . $st_dat .":". $col_name[$sum] .$count_data.")");

                    foreach ( range(100,  119) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '21',  "=SUBTOTAL(9,". $col_name[$sum] . $st_dat .":". $col_name[$sum] .$count_data.")");  
                    foreach ( range(121, 126) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '21',  "=SUBTOTAL(9,". $col_name[$sum] . $st_dat .":". $col_name[$sum] .$count_data.")"); 
                    foreach ( range(128, 139) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '21',  "=SUBTOTAL(9,". $col_name[$sum] . $st_dat .":". $col_name[$sum] .$count_data.")"); 

                    foreach ( range(21, 27) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '19',  "=(". $col_name[$sum] . "21" ."/". '$T$21)*1000000' );
                    foreach ( range(33, 45) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '19',  "=(". $col_name[$sum] . "21" ."/". '$T$21)*1000000' );
                    foreach ( range(47, 83) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '19',  "=(". $col_name[$sum] . "21" ."/". '$T$21)*1000000' );
                    foreach ( range(85, 98) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '19',  "=(". $col_name[$sum] . "21" ."/". '$T$21)*1000000' );

                    foreach ( range(100,  119) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '19',  "=(". $col_name[$sum] . "21" ."/". '$T$21)*1000000' );  
                    foreach ( range(121, 126) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '19',  "=(". $col_name[$sum] . "21" ."/". '$T$21)*1000000' ); 
                    foreach ( range(128, 139) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '19',  "=(". $col_name[$sum] . "21" ."/". '$T$21)*1000000' ); 




                    foreach ( range(5, 13)   as $sum )  

                        {
                            $sum_as = "SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$Y$'  . $st_dat . ":" . '$Y$'  . $count_data . ")";
                            $sum_d4 = "SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$Z$'  . $st_dat . ":" . '$Z$'  . $count_data . ")";
                            $sum_ma = "SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$X$'  . $st_dat . ":" . '$X$'  . $count_data . ")";
                            $sum_pe = "SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$AA$' . $st_dat . ":" . '$AA$' . $count_data . ")";
                            $sum_oh = "SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$AB$' . $st_dat . ":" . '$AB$' . $count_data . ")";
                            $sum_rm = "SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$W$'  . $st_dat . ":" . '$W$'  . $count_data . ")";

                            //echo $sum_ma; exit;

                            $sum_total = " SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$T$' . $st_dat . ":" . '$T$' . $count_data  . ")";

                            $sum_recei = " SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$S$' . $st_dat . ":" . '$S$' . $count_data  . ")";
                           

                            $objPHPExcel->getActiveSheet()->setCellValue( 'X' . $sum,   "=" . $sum_total);
                            $objPHPExcel->getActiveSheet()->setCellValue( 'Y' . $sum,   "=" . $sum_recei);
                            
                            if ($sum == 5)
                                $objPHPExcel->getActiveSheet()->setCellValue( 'AA' . $sum,   "=" . $sum_as);
                            elseif ($sum == 9) 
                                $objPHPExcel->getActiveSheet()->setCellValue( 'AA' . $sum,   "=" . $sum_d4);
                            else
                                $objPHPExcel->getActiveSheet()->setCellValue( 'AA' . $sum,   "=" . $sum_ma);


						


                            $objPHPExcel->getActiveSheet()->setCellValue( 'Z'  . $sum,   "=" . $sum_rm);
                            $objPHPExcel->getActiveSheet()->setCellValue( 'AB'  . $sum,   "=" . $sum_pe);
                            $objPHPExcel->getActiveSheet()->setCellValue( 'AC'  . $sum,   "=" . $sum_oh);

							if ($sum == 12)	$objPHPExcel->getActiveSheet()->setCellValue( 'AD' . $sum,  "=IFERROR(( Z". $sum . "/" . "Y" . $sum ."  ) * 1000000 ,0)" );
							else $objPHPExcel->getActiveSheet()->setCellValue( 'AD' . $sum,  "0" );
                            
                            $objPHPExcel->getActiveSheet()->setCellValue( 'AE' . $sum,  "=IFERROR(( AA". $sum . "/" . "X" . $sum ." ) * 1000000 ,0)" );
                            $objPHPExcel->getActiveSheet()->setCellValue( 'AF' . $sum,  "=IFERROR(( AB". $sum . "/" . "X" . $sum ." ) * 1000000 ,0)" );
                            $objPHPExcel->getActiveSheet()->setCellValue( 'AG' . $sum,  "=IFERROR(( AC". $sum . "/" . "X" . $sum ." ) * 1000000 ,0)" );

                            //echo "=(". $sum_oh . "/" . $sum_total .")*1000000)"; exit;
                        }

                    foreach ( range(25, $count_data) as $sum)

                        {

                                $objPHPExcel->getActiveSheet()->setCellValue( 'W'   . $sum,  "=SUM(" . $style_layout['rm'][0] . $sum . ":" . $style_layout['rm'][1] . $sum . ")" );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'X'   . $sum,  "=SUM(" . $style_layout['ma'][0] . $sum . ":" . $style_layout['ma'][1] . $sum . ")" );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'Y'   . $sum,  "=SUM(" . $style_layout['as'][0] . $sum . ":" . $style_layout['as'][1] . $sum . ")" );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'Z'   . $sum,  "=SUM(" . $style_layout['di'][0] . $sum . ":" . $style_layout['di'][1] . $sum . ")" );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'AA'  . $sum,  "=SUM(" . $style_layout['pe'][0] . $sum . ":" . $style_layout['pe'][1] . $sum . ")" );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'AB'  . $sum,  "=SUM(" . $style_layout['oh'][0] . $sum . ":" . $style_layout['oh'][1] . $sum . ")" );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'AC'  . $sum,  "= X" . $sum . "+" . "AA" . $sum  );



                                $objPHPExcel->getActiveSheet()->setCellValue( 'L'   . $sum,  "=K" . $sum . " * " . "W"  . $sum  );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'M'   . $sum,  "=K" . $sum . " * " . "X"  . $sum  );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'N'   . $sum,  "=K" . $sum . " * " . "Y"  . $sum  );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'O'   . $sum,  "=K" . $sum . " * " . "Z"  . $sum  );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'P'   . $sum,  "=K" . $sum . " * " . "AA" . $sum  );
                                $objPHPExcel->getActiveSheet()->setCellValue( 'Q'   . $sum,  "=K" . $sum . " * " . "AB" . $sum  );                      
                       
                       }

                    foreach ( range(5, 13)   as $sum )  

                        {

                            $objConditional1 = new PHPExcel_Style_Conditional();
                            $objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
                                        ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHAN)
                                        ->addCondition('V' . $sum )
                                        ->getStyle()->applyFromArray(array( 'font' => Style_Font(11,'FF0000',true,false)));

                            $objPHPExcel->getActiveSheet()->getStyle('AE'. $sum)->setConditionalStyles(array($objConditional1));

                            $objConditional2 = new PHPExcel_Style_Conditional();
                            $objConditional2->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
                                        ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHAN)
                                        ->addCondition('W' . $sum )
                                        ->getStyle()->applyFromArray(array( 'font' => Style_Font(11,'FF0000',true,false)));

                            $objPHPExcel->getActiveSheet()->getStyle('AF'.$sum)->setConditionalStyles(array($objConditional2));

                            $objConditional3 = new PHPExcel_Style_Conditional();
                            $objConditional3->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
                                        ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHAN)
                                        ->addCondition('U' . $sum )
                                        ->getStyle()->applyFromArray(array( 'font' => Style_Font(11,'FF0000',true,false)));

                            $objPHPExcel->getActiveSheet()->getStyle('AD'.$sum)->setConditionalStyles(array($objConditional3));                                     
                        }                       
                                            

            #======================================================================== จัดตำแหน่ง ข้อมูล ====================================================================================
                    Style_Alignment('B15'.':'.$col_name[$count_index].($st_col),3, true, $objPHPExcel); 
                    Style_Alignment('C3',9, False, $objPHPExcel);    
                    Style_Alignment('C6',9, False, $objPHPExcel);   
                    Style_Alignment('H12:H13',6, False, $objPHPExcel);

                    Style_Alignment('S2:AG4',3, False, $objPHPExcel);
                    Style_Alignment('S5:S13',9, False, $objPHPExcel);
                    //Style_Alignment('T4:T12',9, False, $objPHPExcel);
                    Style_Alignment('C25:J'.($count_data),9, False, $objPHPExcel); 

                    Style_Alignment('S19:V19',3, False, $objPHPExcel);

                    Style_Alignment('T19',9, False, $objPHPExcel);
                    Style_Alignment('V19',9, False, $objPHPExcel);               
                   // Style_Alignment('B15'.':'.$col_name[$count_index].($st_col),9, true, $objPHPExcel);

                    
            #======================================================================== กำหนด ฟ้อร์น และ สี ==================================================================================== 
                                                     
                   $objPHPExcel->getActiveSheet()->getStyle('B15') ->applyFromArray(array('font' => Style_Font(36,"000000",True,False,'Arial Rounded MT Bold'))); 
                   $objPHPExcel->getActiveSheet()->getStyle('S15:EL15') ->applyFromArray(array('font' => Style_Font(36,"000000",True,False,'Arial Rounded MT Bold'))); 
                   $objPHPExcel->getActiveSheet()->getStyle('K15:Q15')->applyFromArray(array('font' => Style_Font(36,"000000",True,False,'Arial Rounded MT Bold')));
                   $objPHPExcel->getActiveSheet()->getStyle('K19')->applyFromArray(array('font' => Style_Font(12,"000000",True,False,'Arial Rounded MT Bold')));

                   $objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray(array('font' => Style_Font(26,"000000",True,False,'Arial Rounded MT Bold')));
                   $objPHPExcel->getActiveSheet()->getStyle('C6')->applyFromArray(array('font' => Style_Font(18,"000000",True,False,'Arial Rounded MT Bold')));
                   $objPHPExcel->getActiveSheet()->getStyle('H12')->applyFromArray(array('font' => Style_Font(12,"000000",False,False,'Arial Rounded MT Bold')));
                   $objPHPExcel->getActiveSheet()->getStyle('H13')->applyFromArray(array('font' => Style_Font(11,"000000",False,False,'Arial Rounded MT Bold')));

                   $objPHPExcel->getActiveSheet()->getStyle('S2')->applyFromArray(array('font' => Style_Font(12,"000000",True,False)));
                   $objPHPExcel->getActiveSheet()->getStyle('S3:S13')->applyFromArray(array('font' => Style_Font(12,"000000",True,False)));
                   $objPHPExcel->getActiveSheet()->getStyle('S4:S14')->applyFromArray(array('font' => Style_Font(11,"000000",True,False)));
                
                   $objPHPExcel->getActiveSheet()->getStyle('U3:AG13')->applyFromArray(array('font' => Style_Font(11,"000000",True,False)));


                   //$objPHPExcel->getActiveSheet()->getStyle('S19')->applyFromArray(array('font' => Style_Font(12,"000000",True,False,'Arial Rounded MT Bold')));

                   //$objPHPExcel->getActiveSheet()->getStyle('I2')->applyFromArray(array('font' => Style_Font(16,"000000",True,False)));

                   $objPHPExcel->getActiveSheet()->getStyle('S19:EL21')->applyFromArray(array('font' => Style_Font(12,"000000",True,False,'Arial Rounded MT Bold')));

                   $objPHPExcel->getActiveSheet()->getStyle('B'. $st_col . ':'.$col_name[$count_index].($st_col)) ->applyFromArray(array('font' => Style_Font(12,"000000",True,False,'Arial Rounded MT Bold')));

                   $objPHPExcel->getActiveSheet()->getStyle('S'. $st_dat . ':'.$col_name[$count_index].($count_data)) ->applyFromArray(array('font' => Style_Font(11,"000000",True,False,'Arial Rounded MT Bold')));


                   $objPHPExcel->getActiveSheet()->getStyle('AD19:AD21')->applyFromArray(array('font' => Style_Font(24,"000000",True,False,'Wingdings')));
                   $objPHPExcel->getActiveSheet()->getStyle('AG19:AG21')->applyFromArray(array('font' => Style_Font(24,"000000",True,False,'Wingdings')));
            #======================================================================== พิเศษ ========================================================================================

                       $objPHPExcel->getActiveSheet()->setCellValue('AV'  . '15', 'RM' );
                       $objPHPExcel->getActiveSheet()->setCellValue('CH'  . '15', 'MA' );
                       $objPHPExcel->getActiveSheet()->setCellValue('CW'  . '15', 'AS' );
                       $objPHPExcel->getActiveSheet()->setCellValue('DR'  . '15', 'PD4');
                       $objPHPExcel->getActiveSheet()->setCellValue('DY'  . '15', 'PE' );
                       $objPHPExcel->getActiveSheet()->setCellValue('EL'  . '15', 'OTHER' );
                        


                    foreach (array('AV','CH','CW','DR','DY','EL') as $in_c)

                     {
                        Style_Alignment( $in_c . '15' ,2, False, $objPHPExcel);
                        $objPHPExcel->getActiveSheet()->getStyle( $in_c . '15')->getAlignment()->setTextRotation(-90);
                        $objPHPExcel->getActiveSheet()->getStyle( $in_c . '15')->applyFromArray(array('font' => Style_Font(28,"FF0000",True,False,'Arial Rounded MT Bold')));                        
                        $objPHPExcel->getActiveSheet()->getColumnDimension($in_c)->setWidth('8.71');  
                        $objPHPExcel->getActiveSheet()->mergeCells( $in_c . '15' . ':' .$in_c . '23' );# code...
                     }
          
                       


              


                                            
                                    


            #======================================================================== merge crll  ==================================================================================== 
                    $objPHPExcel->getActiveSheet()->mergeCells( 'C3:I5' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'C6:I7' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'H12:J12' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'H13:J13' );

                    $objPHPExcel->getActiveSheet()->mergeCells( 'S2:AG2' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'S3:T4' );
                    //$objPHPExcel->getActiveSheet()->mergeCells( 'AC3:AC4' );
                    foreach ( range(5, 13) as $ro)
                    $objPHPExcel->getActiveSheet()->mergeCells( 'S' . $ro . ':' . 'T' . $ro );


                    $objPHPExcel->getActiveSheet()->mergeCells( 'Z3:AC3' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'AD3:AG3' );                    



                    $objPHPExcel->getActiveSheet()->mergeCells( 'B15:J19' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'B21:J21' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'K19:Q19' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'AE21:AF21' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'AE19:AF19' );
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['cost'][0] . '15' . ':'. $style_layout['cost'][1] . '18'  );
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['rm'][0] . '15' . ':'. $style_layout['rm'][1] . '18'  );
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['ma'][0] . '15' . ':'. $style_layout['ma'][1] . '18'  );  
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['as'][0] . '15' . ':'. $style_layout['as'][1] . '18'  ); 
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['di'][0] . '15' . ':'. $style_layout['di'][1] . '18'  );                     
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['pe'][0] . '15' . ':'. $style_layout['pe'][1] . '18'  );  
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['oh'][0] . '15' . ':'. $style_layout['oh'][1] . '18'  );  
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['summ'][0] . '15' . ':'. $style_layout['summ'][1] . '18'  );

            #======================================================================== กำหนด ฟอรืแมท ข้อมุล  ==================================================================================== 

                    $objPHPExcel->getActiveSheet()->getStyle('U5:AC13')->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');

                    $objPHPExcel->getActiveSheet()->getStyle('AD5:AG13')->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
                                                  
                    $objPHPExcel->getActiveSheet()->getStyle('K'.$st_dat.':'.'Q'.$count_data)->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');

                    $objPHPExcel->getActiveSheet()->getStyle('S'.$st_dat.':'.$col_name[$count_index].$count_data)->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
                    
                    $objPHPExcel->getActiveSheet()->getStyle('AI'.$st_col.':'.$col_name[$count_index].$st_col)->getNumberFormat()->setFormatCode('00#');                                                 

                    $objPHPExcel->getActiveSheet()->getStyle('L'.'21'.':'.'Q'.'21')->getNumberFormat()->setFormatCode('#,##0.00');

                    $objPHPExcel->getActiveSheet()->getStyle('W'.'19'.':'.$col_name[$count_index].'19')->getNumberFormat()->setFormatCode('#,##0.00');

                    $objPHPExcel->getActiveSheet()->getStyle('S'.'21'.':'.$col_name[$count_index].'21')->getNumberFormat()->setFormatCode('#,##0');

            #======================================================================== กรุป คอลัมป์  ==================================================================================== 
                   Style_group_Col($col_name, 5, $objPHPExcel);
                   Style_group_Col($col_name, 7, $objPHPExcel);
                   foreach ( range(9, 15)    as $index) Style_group_Col($col_name, $index, $objPHPExcel);
                   foreach ( range(32,  140) as $index) Style_group_Col($col_name, $index, $objPHPExcel, 1);                  
                   foreach ( range(33,  45)  as $index) Style_group_Col($col_name, $index, $objPHPExcel, 2 );
                   foreach ( range(47,  83)  as $index) Style_group_Col($col_name, $index, $objPHPExcel, 2 );
                   foreach ( range(85,  98)  as $index) Style_group_Col($col_name, $index, $objPHPExcel, 2 );
                   foreach ( range(100, 119) as $index) Style_group_Col($col_name, $index, $objPHPExcel, 2 );
                   foreach ( range(121, 126) as $index) Style_group_Col($col_name, $index, $objPHPExcel, 2 );
                   foreach ( range(128, 139) as $index) Style_group_Col($col_name, $index, $objPHPExcel, 2 );
 
                   foreach ( range(2, 14) as $index) Style_group_lv1_Row($index, $objPHPExcel, true);
            #======================================================================== กรุป คอลัมป์  ====================================================================================                    
             

#========================================================================================================================  Put field ==================================================================================== 
            }
            elseif( $sheetIndex == 'code_detail' ) 
            {
                $objPHPExcel->getActiveSheet()->setTitle( "Code detail" );                
                //$objPHPExcel->getActiveSheet()->setShowGridlines(False);
                $objPHPExcel->setActiveSheetIndex($ind);
                //$objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('A1', 'TEST');
                //$objPHPExcel->getActiveSheet()->getStyle('A10')->getAlignment()->setTextRotation(45);
                $st_col = 2;
                $st_dat = 4;
                $count_index =  count($list_act_report[$sheetIndex][0]) - 1 ;
                $row = $st_dat;
                $i=0;
                $look_data = 0;
                $count_data  =  count($list_act_report[$sheetIndex]) + $row-1;
                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 10 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 30 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 10 );


                $objPHPExcel->getActiveSheet()->freezePane('A'.$row);   
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(100);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('C'.($st_col+1).':'.$col_name[$count_index].($st_col+1)); 
                                
                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {
                        $objPHPExcel->getActiveSheet()->setCellValue($col_name[$i++].($st_col), str_replace("_", " ", $key));
                } // exit;     
#========================================================================================================================  Put data ====================================================================================                

                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {               
                   $col = 0;
                    foreach ($value as $body => $val) 
                    {
                            $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col++].($row), $val);
                                                            
                    }//exit;
                    $row++;               
                }
                $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('2');              #A
                $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('14');     #B no
                $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('50');     #D plnt
                $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('50');     #C pd                
                $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('23');     #E so_no

                $objPHPExcel->getActiveSheet()->getStyle( 'B' . $count_data . ":" . 'B' . $count_data )->getNumberFormat()->setFormatCode('00#');
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
$con = 'Content-Disposition: attachment;filename='.$filename.'.xlsx';
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

function Style_group_Col($cell=null, $index=0, $objPHPExcel=null, $level=1, $vi=false, $co=true)
{
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setOutlineLevel($level);
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