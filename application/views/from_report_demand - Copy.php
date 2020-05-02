<?php
//error_reporting(E_ALL);
error_reporting(E_ALL ^ E_NOTICE);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Asia/Bangkok');
ini_set('max_execution_time', 300); 
ini_set('memory_limit','2048M');
if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/PHPExcel-1.8.1/Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();
$data_col = array();
//var_dump($list_act_report); exit;
$col_name = array();
foreach ( range('C', 'Z') as $cm ) { array_push($col_name, $cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "A".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "B".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "C".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "D".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "E".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "F".$cm); }

$look_month =  31 - date('t');
$limit_dat  =  date('t', strtotime( date('d-m-y') ) ) + 0;
$limit_col  =  ($limit_dat == 31) ? $limit_dat."st" : $limit_dat . "th" ;
$ind = 0;

//echo $limit_col . " " . $look_month . " " . date('t', strtotime( date('d-m-y') ) ); exit;
echo "Start " . date('Y-m-d H:i:s') . "<hr>";
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

                $objPHPExcel->getActiveSheet()->setTitle( "$til" );
                $objPHPExcel->getActiveSheet()->setShowGridlines(False);
                $st_col = 15;
                $st_dat = 17;
                $count_index =  count( $list_act_report[$sheetIndex][0] ) - ($look_month+1) ;
                $row = $st_dat;
                $look_data = 0;
                $count_data  =  count( $list_act_report[$sheetIndex] ) + $row-1;

                $gdImage = dirname(__FILE__) . '/img/NEW-TBKK-LOGO_0.png';

                $opt = array('demand', 'prod. Plan', 'prod. Act.', 'accm. Diff', 'stock', 'stock lvl', 'del. box', 'plan box', 'reserv box', 'Total box' );
// demand
// prod. Plan
// prod. Act.
// accm. Diff
// stock
// stock lvl.
// Del_Box
// Plan_Box
// Reserv_Box
#========================================================================================================================  set head ====================================================================================

                // Add a drawing to the worksheetecho date('H:i:s') . " Add a drawing to the worksheet\n";
                $objDrawing = new PHPExcel_Worksheet_Drawing();
                $objDrawing->setName('Sample image');
                $objDrawing->setDescription('Sample image');
                $objDrawing->setPath($gdImage);
                //$objDrawing->setRenderingFunction(PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG);
                //$objDrawing->setMimeType(PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT);
                $objDrawing->setOffsetX(20); 
                $objDrawing->setOffsetY(10);  
                $objDrawing->setHeight(160);
                $objDrawing->setWidth(145); 
                $objDrawing->setCoordinates('B2');
                $objDrawing->setWorksheet($objPHPExcel->getActiveSheet()); 

                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 7 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 7 );

                foreach (range(4, 13) as $id ) 
                $objPHPExcel->getActiveSheet()->getRowDimension( $id )->setRowHeight( 21 ); 

                $objPHPExcel->getActiveSheet()->getRowDimension( $st_col-1 )->setRowHeight( 8 );               
                $objPHPExcel->getActiveSheet()->getRowDimension( $st_col )->setRowHeight( 35 );
                $objPHPExcel->getActiveSheet()->getRowDimension( $st_col+1 )->setRowHeight( 8 );

                foreach (range( $st_dat, $count_data ) as $id ) 
                $objPHPExcel->getActiveSheet()->getRowDimension( $id )->setRowHeight( 18 );


                $objPHPExcel->getActiveSheet()->freezePane('K'.$row);   
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('C'.($st_col+1).':'. $col_name[$count_index] .($st_col+1));   
#====================================================================== เส้นตารางข้อมูล =============================================================================# 

                $objPHPExcel->getActiveSheet()->getStyle( "C" . $st_col  . ':' . $col_name[$count_index]   . ($st_col+1) )
                                              ->applyFromArray(array(
                                                'borders' => array('allborders'   => Style_border(PHPExcel_Style_Border::BORDER_THIN ,'ffd966')))); 


                $objPHPExcel->getActiveSheet()->getStyle( "J" . '4' . ':' . $col_name[$count_index]   . '13' )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN ,'00cc99'))));     

                $objPHPExcel->getActiveSheet()->getStyle( "C" . $st_dat  . ':' . $col_name[$count_index]   . $count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN ,'00cc99'))));   


                $objPHPExcel->getActiveSheet()->getStyle( $col_name[$count_index+3] . '4'  . ':' . $col_name[$count_index+3]   . '13')
                                              ->applyFromArray(array('borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'000000'))));

                $objPHPExcel->getActiveSheet()->getStyle( $col_name[$count_index+3] . '4'  . ':' . $col_name[$count_index+4]   . '13')
                                              ->applyFromArray(array('borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THIN ,'00cc99'))));      

                $objPHPExcel->getActiveSheet()->getStyle( $col_name[$count_index+3] . '4'  . ':' . $col_name[$count_index+3]   . '13')
                                              ->applyFromArray(array('borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THIN ,'00cc99'))));  
#======================================================================== กำหนดสี fill ========================================================# 

                $objPHPExcel->getActiveSheet()->getStyle("B" . '2'  . ':' . $col_name[$count_index+1] . ($count_data+1) )->applyFromArray(array('fill' => Style_Fill('00cc99')));

                $objPHPExcel->getActiveSheet()->getStyle( $col_name[$count_index+4] . '4'  . ':' . $col_name[$count_index+4]   . '13' )->applyFromArray(array('fill' => Style_Fill('00cc99')));

                $objPHPExcel->getActiveSheet()->getStyle("K" . '4'      . ':' . $col_name[$count_index]   . '13'       )->applyFromArray(array('fill' => Style_Fill('FFFFFF'))); 
                $objPHPExcel->getActiveSheet()->getStyle("K" . $st_dat  . ':' . $col_name[$count_index]   . $count_data)->applyFromArray(array('fill' => Style_Fill('FFFFFF'))); 

                $objPHPExcel->getActiveSheet()->getStyle("C" . $st_col  . ':' . $col_name[$count_index]   . $st_col)->applyFromArray(array('fill' => Style_Fill('f4b084')));

                $objPHPExcel->getActiveSheet()->getStyle("C" . ($st_col+1)  . ':' . $col_name[$count_index]   . ($st_col+1) )->applyFromArray(array('fill' => Style_Fill('ffd966'))); 


                $objPHPExcel->getActiveSheet()->getStyle("J" . '4'  . ':' . 'J'   . '13' )->applyFromArray(array('fill' => Style_Fill('ffe699')));

                $objPHPExcel->getActiveSheet()->getStyle("J" . $st_dat  . ':' . 'J'   . $count_data )->applyFromArray(array('fill' => Style_Fill('ffe699')));

                $objPHPExcel->getActiveSheet()->getStyle("K" . '4'  . ':' . 'K'   . '13' )->applyFromArray(array('fill' => Style_Fill('D5D8DC')));

                $objPHPExcel->getActiveSheet()->getStyle("K" . $st_dat  . ':' . 'K'   . $count_data )->applyFromArray(array('fill' => Style_Fill('D5D8DC')));

                $objPHPExcel->getActiveSheet()->getStyle("C" . $st_dat  . ':' . 'I'   . ($st_dat+9) )->applyFromArray(array('fill' => Style_Fill('EAF2F8')));
#D5D8DC

                $objPHPExcel->getActiveSheet()->getStyle("K" . '4'  . ':' .  $col_name[$count_index]   .  '4' )->applyFromArray(array('fill' => Style_Fill('F5B7B1')));
                
                $objPHPExcel->getActiveSheet()->getStyle("K" . '6'  . ':' .  $col_name[$count_index]   .  '6' )->applyFromArray(array('fill' => Style_Fill('D5F5E3')));

                $objPHPExcel->getActiveSheet()->getStyle("K" . '10'  . ':' .  $col_name[$count_index]   . '10' )->applyFromArray(array('fill' => Style_Fill('FAD7A0')));

                $objPHPExcel->getActiveSheet()->getStyle("K" . '11'  . ':' .  $col_name[$count_index]   . '11' )->applyFromArray(array('fill' => Style_Fill('FCF3CF')));

                $objPHPExcel->getActiveSheet()->getStyle("K" . '12'  . ':' .  $col_name[$count_index]   . '12' )->applyFromArray(array('fill' => Style_Fill('F6DDCC')));

                $objPHPExcel->getActiveSheet()->getStyle("K" . '13'  . ':' .  $col_name[$count_index]   . '13' )->applyFromArray(array('fill' => Style_Fill('2980B9')));
#======================================================================== กำหนดสี merge cell ========================================================# 
                     
                foreach(range('C', 'I') as $cel) $objPHPExcel->getActiveSheet()->mergeCells( $cel . $st_dat  . ':' . $cel   . ($st_dat+9) );


                $objPHPExcel->getActiveSheet()->mergeCells( $col_name[$count_index+4] . '4'  . ':' . $col_name[$count_index+4]   . '13' );
                                          
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['head'][0] . '15' . ':'. $style_layout['head'][1] . '19')->applyFromArray(array('fill' => Style_Fill($color_border['deta'])));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['head'][0] . '21' . ':'. $style_layout['head'][1] . '21')->applyFromArray(array('fill' => Style_Fill($color_border['deta'])));



                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['rm'][0]   . '15' . ':'. $style_layout['rm'][1]   . '18')->applyFromArray(array('fill' => Style_Fill($color_border['rm'])));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['ma'][0]   . '15' . ':'. $style_layout['ma'][1]   . '18')->applyFromArray(array('fill' => Style_Fill($color_border['ma'])));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['summ'][0] . '15' . ':'. $style_layout['summ'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['summ'])));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['cost'][0] . '15' . ':'. $style_layout['cost'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['cost'])));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['cost'][0] . '19' . ':'. $style_layout['cost'][1] . '19')->applyFromArray(array('fill' => Style_Fill($color_border['cost'])));

                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['as'][0] . '15' . ':' . $style_layout['as'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['as'])));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['di'][0] . '15' . ':' . $style_layout['di'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['di'])));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['pe'][0] . '15' . ':' . $style_layout['pe'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['pe'])));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['oh'][0] . '15' . ':' . $style_layout['oh'][1] . '18')->applyFromArray(array('fill' => Style_Fill($color_border['oh'])));

                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['head'][0]  . ($st_col+1) .':'. $style_layout['head'][1]  . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['deta'])));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['cost'][0]  . ($st_col+1) .':'. $style_layout['cost'][1]  . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['cost']  )));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['summ'][0]  . ($st_col+1) .':'. $style_layout['summ'][1]  . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['summ']  )));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['rm'][0]    . ($st_col+1) .':'. $style_layout['rm'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['rm']  )));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['ma'][0]    . ($st_col+1) .':'. $style_layout['ma'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['ma']  )));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['as'][0]    . ($st_col+1) .':'. $style_layout['as'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['as']  )));                
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['di'][0]    . ($st_col+1) .':'. $style_layout['di'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['di']  )));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['pe'][0]    . ($st_col+1) .':'. $style_layout['pe'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['pe']  )));
                // $objPHPExcel->getActiveSheet()->getStyle($style_layout['oh'][0]    . ($st_col+1) .':'. $style_layout['oh'][1]    . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['oh']  )));
#======================================================================== กำหนดขนาด คอลัมป์ ========================================================# 

                 $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('1.71');
                 $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('2.71');              
                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$count_index+1])->setWidth('2.71'); 
                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$count_index+2])->setWidth('2.71');

                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$count_index+3])->setWidth('14.71');
                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$count_index+4])->setWidth('3.71');

                 $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('10.71');    

                 $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('12.71');

                 $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('59.71');                  
                 $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('14.71');    
                 $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('20.71');    
                 $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('10.71');    
                 $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('10.71');    
                 $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('14.71');
                 $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth('14.71'); 
     
                 foreach ( range(9, $count_index-1 ) as $id )
                    $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$id])->setWidth('10.29');
#======================================================================== กำหนด ฟ้อร์น และ สี ==================================================================================== 
                                                     
                   $objPHPExcel->getActiveSheet()->getStyle('C8') ->applyFromArray(array('font' => Style_Font(14,"000000",True,False,'Leelawadee'))); 
                   $objPHPExcel->getActiveSheet()->getStyle('C9') ->applyFromArray(array('font' => Style_Font(8,"000000",True,False,'Arial Rounded MT Bold'))); 
                   $objPHPExcel->getActiveSheet()->getStyle('C13')->applyFromArray(array('font' => Style_Font(8,"000000",True,False,'Arial Rounded MT Bold')));


                   $objPHPExcel->getActiveSheet()->getStyle('C'.$st_col . ':' . 'J'.($st_col) )->applyFromArray(array('font' => Style_Font(12,"000099",True,False,'Calibri')));
                   
                   $objPHPExcel->getActiveSheet()->getStyle('K'.$st_col . ':' . $col_name[$count_index].($st_col) )->applyFromArray(array('font' => Style_Font(10,"000099",True,True,'Bahnschrift Light')));

                   $objPHPExcel->getActiveSheet()->getStyle("J" . '4'  . ':' . 'J'   . '13' )->applyFromArray(array('font' => Style_Font(11,"000000",True,True,'Calibri')));
                   $objPHPExcel->getActiveSheet()->getStyle("K" . '4'  . ':' . 'K'   . '13' )->applyFromArray(array('font' => Style_Font(11,"000000",True,False,'Calibri')));

                   $objPHPExcel->getActiveSheet()->getStyle("J" . $st_dat  . ':' . 'J'   . $count_data )->applyFromArray(array('font' => Style_Font(10,"000000",True,True,'Calibri')));
                   $objPHPExcel->getActiveSheet()->getStyle("K" . $st_dat  . ':' . 'K'   . $count_data )->applyFromArray(array('font' => Style_Font(10,"000000",True,False,'Calibri')));

                   $objPHPExcel->getActiveSheet()->getStyle("C" . $st_dat  . ':' . 'I'   . $count_data )->applyFromArray(array('font' => Style_Font(11,"000000",True,False,'Calibri')));
                   // $objPHPExcel->getActiveSheet()->getStyle('K19')->applyFromArray(array('font' => Style_Font(12,"000000",True,False,'Arial Rounded MT Bold')));

                   // $objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray(array('font' => Style_Font(26,"000000",True,False,'Arial Rounded MT Bold')));
                   // $objPHPExcel->getActiveSheet()->getStyle('C6')->applyFromArray(array('font' => Style_Font(18,"000000",True,False,'Arial Rounded MT Bold')));
                   // $objPHPExcel->getActiveSheet()->getStyle('H12')->applyFromArray(array('font' => Style_Font(12,"000000",False,False,'Arial Rounded MT Bold')));
                   // $objPHPExcel->getActiveSheet()->getStyle('H13')->applyFromArray(array('font' => Style_Font(11,"000000",False,False,'Arial Rounded MT Bold')));

                   // $objPHPExcel->getActiveSheet()->getStyle('S2')->applyFromArray(array('font' => Style_Font(12,"000000",True,False)));
                   // $objPHPExcel->getActiveSheet()->getStyle('S3:S13')->applyFromArray(array('font' => Style_Font(12,"000000",True,False)));
                   // $objPHPExcel->getActiveSheet()->getStyle('S4:S14')->applyFromArray(array('font' => Style_Font(11,"000000",True,False)));
                
                   // $objPHPExcel->getActiveSheet()->getStyle('U3:AG13')->applyFromArray(array('font' => Style_Font(11,"000000",True,False)));


                   //$objPHPExcel->getActiveSheet()->getStyle('S19')->applyFromArray(array('font' => Style_Font(12,"000000",True,False,'Arial Rounded MT Bold')));

                   //$objPHPExcel->getActiveSheet()->getStyle('I2')->applyFromArray(array('font' => Style_Font(16,"000000",True,False)));

                   // $objPHPExcel->getActiveSheet()->getStyle('S19:EL21')->applyFromArray(array('font' => Style_Font(12,"000000",True,False,'Arial Rounded MT Bold')));

                   // $objPHPExcel->getActiveSheet()->getStyle('B'. $st_col . ':'.$col_name[$count_index].($st_col)) ->applyFromArray(array('font' => Style_Font(12,"000000",True,False,'Arial Rounded MT Bold')));

                   // $objPHPExcel->getActiveSheet()->getStyle('S'. $st_dat . ':'.$col_name[$count_index].($count_data)) ->applyFromArray(array('font' => Style_Font(11,"000000",True,False,'Arial Rounded MT Bold')));


                   // $objPHPExcel->getActiveSheet()->getStyle('AD19:AD21')->applyFromArray(array('font' => Style_Font(24,"000000",True,False,'Wingdings')));
                   // $objPHPExcel->getActiveSheet()->getStyle('AG19:AG21')->applyFromArray(array('font' => Style_Font(24,"000000",True,False,'Wingdings')));
#======================================================================== พิเศษ             ========================================================================================
#======================================================================== กำหนด ฟอรืแมท ข้อมุล  ==================================================================================== 

                    $objPHPExcel->getActiveSheet()->getStyle("K" . '4'      . ':' . $col_name[$count_index+3]   . '13'       )->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');


                    $objPHPExcel->getActiveSheet()->getStyle("K" . '9'      . ':' . $col_name[$count_index+3]   . '9'       )->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');


                    $objPHPExcel->getActiveSheet()->getStyle("K" . $st_dat  . ':' . $col_name[$count_index]   . $count_data)->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');


                    $objPHPExcel->getActiveSheet()->getStyle( "J" . '7' . ':' .  $col_name[$count_index]   .  '7' )->applyFromArray(array('font' => Style_Font(10,"FF0000",True,False,'Calibri')));

                    $objPHPExcel->getActiveSheet()->getStyle("K" . '13'  . ':' .  $col_name[$count_index]   . '13' )->applyFromArray(array('font' => Style_Font(11,"FFE5CC",True,False,'Calibri')));



                    $objPHPExcel->getActiveSheet()->getStyle( $col_name[$count_index+3] . '4' . ':' .  $col_name[$count_index+3]   .  '13' )->applyFromArray(array('font' => Style_Font(11,"000000",True,False,'Calibri')));

                    $objPHPExcel->getActiveSheet()->getStyle( $col_name[$count_index+3] . '7' . ':' .  $col_name[$count_index+3]   .  '7' )->applyFromArray(array('font' => Style_Font(11,"FF0000",True,False,'Calibri')));

                    $objPHPExcel->getActiveSheet()->getStyle($col_name[$count_index+4]   . '4' )->applyFromArray(array('font' => Style_Font(18,"000000",True,False,'Calibri')));

                    // $objPHPExcel->getActiveSheet()->getStyle('AD5:AG13')->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
                                                  
                    // $objPHPExcel->getActiveSheet()->getStyle('K'.$st_dat.':'.'Q'.$count_data)->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');

                    // $objPHPExcel->getActiveSheet()->getStyle('S'.$st_dat.':'.$col_name[$count_index].$count_data)->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
                    
                    // $objPHPExcel->getActiveSheet()->getStyle('AI'.$st_col.':'.$col_name[$count_index].$st_col)->getNumberFormat()->setFormatCode('00#');                                                 

                    // $objPHPExcel->getActiveSheet()->getStyle('L'.'21'.':'.'Q'.'21')->getNumberFormat()->setFormatCode('#,##0.00');

                    // $objPHPExcel->getActiveSheet()->getStyle('W'.'19'.':'.$col_name[$count_index].'19')->getNumberFormat()->setFormatCode('#,##0.00');

                    // $objPHPExcel->getActiveSheet()->getStyle('S'.'21'.':'.$col_name[$count_index].'21')->getNumberFormat()->setFormatCode('#,##0');
#======================================================================== การ input ========================================================================================    
                $limit = 0;
                $i     = 0;
                $sub_total = array( array(), array(), array(), array(), array(), array(), array(), array(), array(),  array() ); 
                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {

                    

                        
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$i++].$st_col, str_replace("_", " ", $key));

                    if ( $key == $limit_col ) break;

                } // exit;     
                    
                //echo $list_act_report[$sheetIndex][0]['ITEM_CD']; exit;

                $item_ck = $list_act_report[$sheetIndex][0]['ITEM_CD'];
                $use_item  = array( array(), array(), array(), array(), array(), array(), array(), array(), array(),  array() );

                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {               
                    $col = 0;




                        foreach ($value as $body => $val) 
                        {

                            

                            $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col++].($row), $val);
                              
                            if($val == 3 && $body == 'MODEL')  $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('###"E00"');


                            if ( $body == $limit_col ) break;
                        }

                        if( $item_ck !=  $value['ITEM_CD'])
                        {


                            if( $limit == 0 )
                            {
                                $objPHPExcel->getActiveSheet()->getStyle("C" . $row  . ':' . 'I'   . ($row+9) )->applyFromArray(array('fill' => Style_Fill('fce4f4')));
                                $limit = 1;
                            }
                            else
                            {

                                $objPHPExcel->getActiveSheet()->getStyle("C" . $row  . ':' . 'I'   . ($row+9) )->applyFromArray(array('fill' => Style_Fill('D1F2EB')));  
                                $limit = 0;
                            }

                            $objPHPExcel->getActiveSheet()->getStyle( "C" . ($row-1)  . ':' . $col_name[$count_index]   . ($row-1) )
                                                          ->applyFromArray(array( 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK ,'00cc99'))));

                           foreach(range('C', 'I') as $cel) 
                              $objPHPExcel->getActiveSheet()->mergeCells( $cel . ($row)  . ':' . $cel   . ($row+9) );
                           

                            $item_ck = $value['ITEM_CD'];

                        }

                    //$objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].($row), "=SUM(" . "L" . $row . ":" . $col_name[$count_index] . $count_index . ")" );
                    // Style_group_lv1_Row( 9,  $objPHPExcel, True, False );
                    // Style_group_lv1_Row( 10, $objPHPExcel, True, False );
                    // Style_group_lv1_Row( 11, $objPHPExcel, True, False );
                        switch ( $value['DM_TYPE'] ) 
                        {
                            case '1':
                                $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $row,  $opt[$value['DM_TYPE']-1] );
                                $objPHPExcel->getActiveSheet()->getStyle("K" . $row  . ':' .  $col_name[$count_index]   . ($row) )->applyFromArray(array('fill' => Style_Fill('F5B7B1')));
                                if( dup_item( $use_item[$value['DM_TYPE']-1], $value['ITEM_CD'] ) ){  array_push($sub_total[$value['DM_TYPE']-1], $row);  array_push($use_item[$value['DM_TYPE']-1], $value['ITEM_CD']);     }
                                
                                break;
                            case '2':
                                $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $row,  $opt[$value['DM_TYPE']-1] );
                                array_push($sub_total[$value['DM_TYPE']-1], $row);
                                break;
                            case '3':
                                $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $row,  $opt[$value['DM_TYPE']-1] );
                                $objPHPExcel->getActiveSheet()->getStyle("K" . $row  . ':' .  $col_name[$count_index]   . ($row) )->applyFromArray(array('fill' => Style_Fill('D5F5E3')));
                                array_push($sub_total[$value['DM_TYPE']-1], $row);
                                break;
                            case '4':
                                $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $row,  $opt[$value['DM_TYPE']-1] );
                                $objPHPExcel->getActiveSheet()->getStyle( "J" . $row  . ':' .  $col_name[$count_index]   . ($row) )->applyFromArray(array('font' => Style_Font(10,"FF0000",True,False,'Calibri')));
                                array_push($sub_total[$value['DM_TYPE']-1], $row);
                                //$objPHPExcel->getActiveSheet()->getStyle("C" . $row  . ':' . 'I'   . ($row+9) )->applyFromArray(array('fill' => Style_Fill('fce4f4')));
                                break;
                            case '5':
                                $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $row,  $opt[$value['DM_TYPE']-1] );
                                 if( dup_item( $use_item[$value['DM_TYPE']-1], $value['ITEM_CD'] ) ){  array_push($sub_total[$value['DM_TYPE']-1], $row);  array_push($use_item[$value['DM_TYPE']-1], $value['ITEM_CD']);     }
                                
                                //$objPHPExcel->getActiveSheet()->getStyle("C" . $row  . ':' . 'I'   . ($row+9) )->applyFromArray(array('fill' => Style_Fill('fce4f4')));
                                break;
                            case '6':
                                $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $row,  $opt[$value['DM_TYPE']-1] );
                                $objPHPExcel->getActiveSheet()->getStyle( "K" . $row  . ':' . $col_name[$count_index]   . $row )->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
                                if( dup_item( $use_item[$value['DM_TYPE']-1], $value['ITEM_CD'] ) ){  array_push($sub_total[$value['DM_TYPE']-1], $row);  array_push($use_item[$value['DM_TYPE']-1], $value['ITEM_CD']);     }

                                //$objPHPExcel->getActiveSheet()->getStyle("C" . $row  . ':' . 'I'   . ($row+9) )->applyFromArray(array('fill' => Style_Fill('fce4f4')));
                                break;
                            case '7':
                                $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $row,  $opt[$value['DM_TYPE']-1] );
                                $objPHPExcel->getActiveSheet()->getStyle("K" . $row  . ':' .  $col_name[$count_index]   . ($row) )->applyFromArray(array('fill' => Style_Fill('FAD7A0')));
                                if( dup_item( $use_item[$value['DM_TYPE']-1], $value['ITEM_CD'] ) ){  array_push($sub_total[$value['DM_TYPE']-1], $row);  array_push($use_item[$value['DM_TYPE']-1], $value['ITEM_CD']);     }
                                Style_group_lv1_Row( $row, $objPHPExcel, False, False );
                                break;
                            case '8':
                                $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $row,  $opt[$value['DM_TYPE']-1] );
                                $objPHPExcel->getActiveSheet()->getStyle("K" . $row  . ':' .  $col_name[$count_index]   . ($row) )->applyFromArray(array('fill' => Style_Fill('FCF3CF')));
                                array_push($sub_total[$value['DM_TYPE']-1], $row);
                                Style_group_lv1_Row( $row, $objPHPExcel, False, False );
                                break;
                            case '9':
                                $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $row,  $opt[$value['DM_TYPE']-1] );
                                $objPHPExcel->getActiveSheet()->getStyle("K" . $row  . ':' .  $col_name[$count_index]   . ($row) )->applyFromArray(array('fill' => Style_Fill('F6DDCC')));
                                if( dup_item( $use_item[$value['DM_TYPE']-1], $value['ITEM_CD'] ) ){  array_push($sub_total[$value['DM_TYPE']-1], $row);  array_push($use_item[$value['DM_TYPE']-1], $value['ITEM_CD']);     }
                                Style_group_lv1_Row( $row, $objPHPExcel, False, False );
                                break;
                            case '10':
                                $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $row,  $opt[$value['DM_TYPE']-1] );
                                $objPHPExcel->getActiveSheet()->getStyle("K" . $row  . ':' .  $col_name[$count_index]   . ($row) )->applyFromArray(array('fill' => Style_Fill('2980B9')));
                                $objPHPExcel->getActiveSheet()->getStyle("K" . $row  . ':' .  $col_name[$count_index]   . ($row) )->applyFromArray(array('font' => Style_Font(11,"FFE5CC",True,False,'Calibri')));
                                array_push($sub_total[$value['DM_TYPE']-1], $row);
                                
                                break;                                                                                                                                                                        
                        }
//echo "<br>";
                    $row++;           
                }
//var_dump($sub_total); exit;
#======================================================================== กำหนด สูตร excel  ====================================================================================                
                $p_stock = ( date('d',strtotime( date('Y-m-d') ) ) == '01' ) ?   date('d',strtotime( date('Y-m-d') ) ) : date( 'd', strtotime( "-1 day", strtotime( date('Y-m-d') ) ) ) + 8 ;

                            $objConditional1 = new PHPExcel_Style_Conditional();
                            $objConditional1->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
                                        ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHAN)
                                        ->addCondition( 2 )
                                        ->getStyle()->applyFromArray(array( 'font' => Style_Font(11,'FF0000',true,false),
                                                                            'fill' => Style_Fill_Con('FFFF00','FFFF00')
                                                                          ));

            	
                //$objPHPExcel->getActiveSheet()->getStyle( 'L9:'. $col_name[$count_index]. '9' )->setConditionalStyles( array($objConditional1) );

                //echo $p_stock; exit;
                for ( $rw=$st_dat; $rw < $count_data  ; $rw += 10 ) 
                {

                    $objPHPExcel->getActiveSheet()->setCellValue( 'L'  . ($rw+3),  "=" . "L" .  ($rw+2) . "-" . "L" .  ($rw+1)   );

                    $objPHPExcel->getActiveSheet()->setCellValue( 'L' . ($rw+5),  "=IFERROR( IF(L" . ($rw+4) . "> 0 ," .  'L'  . ($rw+4) . "/" .  "AVERAGEIF(" . 'M' . ($rw) . ":" . $col_name[$count_index] . ($rw) . ',"<>0"),' . '0' . ' ) , 0 )'   );

					           $objPHPExcel->getActiveSheet()->setCellValue( 'L' . ($rw+9),  "=SUM(" . 'L' . ($rw+6) . ":" . 'L' . ($rw+8) . ")"  );


                    $objPHPExcel->getActiveSheet()->setCellValue( 'K'  . ($rw+9),  "=SUM(" . 'K' . ($rw+6) . ":" . 'K' . ($rw+8) . ")"  );

                    $objPHPExcel->getActiveSheet()->setCellValue( 'K'  . ($rw+8),  "=" . "K" . ($rw+4) . "/" . $col_name[6] . ($rw+4)  );




                    if( date('d',strtotime( date('Y-m-d') ) ) == '01' )
                    {

                       // $objPHPExcel->getActiveSheet()->setCellValue( 'K' . ($rw+5),  "=IFERROR( IF(K" . ($rw+4) . "> 0 ," .  'K'  . ($rw+4) . "/" .  "AVERAGEIF(" . $col_name[$cel+1] . ($rw) . ":" . $col_name[$count_index] . ($rw) . ',"<>0"),' . $col_name[$cel-1]  . ($rw+5) . ' ) , 0 )'   );

                    	$objPHPExcel->getActiveSheet()->setCellValue( 'L' . ($rw+4),  "=" . 'K' . ($rw+4) . "-"  . "(" . 'L' . ($rw) . "-" . 'L' . ($rw+1) . ")"  );

                        //$objPHPExcel->getActiveSheet()->setCellValue( 'L'  . ($rw+8),  "=" . "K" .  ($rw+2) . "-" . "K" .  ($rw+1)   );
                        //$objPHPExcel->getActiveSheet()->setCellValue( 'L'  . ($rw+9),  "=" . "K" .  ($rw+2) . "-" . "K" .  ($rw+1)   );
                    }
                    foreach (range( 10, (8+$limit_dat) ) as $cel) 
                    {
                        $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel] . ($rw+3),  "=" . $col_name[$cel-1] . ($rw+3) . "+" . "(" . $col_name[$cel] . ($rw+2) . "-" . $col_name[$cel] .  ($rw+1)  . ")"  );

                        
                        if( $cel >= $p_stock )
                        {
                                if($cel != $p_stock)
                                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel] . ($rw+4),  "=" . $col_name[$cel-1] . ($rw+4) . "-"  . "(" . $col_name[$cel] . ($rw) . "-" . $col_name[$cel] . ($rw+1) . ")"  );

                                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel] . ($rw+5),  "=IFERROR( IF(" .  $col_name[$cel]  . ($rw+4) . ">0 ," . $col_name[$cel]  . ($rw+4) . "/" .  "AVERAGEIF(" . $col_name[$cel+1] . ($rw) . ":" . $col_name[$count_index] . ($rw) . ',"<>0"),' . $col_name[$cel-1]  . ($rw+5) . ' ) , 0)'   );

                                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel] . ($rw+7),  "=" . $col_name[$cel] . ($rw+1) ."/" . $col_name[6] . ($rw+7)   );

                                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel] . ($rw+8),  "=ABS(" . $col_name[$cel] . ($rw+4) .")" ."/" . $col_name[6] . ($rw+8)   );

                                


                        }
                        //$objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel] . ($rw+3),  "=" . $col_name[$cel-1] . ($rw+3) . "+" . "(" . $col_name[$cel] . ($rw+2) . "-" . $col_name[$cel] .  ($rw+1)  . ")"  );

                        //$objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel] . ($rw+3),  "=" . $col_name[$cel-1] . ($rw+3) . "+" . "(" . $col_name[$cel] . ($rw+2) . "-" . $col_name[$cel] .  ($rw+1)  . ")"  );

                        $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel] . ($rw+9),  "=SUM(" . $col_name[$cel] . ($rw+6) . ":" . $col_name[$cel] . ($rw+8) . ")"  );



                    }
                    $objPHPExcel->getActiveSheet()->getStyle( 'L' . ($rw+5) . ":" . $col_name[$count_index] . ($rw+5) )->setConditionalStyles( array($objConditional1) );

                }
//exit;
            #======================================================================== จัดตำแหน่ง ข้อมูล ====================================================================================     


                //echo date('t F Y', strtotime( "$mnt month", strtotime( date('01-m-Y') ) ) ) ; exit;


                 $objPHPExcel->getActiveSheet()->setCellValue( 'C'  . '8',  'Provision List Report fo ' . date('Y F d') );
                 $objPHPExcel->getActiveSheet()->setCellValue( 'C'  . '9',  'TBKK (Thailand) Co., Ltd. ' );
                 $objPHPExcel->getActiveSheet()->setCellValue( 'C'  . '13', 'vol. 1.21  :  Issue by Pc System ' . date('d-m-Y') );

                 $objPHPExcel->getActiveSheet()->setCellValue( 'C'  . $st_col , 'PD' );
                 $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $st_col , 'Option' );
                 $objPHPExcel->getActiveSheet()->setCellValue( 'K'  . $st_col , 'Last Month' );


                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].'4' , "=SUM(" . "L" . '4'  . ":" . $col_name[$count_index] . '4'  . ")" );
                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].'5' , "=SUM(" . "L" . '5'  . ":" . $col_name[$count_index] . '5'  . ")" );
                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].'6' , "=SUM(" . "L" . '6'  . ":" . $col_name[$count_index] . '6'  . ")" );
                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].'7' , "=" . $col_name[$count_index+3] . '6'  . "-" . $col_name[$count_index+3] . '5' );
                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].'8' , "=" . $col_name[$count_index] . '8'   );
                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].'9' , "=" . $col_name[$count_index] . '9'   );
                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].'10', "=SUM(" . "L" . '10' . ":" . $col_name[$count_index] . '10' . ")" );
                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].'11', "=SUM(" . "L" . '11' . ":" . $col_name[$count_index] . '11' . ")" );
                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].'12', "=SUM(" . "L" . '12' . ":" . $col_name[$count_index] . '12' . ")" );
                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+3].'13', "=SUM(" . "L" . '13' . ":" . $col_name[$count_index] . '13' . ")" );


                 $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$count_index+4].'4', "TOTAL" );
                 
                 foreach ( range(4,13) as $key ) 
                 {

                            $objPHPExcel->getActiveSheet()->setCellValue( 'J'  . $key, $opt[$key-4] );              
                 }

                 $demand = array( "=SUBTOTAL(109,", "=SUBTOTAL(109,", "=SUBTOTAL(109,", "=SUBTOTAL(109,", "=SUBTOTAL(109,", "=SUBTOTAL(109,", "=SUBTOTAL(109,", "=SUBTOTAL(109,", "=SUBTOTAL(109,", "=SUBTOTAL(109," ) ;

                 //$str_rw = "";

                            
                            foreach ( range(4,13) as $r ) 
                            { 
                                foreach (range( 8, (8+$limit_dat) ) as $cel) 
                                {
                                    $str_rw = "";
                                    if( $r == 9)
                                    {
                                    	if($cel > 8)
                                        $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel]  . $r, "=IFERROR( " .  $col_name[$cel]  . '8' . "/" .  "AVERAGEIF(" . $col_name[$cel+1] . '4' . ":" . $col_name[$count_index] . '4' . ',"<>0"),' . $col_name[$cel-1]  . $r . ' )'  );
                                    	
                                    }
                                    else if ( $r < 13 )
                                    {
                                        foreach ($sub_total[$r-4] as $key => $value) 
                                        {  

                                            $str_rw .=  $col_name[$cel] . $value . ",";

                                        }

                                        $str_rw  = substr($str_rw, 0,-1) . ")" ;
                                     //echo $col_name[$cel]  . $r .  " =SUBTOTAL(109," . $str_rw . "<br>"; 
                                     
                                        $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel]  . $r, "=SUBTOTAL(109," . $str_rw );   
                                    }
                                    else
                                    {

                                        $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$cel]  . $r, "=SUM(" . $col_name[$cel].($r-3)  . ":" . $col_name[$cel].($r-1) . ")"  );
                                

                                    }
                                }
                              //exit;
                            }                  
                                            
            #======================================================================== จัดตำแหน่ง ข้อมูล ====================================================================================
                    Style_Alignment('C' . $st_col . ':'.$col_name[$count_index].($st_col),3, true, $objPHPExcel); 
                    Style_Alignment('C' . '7'  ,9, False, $objPHPExcel);    
                    Style_Alignment('C' . '9'  ,7, False, $objPHPExcel);   
                    Style_Alignment('C' . '13' ,7, False, $objPHPExcel);


                    Style_Alignment('K' . '3' . ':' . $col_name[$count_index] . '3' ,3, False, $objPHPExcel);  
                    Style_Alignment('C' . $st_dat . ':' . 'I' . $count_data ,3, False, $objPHPExcel);

                    Style_Alignment( $col_name[$count_index+4]. '4' ,3, False, $objPHPExcel);




                    $objPHPExcel->getActiveSheet()->getStyle( 'H' . $st_dat . ':' . 'H' . $count_data)->getAlignment()->setTextRotation(90);

                    $objPHPExcel->getActiveSheet()->getStyle( $col_name[$count_index+4]. '4' )->getAlignment()->setTextRotation(-90);

                    // Style_Alignment('S2:AG4',3, False, $objPHPExcel);
                    // Style_Alignment('S5:S13',9, False, $objPHPExcel);
                    // //Style_Alignment('T4:T12',9, False, $objPHPExcel);
                    // Style_Alignment('C25:J'.($count_data),9, False, $objPHPExcel); 

                    // Style_Alignment('S19:V19',3, False, $objPHPExcel);

                    // Style_Alignment('T19',9, False, $objPHPExcel);
                    // Style_Alignment('V19',9, False, $objPHPExcel);               
                   // Style_Alignment('B15'.':'.$col_name[$count_index].($st_col),9, true, $objPHPExcel);


                $ind_yes = date( 'd', strtotime( "-1 day", strtotime( date('Y-m-d') ) ) ) + 8;

                $ind_tod = date( 'd', strtotime( "-0 day", strtotime( date('Y-m-d') ) ) ) + 8;
#======================================================================== Fill set ====================================================================================                    
                foreach ( $holiday as $ind_cal => $val)
                {

                    $objPHPExcel->getActiveSheet()->getStyle( $col_name[ (int)substr( $val['d_t'],-2 ) + 8 ] . $st_dat  . ':' .  $col_name[ (int)substr( $val['d_t'],-2 ) + 8 ] . $count_data )->applyFromArray(array( 'fill' => Style_Fill('BFC9CA') ));


                    $objPHPExcel->getActiveSheet()->getStyle( $col_name[ (int)substr( $val['d_t'],-2 ) + 8 ] . '4'  . ':' .  $col_name[ (int)substr( $val['d_t'],-2 ) + 8 ] . '13' )->applyFromArray(array( 'fill' => Style_Fill('BFC9CA') ));


                    $objPHPExcel->getActiveSheet()->getStyle( $col_name[ (int)substr( $val['d_t'],-2 ) + 8 ] . $st_dat  . ':' .  $col_name[ (int)substr( $val['d_t'],-2 ) + 8 ] . $count_data )->applyFromArray(array('font' => Style_Font(11,"000FFF",False,False,'Calibri')));

                    $objPHPExcel->getActiveSheet()->getStyle( $col_name[ (int)substr( $val['d_t'],-2 ) + 8 ] . '4'  . ':' .  $col_name[ (int)substr( $val['d_t'],-2 ) + 8 ] . '13' )->applyFromArray(array('font' => Style_Font(11,"000FFF",False,False,'Calibri')));
                }

                if(date('d') != '01')
                $objPHPExcel->getActiveSheet()->getStyle( $col_name[ (int) $ind_yes ] . '3'  . ':' .  $col_name[ (int) $ind_yes ]   . '3' )->applyFromArray(array('fill' => Style_Fill('000FFF') ) );


                $objPHPExcel->getActiveSheet()->getStyle( $col_name[ (int) $ind_tod ] . '3'  . ':' .  $col_name[ (int) $ind_tod ]   . '3' )->applyFromArray(array('fill' => Style_Fill('FF0000') ) );
#======================================================================== Borders set ====================================================================================

                if(date('d') != '01')
                $objPHPExcel->getActiveSheet()->getStyle( $col_name[ (int) $ind_yes ] . '3'  . ':' .  $col_name[ (int) $ind_yes ] . $count_data ) ->applyFromArray(array('borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK ,'000FFF'))));

                $objPHPExcel->getActiveSheet()->getStyle( $col_name[ (int) $ind_tod ] . '3'  . ':' .  $col_name[ (int) $ind_tod ] . $count_data ) ->applyFromArray(array('borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK ,'FF0000'))));
                                            
#======================================================================== Font set ========================================================================================

                if(date('d') != '01')
                $objPHPExcel->getActiveSheet()->getStyle( $col_name[ (int) $ind_yes ] . '3'  . ':' .  $col_name[ (int) $ind_yes ]   . '3' )->applyFromArray(array('font' => Style_Font(11,"FBFF00",True,True,'Calibri')));


                $objPHPExcel->getActiveSheet()->getStyle( $col_name[ (int) $ind_tod ] . '3'  . ':' .  $col_name[ (int) $ind_tod ]   . '3' )->applyFromArray(array('font' => Style_Font(11,"FBFF00",True,True,'Calibri')));
#======================================================================== Font set ========================================================================================
                if(date('d') != '01')
                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[ (int) $ind_yes ] . '3', "Yesterday" );

                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[ (int) $ind_tod ] . '3', "Today" );
#======================================================================== merge crll  ==================================================================================== 
#======================================================================== merge crll  ==================================================================================== 
                    // $objPHPExcel->getActiveSheet()->mergeCells( 'C3:I5' );
                    // $objPHPExcel->getActiveSheet()->mergeCells( 'C6:I7' );
                    // $objPHPExcel->getActiveSheet()->mergeCells( 'H12:J12' );
                    // $objPHPExcel->getActiveSheet()->mergeCells( 'H13:J13' );

                    // $objPHPExcel->getActiveSheet()->mergeCells( 'S2:AG2' );
                    // $objPHPExcel->getActiveSheet()->mergeCells( 'S3:T4' );
                    // //$objPHPExcel->getActiveSheet()->mergeCells( 'AC3:AC4' );
                    // foreach ( range(5, 13) as $ro)
                    // $objPHPExcel->getActiveSheet()->mergeCells( 'S' . $ro . ':' . 'T' . $ro );
#======================================================================== กำหนด ฟอรืแมท ข้อมุล  ==================================================================================== 

                    // $objPHPExcel->getActiveSheet()->getStyle('U5:AC13')->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');

                    // $objPHPExcel->getActiveSheet()->getStyle('AD5:AG13')->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
                                                  
                    // $objPHPExcel->getActiveSheet()->getStyle('K'.$st_dat.':'.'Q'.$count_data)->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');

                    // $objPHPExcel->getActiveSheet()->getStyle('S'.$st_dat.':'.$col_name[$count_index].$count_data)->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
                    
                    // $objPHPExcel->getActiveSheet()->getStyle('AI'.$st_col.':'.$col_name[$count_index].$st_col)->getNumberFormat()->setFormatCode('00#');                                                 

                    // $objPHPExcel->getActiveSheet()->getStyle('L'.'21'.':'.'Q'.'21')->getNumberFormat()->setFormatCode('#,##0.00');

                    // $objPHPExcel->getActiveSheet()->getStyle('W'.'19'.':'.$col_name[$count_index].'19')->getNumberFormat()->setFormatCode('#,##0.00');

                    // $objPHPExcel->getActiveSheet()->getStyle('S'.'21'.':'.$col_name[$count_index].'21')->getNumberFormat()->setFormatCode('#,##0');
#======================================================================== กรุป คอลัมป์  ==================================================================================== 
                    Style_group_Col($col_name, 2, $objPHPExcel);
                    Style_group_Col($col_name, 4, $objPHPExcel);

                    Style_group_lv1_Row( 10,  $objPHPExcel, False, False );
                    Style_group_lv1_Row( 11, $objPHPExcel, False, False );
                    Style_group_lv1_Row( 12, $objPHPExcel, False, False );
                    $p_stock = date( 'd', strtotime( "-1 day", strtotime( date('Y-m-d') ) ) ) + 8 ;
                    $to_date = date( 'd', strtotime( date('Y-m-d')  ) ) + 0;

                    if($to_date >= 8)
                    {

                        foreach ( range(8, ($to_date+8)-7 ) as $hide) 
                        {
                            Style_group_Col($col_name, $hide, $objPHPExcel);
                        }


                    }
                    else
                    {

                        foreach (range( 24, $count_index ) as $hide) 
                        {
                            Style_group_Col($col_name, $hide, $objPHPExcel);
                        }                    	
                    }
#======================================================================== กรุป คอลัมป์  ====================================================================================                    
#========================================================================================================================  Put field ===================================================================================       
#========================================================================================================================  Put data ====================================================================================         
    } else {
                    $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "No data ".$til.".");
                    $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true,false,'Franklin Gothic Book')));
    }
$ind++;
echo $til . "<br>" ;
}
echo "<hr>" ."Stop " . date('Y-m-d H:i:s');

//exit;
$objPHPExcel->setActiveSheetIndex(0);

$objPHPExcel->removeSheetByIndex(count($title));                             
                           
$today = date("My");
// header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
// $con = 'Content-Disposition: attachment;filename='.$filename.'.xlsx';
// header($con);
// header('Cache-Control: max-age=0');
// header('Cache-Control: max-age=1');
// header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
// header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
// header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
// header ('Pragma: public'); // HTTP/1.0
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save( "G:/vbs_demand/bin/" .$filename.'.xlsx');
exit;

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function dup_item($data, $item)
{

    foreach ( $data as $key => $data_val ) 
    {
        if( $data_val == $item )
        { //echo $data_val . "<br>"; 

        return False;

    	} 
    }




    return True;
}






function Style_Fill($color=null) {

    return array( 'type'  => PHPExcel_Style_Fill::FILL_SOLID,                           
                  'color' => array('rgb' => $color)                                    
                );                                   
}
function Style_Fill_Con($color_st=null, $color_en=null) {

    return array( 'type'       => PHPExcel_Style_Fill::FILL_SOLID,                           
                  'startcolor' =>array('argb' => $color_st),
                  'endcolor'   =>array('argb' => $color_en)                             
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