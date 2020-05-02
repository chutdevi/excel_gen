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
                $objDrawing->setOffsetX(10); 
                $objDrawing->setOffsetY(10);  
                $objDrawing->setHeight(250);
                $objDrawing->setWidth(250); 
                $objDrawing->setCoordinates('B2');
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

                $color_border['head'] = '9999ff';
                $color_border['targ'] = '9999ff';
                $color_border['deta'] = '9999ff';
                $color_border['cost'] = 'ffff99';
                $color_border['summ'] = '9999ff';
                $color_border['rm']   = 'b3ff99';
                $color_border['ma']   = 'ff9999';
                $color_border['as']   = 'ffcc99';
                $color_border['di']   = 'ffe699';
                $color_border['pe']   = '99ffcc';
                $color_border['oh']   = 'cc99ff';

                $style_layout['head'] = 'B2:J13';
                $style_layout['targ'] = 'S2:AC13';
                $style_layout['cost'] = 'K15:Q18';
                $style_layout['summ'] = 'S15:AC18';
                $style_layout['rm']   = 'AD15:AP18';
                $style_layout['ma']   = 'AQ15:CA18';
                $style_layout['as']   = 'CB15:CO18';
                $style_layout['di']   = 'CP15:DC18';
                $style_layout['pe']   = 'DD15:DI18';
                $style_layout['oh']   = 'DJ15:DS18';

                $objPHPExcel->getActiveSheet()->freezePane('A'.$row);   
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(59);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('B'.($st_col+1).':'.$col_name[$count_index].($st_col+1));            

                $objPHPExcel->getActiveSheet()->getStyle( 'B' . $st_dat . ':' . $col_name[15].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'3e4241'))));

                $objPHPExcel->getActiveSheet()->getStyle( 'S' . $st_dat . ':' . $col_name[$count_index].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_HAIR ,'3e4241'))));

            #================================================================================================================================================================#   

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['head'] )
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['head'])))); 

                $objPHPExcel->getActiveSheet()->getStyle( $style_layout['targ'] )
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['targ']))));  

                $objPHPExcel->getActiveSheet()->getStyle( "S3:AC4"  ) 
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['targ']),
                                                                   'top'      => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF"),
                                                                   'bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['targ']))));  



                $objPHPExcel->getActiveSheet()->getStyle( "S5:AC13" )
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

                                                

            #================================================================================================================================================================#   


                $objPHPExcel->getActiveSheet()->getStyle('B' . $st_col . ':'. 'J' . $st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['deta'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('K' . $st_col . ':'. 'Q' . $st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['summ'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('S' . $st_col . ':'. 'AC' . $st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['cost'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('AD' . $st_col . ':'. 'AP' . $st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['rm'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('AQ' . $st_col . ':'. 'CC' . $st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['ma'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('CD' . $st_col . ':'. 'CO' . $st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['as'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('CP' . $st_col . ':'. 'DC' . $st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['di'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('DD' . $st_col . ':'. 'DI' . $st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['pe'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('DJ' . $st_col . ':'. 'DS' . $st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,$color_border['oh'])))); 

                                                                                                                                                                                                                                                                                                                                            
            #================================================================================================================================================================#                                  

                // $objPHPExcel->getActiveSheet()->getStyle('B14:'.$col_name[$count_index+1].($count_data+2))
                //                               ->applyFromArray(array(
                //                                 'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['deta'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('B15:J'.($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['deta'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('B21:J21')
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['deta'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('K15:Q'.($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['cost'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('S15:AC'.($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['summ'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('AD15:AP'.($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['rm'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('AQ15:CA'.($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['ma'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('CB15:CO'.($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['as'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('CP15:DC'.($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['di'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('DD15:DI'.($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['pe'])))); 

                $objPHPExcel->getActiveSheet()->getStyle('DJ15:DS'.($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['oh'])))); 


            #================================================================================================================================================================#       

                $objPHPExcel->getActiveSheet()->getStyle('K19:Q19')
                                              ->applyFromArray(array(
                                                'borders' => array('top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 

                $objPHPExcel->getActiveSheet()->getStyle('S19:AC19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['summ']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 

                $objPHPExcel->getActiveSheet()->getStyle('AD19:AP19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['rm']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 

                $objPHPExcel->getActiveSheet()->getStyle('AQ19:CA19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['ma']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 

                $objPHPExcel->getActiveSheet()->getStyle('CB19:CO19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['as']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 

                $objPHPExcel->getActiveSheet()->getStyle('CP19:DC19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['di']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 
                $objPHPExcel->getActiveSheet()->getStyle('DD19:DI19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['pe']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") ))); 
                $objPHPExcel->getActiveSheet()->getStyle('DJ19:DS19')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR,$color_border['oh']),
                                                                   'top'       => Style_border(PHPExcel_Style_Border::BORDER_THIN,"FFFFFF") )));  

            #================================================================================================================================================================#       

                $objPHPExcel->getActiveSheet()->getStyle('K21:Q21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['cost']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['cost']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle('S21:AC21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['summ']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['summ']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle('AD21:AP21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['rm']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['rm']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle('AQ21:CA21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['ma']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['ma']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle('CB21:CO21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['as']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['as']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle('CP21:DC21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['di']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['di']) )));  

                $objPHPExcel->getActiveSheet()->getStyle('DD21:DI21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['pe']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['pe']) ))); 

                $objPHPExcel->getActiveSheet()->getStyle('DJ21:DS21')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'    => Style_border(PHPExcel_Style_Border::BORDER_HAIR, $color_border['oh']),
                                                                   'outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$color_border['oh']) ))); 


                Style_Alignment('B15'.':'.$col_name[$count_index].($st_col),3, true, $objPHPExcel); 

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['head'])->applyFromArray(array('fill' => Style_Fill($color_border['head'])));
                $objPHPExcel->getActiveSheet()->getStyle('S2:AC2')->applyFromArray(array('fill' => Style_Fill($color_border['targ'])));                          

                $objPHPExcel->getActiveSheet()->getStyle('B15:J19')->applyFromArray(array('fill' => Style_Fill($color_border['deta'])));
                $objPHPExcel->getActiveSheet()->getStyle('B21:J21')->applyFromArray(array('fill' => Style_Fill($color_border['deta'])));



                $objPHPExcel->getActiveSheet()->getStyle($style_layout['rm'])->applyFromArray(array('fill' => Style_Fill($color_border['rm'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['ma'])->applyFromArray(array('fill' => Style_Fill($color_border['ma'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['summ'])->applyFromArray(array('fill' => Style_Fill($color_border['summ'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['cost'])->applyFromArray(array('fill' => Style_Fill($color_border['cost'])));
                $objPHPExcel->getActiveSheet()->getStyle('K19:Q19')->applyFromArray(array('fill' => Style_Fill($color_border['cost'])));

                $objPHPExcel->getActiveSheet()->getStyle($style_layout['as'])->applyFromArray(array('fill' => Style_Fill($color_border['as'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['di'])->applyFromArray(array('fill' => Style_Fill($color_border['di'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['pe'])->applyFromArray(array('fill' => Style_Fill($color_border['pe'])));
                $objPHPExcel->getActiveSheet()->getStyle($style_layout['oh'])->applyFromArray(array('fill' => Style_Fill($color_border['oh'])));

                $objPHPExcel->getActiveSheet()->getStyle('B'  . ($st_col+1) .':J'  . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['deta'])));
                $objPHPExcel->getActiveSheet()->getStyle('K'  . ($st_col+1) .':Q'  . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['cost']  )));
                $objPHPExcel->getActiveSheet()->getStyle('S'  . ($st_col+1) .':AC' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['summ']  )));
                $objPHPExcel->getActiveSheet()->getStyle('AD' . ($st_col+1) .':AP' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['rm']  )));
                $objPHPExcel->getActiveSheet()->getStyle('AQ' . ($st_col+1) .':CA' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['ma']  )));
                $objPHPExcel->getActiveSheet()->getStyle('CB' . ($st_col+1) .':CO' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['as']  )));                
                $objPHPExcel->getActiveSheet()->getStyle('CP' . ($st_col+1) .':DC' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['di']  )));
                $objPHPExcel->getActiveSheet()->getStyle('DD' . ($st_col+1) .':DI' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['pe']  )));
                $objPHPExcel->getActiveSheet()->getStyle('DJ' . ($st_col+1) .':DS' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['oh']  )));






                 $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('2.71');              

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
                 $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth('2.71');  
                foreach (range('S', 'V') as $id )
                 $objPHPExcel->getActiveSheet()->getColumnDimension($id)->setWidth('18.71');                 
                 $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth('12.71');   
                 $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth('12.71');
                 $objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth('12.71');   
                 $objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth('12.71');
                 $objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth('12.71');
                 $objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth('12.71');
                 $objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setWidth('12.71');  
                                
                foreach (range(28, 121) as $id )
                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$id])->setWidth('13.29');  



                 $objPHPExcel->getActiveSheet()->getStyle('AD'.$st_col.':'.$col_name[$count_index].$st_col)->getNumberFormat()->setFormatCode('00#');





                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {
                    if ($key != 'IND') 

                    {
                        
                        if ( substr($key, 0,3) == 'CD_') $objPHPExcel->getActiveSheet()->setCellValue($col_name[$i++].$st_col,  intval( substr($key , 3 ,3) ));

                        elseif ( $key == 'BK' ) $objPHPExcel->getActiveSheet()->setCellValue($col_name[$i++].$st_col,  '' ) ;

                        else $objPHPExcel->getActiveSheet()->setCellValue($col_name[$i++].$st_col, str_replace("_", " ", $key));


                    }
                } // exit;     
#========================================================================================================================  Put data ====================================================================================                

                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {               
                   $col = 0;
                    foreach ($value as $body => $val) 
                    {

                        if ($body != 'IND') 

                        {

                                $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col++].($row), $val);

                                if($val == 3 && $body == 'MODEL')  $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('###"E00"');

                        }

                    }
                    $row++;               
                }


                 $objPHPExcel->getActiveSheet()->setCellValue('E'  . '4', 'DEFECT  WEEKLY REPORT');
                 $objPHPExcel->getActiveSheet()->setCellValue('E'  . '7', '( Accumulate Ng ' . date('01 F') . ' to ' . date('d F', strtotime("- 1 day") ) . " )" );

                 $objPHPExcel->getActiveSheet()->setCellValue('H'  . '12', 'TBKK (Thailand) Co., Ltd.' );
                 $objPHPExcel->getActiveSheet()->setCellValue('H'  . '13', 'vol. 1.2  :  Issue by Pc System ' . date('d-m-Y') );


                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '2', 'Control PPM in TBKK process.' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '3', 'Group' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '3', 'MA ( PPM. )' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '3', 'PE ( PPM. )' );

                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '4', 'Target' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '4', 'Target' );
                 $objPHPExcel->getActiveSheet()->setCellValue('W'  . '3', 'ng+act' );

                 $objPHPExcel->getActiveSheet()->setCellValue('X'    . '3', 'NG Pcs.' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AA'   . '3', 'NG PPM.' );
                 
                 $objPHPExcel->getActiveSheet()->setCellValue('W'   . '4', '(Pcs.)' );
                 $objPHPExcel->getActiveSheet()->setCellValue('X'   . '4', 'MA' );
                 $objPHPExcel->getActiveSheet()->setCellValue('Y'   . '4', 'PE' );
                 $objPHPExcel->getActiveSheet()->setCellValue('Z'   . '4', 'Other' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AA'  . '4', 'MA' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AB'  . '4', 'PE' );
                 $objPHPExcel->getActiveSheet()->setCellValue('AC'  . '4', 'Other' );

                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '5',  'PD1 ASSY' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '6',  'PD2 ENGINE PUMP' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '7',  'PD3 BRAKE&FCD' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '8',  'PD3 KUBOTA' );                 
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '9',  'PD4 ADC' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '10', 'PD5 GEAR' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '11', 'PD5 GKN' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '12', 'PD6 BH' );
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '13', 'None' );

                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '5',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '6',  '44' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '7',  '69' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '8',  '22' );                 
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '9',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '10', '365' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '11', '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '12', '99' );
                 $objPHPExcel->getActiveSheet()->setCellValue('U'  . '13', '0' );

                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '5',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '6',  '632' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '7',  '74' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '8',  '486' );                 
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '9',  '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '10',  '348' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '11', '0' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '12', '498' );
                 $objPHPExcel->getActiveSheet()->setCellValue('V'  . '13', '0' );                 



                 $objPHPExcel->getActiveSheet()->setCellValue('B'  . '15', 'DEFECT OF ' . strtoupper(date('F-Y')) );
                 $objPHPExcel->getActiveSheet()->setCellValue('K'  . '15', 'COST');
                 $objPHPExcel->getActiveSheet()->setCellValue('K'  . '19', 'TOTAL Cost( Baht )');
                 $objPHPExcel->getActiveSheet()->setCellValue('S'  . '15', 'SUMMARY NG');
                 $objPHPExcel->getActiveSheet()->setCellValue('AD' . '15', 'NG CODE RM');
                 $objPHPExcel->getActiveSheet()->setCellValue('AQ' . '15', 'NG CODE MA');
                 $objPHPExcel->getActiveSheet()->setCellValue('CB' . '15', 'NG CODE AS');
                 $objPHPExcel->getActiveSheet()->setCellValue('CP' . '15', 'NG CODE PD4');                 
                 $objPHPExcel->getActiveSheet()->setCellValue('DD' . '15', 'NG CODE PE');
                 $objPHPExcel->getActiveSheet()->setCellValue('DJ' . '15', 'NG CODE OTHER');

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

                foreach ( range(21, 121) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '19',  "=(". $col_name[$sum] . "21" ."/". '$T$21)*1000000' );

                foreach ( range(17, 121) as $sum )  $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$sum] . '21',  "=SUBTOTAL(9,". $col_name[$sum] . $st_dat .":". $col_name[$sum] .$count_data.")");

                foreach ( range(5, 13) as $sum)

                {

                    $sum_ma = "SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$X$'  . $st_dat . ":" . '$X$'  . $count_data . ")";
                    $sum_pe = "SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$AA$' . $st_dat . ":" . '$AA$' . $count_data . ")";
                    $sum_oh = "SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$AB$' . $st_dat . ":" . '$AB$' . $count_data . ")";

                    //echo $sum_ma; exit;

                    $sum_total = " SUMIF(". '$B$' . $st_dat .":". '$B$' .$count_data . ", S". $sum . ',$T$' . $st_dat . ":" . '$T$' . $count_data  . ")";
                   
                    $objPHPExcel->getActiveSheet()->setCellValue( 'W' . $sum,   "=" . $sum_total);
                    $objPHPExcel->getActiveSheet()->setCellValue( 'X' . $sum,   "=" . $sum_ma);
                    $objPHPExcel->getActiveSheet()->setCellValue( 'Y' . $sum,   "=" . $sum_pe);
                    $objPHPExcel->getActiveSheet()->setCellValue( 'Z'  . $sum,  "=" . $sum_oh);
                    $objPHPExcel->getActiveSheet()->setCellValue( 'AA' . $sum,  "=( X". $sum . "/" . "W" . $sum ." ) * 1000000 " );
                    $objPHPExcel->getActiveSheet()->setCellValue( 'AB' . $sum,  "=( Y". $sum . "/" . "W" . $sum ." ) * 1000000 " );
                    $objPHPExcel->getActiveSheet()->setCellValue( 'AC' . $sum,  "=( Z". $sum . "/" . "W" . $sum ." ) * 1000000 " );

                    //echo "=(". $sum_oh . "/" . $sum_total .")*1000000)"; exit;
                }

                foreach ( range(25, $count_data) as $sum)

                {

                        $objPHPExcel->getActiveSheet()->setCellValue( 'W'   . $sum,  "=SUM(" . "AD" . $sum . ":" . "AP" . $sum . ")" );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'X'   . $sum,  "=SUM(" . "AQ" . $sum . ":" . "CA" . $sum . ")" );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'Y'   . $sum,  "=SUM(" . "CB" . $sum . ":" . "CO" . $sum . ")" );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'Z'   . $sum,  "=SUM(" . "CP" . $sum . ":" . "DC" . $sum . ")" );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'AA'  . $sum,  "=SUM(" . "DD" . $sum . ":" . "DI" . $sum . ")" );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'AB'  . $sum,  "=SUM(" . "DJ" . $sum . ":" . "DS" . $sum . ")" );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'AC'  . $sum,  "= X" . $sum . "+" . "AA" . $sum  );



                        $objPHPExcel->getActiveSheet()->setCellValue( 'L'   . $sum,  "=K" . $sum . " * " . "W"  . $sum  );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'M'   . $sum,  "=K" . $sum . " * " . "X"  . $sum  );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'N'   . $sum,  "=K" . $sum . " * " . "Y"  . $sum  );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'O'   . $sum,  "=K" . $sum . " * " . "Z"  . $sum  );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'P'   . $sum,  "=K" . $sum . " * " . "AA" . $sum  );
                        $objPHPExcel->getActiveSheet()->setCellValue( 'Q'   . $sum,  "=K" . $sum . " * " . "AB" . $sum  );                      
               
               }


                // $objPHPExcel->getActiveSheet()->getStyle('B'  . ($st_col+1) .':J'  . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['deta'])));
                // $objPHPExcel->getActiveSheet()->getStyle('K'  . ($st_col+1) .':Q'  . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['cost']  )));
                // $objPHPExcel->getActiveSheet()->getStyle('S'  . ($st_col+1) .':AC' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['summ']  )));
                // $objPHPExcel->getActiveSheet()->getStyle('AD' . ($st_col+1) .':AP' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['rm']  )));
                // $objPHPExcel->getActiveSheet()->getStyle('AQ' . ($st_col+1) .':CA' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['ma']  )));
                // $objPHPExcel->getActiveSheet()->getStyle('CB' . ($st_col+1) .':CO' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['as']  )));                
                // $objPHPExcel->getActiveSheet()->getStyle('CP' . ($st_col+1) .':DC' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['di']  )));
                // $objPHPExcel->getActiveSheet()->getStyle('DD' . ($st_col+1) .':DI' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['pe']  )));
                // $objPHPExcel->getActiveSheet()->getStyle('DJ' . ($st_col+1) .':DS' . ($st_col+1))->applyFromArray(array('fill' => Style_Fill($color_border['oh']  )));
//=SUMIF($C$25:$C$1082, T4, $Y$25:$Y$1082)

//                 $objPHPExcel->getActiveSheet()->setCellValue('C5',  $H_lastM);
//                 $objPHPExcel->getActiveSheet()->setCellValue('J3',  'SUMMARY TOTAL');
//                 $objPHPExcel->getActiveSheet()->setCellValue('J4',  'PLAN');
//                 $objPHPExcel->getActiveSheet()->setCellValue('J5',  'ACTUAL');
//                 $objPHPExcel->getActiveSheet()->setCellValue('J6',  'DIFF');
//                 $objPHPExcel->getActiveSheet()->setCellValue('J7',  'PRICE AMOUNT');
//                 $objPHPExcel->getActiveSheet()->setCellValue('J8',  'PLAN NEXT MONTH');
//                 $objPHPExcel->getActiveSheet()->setCellValue('J9',  'PRICE NEXT MONTH');

//                 $objPHPExcel->getActiveSheet()->setCellValue('L4',  "=SUBTOTAL(9,J". $st_dat .":J".$count_data.")");
//                 $objPHPExcel->getActiveSheet()->setCellValue('L5',  "=SUBTOTAL(9,K". $st_dat .":K".$count_data.")");
//                 $objPHPExcel->getActiveSheet()->setCellValue('L6',  "=SUBTOTAL(9,L". $st_dat .":L".$count_data.")");
//                 $objPHPExcel->getActiveSheet()->setCellValue('L7',  "=SUBTOTAL(9,M". $st_dat .":M".$count_data.")");
//                 $objPHPExcel->getActiveSheet()->setCellValue('L8',  "=SUBTOTAL(9,N". $st_dat .":N".$count_data.")");
//                 $objPHPExcel->getActiveSheet()->setCellValue('L9',  "=SUBTOTAL(9,O". $st_dat .":O".$count_data.")");

//                 $objPHPExcel->getActiveSheet()->setCellValue('O4',  "Pcs.");
//                 $objPHPExcel->getActiveSheet()->setCellValue('O5',  "Pcs.");
//                 $objPHPExcel->getActiveSheet()->setCellValue('O6',  "Pcs.");
//                 $objPHPExcel->getActiveSheet()->setCellValue('O7',  "Thb.");
//                 $objPHPExcel->getActiveSheet()->setCellValue('O8',  "Pcs.");
//                 $objPHPExcel->getActiveSheet()->setCellValue('O9',  "Thb.");

//                 $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+3),  'Exchange rate');
//                 $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+4),  'USD');
//                 $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+5),  'EUR');
//                 $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+6),  'JPY');
//                 $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+4),  $ex_usd);
//                 $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+5),  $ex_eur);
//                 $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+6),  $ex_jpy);

                Style_Alignment('E4',9, False, $objPHPExcel);    
                Style_Alignment('E7',9, False, $objPHPExcel);   
                Style_Alignment('H12:H13',6, False, $objPHPExcel);

                Style_Alignment('S2:AC4',3, False, $objPHPExcel);
                Style_Alignment('S5:S13',9, False, $objPHPExcel);
                //Style_Alignment('T4:T12',9, False, $objPHPExcel);
                // Style_Alignment('T4:T12',3, False, $objPHPExcel);

                Style_Alignment('C25:J'.($count_data),9, False, $objPHPExcel);                                                      

                   $objPHPExcel->getActiveSheet()->getStyle('B15') ->applyFromArray(array('font' => Style_Font(36,"000000",True,False,'Arial Rounded MT Bold'))); 
                   $objPHPExcel->getActiveSheet()->getStyle('S15:DS15') ->applyFromArray(array('font' => Style_Font(36,"000000",True,False,'Arial Rounded MT Bold'))); 
                   $objPHPExcel->getActiveSheet()->getStyle('K15:Q15')->applyFromArray(array('font' => Style_Font(36,"000000",True,False,'Arial Rounded MT Bold')));
                   $objPHPExcel->getActiveSheet()->getStyle('K19')->applyFromArray(array('font' => Style_Font(12,"000000",True,False,'Arial Rounded MT Bold')));

                   $objPHPExcel->getActiveSheet()->getStyle('E4')->applyFromArray(array('font' => Style_Font(26,"000000",True,False,'Arial Rounded MT Bold')));
                   $objPHPExcel->getActiveSheet()->getStyle('E7')->applyFromArray(array('font' => Style_Font(18,"000000",True,False,'Arial Rounded MT Bold')));
                   $objPHPExcel->getActiveSheet()->getStyle('H12')->applyFromArray(array('font' => Style_Font(12,"000000",False,False,'Arial Rounded MT Bold')));
                   $objPHPExcel->getActiveSheet()->getStyle('H13')->applyFromArray(array('font' => Style_Font(11,"000000",False,False,'Arial Rounded MT Bold')));

                   $objPHPExcel->getActiveSheet()->getStyle('S2')->applyFromArray(array('font' => Style_Font(12,"000000",True,False)));
                   $objPHPExcel->getActiveSheet()->getStyle('S3:S13')->applyFromArray(array('font' => Style_Font(12,"000000",True,False)));
                   $objPHPExcel->getActiveSheet()->getStyle('S4:S14')->applyFromArray(array('font' => Style_Font(11,"000000",True,False)));
                
                   $objPHPExcel->getActiveSheet()->getStyle('U3:AC13')->applyFromArray(array('font' => Style_Font(12,"000000",True,False)));
                   //$objPHPExcel->getActiveSheet()->getStyle('I2')->applyFromArray(array('font' => Style_Font(16,"000000",True,False)));

                   $objPHPExcel->getActiveSheet()->getStyle('S19:DS21')->applyFromArray(array('font' => Style_Font(12,"000000",False,False,'Arial Rounded MT Bold')));
                   $objPHPExcel->getActiveSheet()->getStyle('B'. $st_col . ':'.$col_name[$count_index].($st_col)) ->applyFromArray(array('font' => Style_Font(12,"000000",True,False,'Arial Rounded MT Bold')));
                   $objPHPExcel->getActiveSheet()->getStyle('S'. $st_dat . ':'.$col_name[$count_index].($count_data)) ->applyFromArray(array('font' => Style_Font(11,"000000",True,False,'Arial Rounded MT Bold')));


                    $objPHPExcel->getActiveSheet()->mergeCells( 'E4:J6' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'E7:J8' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'H12:J12' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'H13:J13' );

                    $objPHPExcel->getActiveSheet()->mergeCells( 'S2:AC2' );
                    //$objPHPExcel->getActiveSheet()->mergeCells( 'S3:T4' );
                    //$objPHPExcel->getActiveSheet()->mergeCells( 'AC3:AC4' );
                    foreach ( range(5, 13) as $ro)
                    $objPHPExcel->getActiveSheet()->mergeCells( 'S' . $ro . ':' . 'T' . $ro );


                    $objPHPExcel->getActiveSheet()->mergeCells( 'X3:Z3' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'AA3:AC3' );                    



                    $objPHPExcel->getActiveSheet()->mergeCells( 'B15:J19' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'B21:J21' );
                    $objPHPExcel->getActiveSheet()->mergeCells( 'K19:Q19' );

                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['cost']);
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['summ']);
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['rm']  );  
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['ma']  ); 
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['as']  );                     
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['di']  );  
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['pe']  );  
                    $objPHPExcel->getActiveSheet()->mergeCells( $style_layout['oh']  );  
//                 $objPHPExcel->getActiveSheet()->getStyle('C5')->applyFromArray(array('font' => Style_Font(21,"000000",true,true)));

//                 $objPHPExcel->getActiveSheet()->getStyle('J3')->applyFromArray(array('font' => Style_Font(14,"ebf1de",true,true)));
//                 $objPHPExcel->getActiveSheet()->getStyle('J4:O9')->applyFromArray(array('font' => Style_Font(12,"974706",true,true)));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.$st_col.':'.'O'.$st_col)->applyFromArray(array('font' => Style_Font(10,"ebf1de",true,true)));
//                 $objPHPExcel->getActiveSheet()->getStyle('B'.$st_dat.':'.'O'.$count_data)->applyFromArray(array('font' => Style_Font(10,"000005",false,false)));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.($count_data+3) )->applyFromArray(array('font' => Style_Font(10,"000000",false,true)));
//                 $objPHPExcel->getActiveSheet()->getStyle( 'B'.($count_data+4).':'.'D'.($count_data+6) )->applyFromArray(array('font' => Style_Font(9,"000000",true,true)));


//                 $objPHPExcel->getActiveSheet()->getStyle('A1:O10')->applyFromArray(array('fill' => Style_Fill('FFFFFF')));
//                 //$objPHPExcel->getActiveSheet()->insertNewRowBefore(3,1);

//                 $objPHPExcel->getActiveSheet()->getStyle('J3'.':'.$col_name[$count_index]."3")->applyFromArray(array('fill' => Style_Fill('004700')));
//                 $objPHPExcel->getActiveSheet()->getStyle('J4'.':'.$col_name[$count_index]."9")->applyFromArray(array('fill' => Style_Fill('c6e0b4')));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.$st_col.':'.$col_name[$count_index].$st_col)->applyFromArray(array('fill' => Style_Fill('004700')));


//                 $objPHPExcel->getActiveSheet()->getStyle('C3:H4')
//                                               ->applyFromArray(array(
//                                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'00000E'))));

//                 $objPHPExcel->getActiveSheet()->getStyle('C5:H6')
//                                               ->applyFromArray(array(
//                                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E'))));

//                 $objPHPExcel->getActiveSheet()->getStyle('J4:'.$col_name[$count_index].'9')
//                                               ->applyFromArray(array(
//                                                 'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));
//                 $objPHPExcel->getActiveSheet()->getStyle('J10:'.$col_name[$count_index].'10')
//                                               ->applyFromArray(array(
//                                                 'borders' => array('top' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.$st_col.':'.$col_name[$count_index].$st_col)
//                                               ->applyFromArray(array(
//                                                 'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.$st_dat.':'.$col_name[$count_index].$count_data)
//                                               ->applyFromArray(array(
//                                                 'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));
//                 $objPHPExcel->getActiveSheet()->getStyle('B'.$count_data.':'.$col_name[$count_index].$count_data)
//                                               ->applyFromArray(array(
//                                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.$look_data.':'.'O'.$look_data)
//                                               ->applyFromArray(array(
//                                                 'borders' => array('top'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E'))));

//                 $objPHPExcel->getActiveSheet()->getStyle( 'B' . ($count_data+3) .':'. 'D' .($count_data+3) )
//                                               ->applyFromArray(array(
//                                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E'))));                                                              
// //echo $look_data; exit;

                    $objPHPExcel->getActiveSheet()->getStyle('U5:Z13')->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');

                    $objPHPExcel->getActiveSheet()->getStyle('AA5:AC13')->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
                                                  
                    $objPHPExcel->getActiveSheet()->getStyle('K'.$st_dat.':'.'Q'.$count_data)->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');

                    $objPHPExcel->getActiveSheet()->getStyle('S'.$st_dat.':'.$col_name[$count_index].$count_data)->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
                                                 

                    $objPHPExcel->getActiveSheet()->getStyle('L'.'21'.':'.'Q'.'21')->getNumberFormat()->setFormatCode('#,##0.00');

                    $objPHPExcel->getActiveSheet()->getStyle('S'.'19'.':'.$col_name[$count_index].'19')->getNumberFormat()->setFormatCode('#,##0.00');

                    $objPHPExcel->getActiveSheet()->getStyle('S'.'21'.':'.$col_name[$count_index].'21')->getNumberFormat()->setFormatCode('#,##0');
                                                  
                    // $objPHPExcel->getActiveSheet()->getStyle('L'.'7')
                    //                               ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');                              
                    // $objPHPExcel->getActiveSheet()->getStyle('L'.'9')
                    //                               ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-'); 
             #A
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[0])->setWidth('5');     #B no
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[1])->setWidth('7');     #D plnt
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[1])->setWidth('8');     #C pd                
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[3])->setWidth('11');    #E so_no
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[4])->setWidth('19');    #F so_nm
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[5])->setWidth('17');    #G it_no
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[6])->setWidth('30');    #H it_nm
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[7])->setWidth('21');    #I model
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[8])->setWidth('12');    #J
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[9])->setWidth('12');    #K
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[10])->setWidth('12');   #L
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[11])->setWidth('14.29');#M
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[12])->setWidth('12');   #N
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[13])->setWidth('14.29');#M

//                 Style_Alignment('C2:C5',3, false, $objPHPExcel);
//                 Style_Alignment('J3',3, false, $objPHPExcel);
//                 Style_Alignment(('B'.$st_col.':'.'O'.$st_col), 3, false, $objPHPExcel);
//                 Style_Alignment(('B'.$st_dat.':'.'I'.$count_data), 9, false, $objPHPExcel);
//                 $objPHPExcel->getActiveSheet()->mergeCells('C3:'.'H4');
//                 $objPHPExcel->getActiveSheet()->mergeCells('C5:'.'H6');
//                 $objPHPExcel->getActiveSheet()->mergeCells('J3:'.'N3');

//                   foreach (range( ($count_data+4) , ($count_data+6) ) as $index) Style_group_lv1_Row($index, $objPHPExcel);

                   Style_group_lv1_Col($col_name, 5, $objPHPExcel);
                   Style_group_lv1_Col($col_name, 7, $objPHPExcel);
                   foreach ( range(9, 15)   as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel);
                   foreach ( range(28, 121) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel);
//                 //echo ($count_data+4); exit;

//                 foreach(range(4, 9) as $r)
//                 {
//                     $objPHPExcel->getActiveSheet()->mergeCells('L'.$r.':'.'N'.$r);
//                     $objPHPExcel->getActiveSheet()->mergeCells('J'.$r.':'.'K'.$r);                    
//                 }

//                 $objPHPExcel->getActiveSheet()->mergeCells('B' . ($count_data+3) .':'. 'D' .($count_data+3));

//                 foreach(range(($count_data+4), ($count_data+6)) as $r)
//                 {
//                     $objPHPExcel->getActiveSheet()->mergeCells('C'.$r.':'.'D'.$r);                 
//                 }               

#========================================================================================================================  Put field ==================================================================================== 
            }
            elseif( $sheetIndex == 'code_detail' ) 
            {
                $objPHPExcel->getActiveSheet()->setTitle( "Code detail" );                
                $objPHPExcel->getActiveSheet()->setShowGridlines(False);
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
                
//                 $objPHPExcel->getActiveSheet()->setCellValue('B5', 'MONTHLY RECEIVING HISTORY REPORT');
// $st = ((date('m')-12) < 1 )  ? date('F-Y', strtotime( (date('Y')-1) . "-" . (12+(date('m')-12)). "-" . '01' ) ) : date('F-Y (ERROR)') ; 

// $en = ((date('m')-1)  < 1 )  ? date('F-Y', strtotime( (date('Y')-1). "-" ."12". "-" . '01' ) ) : date('F-Y', strtotime( (date('Y')+0). "-" .(date('m')-1). "-" . '01' ) ) ;//
//                 $objPHPExcel->getActiveSheet()->setCellValue('B7', 'PERIOD TIME :  '. $st . ' To '. $en);
//                 $objPHPExcel->getActiveSheet()->setCellValue('H3', 'Summary Actual (Pcs.)');
//                 $objPHPExcel->getActiveSheet()->setCellValue('I3', 'Summary Price (Thb.)' );

//                 $objPHPExcel->getActiveSheet()->setCellValue('AH2',  'p' );                
//                 $objPHPExcel->getActiveSheet()->setCellValue('AH6',  'Click button to unhide' );
//                 $objPHPExcel->getActiveSheet()->setCellValue('AH20', 'DATA HISTORY MONTHLY RECEIVE' );

//                 $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+3),  'Exchange rate');
//                 $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+4),  'USD');
//                 $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+5),  'EUR');
//                 $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+6),  'JPY');
//                 $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+4),  $ex_usd);
//                 $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+5),  $ex_eur);
//                 $objPHPExcel->getActiveSheet()->setCellValue('C'.($count_data+6),  $ex_jpy);

//                 $re_mon = 12;
//                 foreach(range(4, 15) as $mon)
//                 {
// $his_month = ((date('m')-($re_mon)) < 1 )  ? date('F-Y', strtotime( (date('Y')-1) . "-" . (12+(date('m')-($re_mon--))). "-" . '01' ) ) : date('F-Y', strtotime( (date('Y')+0). "-" .(date('m')-($re_mon--)). "-" . '01' ) ) ;

//                     $objPHPExcel->getActiveSheet()->setCellValue('G'.$mon , $his_month);
//                 }
//                 $sum_rA = 15;
//                 $sum_rP = 15;
//                 $switch_col = 0;
//                 foreach(range(8, 31) as $de_col)
//                 {
//                     $detail = ($de_col % 2 == 0) ? "Actual (Pcs.)" : "Price (Thb.)" ;
//                         $objPHPExcel->getActiveSheet()->setCellValue($col_name[$de_col].$st_col, $detail);

                    
//                     if($de_col % 2 == 0)
//                     {
// $his_month = ((date('m')-(++$re_mon)) < 1 )  ? date('F-Y', strtotime( (date('Y')-1) . "-" . (12+(date('m')-($re_mon))). "-" . '01' ) ) : date('F-Y', strtotime( (date('Y')+0). "-" .(date('m')-($re_mon)). "-" . '01' ) ) ;
//                         $objPHPExcel->getActiveSheet()->setCellValue($col_name[$de_col].($st_col-1), $his_month);

//                         if($switch_col == 0)
//                         {
//                             $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col-1) )->applyFromArray(array('fill' => Style_Fill('002900')));
//                             $switch_col = 1;
//                         }
//                         else
//                         {
//                             $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col-1) )->applyFromArray(array('fill' => Style_Fill('333300')));
//                             $switch_col = 0;
//                         }


//                         $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) )->applyFromArray(array('fill' => Style_Fill('76933c')));
//                         $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) .":".$col_name[$de_col].($count_data) )->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
//                         $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) .":".$col_name[$de_col].($count_data) )->applyFromArray(array('font' => Style_Font(10,"000005",false,true)));
                              
//                         $objPHPExcel->getActiveSheet()->setCellValue($col_name[6].($sum_rA--), '=SUBTOTAL(9,'.$col_name[$de_col].$st_dat.":".$col_name[$de_col].$count_data.')');
//                     }
//                     else
//                     {
//                        /// $objPHPExcel->getActiveSheet()->setCellValue($col_name[$de_col].($st_col-1), '');
//                          $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col-1) )->applyFromArray(array('fill' => Style_Fill('333300')));
//                          $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) )->applyFromArray(array('fill' => Style_Fill('4f6228')));
//                          $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_dat) .":".$col_name[$de_col].($count_data) )->applyFromArray(array('fill' => Style_Fill('ebf1de')));
//                          $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) .":".$col_name[$de_col].($count_data) )->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
//                          $objPHPExcel->getActiveSheet()->getStyle( $col_name[$de_col].($st_col) .":".$col_name[$de_col].($count_data) )->applyFromArray(array('font' => Style_Font(10,"eb2613",false,true)));                               

//                         $objPHPExcel->getActiveSheet()->setCellValue($col_name[7].($sum_rP--), '=SUBTOTAL(9,'.$col_name[$de_col].$st_dat.":".$col_name[$de_col].$count_data.')');
//                     }
//                 }

//                 $objPHPExcel->getActiveSheet()->getStyle('B5')->applyFromArray(array('font' => Style_Font(18,"000000",true,true)));
//                 $objPHPExcel->getActiveSheet()->getStyle('B7')->applyFromArray(array('font' => Style_Font(12,"000000",true,true)));

//                 $objPHPExcel->getActiveSheet()->getStyle('G3:I15')->applyFromArray(array('font' => Style_Font(11,"ebf1de",true,true)));
//                 $objPHPExcel->getActiveSheet()->getStyle('H4:I15')->applyFromArray(array('font'  => Style_Font(12,"974706",true,true)));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.($st_col-1).':'.'I' .($st_col-1))->applyFromArray(array('font' => Style_Font(10,"ebf1de",true,true)));
//                 $objPHPExcel->getActiveSheet()->getStyle('J'.($st_col-1).':'.'AG'.($st_col-1))->applyFromArray(array('font' => Style_Font(11,"ebf1de",true,true)));
//                 $objPHPExcel->getActiveSheet()->getStyle('J'.($st_col).  ':'.'AG'.($st_col))->applyFromArray(array('font' => Style_Font(10,"ebf1de",true,true)));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.($st_dat).':'.'I'.$count_data)->applyFromArray(array('font' => Style_Font(10,"000005",false,true)));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.($count_data+3) )->applyFromArray(array('font' => Style_Font(11,"000000",false,true)));
//                 $objPHPExcel->getActiveSheet()->getStyle( 'B'.($count_data+4).':'.'D'.($count_data+6) )->applyFromArray(array('font' => Style_Font(10,"000000",false,true)));   

//                 $objPHPExcel->getActiveSheet()->getStyle('AH2')->applyFromArray(array('font' => Style_Font(36,"00b0f0",true,false,'Wingdings 3')));
//                 $objPHPExcel->getActiveSheet()->getStyle('AH6')->applyFromArray(array('font' => Style_Font(14,"00b0f0",true,true,'Arial Unicode MS')));
//                 $objPHPExcel->getActiveSheet()->getStyle('AH20')->applyFromArray(array('font' => Style_Font(26,"00b0f0",true,true)));
                
//                 //$objPHPExcel->getActiveSheet()->getStyle('AH2')->getAlignment()->setTextRotation(90);
//                 $objPHPExcel->getActiveSheet()->getStyle('AH6')->getAlignment()->setTextRotation(-90);
//                 $objPHPExcel->getActiveSheet()->getStyle('AH20')->getAlignment()->setTextRotation(-90);

//                 $objPHPExcel->getActiveSheet()->getStyle('A1:M9')->applyFromArray(array('fill' => Style_Fill('FFFFFF')));
//                 //$objPHPExcel->getActiveSheet()->insertNewRowBefore(3,1);
//                 $objPHPExcel->getActiveSheet()->getStyle('H3'.':'.'I3')->applyFromArray(array('fill' => Style_Fill('002900')));
//                 $objPHPExcel->getActiveSheet()->getStyle('G4'.':'.'G15')->applyFromArray(array('fill' => Style_Fill('002900')));
//                 $objPHPExcel->getActiveSheet()->getStyle('H4'.':'.'I15')->applyFromArray(array('fill' => Style_Fill('c6e0b4')));



//                 $objPHPExcel->getActiveSheet()->getStyle('B'.($st_col-1).':'.$col_name[7].($st_col-1))->applyFromArray(array('fill' => Style_Fill('002900')));
// $objPHPExcel->getActiveSheet()->getStyle( 'H4' .":".'H15' )->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
// $objPHPExcel->getActiveSheet()->getStyle( 'I4' .":".'I15' )->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');



//                 $objPHPExcel->getActiveSheet()->getStyle('B5:F6')
//                                               ->applyFromArray(array(
//                                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'00000E'))));

//                 $objPHPExcel->getActiveSheet()->getStyle('B7:F8')
//                                               ->applyFromArray(array(
//                                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E'))));

//                 $objPHPExcel->getActiveSheet()->getStyle('H3:I15')
//                                               ->applyFromArray(array(
//                                                 'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));
//                 $objPHPExcel->getActiveSheet()->getStyle('G4:G15')
//                                               ->applyFromArray(array(
//                                                 'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));                                              
//                 $objPHPExcel->getActiveSheet()->getStyle('H16:I16')
//                                               ->applyFromArray(array(
//                                                 'borders' => array('top' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.($st_col-1).':'.$col_name[$count_index].$st_col)
//                                               ->applyFromArray(array(
//                                                 'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));

//                 $objPHPExcel->getActiveSheet()->getStyle('B'.$st_dat.':'.$col_name[$count_index].$count_data)
//                                               ->applyFromArray(array(
//                                                 'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));
//                 $objPHPExcel->getActiveSheet()->getStyle('B'.$count_data.':'.$col_name[$count_index].$count_data)
//                                               ->applyFromArray(array(
//                                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));
//                 $objPHPExcel->getActiveSheet()->getStyle('B'.$look_data.':'.$col_name[$count_index].$look_data)
//                                               ->applyFromArray(array(
//                                                 'borders' => array('top'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E'))));

//                 $objPHPExcel->getActiveSheet()->getStyle('B2:'.$col_name[7].($count_data+1))
//                                               ->applyFromArray(array(
//                                                 'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000023')))); 

//                 $objPHPExcel->getActiveSheet()->getStyle('J16:'.$col_name[31].($count_data+1))
//                                               ->applyFromArray(array(
//                                                 'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000023'))));   

//                 $objPHPExcel->getActiveSheet()->getStyle( 'B' . ($count_data+3) .':'. 'D' .($count_data+3) )
//                                               ->applyFromArray(array(
//                                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'00000E')))); 

//                 foreach (range(8, 31) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel);
//                 foreach (range( ($count_data+4) , ($count_data+6) ) as $index) Style_group_lv1_Row($index, $objPHPExcel);


//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[4])->setWidth('19');    #F so_nm
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[5])->setWidth('19');    #G it_no
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[6])->setWidth('30');    #H it_nm
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[7])->setWidth('30');    #I model    
//                 foreach(range(8, 31) as $key)
//                     $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$key])->setWidth('14.71');
//                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[32])->setWidth('15.71');    #I model

//                 Style_Alignment('B2:B5',7, false, $objPHPExcel);
//                 Style_Alignment('H3:I3',3, false, $objPHPExcel);
//                 Style_Alignment('AH2',3, false, $objPHPExcel);
//                 Style_Alignment('AH6',2, false, $objPHPExcel);
//                 Style_Alignment('AH20',2, false, $objPHPExcel);
//                 Style_Alignment(('B'.($st_col-1).':'.$col_name[$count_index].$st_col), 3, false, $objPHPExcel);
//                 Style_Alignment(('B'.$st_dat.':'.'I'.$count_data), 9, false, $objPHPExcel);

//                 foreach(range(0, 7)  as $key) $objPHPExcel->getActiveSheet()->mergeCells($col_name[$key].($st_col-1).':'.$col_name[$key].$st_col);
//                 foreach(range(8, 31) as $key) 
//                     if($key % 2 == 0)
//                         $objPHPExcel->getActiveSheet()->mergeCells($col_name[$key].($st_col-1).':'.$col_name[($key+1)].($st_col-1));   
//                 $objPHPExcel->getActiveSheet()->mergeCells('B5'.':'.'F6');
//                 $objPHPExcel->getActiveSheet()->mergeCells('B7'.':'.'F8');  
//                 $objPHPExcel->getActiveSheet()->mergeCells('AH2'. ':'.'AH5');
//                 $objPHPExcel->getActiveSheet()->mergeCells('AH6'. ':'.'AH16');
//                 $objPHPExcel->getActiveSheet()->mergeCells('AH20'.':'.'AH'.($count_data+1));

//                 $objPHPExcel->getActiveSheet()->mergeCells('B' . ($count_data+3) .':'. 'D' .($count_data+3));

//                 foreach(range(($count_data+4), ($count_data+6)) as $r)
//                 {
//                     $objPHPExcel->getActiveSheet()->mergeCells('C'.$r.':'.'D'.$r);                 
//                 }                                   
//             }
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