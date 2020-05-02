<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');
ini_set('max_execution_time', 300); 
ini_set('memory_limit','10240M');
if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/PHPExcel-1.8.1/Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();
$data_col = array();
$dayA   = date('d');
$monthA = date('m');
$yearA  = date('Y');
$en = ((date('m')-1)  < 1 )  ? date('Y/m/d', strtotime( (date('Y')-1). "-" ."12". "-" . '01' ) ) : date('F-Y', strtotime( (date('Y')+0). "-" .(date('m')-1). "-" . '01' ) ) ;





$lastmount = substr(date('Y/m/t',strtotime($en)),8, 2);
//var_dump($list_act_report); exit;
$col_name = array();

$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
$objPHPExcel->getActiveSheet()->getPageSetup()->setFitToPage(true);
$objPHPExcel->getActiveSheet()->getPageSetup()->setFitToWidth(1);
$objPHPExcel->getActiveSheet()->getPageSetup()->setFitToHeight(0);
    



foreach ( range('A', 'Z') as $cm ) { array_push($col_name, $cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "A".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "B".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "C".$cm); }

$gdImage = dirname(__FILE__) . '/img/NEW-TBKK-LOGO_0.png';
// Add a drawing to the worksheetecho date('H:i:s') . " Add a drawing to the worksheet\n";
$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setName('Sample image');
$objDrawing->setDescription('Sample image');
$objDrawing->setPath($gdImage);
//$objDrawing->setRenderingFunction(PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG);
//$objDrawing->setMimeType(PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT);
$objDrawing->setOffsetX(0); 
$objDrawing->setOffsetY(1);  
$objDrawing->setHeight(195);
$objDrawing->setWidth(195); 
$objDrawing->setCoordinates('B2');
$objDrawing->setWorksheet($objPHPExcel->getActiveSheet()); 
//var_dump($subacc_diff); exit;

$ind = 0;
$i=2;


/* page margins */
$objPHPExcel->getActiveSheet()->getPageMargins()->setBottom(1.778);
$objPHPExcel->getActiveSheet()->getPageMargins()->setLeft(0.2);
$objPHPExcel->getActiveSheet()->getPageMargins()->setRight(0.2);
$objPHPExcel->getActiveSheet()->getPageMargins()->setTop(0.2);

$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(8,9);
 // $objPHPExcel->getActiveSheet()
 //    ->getHeaderFooter()
 //    ->setOddHeader('&C&HPlease treat this document as confidential!');
$objPHPExcel->getActiveSheet()
    ->getHeaderFooter()
    ->setOddFooter('&L&B' . 'บริษัท ทีบีเคเค ประเทศไทย จำกัด'  );   
// $T_lastM = ((date('m')-1) > 12 ) ? date('My', strtotime( (date('Y')-1). "-" ."12". "-" . '01' ) ) : date('My', strtotime( (date('Y')+0). "-" .(date('m')-1). "-" . '01' ) ) ;// exit;
// $H_lastM = ((date('m')-1) > 12 ) ? date('F Y', strtotime( (date('Y')-1). "-" ."12". "-" . '01' ) ) : date('F Y', strtotime( (date('Y')+0). "-" .(date('m')-1). "-" . '01' ) ) ;// exit;
// //$T_lastM = date('My',  strtotime( date('Y'). "-" .(date('m')-1). "-" . 1 ) ) ;// exit;
// //$H_lastM = date('F Y', strtotime( date('Y'). "-" .(date('m')-1). "-" . 1 ) ) ;// exit;

// $ex_usd = $rate[0]['CURRENCY_RATE'];
// $ex_eur = $rate[1]['CURRENCY_RATE'];
// $ex_jpy = $rate[2]['CURRENCY_RATE'];
// echo $ex_usd; 
// echo "<hr>";
// echo $ex_eur; 
// echo "<hr>";
// echo $ex_jpy; 
// echo "<hr>";
//echo $title; exit;
// exit;
foreach ($title as $inTil => $til) 
{
             $objPHPExcel->createSheet();
             $objPHPExcel->setActiveSheetIndex($inTil);
             //$objPHPExcel->setActiveSheetIndex(0);

            $sheetIndex  =  strtolower(str_replace(' ', '_', $title[$inTil])); 
            $count_index = 0;
            $i = 1;   
            // $ind = 0;
            $count_data  =  count($list_act_report[$sheetIndex]);


    if ($count_data > 0) 
    {      
#========================================================================================================================  Put field ====================================================================================        
                $objPHPExcel->getActiveSheet()->setTitle( "$til"  );
                $objPHPExcel->getActiveSheet()->setShowGridlines(False);
                $st_col = 8;
                $st_dat = 10;
                $count_index =   count($list_act_report[$sheetIndex][0]) ;
                $row = $st_dat;
                $count_data  =  count($list_act_report[$sheetIndex]) + $row-1;
                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 10 );
                foreach(range(2, 5) as $r)
                $objPHPExcel->getActiveSheet()->getRowDimension( $r )->setRowHeight( 36 );
                
                $objPHPExcel->getActiveSheet()->getRowDimension( 6  )->setRowHeight( 25 ); 
                $objPHPExcel->getActiveSheet()->getRowDimension( 7  )->setRowHeight( 8 );
                foreach(range(8, 9) as $r)               
                $objPHPExcel->getActiveSheet()->getRowDimension( $r )->setRowHeight( 36 );

                foreach(range($st_dat, $count_data) as $r)               
                $objPHPExcel->getActiveSheet()->getRowDimension( $r )->setRowHeight( 36 );            
                // $objPHPExcel->getActiveSheet()->getRowDimension( 12 )->setRowHeight( 10 ); 
                // $objPHPExcel->getActiveSheet()->getRowDimension( 13 )->setRowHeight( 10 ); 

                //$objPHPExcel->getActiveSheet()->freezePane('G'.$row);   
                //$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(90);    
                //$objPHPExcel->getActiveSheet()->setAutoFilter('C'.($st_dat-1).':'.$col_name[$count_index].($st_dat-1));                


                $objPHPExcel->getActiveSheet()->getStyle( 'B'.$st_col.':'.$col_name[$count_index].($count_data) )
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000')))); 


                $objPHPExcel->getActiveSheet()->getStyle( 'N2:R5' )
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000')))); 


                            
                 Style_Alignment('B'.($st_col).':'.$col_name[$count_index].($st_col+1),3, false, $objPHPExcel);
                             
                 Style_Alignment('H1:R5',3, false, $objPHPExcel);
                 Style_Alignment('G3',3, false, $objPHPExcel);

                 Style_Alignment( 'B'.$st_dat.':'.$col_name[3].$count_data ,3, false, $objPHPExcel);


                                                                                                                              












                $objPHPExcel->getActiveSheet()->getStyle( 'B'.($st_col+1).':'.$col_name[$count_index].($st_col+1) )
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000')))); 






                $objPHPExcel->getActiveSheet()->getStyle( 'B'.$st_col.':'.$col_name[$count_index].($st_col+1) )
                                              ->applyFromArray(array(
                                                'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                $objPHPExcel->getActiveSheet()->getStyle( 'B'.$st_dat.':'.$col_name[$count_index].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                $objPHPExcel->getActiveSheet()->getStyle( 'N2:R5' )
                                              ->applyFromArray(array(
                                                'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                 $objPHPExcel->getActiveSheet()->getStyle('C2')->applyFromArray(array('font' => Style_Font(24,"000000",true,false)));
                 $objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray(array('font' => Style_Font(20,"000000",true,false)));


                 $objPHPExcel->getActiveSheet()->getStyle('C5')->applyFromArray(array('font' => Style_Font(18,"000000",true,false)));

                 $objPHPExcel->getActiveSheet()->getStyle('E5')->applyFromArray(array('font' => Style_Font(18,"000000",true,false)));

                 $objPHPExcel->getActiveSheet()->getStyle('B'.($count_data+2) .":".'B'.($count_data+9))->applyFromArray(array('font' => Style_Font(18,"000000",true,false)));

                 $objPHPExcel->getActiveSheet()->getStyle('G3')->applyFromArray(array('font' => Style_Font(28,"000000",true,false)));


                 $objPHPExcel->getActiveSheet()->getStyle('B8:R8')->applyFromArray(array('font' => Style_Font(18,"000000",true,false)));
                 $objPHPExcel->getActiveSheet()->getStyle('F8:Q8')->applyFromArray(array('font' => Style_Font(12,"000000",true,false)));

                 $objPHPExcel->getActiveSheet()->getStyle( 'B'.$st_dat.':'.$col_name[$count_index].$count_data )->applyFromArray(array('font' => Style_Font(18,"000000",false,false)));


                 foreach( array( array(5,7), array(8,10), array(11,13), array(14,16) )  as  $ty )
                 {
                        $objPHPExcel->getActiveSheet()->getStyle( $col_name[$ty[0]].$st_col.':'.$col_name[$ty[1]].($count_data) )
                                                      ->applyFromArray(array(
                                                        'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));  


                    $objPHPExcel->getActiveSheet()->mergeCells($col_name[$ty[0]].$st_col.':'.$col_name[$ty[1]].$st_col);
                 }
                 foreach( array(1, 2, 3, 4, $count_index)  as  $ty )
                 {

                    $objPHPExcel->getActiveSheet()->mergeCells($col_name[$ty].$st_col.':'.$col_name[$ty].($st_col+1));
                 }

                // $objPHPExcel->getActiveSheet()->getStyle('C'.$st_col.':'.$col_name[$count_index].$st_col)->applyFromArray(array('fill' => Style_Fill('305496')));

                // $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].$st_col.':'.$col_name[($count_index+4)].$st_col)->applyFromArray(array('fill' => Style_Fill('305496')));


                // $objPHPExcel->getActiveSheet()->getStyle('C'.($st_dat-1).':'.$col_name[$count_index].($st_dat-1))->applyFromArray(array('fill' => Style_Fill('e0ebeb')));

                // $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].($st_dat-1).':'.$col_name[($count_index+4)].($st_dat-1))->applyFromArray(array('fill' => Style_Fill('e0ebeb')));





                // $objPHPExcel->getActiveSheet()->getStyle('C'.$st_col.':'.'E'.$st_col)->applyFromArray(array('font' => Style_Font(11,"FFFFFF",true,true)));

                // $objPHPExcel->getActiveSheet()->getStyle('F'.$st_col.':'.$col_name[$count_index+6].($st_col))->applyFromArray(array('font' => Style_Font(10,"FFFFFF",true,true)));

                // $objPHPExcel->getActiveSheet()->getStyle('F'.$st_dat.':'.$col_name[$count_index+6].$count_data)->applyFromArray(array('font' => Style_Font(10,"000000",false,true)));
                // //echo $row; exit;                
                   
                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {    
                    if ( $key != "NO" && $key != "PD" )
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++].$st_col, str_replace("_", " ", $key));
                  
                } // exit;     
#========================================================================================================================  Put data ====================================================================================                

                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {               
                   $col = 1;
                    foreach ($value as $body => $val) 
                    {

                        
                        if ($body != "NO" && $body != "PD"){

                            if( substr($body, 0, (strlen($body)-2) ) == "REMAIN" && $val > 0) $objPHPExcel->getActiveSheet()->getStyle( $col_name[$col].($row) )->applyFromArray(array('fill' => Style_Fill('c8c8c8')));

                            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$col++].($row), $val);
                        }
                        

                        if($val == 3 && $body == 'MODEL')  $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('###"E00"');

                        if($val > 10000 && $body == 'SNP')
                        {  
                            $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col-1].($row), 'Please Update Snp.'); 
                            $objPHPExcel->getActiveSheet()->getStyle( $col_name[$col-1].($row) )->applyFromArray(array('fill' => Style_Fill('c8c8c8')));     
                            $objPHPExcel->getActiveSheet()->getStyle( $col_name[$col-1].($row) )->applyFromArray(array('font' => Style_Font(8,"000000",true,false)));
                            Style_Alignment($col_name[$col-1].($row),3, false, $objPHPExcel);
                        }              

                    }
//exit;
                    $row++;               
                }

                 foreach ( range(5, 16) as $ke) 
                 {


                     if( ($ke % 3) == 2  )
                            $objPHPExcel->getActiveSheet()->setCellValue($col_name[$ke] . 9, "จำนวน");
                     elseif( ($ke % 3) == 0 )
                            $objPHPExcel->getActiveSheet()->setCellValue($col_name[$ke] . 9, "เศษ");
                     else
                            $objPHPExcel->getActiveSheet()->setCellValue($col_name[$ke] . 9, "กล่อง");
                 }
                 $objPHPExcel->getActiveSheet()->setCellValue($col_name[5] . 8, "( PERIOD 1-2 ) 05:00 น. - 10:00 น.");
                 $objPHPExcel->getActiveSheet()->setCellValue($col_name[8] . 8, "( PERIOD 3-4 ) 13:00 น. - 16:00 น.");
                 $objPHPExcel->getActiveSheet()->setCellValue($col_name[11]. 8, "( PERIOD 5-6 ) 18:00 น. - 19:00 น.");
                 $objPHPExcel->getActiveSheet()->setCellValue($col_name[14]. 8, "( PERIOD 7-8 ) 20:00 น. - 00:00 น.");



                 $objPHPExcel->getActiveSheet()->setCellValue('C2', "ใบเบิกงาน-จัดงาน / ใบนำของออกนอกโรงงาน");
                 $objPHPExcel->getActiveSheet()->setCellValue('C3', "( PICKING LIST & DELIVERY NOTE )");

                 $objPHPExcel->getActiveSheet()->setCellValue('C5', "วันที่ออกเอกสาร : ". date('d/m/Y') );
                 $objPHPExcel->getActiveSheet()->setCellValue('E5', "วันที่ส่งของ : ". date('d/m/Y', strtotime($de_date)) );

                 $cust_str = $de_cust;
                 if (  $de_cust == 'MEC-SUM' )  $cust_str = 'MEC-Supplement Free Zone';
                 elseif( $de_cust == 'SKC-SUM' ) $cust_str = 'SKC-200 & 500';
                 else $cust_str = $de_cust;
                 //$cust_str = (  $de_cust == 'MEC-SUM' ) ? 'MEC-Supplement Free Zone' :  (  $de_cust == 'SKC-SUM' ) ? 'SKC-200 & 500': $de_cust;
                 $objPHPExcel->getActiveSheet()->setCellValue('G3', "ลูกค้า  ".$cust_str);

                 $objPHPExcel->getActiveSheet()->setCellValue('N4', 'CHECKDOCUMENTS BY'."\n".'เช็คเอกสารโดย');

                 $objPHPExcel->getActiveSheet()->setCellValue('P4', 'PART / FG PREPARED BY'."\n".'จัดเตรียมงาน / จัดส่งโดย');
                 $objPHPExcel->getActiveSheet()->setCellValue('R4', 'APPROVE BY'."\n".'อนุมติโดย');

                 $objPHPExcel->getActiveSheet()->getStyle('N4:R4')->getAlignment()->setWrapText(true);
                 
                 $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+3), "ผู้นำของออก_________________________________TIME / เวลา______________________________");
                 $objPHPExcel->getActiveSheet()->setCellValue('B'.($count_data+7), "วันที่นำของออก______________________________ทะเบียนรถ_______________________________");






 //echo $til . " = " . $subplan[0] . "<hr>" ;

                //$objPHPExcel->getActiveSheet()->setCellValue('J7',  $subplan[0] . ")" );
                // $objPHPExcel->getActiveSheet()->setCellValue('C2', 'Production of '.date('F Y',strtotime($en)) );
                // $objPHPExcel->getActiveSheet()->setCellValue('C7', 'TOTAL');
                // $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+3)].$st_col, 'Summary');

                // $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+4)].($st_col+2), '=SUM(' .$col_name[6]."7".":".$col_name[(6+$lastmount)]."7".")");
                // $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+4)].($st_col+3), '=SUM(' .$col_name[6]."8".":".$col_name[(6+$lastmount)]."8".")");
                // $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+4)].($st_col+4), '=SUM(' .$col_name[6]."9".":".$col_name[(6+$lastmount)]."9".")");
                // $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+4)].($st_col+5), '=' .$col_name[(6+$lastmount)]."10");
                // $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+4)].($st_col+6), '=SUM(' .$col_name[6]."11".":".$col_name[(6+$lastmount)]."11".")");

                // $objPHPExcel->getActiveSheet()->setCellValue('F5',  '#');
                
                // $objPHPExcel->getActiveSheet()->setCellValue('F7',  '  Plan');
                // $objPHPExcel->getActiveSheet()->setCellValue('F8',  '  Actual');
                // $objPHPExcel->getActiveSheet()->setCellValue('F9',  '  Diff');
                // $objPHPExcel->getActiveSheet()->setCellValue('F10', '  Acc. Diff');
                // $objPHPExcel->getActiveSheet()->setCellValue('F11', '  Ng');

                // $objPHPExcel->getActiveSheet()->setCellValue($col_name[($count_index+3)].'7',  '  Plan');
                // $objPHPExcel->getActiveSheet()->setCellValue($col_name[($count_index+3)].'8',  '  Actual');
                // $objPHPExcel->getActiveSheet()->setCellValue($col_name[($count_index+3)].'9',  '  Diff');
                // $objPHPExcel->getActiveSheet()->setCellValue($col_name[($count_index+3)].'10', '  Acc. Diff');
                // $objPHPExcel->getActiveSheet()->setCellValue($col_name[($count_index+3)].'11', '  Ng');

                //  $objPHPExcel->getActiveSheet()->setCellValue('C' . ($count_data+3),  'TBKK (Thailand) Co.,Ltd. by Pcsystem.');

                // $objPHPExcel->getActiveSheet()->getStyle('C2')->applyFromArray(array('font' => Style_Font(24,"305496",true,true,'Franklin Gothic Heavy')));

                // $objPHPExcel->getActiveSheet()->getStyle('C7')->applyFromArray(array('font' => Style_Font(29,"000000",true,true )));

                // $objPHPExcel->getActiveSheet()->getStyle('F7:F11')->applyFromArray(array('font' => Style_Font(10,"000000",true,true)));

                // $objPHPExcel->getActiveSheet()->getStyle('G'.'7'.':'.$col_name[$count_index+6].'11')->applyFromArray(array('font' => Style_Font(10,"000000",true,true)));              

                // $objPHPExcel->getActiveSheet()->getStyle('C'.$st_dat.':'.'E'.$count_data)->applyFromArray(array('font' => Style_Font(10,"000000",false,true)));

                // $objPHPExcel->getActiveSheet()->getStyle('C' . ($count_data+3) .':'. 'D' .($count_data+3))->applyFromArray(array('font' => Style_Font(9,"000000",true,true)));

                // $objPHPExcel->getActiveSheet()->getStyle('F8' .':'. $col_name[$count_index] . '8')->applyFromArray(array('fill' => Style_Fill('d9ffb3')));
                // $objPHPExcel->getActiveSheet()->getStyle('F8' .':'. $col_name[$count_index] . '8')->applyFromArray(array('font' => Style_Font(10,"0000ff",true,true)));

                // $objPHPExcel->getActiveSheet()->getStyle('F11' .':'. $col_name[$count_index] . '11')->applyFromArray(array('fill' => Style_Fill('eaeae1')));
                // $objPHPExcel->getActiveSheet()->getStyle('F11' .':'. $col_name[$count_index] . '11')->applyFromArray(array('font' => Style_Font(10,"ff0000",true,true)));


                // $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'8'.':'. $col_name[($count_index+4)] . '8')->applyFromArray(array('fill' => Style_Fill('d9ffb3')));
                // $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'8'.':'. $col_name[($count_index+4)] . '8')->applyFromArray(array('font' => Style_Font(10,"0000ff",true,true)));

                // $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'11'.':'. $col_name[($count_index+4)] . '11')->applyFromArray(array('fill' => Style_Fill('eaeae1')));
                // $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'11'.':'. $col_name[($count_index+4)] . '11')->applyFromArray(array('font' => Style_Font(10,"ff0000",true,true)));
                
                // $objPHPExcel->getActiveSheet()->getStyle('C2:'.'E'.'3')
                //                               ->applyFromArray(array(
                //                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'305496'))));

                // $objPHPExcel->getActiveSheet()->getStyle('C6:'.$col_name[$count_index].'6')
                //                               ->applyFromArray(array(
                //                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));

                // $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'6'.':'.$col_name[($count_index+4)].'6')
                //                               ->applyFromArray(array(
                //                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));                                              

                // $objPHPExcel->getActiveSheet()->getStyle('C11:'.$col_name[$count_index].'11')
                //                               ->applyFromArray(array(
                //                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));

                // $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'11'.':'.$col_name[($count_index+4)].'11')
                //                               ->applyFromArray(array(
                //                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000')))); 

                // $objPHPExcel->getActiveSheet()->getStyle('F7:'.$col_name[$count_index].'11')
                //                               ->applyFromArray(array(
                //                                 'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

                // $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'7'.':'.$col_name[($count_index+4)].'11')
                //                               ->applyFromArray(array(
                //                                 'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_DOTTED,'a6a6a6'))));

                // $objPHPExcel->getActiveSheet()->getStyle('F7:'.$col_name[$count_index].'11')
                //                               ->applyFromArray(array(
                //                                 'borders' => array('left'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));

                // $objPHPExcel->getActiveSheet()->getStyle('C'.$st_col.':'.$col_name[$count_index].$st_col)
                //                               ->applyFromArray(array(
                //                                 'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));

                // $objPHPExcel->getActiveSheet()->getStyle( 'F' . $st_dat .':'. 'F' . $count_data )
                //                               ->applyFromArray(array(
                //                                 'borders' => array('left'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));


                // $objPHPExcel->getActiveSheet()->getStyle( 'C' . ($count_data+3) .':'. 'D' .($count_data+3) )
                //                               ->applyFromArray(array(
                //                                 'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));    


                                                              
//echo $col_name[42]; exit;

               //  $objPHPExcel->getActiveSheet()->getStyle('G'.$st_dat.':'.$col_name[$count_index+6].$count_data)
               //                                ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');


                 $objPHPExcel->getActiveSheet()->getStyle( 'E'.$st_dat.':'.$col_name[$count_index].$count_data )
                                               ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');

                $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('1.71');
                //$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('2.71');
                //$objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$count_index])->setWidth('2.71');     #A
                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[1])->setWidth('27.71');     #B no
                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[2])->setWidth('38.71');    #I model
                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[3])->setWidth('20.71');    #J
                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[4])->setWidth('14.71');    #K
                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$count_index])->setWidth('22.71');     #L
               //  $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+1)])->setWidth('3.71');    #M
               //  $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+2)])->setWidth('3.71');    #N
               //  $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+3)])->setWidth('10.71');   #M
               //  $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+4)])->setWidth('16.71');   #M
               //  $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+5)])->setWidth('3.71');    #N                
                 foreach(range(5, ($count_index-1) ) as $r)
                 $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$r])->setWidth('11.71');   #M                 


               
               // // Style_Alignment('J3',3, false, $objPHPExcel);
               //  Style_Alignment(('B'.$st_col.':'.$col_name[$count_index+6].$st_col), 3, false, $objPHPExcel);
               //  Style_Alignment(('G'.'7'.':'.$col_name[$count_index+6].$count_data), 6, false, $objPHPExcel);
               //  Style_Alignment(('B2'.':'.'E'.$count_data), 3, false, $objPHPExcel);
               //  Style_Alignment(($col_name[($count_index+3)].'7'.':'.$col_name[($count_index+3)].$count_data), 9, false, $objPHPExcel);
               //  //Style_Alignment(('B'.$st_dat.':'.'I'.$count_data), 9, false, $objPHPExcel);
                //$objPHPExcel->getActiveSheet()->mergeCells('C2:'.'E3');
               //  $objPHPExcel->getActiveSheet()->mergeCells($col_name[($count_index+3)].$st_col.':'.$col_name[($count_index+4)].$st_col);
               //  $objPHPExcel->getActiveSheet()->mergeCells($col_name[($count_index+4)].'9'.':'.$col_name[($count_index+4)].'10');
               $objPHPExcel->getActiveSheet()->mergeCells('G3:'.'M4');

               $objPHPExcel->getActiveSheet()->mergeCells('N2:'.'O3');
               $objPHPExcel->getActiveSheet()->mergeCells('N4:'.'O5');

               $objPHPExcel->getActiveSheet()->mergeCells('P2:'.'Q3');
               $objPHPExcel->getActiveSheet()->mergeCells('P4:'.'Q5');

               $objPHPExcel->getActiveSheet()->mergeCells('R2:'.'R3');
               $objPHPExcel->getActiveSheet()->mergeCells('R4:'.'R5');

                //$objPHPExcel->getActiveSheet()->duplicateStyle( $objPHPExcel->getActiveSheet()->getStyle( $col_name[44].($st_dat+4) ), ('C14:C18') );


                //Style_group_lv1_Col($col_name, 4, $objPHPExcel );


                // if ( date('d') < 3  )
                //     {
                //         foreach (range( 13 , $count_index ) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel );
                //     }
                // elseif ( date('d') > 3 && date('d') < $lastmount ) 
                //     {
                //         foreach (range( 6 , ( (date('d')-4)+5 ) ) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel );

                //         foreach (range( ((date('d')+4)+5) , $count_index ) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel );
                                                                         
                //     }
                // else
                //     {
                //         foreach (range( 6 , ($count_index-2) ) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel );
                //     }

                //     if( (date('d')-0) != 1 )
                //     {
                //         $objPHPExcel->getActiveSheet()->getStyle( $col_name[( (date('d')-1)+5 )].($st_col-1).':'. $col_name[( (date('d')-1)+5 )].$count_data )
                //                                       ->applyFromArray(array(
                //                                             'borders' => array('outline' => Style_border(PHPExcel_Style_Border::BORDER_THICK,'00cc00'))));     #GREEN IS YESTERDAY

                //         $objPHPExcel->getActiveSheet()->setCellValue( $col_name[( (date('d')-1)+5 )].($st_col-1), 'Yesterday');

                //         $objPHPExcel->getActiveSheet()->getStyle($col_name[( (date('d')-1)+5 )].($st_col-1))->applyFromArray(array('fill' => Style_Fill('00cc00')));



                //         $objPHPExcel->getActiveSheet()->getStyle( $col_name[( (date('d')-1)+5 )].($st_col-1) )
                //                                       ->applyFromArray(array(
                //                                             'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'ffff4d'))));  

                //     }
                //         $objPHPExcel->getActiveSheet()->getStyle( $col_name[( (date('d')-0)+5 )].($st_col-1).':'. $col_name[( (date('d')-0)+5 )].$count_data )
                //                                       ->applyFromArray(array(
                //                                             'borders' => array('outline' => Style_border(PHPExcel_Style_Border::BORDER_THICK,'ff0000'))));   #RED IS TODAY

                //         $objPHPExcel->getActiveSheet()->getStyle( $col_name[( (date('d')-0)+5 )].($st_col-1) )
                //                                       ->applyFromArray(array(
                //                                             'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'ffff4d'))));   
                                                            

                //         $objPHPExcel->getActiveSheet()->setCellValue( $col_name[( (date('d')-0)+5 )].($st_col-1), 'Today'); 


                //         $objPHPExcel->getActiveSheet()->getStyle($col_name[( (date('d')-0)+5 )].($st_col-1))->applyFromArray(array('fill' => Style_Fill('ff0000')));

                //         $objPHPExcel->getActiveSheet()->getStyle($col_name[( (date('d')-1)+5 )].($st_col-1))->applyFromArray(array('font' => Style_Font(9,"ffff4d",true,true)));
                //         $objPHPExcel->getActiveSheet()->getStyle($col_name[( (date('d')-0)+5 )].($st_col-1))->applyFromArray(array('font' => Style_Font(9,"ffff4d",true,true)));

                //         Style_Alignment( ($col_name[( (date('d')-1)+5 )].($st_col-1) . ':' . $col_name[( (date('d')-0)+5 )].($st_col-1) ), 3, false, $objPHPExcel);


                //         $objPHPExcel->getActiveSheet()->getStyle( 'C'. $count_data .':'. $col_name[$count_index] . $count_data )
                //                                           ->applyFromArray(array(
                //                                             'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));                                                                                                                                                                         



                //echo ($count_data+4); exit;

                // foreach(range(4, 9) as $r)
                // {
                //     $objPHPExcel->getActiveSheet()->mergeCells('L'.$r.':'.'N'.$r);
                //     $objPHPExcel->getActiveSheet()->mergeCells('J'.$r.':'.'K'.$r);                    
                // }

                // $objPHPExcel->getActiveSheet()->mergeCells('B' . ($count_data+3) .':'. 'D' .($count_data+3));

                // foreach(range(($count_data+4), ($count_data+6)) as $r)
                // {
                //     $objPHPExcel->getActiveSheet()->mergeCells('C'.$r.':'.'D'.$r);                 
                // }               

#========================================================================================================================  Put data ====================================================================================         
    } else {
                    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('A1', "No data ".$til.".");
                    $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true,false,'Franklin Gothic Book')));
    }
$ind++;


//echo $til; exit;

}
//exit;
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


function holiday($dat, $hol)
{

//echo $dat;
    foreach ($hol as $ld) 
        if ( substr( $ld['d_t'], 8,2 ) == $dat ) 
            return true;

}
?>