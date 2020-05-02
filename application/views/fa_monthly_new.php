<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');
if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/PHPExcel-1.8.1/Classes/PHPExcel.php';

//============================================================================================= date =================================================================================
$dayA = date('d');
$dayB = date('d');
$dayC = date('d');
$monthA = date('M');
$yearA = date('Y');
$lastmount = substr(date('Y-m-t',strtotime('today')),8, 2);
$todayA = date('Y-M-d', strtotime($yearA."-".$monthA."-".$dayA));
$monthYes = date('m');
if (date('d') == "01"){
    $monthYes = date('m')-1;
    $dayA = substr(date('Y-m-t',strtotime($yearA."-".$monthYes."-".$dayA)),8, 2);
} else { 
    $monthYes = date('m'); $dayA = date('d') - 1; 
}

//-----------------------------------------------------------------------------------------------------------------------------
if(strlen($dayA) < 2) $dayA = "0".$dayA;
$yesterdayA = date('Y-M-d', strtotime($yearA."-".$monthYes."-".$dayA));
$dayA = date('d')-2;
if(strlen($dayA) < 2)  $dayA = "0".$dayA;
if(date('d') == "01")  $dayA = intval(substr(date('Y-m-t',strtotime($yearA."-".$monthYes."-".$dayA)),8, 2))-1;
$yesterdayB = date('Y-M-d', strtotime($yearA."-".$monthYes."-".$dayA));
$dayA = date('d')-3;
if(strlen($dayA) < 2) $dayA = "0".$dayA;
if(date('d') == "01")  $dayA = intval(substr(date('Y-m-t',strtotime($yearA."-".$monthYes."-".$dayA)),8, 2))-2;
$yesterdayC = date('Y-M-d', strtotime($yearA."-".$monthYes."-".$dayA));
//-----------------------------------------------------------------------------------------------------------------------------

//echo $yesterdayA."<br>".$yesterdayB."<br>".$yesterdayC; exit;

$dayA = ( $lastmount  == date('d') )   ? date('Y-M-d', strtotime(date('Y')."-".(date('m')+1)."-". '1')) : date('Y-M-d', strtotime( date('Y') . "-" . date('m') . "-" . (date('d')+1) )) ;

$dayB = ( $lastmount  == date('d') )   ? date('Y-M-d', strtotime(date('Y')."-".(date('m')+1)."-". '2' )): date('Y-M-d', strtotime( date('Y') . "-" . date('m') . "-" . (date('d')+2) )) ;
$tomorrowA = $dayA;
$tomorrowB = $dayB;
//============================================================================================= date =================================================================================

$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
                             ->setLastModifiedBy("Maarten Balliauw")
                             ->setTitle("Office 2007 XLSX Test Document")
                             ->setSubject("Office 2007 XLSX Test Document")
                             ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
                             ->setKeywords("office 2007 openxml php")
                             ->setCategory("Test result file");

$objPHPExcel->setActiveSheetIndex(0);
$col_name = array();
foreach ( range('A', 'Z') as $cm ) {array_push($col_name, $cm);}
    

foreach ( range('A', 'Z') as $cm ) {array_push($col_name, "A$cm");}
    
foreach ( range('A', 'Z') as $cm ) {array_push($col_name, "B$cm");}
$Today = ( (date('d')+0) == 1 ) ? date('d-M-Y', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : date('d-M-Y', strtotime(date('Y')."-".(date('m')+0)."-".(date('d')-1)));

$Onday = ( (date('d')+0) == 1 ) ? 32   : date('d', strtotime(date('Y')."-".(date('m')+0)."-".(date('d'))));
$Yeday = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : date('d', strtotime(date('Y')."-".(date('m')+0)."-".(date('d')-1)));


$MontCol = ( (date('d')+0) == 1 ) ? date('M', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('M'));
$MontFul = ( (date('d')+0) == 1 ) ? date('M', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('F'));
$YearCol = ( (date('d')+0) == 1 ) ? date('Y', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('Y'));
//echo $dateCol . '/' . $MontCol; exit;


//var_dump($col_index); exit;
//var_dump($head2); exit;

//=======================================================================================  config Style ================================================================================
$indSheet = 0;
foreach ($title as $inTil => $til) 
{
        $sheetIndex =  strtolower(str_replace(' ', '_', $title[$inTil]));
        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex($inTil);
        $objPHPExcel->getActiveSheet()->setTitle("$til");
        $objPHPExcel->getActiveSheet()->setShowGridlines(False);
        $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 30 );


        $objPHPExcel->getActiveSheet()->freezePane('A9');
        $objPHPExcel->getActiveSheet()->freezePane('I9');
      //  $objPHPExcel->getActiveSheet()->freezePane('M5');   
      //  $objPHPExcel->getActiveSheet()->freezePane('M5');            
        // $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('5.5');
        // $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('6.5');
        // $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('8');         
        // $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('10');
        // $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('65');
     $i = 3;
     $day = 1;
     if(count($list_act_report[$sheetIndex]) > 0 )
     {
        foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
        { 
            $objPHPExcel->getActiveSheet()
                    ->getStyle('1:3') 
                    ->getAlignment()
                    ->setWrapText(true)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            //echo $key . "<hr>";
            $key = str_replace("_REV", ".", $key);
            if($key != 'NO')
            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."6", str_replace("_", " ", strtoupper($key)));
                 

        }
         //foreach(range('A','Z') as $columnID) { $objPHPExcel->getActiveSheet()->getColumnDimension('B'.$columnID)->setAutoSize(true); }         
     } else { 
            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('A1', "No Data ".$til.".");
            $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(22,'000000',true)));
            $objPHPExcel->getActiveSheet()
                    ->getStyle('1:3') 
                    ->getAlignment()
                    ->setWrapText(false)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            }

//exit;
//=======================================================================================  Input data ================================================================================
$row = 3;

// foreach ($list_act_report as $key => $value) 
    {
                //echo substr('DATE1',4,2); exit;
             //   var_dump($key); exit;
     if(count($list_act_report[$sheetIndex]) > 0 )
         { 
                   // if ($key == 'fa_supply_list') 
                   // {
                   $objPHPExcel->setActiveSheetIndex($indSheet);
                    $st_cal = 6 ; 
                    $cu_cal = 9 ;
                    $st_dat = 9 ;
                    $cu_dat = count($list_act_report[$sheetIndex]) ;
                    //$objPHPExcel->getActiveSheet()->insertNewRowBefore(1,2);
                    //$objPHPExcel->getActiveSheet()->freezePane('M4');
                    $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 5 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 30 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 35 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 32 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 5 )->setRowHeight( 25 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 6 )->setRowHeight( 35 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 7 )->setRowHeight( 50 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 8 )->setRowHeight( 12 );                  
                    $objPHPExcel->getActiveSheet()
                                ->getStyle(('D'.$st_cal.':'.$col_name[55].$st_cal))
                                ->getAlignment()
                                ->setWrapText(true)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 


                    $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80); 
                     
                    $objPHPExcel->getActiveSheet()->setAutoFilter('D8:'.$col_name[55].'8');


                    $objPHPExcel->getActiveSheet()
                                ->getStyle('D5:'.$col_name[55].(count( $list_act_report[$sheetIndex] )+8))
                                ->getAlignment()
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)   //*** set left
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);   //*** set center
                    $objPHPExcel->getActiveSheet()
                                ->getStyle('D6:'.$col_name[55].(count( $list_act_report[$sheetIndex] )+8))
                                ->getAlignment()
                                ->setHorizontal(PHPExcel_Style_Alignment::VERTICAL_CENTER)   //*** set left
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);   //*** set center     
                    $objPHPExcel->getActiveSheet()
                                ->getStyle('N1:P1')
                                ->getAlignment()
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);  

            //        $objPHPExcel->getActiveSheet()->setAutoFilter('D8:'.$col_name[53].'8');

                    $objPHPExcel->getActiveSheet()->getStyle('D6:'.$col_name[55].'6')->applyFromArray(array('fill'    => Style_Fill($colhead)));
                 //   $objPHPExcel->getActiveSheet()->getStyle('V6:'.$col_name[52].'6')->applyFromArray(array('fill'    => Style_Fill('00ffcc')));

                    $objPHPExcel->getActiveSheet()->getStyle('V6:'.$col_name[55].'6')->applyFromArray(array('font'    => Style_Font(10, '000000', true, 'Calibri'))); 


                    $objPHPExcel->getActiveSheet()->getStyle('D6:'.$col_name[55].'6')->applyFromArray(array('font'    => Style_Font(10, 'FFFFFF', true, 'Calibri')));  

                    $objPHPExcel->getActiveSheet()->getStyle('J5:BD5')->applyFromArray(array('font'    => Style_Font(12, '000000', true, 'Calibri')));      
                    $objPHPExcel->getActiveSheet()->getStyle('D9:'.$col_name[55].(count( $list_act_report[$sheetIndex] )+8))->applyFromArray(array('font'    => Style_Font(11, '000000', false, 'Calibri')));
                //   $objPHPExcel->getActiveSheet()->setCellValue('J6', 'STOCK TODAY');
                    $objPHPExcel->getActiveSheet()->getStyle('D6:'.$col_name[55].'7')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                    $objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true)
                                ->setName('Consolas')
                                ->setSize(11)
                                ->getColor()->setRGB('FFFFFF');


                      $objPHPExcel->getActiveSheet()->getStyle('Y3:BA3')->applyFromArray(array('fill' => Style_Fill('b3d9ff')));

                      $objPHPExcel->getActiveSheet()->getStyle('N3:Q3')->applyFromArray(array('fill' => Style_Fill('ccccff')));  
                      $objPHPExcel->getActiveSheet()->getStyle('N4:Q4')->applyFromArray(array('fill' => Style_Fill('ccccff'))); 

                      $objPHPExcel->getActiveSheet()->getStyle('M3:M4')->applyFromArray(array('fill' => Style_Fill('ffcccc'))); 
                      $objPHPExcel->getActiveSheet()->getStyle('R3:R4')->applyFromArray(array('fill' => Style_Fill('ffcc99')));
                    // $objPHPExcel->getActiveSheet()->getStyle('Q5:W5')->applyFromArray(array('fill' => Style_Fill('b3d9ff')));
                    
                    //====================================================================================================================================//

                      $TOTAL_TIME_ALLMAN = 'SUBTOTAL(9,P9:P'. (count( $list_act_report[$sheetIndex] )+5) . ")";
                      $ACTUAL = 'SUBTOTAL(9,M9:M'. (count( $list_act_report[$sheetIndex] )+5) . ")";
                      $objPHPExcel->getActiveSheet()->setCellValue('R4', '=('.$TOTAL_TIME_ALLMAN.')/'.$ACTUAL);
           
         
                    //====================================================================================================================================//


                     // $objPHPExcel->getActiveSheet()->setCellValue('T5', '=('.$sumtotaltime.'/'.$sumactual.')');

                        
                     $objPHPExcel->getActiveSheet()->getStyle('D8:'.$col_name[55].'8')->applyFromArray(array('fill'    => Style_Fill('b3cccc')));
                    
                    $startData = 9;
                    $r = 9;
                            foreach ($list_act_report[$sheetIndex] as $nr => $val) 
                            {
                                $indCol = 3;
                                        foreach ($val as $rowData => $data) 
                                        {
                                           if($rowData != 'NO')
                                           {
                                            if ($rowData == 'MODEL') 
                                            {
                                                if ($data == '3E00') 
                                                {
                                                   $objPHPExcel->getActiveSheet()->getStyle($col_name[$indCol].($r))->getNumberFormat()->setFormatCode('###"E00"');
                                                   $objPHPExcel->getActiveSheet()->setCellValue($col_name[$indCol++].($r), $data);
                                                }
                                                else
                                                {
                                                   $objPHPExcel->getActiveSheet()->setCellValue($col_name[$indCol++].($r), $data);
                                                }
                                                    
                                            }
                                           
                                            else
                                            {
                                               $objPHPExcel->getActiveSheet()->setCellValue($col_name[$indCol++].($r), $data);
                                            }

                                            // if ($body == 'PD' && $val == 'PD04')
                                            //       {
                                            //            $eff = '(R'.$row.'/'.'K'.$row.')';
                                            //            $objPHPExcel->getActiveSheet()->setCellValue('U'.$row, '=IFERROR(IF(R'.$row.'="",0,'.$eff.'),0)');
                                            //           //echo $minusTime. "<hr>" ; 
                                            //           //echo $minusTime ; exit;                          
                                                   
                                            //       }



                                            // if ($val['STOCK_TODAY'] < $val['PLAN_QTY'] )
                                            // {

                                            // 	$objPHPExcel->getActiveSheet()->getStyle('D'.($r).':'.$col_name[15].($r))->applyFromArray(array('fill'    => Style_Fill('ffffcc')));
                                            // 	$objPHPExcel->getActiveSheet()->getStyle('D'.($r).':'.$col_name[15].($r))->applyFromArray(array('font'    => Style_Font(11, 'ff0000', false, 'Calibri')));
                                            // }
                                            // if ($val['SUP_FROM'] == '')
                                            // {
                                            // 	$objPHPExcel->getActiveSheet()->getStyle('D'.($r).':'.$col_name[15].($r))->applyFromArray(array('fill'    => Style_Fill('d9d9d9')));
                                            // 	$objPHPExcel->getActiveSheet()->getStyle('D'.($r).':'.$col_name[15].($r))->applyFromArray(array('font'    => Style_Font(11, '000000', true, 'Calibri')));

                                            // }
                                            // if ($rowData == 'PRODUCT_TYP' && $data != '10' && $val['SUP_FROM'] == '') 
                                            // {

                                            //     $objPHPExcel->getActiveSheet()->setCellValue($col_name[$indCol -5 ].($r),'');
                                            // }

                                         }   

                                        } 
                                $r++;
                            }
        $Montlast = date('F Y', strtotime(date('Y')."-".(date('m')-1)."-".'1'));  
        $M = date('M', strtotime(date('Y')."-".(date('m')-1)."-".'1'));
        $Daylast  = substr(date('Y-m-t',strtotime(date('Y')."-".(date('m')-1)."-".'1')),8, 2);
        //echo $Montlast.$Daylast;exit;
                            $objPHPExcel->getActiveSheet()->setCellValue('D2', "FA SUMMARY REPORT OF ".$Montlast);
                            $objPHPExcel->getActiveSheet()->setCellValue('D4', "Data of : (  01 ".strtoupper($M)." - ".$Daylast." ".strtoupper($M)."  )");

                        //    $objPHPExcel->getActiveSheet()->setCellValue('D3', "FA SUMMARY REPORT OF " .strtoupper(date('d-M-Y',  strtotime((date('d')+1) . "-" . date('M') . "-" . date('Y')) ))); //('D2', "FA Supply list of ".$Montlast);

                             $objPHPExcel->getActiveSheet()->setCellValue('Y3', "Accum Loss Time [min] ");
                             $objPHPExcel->getActiveSheet()->setCellValue('J4', "Total >>");
                             $objPHPExcel->getActiveSheet()->setCellValue('U4', "Total Loss >>");
                             $objPHPExcel->getActiveSheet()->setCellValue('U5', "%Loss Time >>");
                             $objPHPExcel->getActiveSheet()->setCellValue('U7', "Detail Loss >>");
                            $objPHPExcel->getActiveSheet()->setCellValue('M3', "Actual QTY");
                            $objPHPExcel->getActiveSheet()->setCellValue('N3', "Production Time (Min)");
                            $objPHPExcel->getActiveSheet()->setCellValue('R3', "MANHOUR/PCS");


                             

                            $objPHPExcel->getActiveSheet()
                                ->getStyle(('D3'.':'.$col_name[55].'3'))
                                ->getAlignment()
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 
                      //      $objPHPExcel->getActiveSheet()->setCellValue('D4', "Period : (  01 ".strtoupper($M)." - ".$Daylast." ".strtoupper($M)."  )");
                      //      $objPHPExcel->getActiveSheet()->setCellValue('D4', "Data of :(  01 ".strtoupper($M)." - ".$Daylast." ".strtoupper($M).")";
                            $objPHPExcel->getActiveSheet()
                                ->getStyle(('D4'.':'.$col_name[55].'4'))
                                ->getAlignment()
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // example title 
                            $objPHPExcel->getActiveSheet()
                                ->getStyle(('D5'.':'.$col_name[55].'4'))
                                ->getAlignment()
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


                         //   $objPHPExcel->getActiveSheet()->setCellValue('P3', "PRODUCTION DATE \n".strtoupper(date('d-M-Y',  strtotime((date('d')+1) . "-" . date('M') . "-" . date('Y')) )));
                            

        foreach (range(12, 10) as $col) 
        //    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col].'2', '=SUBTOTAL(9,'. $col_name[$col] .'5:'. $col_name[$col] . (count( $list_act_report[$key] )+5) . ")" );
#126180
                      //      $objPHPExcel->getActiveSheet()->getStyle('G2:L2')->applyFromArray(array('font'    => Style_Font(14, '000000', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('D2')->applyFromArray(array('font'    => Style_Font(24, '002d4d', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('Y3')->applyFromArray(array('font'    => Style_Font(24, '000000', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('D4')->applyFromArray(array('font'    => Style_Font(20, '000000', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('J4')->applyFromArray(array('font'    => Style_Font(16, 'ff0000', true, 'Calibri')));

                            $objPHPExcel->getActiveSheet()->getStyle('R3')->applyFromArray(array('font'    => Style_Font(11, 'ff0000', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('U4')->applyFromArray(array('font'    => Style_Font(11, 'ff0000', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('U5')->applyFromArray(array('font'    => Style_Font(11, 'ff0000', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('U7')->applyFromArray(array('font'    => Style_Font(11, 'ff0000', true, 'Calibri')));


                            $objPHPExcel->getActiveSheet()->getStyle('Q3')->applyFromArray(array('font'    => Style_Font(11, '006699', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('L3:P3')->applyFromArray(array('font'    => Style_Font(16, '006699', true, 'Browallia New')));  

                            $objPHPExcel->getActiveSheet()->getStyle('L4:R4')->applyFromArray(array('font'    => Style_Font(16, '000000', true, 'Browallia New')));     
                            $objPHPExcel->getActiveSheet()->getStyle('Y4:BD4')->applyFromArray(array('font'    => Style_Font(16, '000000', true, 'Browallia New')));
                            $objPHPExcel->getActiveSheet()->getStyle('Y5:BD5')->applyFromArray(array('font'    => Style_Font(16, '006699', true, 'Browallia New'))); 
                            $objPHPExcel->getActiveSheet()->getStyle('Y7:BD7')->applyFromArray(array('font'    => Style_Font(12, '009999', true, 'Browallia New')));     
                            // $objPHPExcel->getActiveSheet()->getStyle('L4')->applyFromArray(array('font'    => Style_Font(14, '002d4d', true, 'Calibri')));
                            // $objPHPExcel->getActiveSheet()->getStyle('P4')->applyFromArray(array('font'    => Style_Font(14, '002d4d', true, 'Calibri')));




                          
                           
                            // $objPHPExcel->getActiveSheet()->getStyle('L4:'.'R4'.($cu_dat+10))->applyFromArray(array(
                            //                                         'borders' => array('outline' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                            $objPHPExcel->getActiveSheet()->getStyle('D9:'.'BD'.($cu_dat+9))->applyFromArray(array(
                                                                     'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'c2c2d6'))));


                            $objPHPExcel->getActiveSheet()->getStyle('Y7:BD7')->getAlignment()->setTextRotation(90); 
                            $objPHPExcel->getActiveSheet()->getStyle('L4:R4')->getFont()->setUnderline(true);
                            $objPHPExcel->getActiveSheet()->getStyle('Y4:BD4')->getFont()->setUnderline(true);
                          
                            
                    
                          $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setVisible(false); //hide
                       //   $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setVisible(false); //hide
                          $objPHPExcel->getActiveSheet()->getColumnDimension('S')->setVisible(false); //hide
                          $objPHPExcel->getActiveSheet()->getColumnDimension('T')->setVisible(false); //hide
                          $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setVisible(false); //hide
                          $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setVisible(false); //hide
                          $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setVisible(false); //hide

              
                      //    $objPHPExcel->getActiveSheet()->mergeCells('Q3:Q4');
                          $objPHPExcel->getActiveSheet()->mergeCells('N3:Q3');
                          $objPHPExcel->getActiveSheet()->mergeCells('D2:H3');
                          $objPHPExcel->getActiveSheet()->mergeCells('D4:H4');
                          $objPHPExcel->getActiveSheet()->mergeCells('J4:K4');
                          $objPHPExcel->getActiveSheet()->mergeCells('Y3:BD3');


                          // $objPHPExcel->getActiveSheet()->getStyle('J9'.':J'.( count( $list_act_report[$sheetIndex] )+8) )
                          //                                 ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-'); // //_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-

                          $objPHPExcel->getActiveSheet()->getStyle('K9'.':K'.( count( $list_act_report[$sheetIndex] )+8) )
                                                          ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');

                       	  // $objPHPExcel->getActiveSheet()->getStyle('L5'.':L'.( count( $list_act_report[$sheetIndex] )+8) )
                          //                                 ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');

                          $objPHPExcel->getActiveSheet()->getStyle('L4'.':P4'.( count( $list_act_report[$sheetIndex] )+8) )
                                                          ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');                              

                          $objPHPExcel->getActiveSheet()->getStyle('Q4'.':Q'.( count( $list_act_report[$sheetIndex] )+8) )
                                                          ->getNumberFormat()->setFormatCode('_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)');

                          $objPHPExcel->getActiveSheet()->getStyle('R4'.':U4'.( count( $list_act_report[$sheetIndex] )+8) )
                                                          ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');

                          $objPHPExcel->getActiveSheet()->getStyle('V4'.':BD4'.( count( $list_act_report[$sheetIndex] )+8) )
                                                          ->getNumberFormat()->setFormatCode('_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)');

                          $objPHPExcel->getActiveSheet()->getStyle('U9'.':U'.( count( $list_act_report[$sheetIndex] )+8) )
                                                         ->getNumberFormat()->setFormatCode('_*#,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-'); 

                          $objPHPExcel->getActiveSheet()->getStyle('X9'.':X'.( count( $list_act_report[$sheetIndex] )+8) )
                                                          ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');

                          // $objPHPExcel->getActiveSheet()->getStyle('U1'.':U4'.( count( $list_act_report[$sheetIndex] )+8) )
                          //                                 ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');

                          // $objPHPExcel->getActiveSheet()->getStyle('S9'.':U'.( count( $list_act_report[$sheetIndex] )+8) )
                          //                                 ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-'); 

                          // $objPHPExcel->getActiveSheet()->getStyle('W1'.':W4'.( count( $list_act_report[$sheetIndex] )+8) )
                          //                                 ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
                          // // //_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-

                          //  $objPHPExcel->getActiveSheet()->getStyle('S5'.':T5'.( count( $list_act_report[$sheetIndex] )+8) )
                          //                                 ->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');  

                          // $objPHPExcel->getActiveSheet()->getStyle('J5'.':R5'.( count( $list_act_report[$sheetIndex] )+8) )
                          //                                 ->getNumberFormat()->setFormatCode('_* #,##0_-;[RED](#,##0)_-;_* [BLACK]"-"??_-;_-@_-');
             //============================================================ Uhide accumm loss ======================================================================//

      //       foreach (range(12, 22) as $index) Style_group_lv1();                                              
             // $objPHPExcel->setActiveSheetIndex()->setCellValue('BB1',"BB");
             // $objPHPExcel->setActiveSheetIndex()->setCellValue('BC1', "↢ Unhide to view important loss time code");
             // $objPHPExcel->getActiveSheet()->getStyle('BB1')->applyFromArray(array('font' => Style_Font(36,'FF0000',true,'Wingdings 3')));
             // $objPHPExcel->getActiveSheet()->getStyle('BC1')->applyFromArray(array('font' => Style_Font(18,'FF0000',true,'Franklin Gothic Book')));

            //============================================================ Uhide accumm loss ======================================================================//



                            $objPHPExcel->getActiveSheet()->setCellValue('M4', '=SUBTOTAL(9,M9:M'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('N4', '=SUBTOTAL(9,N9:N'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('O4', '=SUBTOTAL(9,O9:O'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('P4', '=SUBTOTAL(9,P9:P'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('Q4', '=SUBTOTAL(9,Q9:Q'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                        //    $objPHPExcel->getActiveSheet()->setCellValue('Q4', '=SUBTOTAL(9,Q9:Q'. (count( $list_act_report[$sheetIndex] )+8) . ")" );






                            $objPHPExcel->getActiveSheet()->setCellValue('Y4', '=SUBTOTAL(9,Y9:Y'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('Z4', '=SUBTOTAL(9,Z9:Z'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AA4', '=SUBTOTAL(9,AA9:AA'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AB4', '=SUBTOTAL(9,AB9:AB'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AC4', '=SUBTOTAL(9,AC9:AC'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AD4', '=SUBTOTAL(9,AD9:AD'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AE4', '=SUBTOTAL(9,AE9:AE'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AF4', '=SUBTOTAL(9,AF9:AF'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AG4', '=SUBTOTAL(9,AG9:AG'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AH4', '=SUBTOTAL(9,AH9:AH'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AI4', '=SUBTOTAL(9,AI9:AI'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AJ4', '=SUBTOTAL(9,AJ9:AJ'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AK4', '=SUBTOTAL(9,AK9:AK'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AL4', '=SUBTOTAL(9,AL9:AL'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AM4', '=SUBTOTAL(9,AM9:AM'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AN4', '=SUBTOTAL(9,AN9:AN'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AO4', '=SUBTOTAL(9,AO9:AO'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AP4', '=SUBTOTAL(9,AP9:AP'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AQ4', '=SUBTOTAL(9,AQ9:AQ'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AR4', '=SUBTOTAL(9,AR9:AR'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AS4', '=SUBTOTAL(9,AS9:AS'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AT4', '=SUBTOTAL(9,AT9:AT'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AU4', '=SUBTOTAL(9,AU9:AU'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AV4', '=SUBTOTAL(9,AV9:AV'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AW4', '=SUBTOTAL(9,AW9:AW'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AX4', '=SUBTOTAL(9,AX9:AX'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('AY4', '=SUBTOTAL(9,AY9:AY'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('AZ4', '=SUBTOTAL(9,AZ9:AZ'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('BA4', '=SUBTOTAL(9,BA9:BA'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('BB4', '=SUBTOTAL(9,BB9:BB'. (count( $list_act_report[$sheetIndex] )+8) . ")" ); 
                            $objPHPExcel->getActiveSheet()->setCellValue('BC4', '=SUBTOTAL(9,BC9:BC'. (count( $list_act_report[$sheetIndex] )+8) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('BD4', '=SUBTOTAL(9,BD9:BD'. (count( $list_act_report[$sheetIndex] )+8) . ")" );                                      

                            $objPHPExcel->getActiveSheet()->getStyle('Y7:'.'BD'.($cu_dat+9))->applyFromArray(array(
                                                                     'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'c2c2d6'))));

                            $objPHPExcel->getActiveSheet()->getStyle('Y2:'.'BD'.($cu_dat+11))->applyFromArray(array(
                                                                     'borders' => array('outline' => Style_border(PHPExcel_Style_Border::BORDER_THICK,'009999'))));

                             // $objPHPExcel->getActiveSheet()->getStyle('L3:'.'Q3'.($cu_dat+9))->applyFromArray(array(
                             //                                         'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'e0ebeb'))));
                             //  $objPHPExcel->getActiveSheet()->getStyle('L4:'.'Q4'.($cu_dat+9))->applyFromArray(array(
                             //                                         'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'e0ebeb'))));




                              $objPHPExcel->getActiveSheet()->setCellValue('Y7' , "Meeting"); 
                              $objPHPExcel->getActiveSheet()->setCellValue('Z7' , "Cleaning, 5S");
                              $objPHPExcel->getActiveSheet()->setCellValue('AA7' , "Planned Maintenance (PM) - in Plan");
                              $objPHPExcel->getActiveSheet()->setCellValue('AB7' , "Planned Maintenance (PM) - out Plan");
                              $objPHPExcel->getActiveSheet()->setCellValue('AC7' , "Training Operator");
                              $objPHPExcel->getActiveSheet()->setCellValue('AD7', "Regular Machine Daily Check / Lot no. change/ Safety co-check");
                              $objPHPExcel->getActiveSheet()->setCellValue('AE7', "Stop adding Flux. / Melting furnance problem");
                              $objPHPExcel->getActiveSheet()->setCellValue('AF7', "Stop Heat up.");
                              $objPHPExcel->getActiveSheet()->setCellValue('AG7', "Mold Clearning ");
                              $objPHPExcel->getActiveSheet()->setCellValue('AH7', " Model,Mold Change - in Plan ");
                              $objPHPExcel->getActiveSheet()->setCellValue('AI7', " Model,Mold Change - out Plan");
                              $objPHPExcel->getActiveSheet()->setCellValue('AJ7', " Tooling/Tip/Sleeve - on Plan");
                              $objPHPExcel->getActiveSheet()->setCellValue('AK7', "Waiting for material from inhouse");
                              $objPHPExcel->getActiveSheet()->setCellValue('AL7', "Coolant, Lubrication, Metal Chip Cleaning");
                              $objPHPExcel->getActiveSheet()->setCellValue('AM7', "Break Down - Machine ");
                              $objPHPExcel->getActiveSheet()->setCellValue('AN7', "Break Down - Mold (Die Casting)");
                              $objPHPExcel->getActiveSheet()->setCellValue('AO7', "Break Down - JIG /Tip/Sleeve");
                              $objPHPExcel->getActiveSheet()->setCellValue('AP7', "Break Down - Tooling");
                              $objPHPExcel->getActiveSheet()->setCellValue('AQ7', "Adjust Program and Condition.");
                              $objPHPExcel->getActiveSheet()->setCellValue('AR7', "Adjustment - Mold (Die Casting)");
                              $objPHPExcel->getActiveSheet()->setCellValue('AS7', "Electricity, Air, Utility Break Down");
                              $objPHPExcel->getActiveSheet()->setCellValue('AT7', "Waiting for material from store");
                              $objPHPExcel->getActiveSheet()->setCellValue('AU7', "Material/Part Quality Problem");
                              $objPHPExcel->getActiveSheet()->setCellValue('AV7', "Personel Reason (toilet,sick,late other)");
                              $objPHPExcel->getActiveSheet()->setCellValue('AW7', "Quality Judgement");
                              $objPHPExcel->getActiveSheet()->setCellValue('AX7', "Waiting data from QC/4M");
                              $objPHPExcel->getActiveSheet()->setCellValue('AY7', "Waiting for Packaging/Box/Partition");
                              $objPHPExcel->getActiveSheet()->setCellValue('AZ7', "PE/CE Trial request (CE, PE)/Audit");
                              $objPHPExcel->getActiveSheet()->setCellValue('BA7', "TPM, Line Kaizen");
                              $objPHPExcel->getActiveSheet()->setCellValue('BB7', "Waiting for supervisor");
                              $objPHPExcel->getActiveSheet()->setCellValue('BC7', "Production of work other than direct production.");
                              $objPHPExcel->getActiveSheet()->setCellValue('BD7', "Re - washing FG part");


             
                            $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('1');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('1');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('3');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('7');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('9');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('17');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('38');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('15');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('5');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth('12');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth('13');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth('14');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth('16');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth('15');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth('14');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth('16');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth('12');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth('14');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth('13');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth('10');

            //============================================================ % loss time ======================================================================//


                         foreach (range(24, 55) as $key) 

                           $objPHPExcel->getActiveSheet()->setCellValue($col_name[$key].'5', '=IFERROR( '.$col_name[$key].'4'.'/$'.'P4'. ',0)');   //Iferror 0/0
                           $objPHPExcel->getActiveSheet()->getStyle('Y5:BD5')->getNumberFormat()->setFormatCode('_*#,##0.00%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');   


              //============================================================ % loss time ======================================================================//




 } 




   //  var_dump($list_act_report); exit;        
        
$row = 5;
 } 
 $indSheet++;


} 

//   if ($indSheet == 'loss code'){
               
//                     $startData = 2;
//                     $r = 2;
//                             foreach ($value as $nr => $val) 
//                             {
//                                 $indCol = 0;
//                                         foreach ($val as $rowData => $data) 
//                                         {

//                                                $objPHPExcel->setActiveSheetIndex($indSheet)->setCellValue($col_name[$indCol++].($r), $data);


//                                         }
//                                 $r++;
//                             }
//                         #========================================format_loss_code
//                             $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('8');
//                             $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('8');
//                             $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('36');
//                             $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('36');
//                             $objPHPExcel->getActiveSheet()->getStyle('A1:D1')->applyFromArray(array('fill'    => Style_Fill('B8CCE4')));                           
  
// }

//  else 
// {

//             $objPHPExcel->setActiveSheetIndex($cu_dat)->setCellValue('A1', "No Production Plan".$til.".");
//             $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true)));
//             //echo "Non data."; exit;
// }


$objPHPExcel->setActiveSheetIndex(0);
  
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





function Style_Fill($color=null) {

    return array( 'type'  => PHPExcel_Style_Fill::FILL_SOLID,                           
                  'color' => array('rgb' => $color)                                    
                );                                   
}

function Style_Font($size=11, $color='FFFFFF', $bol=false, $fname='Consolas') {

    return  array(
                    'name' => $fname,
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



function sunday($dat = '01-01-2018',  $focus = 1, $start_row = 3, $end_row = 5, $col = null, $ind = 0,  $objPHPExcel = nul)
{
    $objPHPExcel->setActiveSheetIndex($ind);
    $d = date('d', strtotime($dat));
    //echo $d; exit;
    $fillSum   = array();
    $fillTotal = array();

    $fillhide = array();
    $indexSun = 12;
    $kla = 0;
    $merge_nosun = '';
    $merge_nohed = '';
    $Sunday = 0;
    foreach (range(1, $d) as $valDay) 
    {        
        $MontCol = ( (date('d')+0) == 1 ) ? date('m', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('m'));
        $YearCol = ( (date('d')+0) == 1 ) ? date('Y', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('Y'));
        $Tday = date('d-M-Y', strtotime($YearCol."-".$MontCol."-".$valDay));
       // echo $Tday; exit;
        $Nday = date('l', strtotime($Tday));
//echo date('d', strtotime($Tday)) . "-" . $Nday . " ---> " . $col[$indexSun] . ":" . $col[($d+11)] . "<hr>";  
        if($Nday == 'Sunday')
        {            
            $Sunday = $indexSun;
            if($indexSun > 12 && $indexSun-6 > 11)
            {
                // $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun-6].$focus)->applyFromArray(array('fill'    => Style_Fill($colhead)));
                $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun-6].$focus, '=SUM(' . $col[$indexSun-6] . $start_row . ":" . $col[$indexSun] . $start_row . ")" );
                $objPHPExcel->getActiveSheet()->mergeCells($col[$indexSun-6] . $focus .":" . $col[$indexSun] . $focus); 

                array_push($fillhide, array('st' => $indexSun-6, 'ed' => $indexSun));
                array_push($fillSum,  $col[$indexSun-6]  . $focus);
                array_push($fillTotal, $col[$indexSun-6] . $start_row . ":" . $col[$indexSun] . $start_row );
                if( $indexSun + 6 > $d+10 && $valDay < $d-1)
                {
                        $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun+1].$focus, '=SUM(' . $col[$indexSun+1] . $start_row . ":" . $col[$d+11] . $start_row . ")" );
                        $objPHPExcel->getActiveSheet()->mergeCells($col[$indexSun+1] . $focus .":" . $col[$d+11] . $focus);

                        array_push($fillhide, array('st' => $indexSun+1, 'ed' => $d+11) );
                        array_push($fillSum, $col[$indexSun+1] . $focus);
                        array_push($fillTotal, $col[$indexSun+1] . $start_row . ":" . $col[$d+11] . $start_row );
                        //break;
                }

               // echo $Sunday; exit;

            }
            //FFC2C2

            elseif($indexSun > 12 && $indexSun-6 < 12)
            {
                $objPHPExcel->getActiveSheet()->setCellValue( $col[12].$focus, '=SUM(' . $col[12] . $start_row . ":" . $col[$indexSun] . $start_row . ")" );
                $objPHPExcel->getActiveSheet()->mergeCells($col[12] . $focus .":" . $col[$indexSun] . $focus);

                array_push($fillhide, array('st' => 12, 'ed' => $indexSun) );
                array_push($fillSum, $col[12] . $focus);
                array_push($fillTotal, $col[12] . $start_row . ":" . $col[$indexSun] . $start_row );
            }
        

            // echo date('d', strtotime($Tday)) . "-" . $Nday . $col[$indexSun] . "<hr>";
            //echo date('d', strtotime($Tday)) . "<hr>";

            $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun].'6'. ":" . $col[$indexSun].$end_row)->applyFromArray(array('fill'    => Style_Fill('FFC2C2'))); 
            $kla = 99;       
        }
        else
        {            
            if($indexSun == 12 && $indexSun+6 > $d)
            {
                $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun].$focus, '=SUM(' . $col[$indexSun] . '3)' );

                $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun].$focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun].'2')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) ); 
                $kla = 99;
            } 
            elseif($indexSun > 12 &&  $indexSun-1 == $Sunday)
            {
                $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun].$focus, '=SUM(' . $col[$indexSun] . '3)' );
                $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun].$focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun].'3')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) ); 
                $kla = 99;
            }             
            elseif($indexSun > 12 && $indexSun-7 < $Sunday)
            {
                $objPHPExcel->getActiveSheet()->setCellValue( $col[$Sunday+1].$focus, '=SUM(' . $col[$Sunday+1] . '3' . ":" . $col[$indexSun] . '3' . ")" );
                $objPHPExcel->getActiveSheet()->getStyle($col[$Sunday+1].$focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                $objPHPExcel->getActiveSheet()->getStyle($col[$Sunday+1] . '3' . ":" . $col[$indexSun] . '3')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) );
                $merge_nosun = $col[$Sunday+1] . $focus . ":" . $col[$indexSun] . $focus;
                $kla = 0;
            }   
            elseif($indexSun > 12 && $indexSun-6 < 12)
            {
                $objPHPExcel->getActiveSheet()->setCellValue( $col[12].$focus, '=SUM(' . $col[$Sunday+1] . '3' . ":" . $col[$indexSun] . '3' . ")" );

                $objPHPExcel->getActiveSheet()->getStyle($col[12].$focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                $objPHPExcel->getActiveSheet()->getStyle($col[12] . '3' . ":" . $col[$indexSun] . '3')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) );
                $merge_nosun = $col[12] . $focus . ":" . $col[$indexSun] . $focus;
                $kla = 0;
            }    
            //else{ echo $Sunday.' Game '.$indexSun; exit;}
            // elseif($indexSun > 12 && $indexSun-6 > 11)
            // {
            //     // $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun-6].$focus)->applyFromArray(array('fill'    => Style_Fill($colhead)));
            //     $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun-6].$focus, '=SUM(' . $col[$indexSun-6] . '2' . ":" . $col[$indexSun] . '2' . ")" );
            //     $objPHPExcel->getActiveSheet()->mergeCells($col[$indexSun-6] . $focus .":" . $col[$indexSun] . $focus); 
            //     $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun-6] . $focus .":" . $col[$indexSun] . $focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
            //     $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun-6] . '2' .":" . $col[$indexSun] . '2')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) ); 
            //     $kla = 0; 
            //     $merge_nosun = $col[$indexSun-6] . $focus .":" . $col[$indexSun] . $focus;
            //     if( $indexSun + 6 > $d+10 && $valDay < $d-1)
            //     {
            //             $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun+1].$focus, '=SUM(' . $col[$indexSun+1] . '2' . ":" . $col[$d+11] . '2' . ")" );
            //              $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun+1] . $focus .":" . $col[$d+11] . $focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
            //              $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun+1] . '2' .":" . $col[$d+11] . '2')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) ); 
            //              $merge_nosun = $col[$indexSun+1] . $focus .":" . $col[$d+11] . $focus;
            //              $kla = 0;
            //             //break;
            //     }


            // }            
        }
    $indexSun++;
    }
//exit;
    if($kla == 0) $objPHPExcel->getActiveSheet()->mergeCells($merge_nosun);

    foreach ($fillSum as $key => $value) 
    {
                                if ($key == 0)
                                {
                                                    //echo $col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1"; exit;  
                                                    

                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('4F6228') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('D8E4BC') ) );//D8E4BC


                                            $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 1' );
                                            $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['ed']]."1" );

                                            foreach (range($fillhide[$key]['st'], $fillhide[$key]['ed']) as $hid ) 
                                            {
                                                $objPHPExcel->getActiveSheet()->getColumnDimension($col[$hid])->setVisible(false);
                                            }
                                $objPHPExcel->getActiveSheet()->getStyle( $col[$fillhide[$key]['st']]."1")
                                                              ->applyFromArray(array('font' => Style_Font(12, '030c96', true, 'Consolas'))); 


                                $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+1]."1", 'z' ); 
                                $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+2]."1", 'Unhide to view data last week' );

                                $objPHPExcel->getActiveSheet()->getStyle($col[$fillhide[$key]['ed']+1]."1")
                                                              ->applyFromArray(array('font' => Style_Font(21, 'cc0001', true, 'Wingdings 3')));  

                                $objPHPExcel->getActiveSheet()->getStyle($col[$fillhide[$key]['ed']+2]."1")
                                                              ->applyFromArray(array('font' => Style_Font(13, '000000', true, 'Bodoni MT')));                                                              
                                $objPHPExcel->getActiveSheet()
                                            ->getStyle($col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1")
                                            ->getAlignment()
                                            ->setWrapText(false)
                                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);      

                                $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['ed']+2]."1" . ":" . $col[11+$d]."1" );                                        
                                }
                                elseif($key == 1)
                                {                   
//                                    echo $Sunday-12 . " " .$d; exit;


                                                    $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('0F243E') ) );
                                                    $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('B8CCE4') ) );

                                    if($d-($Sunday-12) > 1)
                                    {
                                                        $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 2' );
                                                        $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['st']]."1" );
                                                foreach (range($fillhide[$key]['st'], $fillhide[$key]['ed']) as $hid ) 
                                                {
                                                    $objPHPExcel->getActiveSheet()->getColumnDimension($col[$hid])->setVisible(false);
                                                }
                                    $objPHPExcel->getActiveSheet()->getStyle( $col[$fillhide[$key]['st']]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(12, '030c96', true, 'Consolas'))); 


                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+1]."1", 'z' ); 
                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+2]."1", 'Unhide to view data last week' );

                                    $objPHPExcel->getActiveSheet()->getStyle($col[$fillhide[$key]['ed']+1]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(21, 'cc0001', true, 'Wingdings 3')));  

                                    $objPHPExcel->getActiveSheet()->getStyle($col[$fillhide[$key]['ed']+2]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(13, 'cc0001', true, 'Bodoni MT')));                                                              
                                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1")
                                                ->getAlignment()
                                                ->setWrapText(false)
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);      

                                    $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['ed']+2]."1" . ":" . $col[11+$d]."1" );                                                      

                                    }    

                                }
                                elseif($key == 2)
                                {                                    
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('512603') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('FCD5B4') ) );
                                    if($d-($Sunday-12) > 1)
                                    {
                                                        $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 3' );
                                                        $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['st']]."1" );
                                                foreach (range($fillhide[$key]['st'], $fillhide[$key]['ed']) as $hid ) 
                                                {
                                                    $objPHPExcel->getActiveSheet()->getColumnDimension($col[$hid])->setVisible(false);
                                                }
                                    $objPHPExcel->getActiveSheet()->getStyle( $col[$fillhide[$key]['st']]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(12, '030c96', true, 'Consolas'))); 


                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+1]."1", 'z' ); 
                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+2]."1", 'Unhide to view data last week' );

                                    $objPHPExcel->getActiveSheet()->getStyle($col[$fillhide[$key]['ed']+1]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(21, 'cc0001', true, 'Wingdings 3')));  

                                    $objPHPExcel->getActiveSheet()->getStyle($col[$fillhide[$key]['ed']+2]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(13, 'cc0001', true, 'Bodoni MT')));                                                              
                                    $objPHPExcel->getActiveSheet()
                                                ->getStyle($col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1")
                                                ->getAlignment()
                                                ->setWrapText(false)
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);      

                                    $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['ed']+2]."1" . ":" . $col[11+$d]."1" );                                                      

                                    }

                                }
                                elseif($key == 3)
                                {
                                    
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('4B4B4B') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('D9D9D9') ) );

                                    if($d-($Sunday-12) > 1)
                                    {
                                                        $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 4' );
                                                        $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['st']]."1" );
                                                foreach (range($fillhide[$key]['st'], $fillhide[$key]['ed']) as $hid ) 
                                                {
                                                    $objPHPExcel->getActiveSheet()->getColumnDimension($col[$hid])->setVisible(false);
                                                }
                                        $objPHPExcel->getActiveSheet()->getStyle( $col[$fillhide[$key]['st']]."1")
                                                                      ->applyFromArray(array('font' => Style_Font(12, '030c96', true, 'Consolas'))); 


                                        $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+1]."1", 'z' ); 
                                        $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+2]."1", 'Unhide to view data last week' );

                                        $objPHPExcel->getActiveSheet()->getStyle($col[$fillhide[$key]['ed']+1]."1")
                                                                      ->applyFromArray(array('font' => Style_Font(21, 'cc0001', true, 'Wingdings 3')));  

                                        $objPHPExcel->getActiveSheet()->getStyle($col[$fillhide[$key]['ed']+2]."1")
                                                                      ->applyFromArray(array('font' => Style_Font(13, 'cc0001', true, 'Bodoni MT')));                                                              
                                        $objPHPExcel->getActiveSheet()
                                                    ->getStyle($col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1")
                                                    ->getAlignment()
                                                    ->setWrapText(false)
                                                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                                                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);      

                                        $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['ed']+2]."1" . ":" . $col[11+$d]."1" );                                                      

                                    }   

                                }
                                elseif($key == 4)
                                {
                                                     $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 5' );
                                                     $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['st']]."1" );                                    
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('193300') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('66CC00') ) );
                                        if($d-($Sunday-12) > 1)
                                        {
                                                            $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 5' );
                                                            $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['st']]."1" );
                                                    foreach (range($fillhide[$key]['st'], $fillhide[$key]['ed']) as $hid ) 
                                                    {
                                                        $objPHPExcel->getActiveSheet()->getColumnDimension($col[$hid])->setVisible(false);
                                                    }
                                        $objPHPExcel->getActiveSheet()->getStyle( $col[$fillhide[$key]['st']]."1")
                                                                      ->applyFromArray(array('font' => Style_Font(12, '030c96', true, 'Consolas'))); 


                                        $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+1]."1", 'z' ); 
                                        $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+2]."1", 'Unhide to view data last week' );

                                        $objPHPExcel->getActiveSheet()->getStyle($col[$fillhide[$key]['ed']+1]."1")
                                                                      ->applyFromArray(array('font' => Style_Font(21, 'cc0001', true, 'Wingdings 3')));  

                                        $objPHPExcel->getActiveSheet()->getStyle($col[$fillhide[$key]['ed']+2]."1")
                                                                      ->applyFromArray(array('font' => Style_Font(13, 'cc0001', true, 'Bodoni MT')));                                                              
                                        $objPHPExcel->getActiveSheet()
                                                    ->getStyle($col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1")
                                                    ->getAlignment()
                                                    ->setWrapText(false)
                                                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                                                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);      

                                        $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['ed']+2]."1" . ":" . $col[11+$d]."1" );                                                      

                                        }                                                     
                                }
                                else
                                {
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('FF99CC') ) );
                                }

     
    }

}
 ?>
