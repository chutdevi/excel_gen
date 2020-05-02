<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
ini_set('max_execution_time', 0); 
ini_set('memory_limit','2048M');
date_default_timezone_set('Asia/Bangkok');
if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');

require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

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
foreach ( range('A', 'Z') as $cm ) {
    array_push($col_name, $cm);
}
foreach ( range('A', 'Z') as $cm ) {
    array_push($col_name, "A$cm");
}
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

foreach ($title as $inTil => $til) 
{
        $sheetIndex =  strtolower(str_replace(' ', '_', $title[$inTil]));
        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex($inTil);
        $objPHPExcel->getActiveSheet()->setTitle("$til");
        $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 30 );
        $objPHPExcel->getActiveSheet()
                    ->getStyle('1:4')
                    ->getAlignment()
                    ->setWrapText(true)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $objPHPExcel->getActiveSheet()->freezePane('A6');            
        // $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('5.5');
        // $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('6.5');
        // $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('8');         
        // $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('10');
        // $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('65');
     $i = 0;
     $day = 1;
     if(count($list_act_report[$sheetIndex]) > 0 )
     {
        foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
        { 
            //echo $key . "<hr>";
            $key = str_replace("_REV", ".", $key);
            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."4", str_replace("_", " ", strtoupper($key)));
                 
            if ( substr($key,0,1) == 'D' && strlen($key) < 4) 
            {
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i-1]."4", str_replace("_", " ", date('d-M', strtotime( ($day++) . "-" .$MontCol . "-" . $YearCol ))));
                $dayCh = substr($key,1,strlen($key)-1);
                
                if ($dayCh == $Yeday) 
                {
                    break;
                }
            }
        }
         //foreach(range('A','Z') as $columnID) { $objPHPExcel->getActiveSheet()->getColumnDimension('B'.$columnID)->setAutoSize(true); }         
     }
}
//exit;
//=======================================================================================  Input data ================================================================================
$row = 6;
$indSheet = 0;
foreach ($list_act_report as $key => $value) 
    {
                //echo substr('DATE1',4,2); exit;
                //var_dump($key); exit;
     if(count($list_act_report[$key]) > 0 )
         { 
                   if ($key == 'receiving_weekly') 
                   {
                    $objPHPExcel->setActiveSheetIndex($indSheet);
                    //$objPHPExcel->getActiveSheet()->insertNewRowBefore(1,2);
                    //$objPHPExcel->getActiveSheet()->freezePane('M4');
                    $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 30 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 30 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 27 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 35 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 5 )->setRowHeight( 12 );
                    $objPHPExcel->getActiveSheet()
                                ->getStyle(('A4:'.$col_name[$Onday+10].'4'))
                                ->getAlignment()
                                ->setWrapText(true)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 

                    $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80); 
                     
                    $objPHPExcel->getActiveSheet()->setAutoFilter('A5:'.$col_name[$Onday+10].'5');
                    $objPHPExcel->getActiveSheet()
                                ->getStyle('A6:'.$col_name[6].(count( $list_act_report[$key] )+5))
                                ->getAlignment()
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);   

                    $objPHPExcel->getActiveSheet()->getStyle('A4:'.$col_name[$Onday+10].'4')->applyFromArray(array('fill'    => Style_Fill($colhead)));
                    $objPHPExcel->getActiveSheet()->getStyle('A4:'.$col_name[$Onday+10].'4')->applyFromArray(array('font'    => Style_Font(11, 'FFFFFF', true, 'Roboto')));
                    $objPHPExcel->getActiveSheet()->getStyle('A6:'.$col_name[$Onday+10].(count( $list_act_report[$key] )+5))->applyFromArray(array('font'    => Style_Font(10, '000000', false, 'Consolas')));
                    $objPHPExcel->getActiveSheet()->getStyle('A4:'.$col_name[$Onday+10].'5')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFF99'))));
                    // $objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true)
                    //             ->setName('Consolas')
                    //             ->setSize(11)
                    //             ->getColor()->setRGB('FFFFFF');
                    $objPHPExcel->getActiveSheet()->getStyle('A5:'.$col_name[$Onday+10].'5')->applyFromArray(array('fill'    => Style_Fill('D69389')));
                    $startData = 6;
                    $r = 6;
                            foreach ($value as $nr => $val) 
                            {
                                $indCol = 0;
                                        foreach ($val as $rowData => $data) 
                                        {
                                           
                                            if ($rowData == 'MODEL') 
                                            {
                                                if ($data == '3E00') 
                                                {
                                                   $objPHPExcel->getActiveSheet()->getStyle($col_name[$indCol].($r))->getNumberFormat()->setFormatCode('###"E00"');
                                                   $objPHPExcel->setActiveSheetIndex($indSheet)->setCellValue($col_name[$indCol++].($r), $data);
                                                }
                                                else
                                                {
                                                   $objPHPExcel->setActiveSheetIndex($indSheet)->setCellValue($col_name[$indCol++].($r), $data);
                                                }
                                                    
                                            }
                                            else
                                            {
                                               $objPHPExcel->setActiveSheetIndex($indSheet)->setCellValue($col_name[$indCol++].($r), $data);
                                            }
                                                    if ( substr($rowData,0,1) == 'D' ) 
                                                    {
                                                        if (substr($rowData,1,strlen($key)-1) == $Yeday)  
                                                        {
                                                            break;
                                                        }
                                                        else
                                                        {
                                                         if ( holiday(substr($rowData,1,strlen($key)-1), $holiday) )
                                                                $objPHPExcel->getActiveSheet()
                                                                            ->getStyle($col_name[$indCol-1]. '6:' . $col_name[$indCol-1].(count( $list_act_report[$key] )+5))
                                                                            ->applyFromArray( array( 'fill' => Style_Fill('FFDEAD') ) );                                                           
                                                        }
                                                    } 
                                           // if ($rowData == 'STOCK' && $data < 1) 
                                           //  {
                                           //      $UNREADFontStyle = array( 'font' => Style_Font(11,'FF0000',TRUE, 'Consolas') );          
                                           //      $objPHPExcel->getActiveSheet()->getStyle('L'.$row)->applyFromArray($UNREADFontStyle);
                                           //  }
                                           //  elseif ($rowData == 'ACCUM_DIFF' && $data < 0)
                                           //  {
                                           //      $UNREADFillStyle = array( 'fill' => Style_Fill('FFCCFF') );   
                                           //      $UNREADFontStyle = array( 'font' => Style_Font(12, 'fe0104', true, 'Consolas') );          
                                           //      $objPHPExcel->getActiveSheet()->getStyle($col_name[0].$row.":".$col_name[count($list_act_report[$sheetIndex][0])-1].$row)
                                           //                                    ->applyFromArray($UNREADFillStyle);
                                           //      $objPHPExcel->getActiveSheet()->getStyle($col_name[count($list_act_report[$sheetIndex][0])-5].$row)
                                           //                                    ->applyFromArray($UNREADFontStyle);
                                           //  }
                                        }
                                $r++;
                            }
                            sunday($Today, 2, 3, (count( $list_act_report[$key] )+5), $col_name, $indSheet, $objPHPExcel);      
                            $objPHPExcel->getActiveSheet()->setCellValue('A1', "WEEKLY RM AND PART RECEIVE ".strtoupper(date('Y')));
                            $objPHPExcel->getActiveSheet()->setCellValue('A3', "ACCUMULATE FROM (  01 ".strtoupper($MontCol)." - ".strtoupper($Yeday." ".$MontCol)."  )");
                            $objPHPExcel->getActiveSheet()->setCellValue('I1', "STOCK ON");
                            $objPHPExcel->getActiveSheet()->setCellValue('I2', 'q');
                            $objPHPExcel->getActiveSheet()->setCellValue('I3', $todayA);
                            $objPHPExcel->getActiveSheet()->setCellValue('J1', "ACCUMULATE ".strtoupper($MontFul." ".$YearCol));
                            $objPHPExcel->getActiveSheet()->setCellValue('J2', '=SUBTOTAL(9,J6:J' . (count( $list_act_report[$key] )+5) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('K2', '=SUBTOTAL(9,K6:K' . (count( $list_act_report[$key] )+5) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('L2', '=SUBTOTAL(9,L6:L' . (count( $list_act_report[$key] )+5) . ")" ); 

        foreach (range(12, $Onday+10) as $col) 
            $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col].'3', '=SUBTOTAL(9,'. $col_name[$col] .'6:'. $col_name[$col] . (count( $list_act_report[$key] )+5) . ")" );
#126180
                            $objPHPExcel->getActiveSheet()->getStyle('I2')->applyFromArray(array('font'    => Style_Font(14, 'FFFFFF', true, 'Wingdings 3')));
                            $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font'    => Style_Font(22, 'FFFFFF', true, 'Roboto')));
                            $objPHPExcel->getActiveSheet()->getStyle('A3')->applyFromArray(array('font'    => Style_Font(16, 'FFFFFF', true, 'Roboto')));
                            $objPHPExcel->getActiveSheet()->getStyle('I1')->applyFromArray(array('font'    => Style_Font(16, 'FFFFFF', true, 'Consolas')));
                            $objPHPExcel->getActiveSheet()->getStyle('I3')->applyFromArray(array('font'    => Style_Font(14, 'FFFFFF', true, 'Consolas')));
                            $objPHPExcel->getActiveSheet()->getStyle('J1')->applyFromArray(array('font'    => Style_Font(14, '001933', true, 'Consolas')));
                            $objPHPExcel->getActiveSheet()->getStyle('J2:L2')->applyFromArray(array('font' => Style_Font(12, 'FFFFFF', true, 'Consolas')));
                            //$objPHPExcel->getActiveSheet()->getStyle('M1:'.$col_name[$Onday+10].'1')->applyFromArray(array('font' => Style_Font(13, 'FFFFFF', true, 'Consolas')));
                            $objPHPExcel->getActiveSheet()->getStyle('M2:'.$col_name[$Onday+10].'2')->applyFromArray(array('font' => Style_Font(13, 'FFFFFF', true, 'Consolas')));
                            $objPHPExcel->getActiveSheet()->getStyle('M3:'.$col_name[$Onday+10].'3')->applyFromArray(array('font' => Style_Font(12, '001933', true, 'Consolas')));

                            $objPHPExcel->getActiveSheet()->getStyle('A1:A3')->applyFromArray(array('fill'    => Style_Fill('126180')));
                            $objPHPExcel->getActiveSheet()->getStyle('I1:I3')->applyFromArray(array('fill'    => Style_Fill('00332B')));
                            $objPHPExcel->getActiveSheet()->getStyle('J1:L1')->applyFromArray(array('fill'    => Style_Fill('FDE3A1')));
                            $objPHPExcel->getActiveSheet()->getStyle('J2:L2')->applyFromArray(array('fill'    => Style_Fill('A9A9A9')));

                            $objPHPExcel->getActiveSheet()->getStyle('A1:'.'H3')->applyFromArray(array(
                                                                     'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFF99'))));
                            $objPHPExcel->getActiveSheet()->getStyle('J1:'.'L3')->applyFromArray(array(
                                                                     'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFF99'))));
                            $objPHPExcel->getActiveSheet()->getStyle('M1:'.$col_name[$Onday+10].'3')->applyFromArray(array(
                                                                     'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFF99'))));                            
                    //         $objPHPExcel->getActiveSheet()->getStyle('K1:K2')->applyFromArray(array('font' => Style_Font(16,'FDE9D9',true)));
                    //         $objPHPExcel->getActiveSheet()->getStyle('L1')->applyFromArray(array('fill'    => Style_Fill('202020')));

                    //         //sunday($Today, 1, 4, (count( $list_act_report[$key] )+3), $col_name, $indSheet, $objPHPExcel);  #Function sunday -----------------------------------------------

                    //         $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[$Onday+10].'2')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));

                    //         $objPHPExcel->getActiveSheet()
                    //                     ->getStyle('A1:'.$col_name[$Onday+10].'2')
                    //                     ->getAlignment()
                    //                     ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
                    //                     ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);                            
                           $objPHPExcel->getActiveSheet()
                                ->getStyle('J1:L3')
                                ->getAlignment()
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 


                           $objPHPExcel->getActiveSheet()
                                ->getStyle('M2:'.$col_name[$Onday+10].'3')
                                ->getAlignment()
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 



                            $objPHPExcel->getActiveSheet()->mergeCells('A1:H2');
                            $objPHPExcel->getActiveSheet()->mergeCells('A3:H3');

                            $objPHPExcel->getActiveSheet()->mergeCells('J1:L1');
                            $objPHPExcel->getActiveSheet()->mergeCells('J2:J3');
                            $objPHPExcel->getActiveSheet()->mergeCells('K2:K3');
                            $objPHPExcel->getActiveSheet()->mergeCells('L2:L3');

                            $objPHPExcel->getActiveSheet()->getStyle('I6:'.$col_name[$Onday+10].( count( $list_act_report[$key] )+5) )
                                                          ->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');

                            $objPHPExcel->getActiveSheet()->getStyle('J2:L3')
                                                          ->getNumberFormat()->setFormatCode('_* #,##0_-;[RED](#,##0)_-;_* [BLACK]"-"??_-;_-@_-');

                            $objPHPExcel->getActiveSheet()->getStyle('M2:'.$col_name[$Onday+10].'3')
                                                          ->getNumberFormat()->setFormatCode('_* #,##0_-;[RED](#,##0)_-;_*"-"??_-;_-@_-');                                                          
                    //         $objPHPExcel->getActiveSheet()->getStyle('L2:'.$col_name[$Onday+10].(count( $list_act_report[$key] )+5) )
                    //                                       ->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');

                    //         $objPHPExcel->getActiveSheet()->getStyle('M4:'.$col_name[$Onday+10].(count( $list_act_report[$key] )+5))->applyFromArray(array('font' => Style_Font(10,'FF0000',true)));

                    //         $objPHPExcel->getActiveSheet()->getStyle('L1:'.$col_name[$Onday+10].'1')->applyFromArray(array('font' => Style_Font(16,'FFFFFF',true)));
                    //         $objPHPExcel->getActiveSheet()->getStyle('L1:'.$col_name[$Onday+10].'1')->getFont()->setUnderline(true);

                    //         $objPHPExcel->getActiveSheet()->getStyle('L2:'.'L'.$col_name[$Onday+10].'2')->applyFromArray(array('font' => Style_Font(11,'FF0000',true)));


                    //         $objPHPExcel->getActiveSheet()->getStyle('L4:'.'L'.(count( $list_act_report[$key] )+5))->applyFromArray(array('font' => Style_Font(11,'000066',true)));

                    //         foreach (range('B', 'E') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('13');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('20');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('20');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('40');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('22');
                            foreach (range('I', 'L') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('17');
                            foreach (range('M', 'Z') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('15');
                            foreach (range('A', 'Z') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension('A'.$key)->setWidth('15');

  
                    // $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setVisible(false);
                    // $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setVisible(false);
                    // $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setVisible(false);
                    // $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(28,'000000',true)));



                    } 
             } 
$indSheet++;
    //echo $indSheet; exit;
        
$row = 6;
    
} 













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
            }
                
            elseif( $indexSun + 6 > $d+10 && $valDay < $d-1)
                {
                        $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun+1].$focus, '=SUM(' . $col[$indexSun+1] . $start_row . ":" . $col[$d+11] . $start_row . ")" );
                        $objPHPExcel->getActiveSheet()->mergeCells($col[$indexSun+1] . $focus .":" . $col[$d+11] . $focus);

                        array_push($fillhide, array('st' => $indexSun+1, 'ed' => $d+11) );
                        array_push($fillSum, $col[$indexSun+1] . $focus);
                        array_push($fillTotal, $col[$indexSun+1] . $start_row . ":" . $col[$d+11] . $start_row );
                        //break;
                }

               // echo $Sunday; exit;

           
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
//var_dump($fillSum); exit;
    foreach ($fillSum as $key => $value) 
    {
                                if ($key == 0)
                                {
                                                    //echo $col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1"; exit;  
                                                    

                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('4F6228') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('D8E4BC') ) );//D8E4BC
                                    if($d-($Sunday-12) > 1 || $d > 7)
                                    {

                                                $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 1' );
                                                $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['ed']]."1" );

                                                foreach (range($fillhide[$key]['st'], $fillhide[$key]['ed']) as $hid ) 
                                                {
                                                    $objPHPExcel->getActiveSheet()->getColumnDimension($col[$hid])->setVisible(false);
                                                }

                                    $objPHPExcel->getActiveSheet()->getStyle( $col[$fillhide[$key]['st']]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(12, '030c96', true, 'Consolas'))); 


                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+1]."1", 'z' ); 
                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+2]."1", '<< Unhide to view data last week' );

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

                                    $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['ed']+2]."1" . ":" . $col[$fillhide[$key]['ed']+7]."1" );
                                    }                                        
                                }
                                elseif($key == 1)
                                {                   
                                    //echo $Sunday-12 . " " .$d; exit;

                                    
                                                    $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('0F243E') ) );
                                                    $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('B8CCE4') ) );

                                    if($d-($Sunday-12) > 1 || $d > 14)
                                    {
                                     //echo $col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1"; exit; 
                                      $objPHPExcel->getActiveSheet()->unmergeCells($col[$fillhide[$key]['st']+1]."1" . ":" . $col[$fillhide[$key]['ed']]."1" );
                                      $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 2' );
                                      $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['ed']]."1" );
                                    $objPHPExcel->getActiveSheet()
                                                    ->getStyle($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['ed']]."1" )
                                                    ->getAlignment()
                                                    ->setWrapText(false)
                                                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);                                        
                                                foreach (range($fillhide[$key]['st'], $fillhide[$key]['ed']) as $hid ) 
                                                {
                                                    $objPHPExcel->getActiveSheet()->getColumnDimension($col[$hid])->setVisible(false);
                                                }
                                    $objPHPExcel->getActiveSheet()->getStyle( $col[$fillhide[$key]['st']]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(12, '030c96', true, 'Consolas'))); 


                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+1]."1", 'z' ); 
                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+2]."1", '<< Unhide to view data last week' );

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

                                    $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['ed']+2]."1" . ":" .$col[$fillhide[$key]['ed']+7]."1" );                                                      

                                    }    

                                }
                                elseif($key == 2)
                                {                                    
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('512603') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('FCD5B4') ) );
                                    if($d-($Sunday-12) > 1 || $d > 21)
                                    {
                                     //echo $col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1"; exit; 
                                      $objPHPExcel->getActiveSheet()->unmergeCells($col[$fillhide[$key]['st']+1]."1" . ":" . $col[$fillhide[$key]['ed']]."1" );
                                      $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 3' );
                                      $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['ed']]."1" );
                                    $objPHPExcel->getActiveSheet()
                                                    ->getStyle($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['ed']]."1" )
                                                    ->getAlignment()
                                                    ->setWrapText(false)
                                                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);                                        
                                                foreach (range($fillhide[$key]['st'], $fillhide[$key]['ed']) as $hid ) 
                                                {
                                                    $objPHPExcel->getActiveSheet()->getColumnDimension($col[$hid])->setVisible(false);
                                                }
                                    $objPHPExcel->getActiveSheet()->getStyle( $col[$fillhide[$key]['st']]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(12, '030c96', true, 'Consolas'))); 


                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+1]."1", 'z' ); 
                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+2]."1", '<< Unhide to view data last week' );

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

                                    $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['ed']+2]."1" . ":" .$col[$fillhide[$key]['ed']+7]."1" );                                                      

                                    }

                                }
                                elseif($key == 3)
                                {
                                    
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('4B4B4B') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('D9D9D9') ) );

                                    if($d-($Sunday-12) > 1 || $d > 28)
                                    {
                                     //echo $col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1"; exit; 
                                      $objPHPExcel->getActiveSheet()->unmergeCells($col[$fillhide[$key]['st']+1]."1" . ":" . $col[$fillhide[$key]['ed']]."1" );
                                      $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 4' );
                                      $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['ed']]."1" );
                                    $objPHPExcel->getActiveSheet()
                                                    ->getStyle($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['ed']]."1" )
                                                    ->getAlignment()
                                                    ->setWrapText(false)
                                                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);                                        
                                                foreach (range($fillhide[$key]['st'], $fillhide[$key]['ed']) as $hid ) 
                                                {
                                                    $objPHPExcel->getActiveSheet()->getColumnDimension($col[$hid])->setVisible(false);
                                                }
                                    $objPHPExcel->getActiveSheet()->getStyle( $col[$fillhide[$key]['st']]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(12, '030c96', true, 'Consolas'))); 


                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+1]."1", 'z' ); 
                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+2]."1", '<< Unhide to view data last week' );

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

                                    $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['ed']+2]."1" . ":" .$col[$fillhide[$key]['ed']+7]."1" );                                                      

                                    }

                                }
                                elseif($key == 4)
                                {                                   
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('193300') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('66CC00') ) );
                                    if($d-($Sunday-12) > 1 )
                                    {
                                     //echo $col[$fillhide[$key]['ed']+1]."1".':'.$col[$fillhide[$key]['ed']+2]."1"; exit; 
                                      $objPHPExcel->getActiveSheet()->unmergeCells($col[$fillhide[$key]['st']+1]."1" . ":" . $col[$fillhide[$key]['ed']]."1" );
                                      $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['st']]."1", 'Act. Week 5' );
                                      $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['ed']]."1" );
                                    $objPHPExcel->getActiveSheet()
                                                    ->getStyle($col[$fillhide[$key]['st']]."1" . ":" . $col[$fillhide[$key]['ed']]."1" )
                                                    ->getAlignment()
                                                    ->setWrapText(false)
                                                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);                                        
                                                foreach (range($fillhide[$key]['st'], $fillhide[$key]['ed']) as $hid ) 
                                                {
                                                    $objPHPExcel->getActiveSheet()->getColumnDimension($col[$hid])->setVisible(false);
                                                }
                                    $objPHPExcel->getActiveSheet()->getStyle( $col[$fillhide[$key]['st']]."1")
                                                                  ->applyFromArray(array('font' => Style_Font(12, '030c96', true, 'Consolas'))); 


                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+1]."1", 'z' ); 
                                    $objPHPExcel->getActiveSheet()->setCellValue( $col[$fillhide[$key]['ed']+2]."1", '<< Unhide to view data last week' );

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

                                    $objPHPExcel->getActiveSheet()->mergeCells($col[$fillhide[$key]['ed']+2]."1" . ":" .$col[$fillhide[$key]['ed']+7]."1" );                                                      
                                                    
                                }
                                else
                                {
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('FF99CC') ) );
                                }

     
    }
}

}
 ?>
