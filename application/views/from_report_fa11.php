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
            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."6", str_replace("_", " ", $key));
                 

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
                    $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 30 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 10 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 35 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 32 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 5 )->setRowHeight( 32 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 6 )->setRowHeight( 35 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 7 )->setRowHeight( 3 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 8 )->setRowHeight( 12 );                  
                    $objPHPExcel->getActiveSheet()
                                ->getStyle(('D'.$st_cal.':'.$col_name[41].$st_cal))
                                ->getAlignment()
                                ->setWrapText(true)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 


                    $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80); 
                     
                    $objPHPExcel->getActiveSheet()->setAutoFilter('D8:'.$col_name[10].'8');
                    $objPHPExcel->getActiveSheet()
                                ->getStyle('D9:'.$col_name[41].(count( $list_act_report[$sheetIndex] )+8))
                                ->getAlignment()
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)   //*** set left
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);   //*** set center

                    $objPHPExcel->getActiveSheet()->getStyle('D6:'.$col_name[41].'6')->applyFromArray(array('fill'    => Style_Fill($colhead)));
                    $objPHPExcel->getActiveSheet()->getStyle('D6:'.$col_name[41].'6')->applyFromArray(array('font'    => Style_Font(12, '000000', true, 'Arial Narrow')));
                    $objPHPExcel->getActiveSheet()->getStyle('D9:'.$col_name[41].(count( $list_act_report[$sheetIndex] )+8))->applyFromArray(array('font'    => Style_Font(12, '000000', false, 'Calibri')));
                
                           
                    // $objPHPExcel->getActiveSheet()->getStyle('D6:'.$col_name[41].'6')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                    // $objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true)
                    //              ->setName('Consolas')
                    //              ->setSize(11)
                    //             ->getColor()->setRGB('FFFFFF');
                   
                    $objPHPExcel->getActiveSheet()->getStyle('D8:'.$col_name[41].'8')->applyFromArray(array('fill'    => Style_Fill('f0f5f5')));
                    
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

                                            if ($rowData == 'SUP_FROM' && $data == '') 
                                            {

                                            $objPHPExcel->getActiveSheet()->getStyle('D'.($r).':'.$col_name[41].($r))->applyFromArray(array('fill'    => Style_Fill('d9d9d9'))); //color sheet
                                            $objPHPExcel->getActiveSheet()->getRowDimension( $r )->setRowHeight( 20 );
                                            $objPHPExcel->getActiveSheet()->getStyle('D'.($r).':'.$col_name[41].($r))->applyFromArray(array('font'    => Style_Font(12, '000000', true, 'Calibri')));
                                            }
                                         }   

                                        } 
                                $r++;
                            }
        $Montlast = date('F Y', strtotime(date('Y')."-".(date('m')-1)."-".'1'));  
        $M = date('M', strtotime(date('Y')."-".(date('m')-1)."-".'1'));
        $Daylast  = substr(date('Y-m-t',strtotime(date('Y')."-".(date('m')-1)."-".'1')),8, 2);
        //echo $Montlast.$Daylast;exit;
                            $objPHPExcel->getActiveSheet()->setCellValue('D3',"PRODUCTION PLAN REMAIN REPORT "); //('D2', "FA Supply list of ".$Montlast);
                            $objPHPExcel->getActiveSheet()
                                ->getStyle(('D3'.':'.$col_name[11].'3'))
                                ->getAlignment()
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 
                            $objPHPExcel->getActiveSheet()->setCellValue('D4', "Difference on : ".strtoupper(date('d-M-Y',  strtotime((date('d')-1) . "-" . date('M') . "-" . date('Y')) )));
                            $objPHPExcel->getActiveSheet()
                                ->getStyle(('D4'.':'.$col_name[11].'4'))
                                ->getAlignment()
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // example title 
                            $objPHPExcel->getActiveSheet()->setCellValue('D5', "");
                            $objPHPExcel->getActiveSheet()
                                ->getStyle(('D5'.':'.$col_name[41].'5'))
                                ->getAlignment()
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // example title 
                      
                            

        foreach (range(12, 10) as $col) 
                      
                            $objPHPExcel->getActiveSheet()->getStyle('D3')->applyFromArray(array('font'    => Style_Font(20, '000000', true, 'Arial Narrow')));
                            $objPHPExcel->getActiveSheet()->getStyle('D4')->applyFromArray(array('font'    => Style_Font(16, '000000', true, 'Arial Narrow')));
                            $objPHPExcel->getActiveSheet()->getStyle('M3')->applyFromArray(array('font'    => Style_Font(16, '002d4d', true, 'Arial Narrow')));
                            $objPHPExcel->getActiveSheet()->getStyle('K5:AQ5')->applyFromArray(array('font'    => Style_Font(12, '000000', true, 'Arial Narrow')));
                            $objPHPExcel->getActiveSheet()->getStyle('D3:J3')->applyFromArray(array('fill'    => Style_Fill('fdfdd9')));
                            $objPHPExcel->getActiveSheet()->getStyle('D4:J4')->applyFromArray(array('fill'    => Style_Fill('fdfdd9')));
                            $objPHPExcel->getActiveSheet()->getStyle('D5:AP5')->applyFromArray(array('fill'    => Style_Fill('dce6f1')));
                             $objPHPExcel->getActiveSheet()->setCellValue('K5', "TOTAL [PCS.]");

                            
                           
                            $objPHPExcel->getActiveSheet()->getStyle('D6:'.'AP'.($st_cal))->applyFromArray(array(
                                                                     'borders' => array('top' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                            $objPHPExcel->getActiveSheet()->getStyle('D6:'.'AP'.($st_cal))->applyFromArray(array(
                                                                     'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                           
                            $objPHPExcel->getActiveSheet()->getStyle('D5:'.'AP'.($st_cal))->applyFromArray(array(
                                                                     'borders' => array('top' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                            $objPHPExcel->getActiveSheet()->getStyle('D5:'.'AP'.($st_cal))->applyFromArray(array(
                                                                     'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                            $objPHPExcel->getActiveSheet()->getStyle('B2:'.'AQ'.($cu_dat+10))->applyFromArray(array(
                                                                     'borders' => array('outline' => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));
                            $objPHPExcel->getActiveSheet()->getStyle('D9:'.'AP'.($cu_dat+8))->applyFromArray(array(
                                                                     'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'b3d9ff'))));





        foreach (range(11, 40) as $col) 

            $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col].'5', '=SUBTOTAL(9,'. $col_name[$col] .'9:'. $col_name[$col] . ($cu_dat+9). ")" );

                         
                     

                          $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setVisible(false); //hide

                          $objPHPExcel->getActiveSheet()->mergeCells('D3:J3');
                          $objPHPExcel->getActiveSheet()->mergeCells('D4:J4');
                        //  $objPHPExcel->getActiveSheet()->mergeCells('M3:N4');

                          $objPHPExcel->getActiveSheet()->getStyle('L5'.':AQ'.( count( $list_act_report[$sheetIndex] )+8) ) 
                             ->getNumberFormat()->setFormatCode('_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)');
//------------------------------------------------------------------------------------------------------------ hide date -------------------------------------------------------------------------------//

                            foreach (range(((date('d')+1)+9),41) as $key) 
                            $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$key])->setVisible(false);


             
                            $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('1');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('1');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('3');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('9');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('12');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('9');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('15');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('25');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('29');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth('19');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth('9');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth('9');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth('9');





#-------------------------------------------------------------------------------------------------------------------------------------------


                    } 
              

 // echo $indSheet; exit;
        
$row = 5;
 } 
 $indSheet++;  
} 

//  else 
// {

//             $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "No Production Plan".$til.".");
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
