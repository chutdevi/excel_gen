<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
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
// if(strlen($dayA) < 2) $dayA = "0".$dayA;
// $yesterda dayA =te('Y-M-d', strtotime($yearA."-".$monthYes."-".$dayA));
// $dayA = date('d')-2;
// if(strlen($dayA) < 2)  $dayA = "0".$dayA;
// if(date('d') == "01")  $dayA = intval(substr(date('Y-m-t',strtotime($yearA."-".$monthYes."-".$dayA)),8, 2))-1;
// $yesterdayB = date('Y-M-d', strtotime($yearA."-".$monthYes."-".$dayA));
// $dayA = date('d')-3;
// if(strlen($dayA) < 2) $dayA = "0".$dayA;
// if(date('d') == "01")  $dayA = intval(substr(date('Y-m-t',strtotime($yearA."-".$monthYes."-".$dayA)),8, 2))-2;
// $yesterdayC = date('Y-M-d', strtotime($yearA."-".$monthYes."-".$dayA));
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

foreach ($title as $inTil => $til) 
{
        $sheetIndex =  strtolower(str_replace(' ', '_', $title[$inTil]));
        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex($inTil);
        $objPHPExcel->getActiveSheet()->setTitle("$til");



           
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
            if ($til == "FA Summary") //sheet1
            {
                $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 30 );
                $objPHPExcel->getActiveSheet()
                            ->getStyle('1:3') 
                            ->getAlignment()
                            ->setWrapText(true)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);                
                $key = str_replace("_REV", ".", $key);
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."3", str_replace("_", " ", strtoupper($key)));
                $objPHPExcel->getActiveSheet()->freezePane('M5');     

            }
            elseif ($til == "Fa Summary loss")  //sheet2
            {
                $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 30 );
                $objPHPExcel->getActiveSheet()
                            ->getStyle('1:4') 
                            ->getAlignment()
                            ->setWrapText(true)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);                
                $key = str_replace("_REV", ".", $key);
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."1", str_replace("_", " ", strtoupper($key)));        
            }
            elseif ($til == "LOSS CODE") //sheet3
            {
                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 30 );
                $objPHPExcel->getActiveSheet()
                            ->getStyle('1') 
                            ->getAlignment()
                            ->setWrapText(true)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);                
                $key = str_replace("_REV", ".", $key);
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."1", str_replace("_", " ", strtoupper($key)));        
            }
        }
         //foreach(range('A','Z') as $columnID) { $objPHPExcel->getActiveSheet()->getColumnDimension('B'.$columnID)->setAutoSize(true); }         
     }
}
//exit;
//=======================================================================================  Input data ================================================================================
$row = 5;
//var_dump($list_act_report);
$indSheet = 0;
//exit;
foreach ($list_act_report as $key => $value) 
    {
                //echo substr('DATE1',4,2); exit;
                //var_dump($value); exit;
     if(count($list_act_report[$key]) > 0 )
         { 
            if ($title[$indSheet] == 'FA Summary') {
                    $objPHPExcel->setActiveSheetIndex($indSheet);
                    //$objPHPExcel->getActiveSheet()->insertNewRowBefore(1,2);
                    //$objPHPExcel->getActiveSheet()->freezePane('M4');		
                    					//Rowsize
                    $sheetIndex = 'fa_summary';
                    $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 30 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 30 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 27 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 27 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 5 )->setRowHeight( 15 );
                    $objPHPExcel->getActiveSheet()
                                ->getStyle(('A3:'.$col_name[45].'3'))
                                ->getAlignment()
                                ->setWrapText(false)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 
                    $objPHPExcel->getActiveSheet()
                                ->getStyle(('A4:'.$col_name[45].'4'))
                                ->getAlignment()
                                ->setWrapText(true)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                                //var_dump($col_name[58]); exit;    
                     $objPHPExcel->getActiveSheet()
                                ->getStyle(('O2:'.$col_name[45].'2'))
                                ->getAlignment()
                                ->setWrapText(false)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                                  

                    //$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80); 
                     
                    
                    $objPHPExcel->getActiveSheet()
                                ->getStyle('A5:'.$col_name[5].(count( $list_act_report[$key] )+5))
                                ->getAlignment()
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);   

                    $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80); 

                    $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[45].'3')->applyFromArray(array('fill'    => Style_Fill($colhead)));
                    $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[45].'3')->applyFromArray(array('font'    => Style_Font(11, '000000', true, 'Calibri')));
                    $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[45].'4')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                    $objPHPExcel->getActiveSheet()->getStyle('A5:'.$col_name[45].(count( $list_act_report[$key] )+5))->applyFromArray(array('font'    => Style_Font(8, '000000', false, 'Calibri')));
                    // $objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true)
                    //             ->setName('Consolas')
                    //             ->setSize(11)
                    //             ->getColor()->setRGB('FFFFFF');
                    $objPHPExcel->getActiveSheet()->getStyle('A4:'.$col_name[45].'4')->applyFromArray(array('fill'    => Style_Fill('e6f2ff')));
                    $startData = 5;
                 // echo $usetime;exit;
                    $r = 5;
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

                                        }

                                $r++;
                            }




        $Montlast = date('F Y', strtotime(date('Y')."-".(date('m')-1)."-".'1'));  
        $M = date('M', strtotime(date('Y')."-".(date('m')-1)."-".'1'));
        $Daylast  = substr(date('Y-m-t',strtotime(date('Y')."-".(date('m')-1)."-".'1')),8, 2);
        //echo $Montlast.$Daylast;exit;
                            $objPHPExcel->getActiveSheet()->setCellValue('A1', "Summary FA Report of ".$Montlast);
                            $objPHPExcel->getActiveSheet()->setCellValue('A2', "ACCUMULATE FROM (  01 ".strtoupper($M)." - ".$Daylast." ".strtoupper($M)."  )");
                            $objPHPExcel->getActiveSheet()->setCellValue('G1', "Aver Man/Month [man.]");
                            $objPHPExcel->getActiveSheet()->setCellValue('G2', '=SUBTOTAL(9,G5:G'. (count( $list_act_report[$key] )+5) .')/'.$usetime ); //.')/'.$usetime
                            $objPHPExcel->getActiveSheet()->setCellValue('H1', "Accum \n Production [pcs.] ");
                            $objPHPExcel->getActiveSheet()->setCellValue('H2', '=SUBTOTAL(9,H5:H'. (count( $list_act_report[$key] )+5) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('I2', '=SUBTOTAL(9,I5:I'. (count( $list_act_report[$key] )+5) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('J2', '=SUBTOTAL(9,J5:J'. (count( $list_act_report[$key] )+5) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('K1', "Accum Time [min.] ");
                            $objPHPExcel->getActiveSheet()->setCellValue('K2', '=SUBTOTAL(9,K5:K'. (count( $list_act_report[$key] )+5) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('L2', '=SUBTOTAL(9,L5:L'. (count( $list_act_report[$key] )+5) . ")" );
                            $objPHPExcel->getActiveSheet()->setCellValue('M1', "Accum Loss Time (Min)");


        foreach (range(12, 43) as $col) 
            $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col].'2', '=SUBTOTAL(9,'. $col_name[$col] .'5:'. $col_name[$col] . (count( $list_act_report[$key] )+5) . ")" );
#126180
    				 		$objPHPExcel->getActiveSheet()->getStyle('G5:'.$col_name[45].( count( $list_act_report[$key] )+5) )
                                                          ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');

                            $objPHPExcel->getActiveSheet()->getStyle('G2:'.$col_name[45].'2')
                                                          ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');                        

//============================================================================DETAIL LOSS ==========================================================================================================

                              
                              $objPHPExcel->getActiveSheet()->setCellValue('M4', "Meeting"); 
					          $objPHPExcel->getActiveSheet()->setCellValue('N4', "Cleaning, 5S");
					          $objPHPExcel->getActiveSheet()->setCellValue('O4', "Planned Maintenance (PM) - in Plan");
					          $objPHPExcel->getActiveSheet()->setCellValue('P4', "Planned Maintenance (PM) - out Plan");
					          $objPHPExcel->getActiveSheet()->setCellValue('Q4', "Training Operator");
					          $objPHPExcel->getActiveSheet()->setCellValue('R4', "Regular Machine Daily Check");
					          $objPHPExcel->getActiveSheet()->setCellValue('S4', "Stop adding Flux.");
					          $objPHPExcel->getActiveSheet()->setCellValue('T4', "Stop Heat up.");
					          $objPHPExcel->getActiveSheet()->setCellValue('U4', "Mold Clearning ");
					          $objPHPExcel->getActiveSheet()->setCellValue('V4', "Model Change, Mold Change - in Plan");
					          $objPHPExcel->getActiveSheet()->setCellValue('W4', "Model Change, Mold Change - out Plan");
					          $objPHPExcel->getActiveSheet()->setCellValue('X4', "Tooling Change / Lot no. change/Tip/Sleeve - on Plan");
					          $objPHPExcel->getActiveSheet()->setCellValue('Y4', "Waiting for material from inhouse");
					          $objPHPExcel->getActiveSheet()->setCellValue('Z4', "Coolant, Lubrication, Metal Chip Cleaning");
					          $objPHPExcel->getActiveSheet()->setCellValue('AA4', "Break Down - Machine ");
					          $objPHPExcel->getActiveSheet()->setCellValue('AB4', "Break Down - Mold (Die Casting)");
					          $objPHPExcel->getActiveSheet()->setCellValue('AC4', "Break Down - JIG /Tip/Sleeve");
					          $objPHPExcel->getActiveSheet()->setCellValue('AD4', "Break Down - Tooling");
					          $objPHPExcel->getActiveSheet()->setCellValue('AE4', "Adjust Program and Condition.");
					          $objPHPExcel->getActiveSheet()->setCellValue('AF4', "Adjustment - Mold (Die Casting)");
					          $objPHPExcel->getActiveSheet()->setCellValue('AG4', "Electricity, Air, Utility Break Down");
					          $objPHPExcel->getActiveSheet()->setCellValue('AH4', "Waiting for material from store");
					          $objPHPExcel->getActiveSheet()->setCellValue('AI4', "Material / Part Quality Problem");
					          $objPHPExcel->getActiveSheet()->setCellValue('AJ4', "Personel Reason (toilet, sick, late other)");
					          $objPHPExcel->getActiveSheet()->setCellValue('AK4', "Quality Judgement");
					          $objPHPExcel->getActiveSheet()->setCellValue('AL4', "Waiting data from QC / 4M");
					          $objPHPExcel->getActiveSheet()->setCellValue('AM4', "Waiting for Packaging/Box/Partition");
					          $objPHPExcel->getActiveSheet()->setCellValue('AN4', "PE/CE New Model Trial request (CE, PE)/Audit");
					          $objPHPExcel->getActiveSheet()->setCellValue('AO4', "TPM, Line Kaizen");
					          $objPHPExcel->getActiveSheet()->setCellValue('AP4', "Waiting for supervisor");
					          $objPHPExcel->getActiveSheet()->setCellValue('AQ4', "Production of work other than direct production.");
					          $objPHPExcel->getActiveSheet()->setCellValue('AR4', "Re - washing FG part");






                    //         $objPHPExcel->getActiveSheet()->getStyle('L2:'.$col_name[$Onday+10].(count( $list_act_report[$key] )+5) )
                    //                                       ->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');

                    //         $objPHPExcel->getActiveSheet()->getStyle('M4:'.$col_name[$Onday+10].(count( $list_act_report[$key] )+5))->applyFromArray(array('font' => Style_Font(10,'FF0000',true)));

                    //         $objPHPExcel->getActiveSheet()->getStyle('L1:'.$col_name[$Onday+10].'1')->applyFromArray(array('font' => Style_Font(16,'FFFFFF',true)));
                    //         $objPHPExcel->getActiveSheet()->getStyle('L1:'.$col_name[$Onday+10].'1')->getFont()->setUnderline(true);

                    //         $objPHPExcel->getActiveSheet()->getStyle('L2:'.'L'.$col_name[$Onday+10].'2')->applyFromArray(array('font' => Style_Font(11,'FF0000',true)));


                    //         $objPHPExcel->getActiveSheet()->getStyle('L4:'.'L'.(count( $list_act_report[$key] )+5))->applyFromArray(array('font' => Style_Font(11,'000066',true)));

                    //         foreach (range('B', 'E') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('4');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('9');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('17');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('30');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('11');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('9');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('9');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('9');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('5');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth('8');
                           
                            foreach (range('H', 'L') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('17');
                            foreach (range('M', 'Z') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('12');
                            foreach (range('A', 'Z') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension('A'.$key)->setWidth('10');
                            foreach (range('A', 'Z') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension('B'.$key)->setWidth('10');


  
                    // $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setVisible(false);
                    // $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setVisible(false);
                    // $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setVisible(false);
                    // $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(28,'000000',true)));
#---------------------------------------------------------------------------------------------------------------------------------------------


          $objPHPExcel->getActiveSheet()->insertNewColumnBefore('M', 1);

          $objPHPExcel->getActiveSheet()->setCellValue('M1', "Ave \n ManHrPcs [min]");
          $objPHPExcel->getActiveSheet()->setCellValue('M3', ' Man Hr Pcs (min)');


          for($i=5 ; $i < (count( $list_act_report[$sheetIndex] )+5) ; $i++)
          {
            $minusTime = '((K'.$i.'+'.'L'.$i.')*G'.$i.')/I'.$i;
            $objPHPExcel->getActiveSheet()->setCellValue('M'.$i, '='.$minusTime);
          }
         	$SumAct = '(SUBTOTAL(9,I5:I'. (count( $list_act_report[$sheetIndex] )+5) . '))'; //SumActual
          	$SumStaff = '(SUBTOTAL(9,G5:G'. (count( $list_act_report[$sheetIndex] )+5) . '))'; //SumMan
           	$SumUse = '(SUBTOTAL(9,K5:K'. (count( $list_act_report[$sheetIndex] )+5) . '))';  //SumUse
       	  	$Sumloss = '(SUBTOTAL(9,L5:L'. (count( $list_act_report[$sheetIndex] )+5) . '))'; //Sumloss
       	 	

          $objPHPExcel->getActiveSheet()->setCellValue('M2', '=(('.$SumUse.'+'.$Sumloss.')*'.$SumStaff.')/'.$SumAct);
          $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth('19');
          $objPHPExcel->getActiveSheet()->getStyle('M2')->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-'); 
          $objPHPExcel->getActiveSheet()->getStyle('M5:M'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');


#---------------------------------------------------------------------------------------------------------------------------------------------
//var_dump($list_act_report[$sheetIndex]);exit;

          $objPHPExcel->getActiveSheet()->insertNewColumnBefore('M', 1);
          $objPHPExcel->getActiveSheet()->setCellValue('M1', "Ave EFF.(%)");
          $objPHPExcel->getActiveSheet()->setCellValue('M3', 'EFF.(%)');

          for($i=5 ; $i < (count( $list_act_report[$sheetIndex] )+5) ; $i++)
          {
            $minusTime = '(K'.$i.'-'.'L'.$i.')';
            $objPHPExcel->getActiveSheet()->setCellValue('M'.$i, '=IF(K'.$i.'<1,0,'.$minusTime.'/'.'K'.$i.')');
           // echo('=IF(K'.$i.'<1,0,'.$minusTime.'/'.'K'.$i.')').(count( $list_act_report[$sheetIndex] )+5)."<hr>";
          }
         // exit;

          $SumTime = '(SUBTOTAL(9,K5:K'. (count( $list_act_report[$sheetIndex] )+5) . '))';
          $Sumloss = '(SUBTOTAL(9,L5:L'. (count( $list_act_report[$sheetIndex] )+5) . '))';
          $objPHPExcel->getActiveSheet()->setCellValue('M2', '=('.$SumTime.'-'.$Sumloss.')/'.$SumTime);
          $objPHPExcel->getActiveSheet()->getStyle('M2:M'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-'); 
          $objPHPExcel->getActiveSheet()->setAutoFilter('A4:'.$col_name[13].'4');

 //insert Row !!!!!

 		  $objPHPExcel->getActiveSheet()->insertNewRowBefore(2, 1);
  		  $objPHPExcel->getActiveSheet()->getStyle('G3:N3')->applyFromArray(array('font'    => Style_Font(11, '000000', true, 'Calibri')));
                             $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font'    => Style_Font(20, '000000', true, 'Calibri')));
                             $objPHPExcel->getActiveSheet()->getStyle('A3')->applyFromArray(array('font'    => Style_Font(16, '000000', false, 'Calibri')));
                             $objPHPExcel->getActiveSheet()->getStyle('G1:N2')->applyFromArray(array('font'    => Style_Font(10, '000000', true, 'Calibri')));
                             $objPHPExcel->getActiveSheet()->getStyle('O1:AT1')->applyFromArray(array('font'    => Style_Font(20, '000000', true, 'Calibri')));
                           //$objPHPExcel->getActiveSheet()->getStyle('M1:'.$col_name[$Onday+10].'1')->applyFromArray(array('font' => Style_Font(13, 'FFFFFF', true, 'Consolas')));
                             $objPHPExcel->getActiveSheet()->getStyle('A1:F2')->applyFromArray(array('fill'    => Style_Fill('ffedcc')));
                             $objPHPExcel->getActiveSheet()->getStyle('A3:F3')->applyFromArray(array('fill'    => Style_Fill('ffedcc')));
                             $objPHPExcel->getActiveSheet()->getStyle('A4:F4')->applyFromArray(array('fill'    => Style_Fill('ffedcc')));
                             $objPHPExcel->getActiveSheet()->getStyle('G1:N2')->applyFromArray(array('fill'    => Style_Fill('ffedcc')));
                             $objPHPExcel->getActiveSheet()->getStyle('G3:N3')->applyFromArray(array('fill'    => Style_Fill('ffedcc')));
                             $objPHPExcel->getActiveSheet()->getStyle('G4:N4')->applyFromArray(array('fill'    => Style_Fill('ffedcc')));
                             $objPHPExcel->getActiveSheet()->getStyle('O1:AT1')->applyFromArray(array('fill'    => Style_Fill('F2F2F2')));
                             $objPHPExcel->getActiveSheet()->getStyle('O2:AT2')->applyFromArray(array('fill'    => Style_Fill('F2F2F2')));
                             $objPHPExcel->getActiveSheet()->getStyle('O3:AT3')->applyFromArray(array('fill'    => Style_Fill('F2F2F2')));
                             $objPHPExcel->getActiveSheet()->getStyle('O4:AT4')->applyFromArray(array('fill'    => Style_Fill('F2F2F2')));
//                             $objPHPExcel->getActiveSheet()->getStyle('M4:BE4')->applyFromArray(array('fill'    => Style_Fill('F2F2F2')));
                             $objPHPExcel->getActiveSheet()->getStyle('A1:'.'N4')->applyFromArray(array(
                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                             $objPHPExcel->getActiveSheet()->getStyle('O1:'.'BG1')->applyFromArray(array(
                                                                       'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                             $objPHPExcel->getActiveSheet()->getStyle('O2:'.'BG2')->applyFromArray(array(
                                                                       'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                             $objPHPExcel->getActiveSheet()->getStyle('O3:'.'BG3')->applyFromArray(array(
                                                                       'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
//                             $objPHPExcel->getActiveSheet()->getStyle('P3:'.'BE3')->applyFromArray(array(
//                                                                       'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
//                             $objPHPExcel->getActiveSheet()->getStyle('M1:'.$col_name[$Onday+10].'3')->applyFromArray(array(
//                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF')))
//                         												);

// //=========================================================================================================================================================================================================
                             $objPHPExcel->getActiveSheet()->mergeCells('A1:F2');
                             $objPHPExcel->getActiveSheet()->mergeCells('G1:G2');
                             $objPHPExcel->getActiveSheet()->mergeCells('A3:F3');
                             $objPHPExcel->getActiveSheet()->mergeCells('H1:J2');
                             $objPHPExcel->getActiveSheet()->mergeCells('K1:L2');
                             $objPHPExcel->getActiveSheet()->mergeCells('M1:M2');
                             $objPHPExcel->getActiveSheet()->mergeCells('N1:N2');
                             $objPHPExcel->getActiveSheet()->mergeCells('O1:AT1');                                               
//==========================================================================Accumulate Loss Time (Min) %=========================================================================================

          foreach (range(14, 45) as $key) 
           $objPHPExcel->getActiveSheet()->setCellValue($col_name[$key].'2', '=IFERROR( '.$col_name[$key].'3'.'/$'.'L3'. ',0)');   //Iferror 0/0

  		   $objPHPExcel->getActiveSheet()->getStyle('O2:BG2')->getNumberFormat()->setFormatCode('_*#,##0.00%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');                                              


#-------------------------------------------------------------------------------------------------------------------------------------------


            }
            elseif ($title[$indSheet] == 'Fa Summary loss') 
            {
                    $objPHPExcel->setActiveSheetIndex($indSheet);
                    //$objPHPExcel->getActiveSheet()->insertNewRowBefore(1,2);
                  //  $objPHPExcel->getActiveSheet()->freezePane('D4'); feed ด้านข้าง
                    $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 52 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 30 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 27 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 12 );
                   
                    $objPHPExcel->getActiveSheet()
                                ->getStyle(('A3:'.$col_name[56].'3'))
                                ->getAlignment()
                                ->setWrapText(true)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 

                    $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80); 
                     
                    $objPHPExcel->getActiveSheet()->setAutoFilter('A4:'.$col_name[34].'4');
                    $objPHPExcel->getActiveSheet()
                                ->getStyle('A5:'.$col_name[5].(count( $list_act_report[$key] )+5))
                                ->getAlignment()
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);   

                    $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[34].'3')->applyFromArray(array('fill'    => Style_Fill($colhead)));
                    $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[34].'3')->applyFromArray(array('font'    => Style_Font(12, '000000', true, 'Calibri')));
                    $objPHPExcel->getActiveSheet()->getStyle('A5:'.$col_name[34].(count( $list_act_report[$key] )+5))->applyFromArray(array('font'    => Style_Font(10, '000000', false, 'Calibri')));
                    $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[34].'4')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                    // $objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setBold(true)
                    //             ->setName('Consolas')
                    //             ->setSize(11)
                    //             ->getColor()->setRGB('FFFFFF');
                    $objPHPExcel->getActiveSheet()->getStyle('A4:'.$col_name[34].'4')->applyFromArray(array('fill'    => Style_Fill('e0ebeb'))); //fillter
                    $startData = 5;
                    $r = 5;
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

                                        }
                                $r++;
                            }
        $Montlast = date('F Y', strtotime(date('Y')."-".(date('m')-1)."-".'1'));  
        $M = date('M', strtotime(date('Y')."-".(date('m')-1)."-".'1'));
        $Daylast  = substr(date('Y-m-t',strtotime(date('Y')."-".(date('m')-1)."-".'1')),8, 2);
      //  echo $Montlast.$Daylast;exit;
                            $objPHPExcel->getActiveSheet()->setCellValue('A1', " FA SUMMARY LOSS REPORT OF " .$Montlast);
                            $objPHPExcel->getActiveSheet()->setCellValue('A2', "ACCUMULATE FROM (  01 ".strtoupper($M)." - ".$Daylast." ".strtoupper($M)."  )");
                            $objPHPExcel->getActiveSheet()->setCellValue('D1', "Accumulate Loss Time (Min)");

        foreach (range(3, 34) as $col) 
            $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col].'2', '=SUBTOTAL(9,'. $col_name[$col] .'5:'. $col_name[$col] . (count( $list_act_report[$key] )+5) . ")" );
#126180
                            
                            $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font'    => Style_Font(20, '000000', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('D1:AI1')->applyFromArray(array('font'    => Style_Font(20, 'FFFFFF', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('A2:C2')->applyFromArray(array('font'    => Style_Font(18, '000000', true, 'Calibri')));
                            $objPHPExcel->getActiveSheet()->getStyle('A3:AI3')->applyFromArray(array('font'    => Style_Font(12, 'FFFFFF', true, 'Calibri')));

                            $objPHPExcel->getActiveSheet()->setCellValue('A3', " PD " );
                            $objPHPExcel->getActiveSheet()->setCellValue('B3', " LINE CD" );
                            $objPHPExcel->getActiveSheet()->setCellValue('C3', " LINE NAME " );
                            $objPHPExcel->getActiveSheet()->setCellValue('D3',  " A " );
                            $objPHPExcel->getActiveSheet()->setCellValue('E3',  " B " );
                            $objPHPExcel->getActiveSheet()->setCellValue('F3',  " C " );
                            $objPHPExcel->getActiveSheet()->setCellValue('G3', " C1 " );
                            $objPHPExcel->getActiveSheet()->setCellValue('H3', " D " );
                            $objPHPExcel->getActiveSheet()->setCellValue('I3', " E " );
                            $objPHPExcel->getActiveSheet()->setCellValue('J3'," F " );
                            $objPHPExcel->getActiveSheet()->setCellValue('K3', " F1 " );
                            $objPHPExcel->getActiveSheet()->setCellValue('L3', " F2 " );
                            $objPHPExcel->getActiveSheet()->setCellValue('M3', " G" );
                            $objPHPExcel->getActiveSheet()->setCellValue('N3', " G1 " );
                            $objPHPExcel->getActiveSheet()->setCellValue('O3', " H" );
                            $objPHPExcel->getActiveSheet()->setCellValue('P3', " I " );
                            $objPHPExcel->getActiveSheet()->setCellValue('Q3', " J " );
                            $objPHPExcel->getActiveSheet()->setCellValue('R3', " K " );
                            $objPHPExcel->getActiveSheet()->setCellValue('S3', " K1 " );
                            $objPHPExcel->getActiveSheet()->setCellValue('T3', " K2 " );
                            $objPHPExcel->getActiveSheet()->setCellValue('U3', " K3 " );
                            $objPHPExcel->getActiveSheet()->setCellValue('V3', " L " );
                            $objPHPExcel->getActiveSheet()->setCellValue('W3', " L1 " );
                            $objPHPExcel->getActiveSheet()->setCellValue('X3', " M " );
                            $objPHPExcel->getActiveSheet()->setCellValue('Y3', " N " );
                            $objPHPExcel->getActiveSheet()->setCellValue('Z3', " O " );
                            $objPHPExcel->getActiveSheet()->setCellValue('AA3', " P " );
                            $objPHPExcel->getActiveSheet()->setCellValue('AB3', " Q " );
                            $objPHPExcel->getActiveSheet()->setCellValue('AC3', " Q1 " );
                             $objPHPExcel->getActiveSheet()->setCellValue('AD3', " R " );
                            $objPHPExcel->getActiveSheet()->setCellValue('AE3', " S " );
                            $objPHPExcel->getActiveSheet()->setCellValue('AF3', " T " );
                            $objPHPExcel->getActiveSheet()->setCellValue('AG3', " U " );
                            $objPHPExcel->getActiveSheet()->setCellValue('AH3', " V " );
                            $objPHPExcel->getActiveSheet()->setCellValue('AI3', " W " );
                                                                          

//====================================================================Code color================================================================================================================================
         
                            $objPHPExcel->getActiveSheet()->getStyle('A1:C1')->applyFromArray(array('fill'    => Style_Fill('FFFFFF')));
                            $objPHPExcel->getActiveSheet()->getStyle('A2:F2')->applyFromArray(array('fill'    => Style_Fill('FFFFFF')));
                            $objPHPExcel->getActiveSheet()->getStyle('A3:F3')->applyFromArray(array('fill'    => Style_Fill('FFFFFF')));
                            $objPHPExcel->getActiveSheet()->getStyle('G1:N1')->applyFromArray(array('fill'    => Style_Fill('FFFFFF')));
                            $objPHPExcel->getActiveSheet()->getStyle('G2:N2')->applyFromArray(array('fill'    => Style_Fill('FFFFFF')));
                            $objPHPExcel->getActiveSheet()->getStyle('G3:N3')->applyFromArray(array('fill'    => Style_Fill('FFFFFF')));
                            $objPHPExcel->getActiveSheet()->getStyle('K1:N1')->applyFromArray(array('fill'    => Style_Fill('FFFFFF')));
                            $objPHPExcel->getActiveSheet()->getStyle('K2:L2')->applyFromArray(array('fill'    => Style_Fill('FFFFFF')));
                            $objPHPExcel->getActiveSheet()->getStyle('D1:AI1')->applyFromArray(array('fill'    => Style_Fill('0f243e')));
                            $objPHPExcel->getActiveSheet()->getStyle('D2:AI2')->applyFromArray(array('fill'    => Style_Fill('e0ebeb')));
                            $objPHPExcel->getActiveSheet()->getStyle('A3:AI3')->applyFromArray(array('fill'    => Style_Fill('0f243e')));

                            $objPHPExcel->getActiveSheet()->getStyle('A1:'.'H3')->applyFromArray(array(
                                                                     'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                            $objPHPExcel->getActiveSheet()->getStyle('J1:'.'L3')->applyFromArray(array(
                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                            $objPHPExcel->getActiveSheet()->getStyle('H1:'.'J1')->applyFromArray(array(
                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                            $objPHPExcel->getActiveSheet()->getStyle('Z2:'.'AI2')->applyFromArray(array(
                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
                            $objPHPExcel->getActiveSheet()->getStyle('M1:'.$col_name[$Onday+10].'3')->applyFromArray(array(
                                                                     'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF')))
                                                );


//============================================================================== mergeCells ===========================================================================================================


                            $objPHPExcel->getActiveSheet()->mergeCells('A1:C1');
                            $objPHPExcel->getActiveSheet()->mergeCells('A2:C2');
                            $objPHPExcel->getActiveSheet()->mergeCells('D1:AI1');

                            $objPHPExcel->getActiveSheet()->getStyle('D2:'.$col_name[34].( count( $list_act_report[$key] )+5) )
                                                          ->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');

                            $objPHPExcel->getActiveSheet()->getStyle('D2:AI2')
                                                          ->getNumberFormat()->setFormatCode('_* #,##0_-;[RED](#,##0)_-;_* [BLACK]"-"??_-;_-@_-');

//============================================================================== setWidth ===========================================================================================================
                                                          
                            $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('67');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth('8');    
                            $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth('8');
                        

#-------------------------------------------------------------------------------------------------------------------------------------------

                            



            }              
            elseif ($title[$indSheet] == 'LOSS CODE') 
            {
                    $startData = 2;
                    $r = 2;
                            foreach ($value as $nr => $val) 
                            {
                                $indCol = 0;
                                        foreach ($val as $rowData => $data) 
                                        {

                                               $objPHPExcel->setActiveSheetIndex($indSheet)->setCellValue($col_name[$indCol++].($r), $data);


                                        }
                                $r++;
                            }
#========================================format_loss_code

                            $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('36');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('36');
                            $objPHPExcel->getActiveSheet()->getStyle('A1:D1')->applyFromArray(array('fill'    => Style_Fill('B8CCE4')));
                            



            }       
         } 
$indSheet++;
    //echo $indSheet; exit;
        
//$row = 5;
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
