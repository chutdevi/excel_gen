<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');
if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/PHPExcel-1.8.1/Classes/PHPExcel.php';

//============================================================================================= date =================================================================================

$freez = 'A4';
$start_col = 3; 
// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$data_col = array();

//var_dump($list_act_report); exit;

//exit;
$ind = 0;
foreach ($title as $inTil => $til) {
         $objPHPExcel->createSheet();
         $objPHPExcel->setActiveSheetIndex($ind);
         $objPHPExcel->getActiveSheet()->setTitle("$til");

//echo $til; exit;
$sheetIndex =  strtolower(str_replace(' ', '_', $title[$ind])); 
$end_row = count($list_act_report[$sheetIndex])+$start_col;
//echo $sheetIndex . " " . count($list_act_report[$sheetIndex][0]); exit;
if (count($list_act_report[$sheetIndex]) > 0) {  

if ($til == 'Fa report') {

          $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 53 );
          $objPHPExcel->getActiveSheet()
              ->getStyle('1')
              ->getAlignment()
              ->setWrapText(true)
              ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
              ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);       
           $style =   array(  
                              'font'    => array( 'size' => 11, 
                                                  'bold' => true,
                                                  'color' => array('rgb' => '000000')), 
                              'borders' => array(                                 
                                                  'allborders' => array(
                                                                         'style' => PHPExcel_Style_Border::BORDER_THIN,
                                                                         'color' => array('rgb' => 'FFFFFF')
                                                                       )
                                                )
                          );                        
          // echo count($list_act_report[$sheetIndex]); exit;    


          $col_name = array();
          $i = 0;
          foreach ( range('A', 'Z') as $cm ) { array_push($col_name, $cm); }
          foreach ( range('A', 'Z') as $cm ) { array_push($col_name, 'A'.$cm); }
          foreach ( range('A', 'Z') as $cm ) { array_push($col_name, 'B'.$cm); }  
          
          //echo count($list_act_report[$sheetIndex][0]); exit;
          //var_dump($col_name); exit;
        //  $objPHPExcel->getActiveSheet()->getStyle('K1:M'.(count( $list_act_report[$sheetIndex] )+5) )->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');

         $objPHPExcel->getActiveSheet()->getStyle('T1:AQ'.(count( $list_act_report[$sheetIndex] )+5) )->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');

          foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) {
              if($key == 'WORK_TIME') $key = "WORK_TIME";
              elseif($key == 'LOSS') $key = "LOSS"; 

              //echo strtoupper(str_replace('_', ' ', $key)) . "<hr>"; 
              $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."1",strtoupper(str_replace('_', ' ', $key)));       
          }
//exit;
          $objPHPExcel->getActiveSheet()->getStyle($col_name[0]."1:".$col_name[count($list_act_report[$sheetIndex][0])-1]."1")->applyFromArray($style);

          $objPHPExcel->getActiveSheet()
              ->getStyle($col_name[0]."1:".$col_name[count($list_act_report[$sheetIndex][0])-1]."1")
              ->applyFromArray(
                  array(
                      'fill' => array(
                                      'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                      'color' => array('rgb' => $colhead)
                                     )
                       )
              );


          foreach(range('A','J') as $columnID) {
              $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)
                  ->setAutoSize(true);
          }





        //  foreach (range('K', 'S') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('13');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth('12');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth('10');
                               $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth('12');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth('12');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth('12');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth('8');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth('7');
                              $objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth('7');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth('15');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth('12');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setWidth('12');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('AD')->setWidth('15');

       foreach (range('AC', 'AS') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('12');
       foreach (range('I', 'J') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('12'); #ขนาดคอลัม    


          for($i=2 ; $i < (count( $list_act_report[$sheetIndex] )+2) ; $i++)
          {
            $minusTime = '(L'.$i.'*'.'O'.$i.')/U'.$i;
            $objPHPExcel->getActiveSheet()->setCellValue('AB'.$i, '=IFERROR(IF(L'.$i.'="",0,'.$minusTime.'),0)');
          }




          for($i=2 ; $i < (count( $list_act_report[$sheetIndex] )+2) ; $i++)
          {
            $Mantotaltime = '(U'.$i.'*'.'E'.$i.')';
            $objPHPExcel->getActiveSheet()->setCellValue('AD'.$i, '='.$Mantotaltime);
          }


          $row = 2;
          foreach ($list_act_report[$sheetIndex] as $key => $value) {
              
              $col = 0;
              foreach ($value as $body => $val) 
              {

                      if ($body == 'PD' && $val == 'PD04')
                      {
                           $minusTime = '(L'.$row.'*'.'O'.$row.')/U'.$row.'/Y'.$row;
                           $objPHPExcel->getActiveSheet()->setCellValue('AB'.$row, '=IFERROR(IF(L'.$row.'="",0,'.$minusTime.'),0)');
                          //echo $minusTime. "<hr>" ; 
                          //echo $minusTime ; exit;                          
                       
                      }

                      if ($body == 'LOSS') {
                        $ckminustime = intval ($value['WORK_TIME']) - intval ($value['LOSS']); 
                        $ckZero = (intval ($value['WORK_TIME']) == 0 ) ?  1 : intval ($value['WORK_TIME']);
                        $EFF = ($ckminustime==0) ? 1 : $ckminustime / $ckZero;

                      }

                      elseif($val == 'READ') {

                          $objPHPExcel->getActiveSheet()->getStyle($col_name[count($list_act_report[$sheetIndex][0])-1].$row)->applyFromArray(array( 'font' => Style_Font(11,'4C9900',true)));
                      }
                      elseif ($val == 'UNCOMPLETE') {
                          $UNREADFillStyle = array( 'fill' => Style_Fill('FFE900') );   
                          $UNREADFontStyle = array( 'font' => Style_Font(12,'FF0000',true) );          
                          $objPHPExcel->getActiveSheet()->getStyle($col_name[0].$row.":".$col_name[count($list_act_report[$sheetIndex][0])-2].$row)->applyFromArray($UNREADFillStyle);
                          $objPHPExcel->getActiveSheet()->getStyle($col_name[count($list_act_report[$sheetIndex][0])-1].$row)->applyFromArray($UNREADFontStyle);
                          $objPHPExcel->getActiveSheet()->getStyle('K'.$row)->applyFromArray(array('font' => Style_Font(11,'FF0000',true)));
                      }


                      // else
                      // {
                      //echo . "<hr>"; 
                      if( substr($body, 0,4) != 'BANK')  $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col++].($row), $val);  
                      else $col++;

                      if($body == 'SEQ')
                      {
                        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$row)->getNumberFormat()->setFormatCode('000');
                        $objPHPExcel->getActiveSheet()->setCellValue($col_name[$col-1].$row, intval($val));
                      }

                     // var_dump($val); 
              }
             // exit;
             // var_dump($value); 
            //  exit;
              $row++;
              $objPHPExcel->getActiveSheet()->setAutoFilter($col_name[0]."1:K1");//.$col_name[count($list_act_report[$sheetIndex][0])-1]."1");
              $objPHPExcel->getActiveSheet()->freezePane('L2');
          }
//exit;
                    $objPHPExcel->getActiveSheet()->insertNewRowBefore(1,2);
                            $objPHPExcel->getActiveSheet()->setCellValue('A1', "DAILY FA REPORT  OF ".strtoupper(date('F Y')));
                            $objPHPExcel->getActiveSheet()->setCellValue('K1', "DATE :");
                            $objPHPExcel->getActiveSheet()->setCellValue('K2', "TOTAL :");
                            $objPHPExcel->getActiveSheet()->setCellValue('L1', "DETAIL OF: ".strtoupper(date('d-M-Y',  strtotime((date('d')-1) . "-" . date('M') . "-" . date('Y'))    )));
                            $objPHPExcel->getActiveSheet()->setCellValue('AE1', "IMPORTANT LOSS TIME CODE");

                            $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('L1')->applyFromArray(array('font' => Style_Font(18,'000000',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('AE1:AQ1')->applyFromArray(array('font' => Style_Font(18,'000000',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('M1:X1')->applyFromArray(array('font' => Style_Font(18,'000000',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('K1')->applyFromArray(array('font' => Style_Font(12,'000000',true)));            
                            $objPHPExcel->getActiveSheet()->getStyle('K2:Z2')->applyFromArray(array('font' => Style_Font(11,'000000',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('X2:AP2')->applyFromArray(array('font' => Style_Font(11,'000000',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('A3:Y3')->applyFromArray(array('font' => Style_Font(11,'000000',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('A3:Y3')->applyFromArray(array('font' => Style_Font(11,'000000',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('K1:AD1')->applyFromArray(array('fill' => Style_Fill('e6ffe6'))); //33cccc
                            $objPHPExcel->getActiveSheet()->getStyle('AE1:AQ1')->applyFromArray(array('fill'=> Style_Fill('b3cce6'))); //00b3b3
                            $objPHPExcel->getActiveSheet()->getStyle('AE2:AQ2')->applyFromArray(array('fill' => Style_Fill('b3cce6'))); //33cccc
                         //   $objPHPExcel->getActiveSheet()->getStyle('Z2:AA2')->applyFromArray(array('font' => Style_Font(11,'000000',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('AB2')->applyFromArray(array('font' => Style_Font(11,'ff0000',true)));

                            $objPHPExcel->getActiveSheet()->getStyle('Y3:S3')->applyFromArray(array('fill' => Style_Fill('ccffff'))); 
                            $objPHPExcel->getActiveSheet()->getStyle('AB2:AD2')->applyFromArray(array('fill' => Style_Fill('e6ffe6')));
                            $objPHPExcel->getActiveSheet()->getStyle('Z2:AA2')->applyFromArray(array('fill' => Style_Fill('e6ffe6')));
                            $objPHPExcel->getActiveSheet()->getStyle('K2:Y2')->applyFromArray(array('fill' => Style_Fill('e6ffe6'))); //33cccc
                            $objPHPExcel->getActiveSheet()->getStyle('T2:X2')->applyFromArray(array('fill'=> Style_Fill('e6ffe6')));
                         //    $objPHPExcel->getActiveSheet()->getStyle('Z2:AO2')->applyFromArray(array('fill'=> Style_Fill('b3cce6')));
                            $objPHPExcel->getActiveSheet()->getStyle('A3:AD3')->applyFromArray(array('fill'=> Style_Fill('ccffff'))); //ccffff
                        //    $objPHPExcel->getActiveSheet()->getStyle('Y3:Z3')->applyFromArray(array('fill'=> Style_Fill('OOOOOO')));
                            $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80); //ZOOM


                            $objPHPExcel->getActiveSheet()
                                        ->getStyle('A1:'.$col_name[date('d')+10].'2')
                                        ->getAlignment()
                                        ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
                                        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);                            

                              $objPHPExcel->getActiveSheet()->mergeCells('A1:J2');    
                     
                                              
                    $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 35 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 35 );
                    $objPHPExcel->getActiveSheet()
                                ->getStyle(('A3:'.$col_name[date('d')+10].'3'))
                                ->getAlignment()
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 

            //        $objPHPExcel->getActiveSheet()->setCellValue('M2', '=SUBTOTAL(9,M4:M'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('N2', '=SUBTOTAL(9,N4:N'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('O2', '=SUBTOTAL(9,O4:O'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                   // $objPHPExcel->getActiveSheet()->setCellValue('K2', '=SUBTOTAL(9,K4:K'. (count( $list_act_report[$sheetIndex] )+5) . ")");        

            

                  //  $objPHPExcel->getActiveSheet()->setCellValue('S2', '=ROUNDDOWN('.$subUsetime.'/60,0) & ":" & IF(LEN(MOD('.$subUsetime.',60)) = 1,"0"&'.'MOD('.$subUsetime.',60),MOD('.$subUsetime.',60))');
                    // $objPHPExcel->getActiveSheet()->setCellValue('T2', '=ROUNDDOWN('.$subworktime.'/60,0) & ":" & IF(LEN(MOD('.$subworktime.',60)) = 1,"0"&'.'MOD('.$subworktime.',60),MOD('.$subworktime.',60))');
                    // $objPHPExcel->getActiveSheet()->setCellValue('W2', '=ROUNDDOWN('.$subworktime.'/60,0) & ":" & IF(LEN(MOD('.$subworktime.',60)) = 1,"0"&'.'MOD('.$subworktime.',60),MOD('.$subworktime.',60))');

                //   $objPHPExcel->getActiveSheet()->setCellValue('S2', '=SUBTOTAL(9,S4:S'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('M2', '=SUBTOTAL(101,M4:M'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('L2', '=SUBTOTAL(101,L4:L'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('U2', '=SUBTOTAL(9,U4:U'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('V2', '=SUBTOTAL(9,V4:V'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                   $objPHPExcel->getActiveSheet()->setCellValue('W2', '=SUBTOTAL(9,W4:W'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('X2', '=SUBTOTAL(9,X4:X'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                 //    $objPHPExcel->getActiveSheet()->setCellValue('Y2', '=SUBTOTAL(9,Y4:Y'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('Z2', '=SUBTOTAL(9,Z4:Z'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    // $objPHPExcel->getActiveSheet()->setCellValue('AA2', '=SUBTOTAL(9,AA4:AA'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    // $objPHPExcel->getActiveSheet()->setCellValue('AB2', '=SUBTOTAL(9,AB4:AB'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    // $objPHPExcel->getActiveSheet()->setCellValue('AC2', '=SUBTOTAL(9,AC4:AC'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AD2', '=SUBTOTAL(9,AD4:AD'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AE2', '=SUBTOTAL(9,AE4:AE'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AF2', '=SUBTOTAL(9,AF4:AF'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AG2', '=SUBTOTAL(9,AG4:AG'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AH2', '=SUBTOTAL(9,AH4:AH'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AI2', '=SUBTOTAL(9,AI4:AI'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AJ2', '=SUBTOTAL(9,AJ4:AJ'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AK2', '=SUBTOTAL(9,AK4:AK'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AL2', '=SUBTOTAL(9,AL4:AL'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AM2', '=SUBTOTAL(9,AM4:AM'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AN2', '=SUBTOTAL(9,AN4:AN'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AD2', '=SUBTOTAL(9,AD4:AD'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AO2', '=SUBTOTAL(9,AO4:AO'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AP2', '=SUBTOTAL(9,AP4:AP'. (count( $list_act_report[$sheetIndex] )+5) . ")");
                    $objPHPExcel->getActiveSheet()->setCellValue('AQ2', '=SUBTOTAL(9,AQ4:AQ'. (count( $list_act_report[$sheetIndex] )+5) . ")");

                    $objPHPExcel->getActiveSheet()->getStyle('A1:AQ2')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));

                    $objPHPExcel->getActiveSheet()->getStyle('T2:AR2')->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); //IMPORT
                    $objPHPExcel->getActiveSheet()->getStyle('L2')->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[RED](#,##0.00)_-;_-* "-"??_-;_-@_-'); //IMPORT
                    $objPHPExcel->getActiveSheet()->getStyle('AB2'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');

          // $objPHPExcel->getActiveSheet()->insertNewColumnBefore('Y', 4);

          $objPHPExcel->getActiveSheet()->setCellValue('AE3', "[G]\n Model,Mold Change - in Plan");
          $objPHPExcel->getActiveSheet()->setCellValue('AF3', "[H]\n Tooling/Lot no.change/Tip/Sleeve - on Plan");
          $objPHPExcel->getActiveSheet()->setCellValue('AG3', "[I]\n Waiting for material from inhouse");
          $objPHPExcel->getActiveSheet()->setCellValue('AH3', "[K]\n Break Down - Machine ");
          $objPHPExcel->getActiveSheet()->setCellValue('AI3', "[K1]\n Break Down - Mold");
          $objPHPExcel->getActiveSheet()->setCellValue('AJ3', "[K2]\n Break Down - JIG /Tip/Sleeve");
          $objPHPExcel->getActiveSheet()->setCellValue('AK3', "[K3]\n Break Down - Tooling");
          $objPHPExcel->getActiveSheet()->setCellValue('AL3', "[L]\n Adjust Program and Condition.");
          $objPHPExcel->getActiveSheet()->setCellValue('AM3', "[N]\n Waiting for material from store");
          $objPHPExcel->getActiveSheet()->setCellValue('AN3', "[O]\n Material/Part Quality Problem");
          $objPHPExcel->getActiveSheet()->setCellValue('AO3', "[Q]\n Quality Judgement");
          $objPHPExcel->getActiveSheet()->setCellValue('AP3', "[Q1]\n Waiting data from QC/4M");
          $objPHPExcel->getActiveSheet()->setCellValue('AQ3', "[S]\n PE/CE Trial");

// #--------------------------------------------------------------------------------------------------Man Hr Pcs (TOTAL TIME)------------------------------------------
          $objPHPExcel->getActiveSheet()->setCellValue('AA3', 'Man Hr Pcs (min)'); 
          for($i=4 ; $i < (count( $list_act_report[$sheetIndex] )+4) ; $i++)
          {
            $minusTime = '(U'.$i.'*'.'E'.$i.')/O'.$i;
            $objPHPExcel->getActiveSheet()->setCellValue('AA'.$i, '='.$minusTime);
          }

// #--------------------------------------------------------------------------------------------------Man Hr Pcs (WORKING TIME)------------------------------------------

          $objPHPExcel->getActiveSheet()->setCellValue('AC3', 'Man Hr Pcs (min)'); 
          for($i=4 ; $i < (count( $list_act_report[$sheetIndex] )+4) ; $i++)
          {
            $minusTime = '(W'.$i.'*'.'E'.$i.')/O'.$i;
            $objPHPExcel->getActiveSheet()->setCellValue('AC'.$i, '='.$minusTime);
          }

// #-------------------------------------------------------------------------------------------------------------------------------------------
         
        //  $objPHPExcel->getActiveSheet()->setCellValue('AC3', 'Man Hr Pcs (min)'); 
          // for($=2 ; $i < (count( $list_act_report[$sheetIndex] )+4) ; $i++)
          // {
          //   $EFF = '(L2'.$i.'*'.'O2'.$i.')/U2'.$i;
          //   $objPHPExcel->getActiveSheet()->setCellValue('AB2'.$i, '='.$EFF);
          // }
          $SumACTUAL_CT = '(SUBTOTAL(101,M4:M307'. (count( $list_act_report[$sheetIndex] )+5) . '))';
          $SumSTDCT = '(SUBTOTAL(101,L4:L307'. (count( $list_act_report[$sheetIndex] )+5) . '))';
          $SumACTUAL = '(SUBTOTAL(9,O4:O307'. (count( $list_act_report[$sheetIndex] )+5) . '))';
          $Sumtotaltime = '(SUBTOTAL(9,U4:U307'. (count( $list_act_report[$sheetIndex] )+5) . '))';
          $Sumworktime = '(SUBTOTAL(9,W4:W307'. (count( $list_act_report[$sheetIndex] )+5) . '))';
          $Summan = '(SUBTOTAL(9,E4:E307'. (count( $list_act_report[$sheetIndex] )+5) . '))';


          $objPHPExcel->getActiveSheet()->setCellValue('AA2', '=('.$Sumtotaltime.'*'.$Summan.')/'.$SumACTUAL);
          $objPHPExcel->getActiveSheet()->setCellValue('AB2', '=('.$SumSTDCT.'*'.$SumACTUAL.')/'.$Sumtotaltime);
          $objPHPExcel->getActiveSheet()->setCellValue('AC2', '=('.$Sumworktime.'*'.$Summan.')/'.$SumACTUAL);

          
         
  



            $minusTime = '(L'.$row.'*'.'O'.$row.')/U'.$row.'/Y'.$row;

        //$objPHPExcel->getActiveSheet()->setCellValue('AB'.$row, '=IFERROR(IF(L'.$row.'="",0,'.$minusTime.'),0)');
        //   $minusTime = '(L'.$row.'*'.'O'.$row.')/U'.$row.'/Y'.$row;
        //    $objPHPExcel->getActiveSheet()->setCellValue('AB2', '=IFERROR(IF('.$SumSTDCT.'="",0,'.$SumSTDCT.'*'.$SumACTUAL.')/'.$Sumtotaltime);



          $SumTime = '(SUBTOTAL(9,T4:T272'. (count( $list_act_report[$sheetIndex] )+5) . '))';
          $Sumcycle = '(SUBTOTAL(9,L4:L272'. (count( $list_act_report[$sheetIndex] )+5) . '))';
     //     $Sumactual = '(SUBTOTAL(9,N4:N272'. (count( $list_act_report[$sheetIndex] )+5) . '))';

          $objPHPExcel->getActiveSheet()->setCellValue('T2', '='.$SumTime);
          
          $objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth('10');
          $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth('10');
          $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth('21');
          $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth('19');
          $objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth('12');
          $objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth('10');
          $objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth('10');
          $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth('10');
          $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth('10');
          $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth('10');
          $objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth('10');
          $objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth('10');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth('10');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AD')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AF')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AH')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AJ')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AK')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AL')->setWidth('14');

          $objPHPExcel->getActiveSheet()->getColumnDimension('AM')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setWidth('14');
          $objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setWidth('14');
		  $objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setWidth('14');
       //   foreach (range('K', 'T') as $key) { $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('12'); }
            # code...
          //'_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-'
       //   $objPHPExcel->getActiveSheet()->getStyle('AA2:AA'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
       //  $objPHPExcel->getActiveSheet()->getStyle('AD2:AN'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
       // //   $objPHPExcel->getActiveSheet()->getStyle('K2:K'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
       //   $objPHPExcel->getActiveSheet()->getStyle('O2:O'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');  
       //    $objPHPExcel->getActiveSheet()->getStyle('T2:T'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
      //    $objPHPExcel->getActiveSheet()->getStyle('S2:S'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');  

// #----------------------------------------------------------------------------EFF % TOTAL TIME---------------------------------------------------------------

      //     $objPHPExcel->getActiveSheet()->insertNewColumnBefore('Y', 1);
         
          $objPHPExcel->getActiveSheet()->setCellValue('AB3', 'EFF.(%)');


          // for($i=4 ; $i < (count( $list_act_report[$sheetIndex] )+4) ; $i++)
          // {
          //   $minusTime = '(L'.$i.'*'.'N'.$i.')/T'.$i;
          //   $objPHPExcel->getActiveSheet()->setCellValue('AA'.$i, '=IFERROR(IF(L'.$i.'="",0,'.$minusTime.'),0)');
          // }

// #---------------------------------------------------------------------------EFF % WORKING TIME----------------------------------------------------------------

          $objPHPExcel->getActiveSheet()->setCellValue('AD3', 'TOTALTIME ALL MAN');
          // for($i=4 ; $i < (count( $list_act_report[$sheetIndex] )+4) ; $i++)
          // {
          //   $minusTime = '(L'.$i.'*'.'N'.$i.')/V'.$i;
          //   $objPHPExcel->getActiveSheet()->setCellValue('AC'.$i, '=IFERROR(IF(L'.$i.'="",0,'.$minusTime.'),0)');
          // }


          $SumTime = '(SUBTOTAL(9,S4:S'. (count( $list_act_report[$sheetIndex] )+5) . '))';

          // $objPHPExcel->getActiveSheet()->setCellValue('Y2', '=('.$Sumcycle.'*'.$Sumactual.')/'.$SumTime);


      //    foreach (range('S', 'Z') as $key) { $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('12'); }
            # code...

    $objPHPExcel->getActiveSheet()
    ->getStyle('A1:AQ2')
    ->getAlignment()
    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);      

       $objPHPExcel->getActiveSheet()->getStyle('N2:O2'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
       $objPHPExcel->getActiveSheet()->getStyle('M2:M'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
       $objPHPExcel->getActiveSheet()->getStyle('P2:P'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');
       $objPHPExcel->getActiveSheet()->getStyle('AA2:AA'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-'); 
      $objPHPExcel->getActiveSheet()->getStyle('AB2:AB'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');
       $objPHPExcel->getActiveSheet()->getStyle('AC2:AC'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
     //   $objPHPExcel->getActiveSheet()->getStyle('AC2'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');
        //  $objPHPExcel->getActiveSheet()->getStyle('AL2:AL'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-');
          // $objPHPExcel->getActiveSheet()->getStyle('L'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-'); 

          // $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80);  
        //   $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setVisible(false); //hide
          $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setVisible(false);
          $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setVisible(false);      
          $objPHPExcel->getActiveSheet()->getColumnDimension('S')->setVisible(false); 
          $objPHPExcel->getActiveSheet()->getColumnDimension('T')->setVisible(false);
           $objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setVisible(false);
          $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setVisible(false);
    //      $objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setVisible(false);
          $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(28,'000000',true)));



          $objPHPExcel->getActiveSheet()->mergeCells('L1:X1');
       //   $objPHPExcel->getActiveSheet()->mergeCells('AB2:AC2');
       //   $objPHPExcel->getActiveSheet()->mergeCells('Z2:AA2');
          $objPHPExcel->getActiveSheet()->mergeCells('AE1:AQ1');

   }
   else
   {
//  var_dump($list_act_report[$sheetIndex]);
// exit();
    $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 28 );
    $objPHPExcel->getActiveSheet()
    ->getStyle('1')
    ->getAlignment()
    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);       
 $style =   array(  
                    'font'    => array( 'size' => 11, 
                                        'bold' => true,
                                        'color' => array('rgb' => '000000')), 
                    'borders' => array(                                 
                                        'allborders' => array(
                                                               'style' => PHPExcel_Style_Border::BORDER_THIN,
                                                               'color' => array('rgb' => 'FFFFFF')
                                                             )
                                      )
                );                        
// echo count($list_act_report[$sheetIndex]); exit;    


$col_name = array();
$i = 0;
foreach ( range('A', 'Z') as $cm ) {
    array_push($col_name, $cm);
}
// echo count($list_act_report[$sheetIndex][0]); exit;
// var_dump($col_name); exit;

foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) {
    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."1", strtoupper($key));       
}

$objPHPExcel->getActiveSheet()->getStyle($col_name[0]."1:".$col_name[count($list_act_report[$sheetIndex][0])-1]."1")->applyFromArray($style);

$objPHPExcel->getActiveSheet()
    ->getStyle($col_name[0]."1:".$col_name[count($list_act_report[$sheetIndex][0])-1]."1")
    ->applyFromArray(
        array(
            'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => $colhead)
                           )
             )
    );


foreach(range('A',$col_name[count($list_act_report[$sheetIndex][0])]) as $columnID) {
    $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)
        ->setAutoSize(true);
}

$row = 2;
foreach ($list_act_report[$sheetIndex] as $key => $value) {
    
    $col = 0;
    foreach ($value as $body => $val) {
            if ($val == 'UNREAD') {
                $UNREADFillStyle = array( 'fill' => Style_Fill('FF6666') );   
                $UNREADFontStyle = array( 'font' => Style_Font(12,'FF0000',true) );          
                $objPHPExcel->getActiveSheet()->getStyle($col_name[0].$row.":".$col_name[count($list_act_report[$sheetIndex][0])-2].$row)->applyFromArray($UNREADFillStyle);
                $objPHPExcel->getActiveSheet()->getStyle($col_name[count($list_act_report[$sheetIndex][0])-1].$row)->applyFromArray($UNREADFontStyle);
            }

            elseif($val == 'READ') {

                $objPHPExcel->getActiveSheet()->getStyle($col_name[count($list_act_report[$sheetIndex][0])-1].$row)->applyFromArray(array( 'font' => Style_Font(11,'4C9900',true)));
            }
            elseif ($val == 'UNCOMPLETE') {
                $UNREADFillStyle = array( 'fill' => Style_Fill('FFE900') );   
                $UNREADFontStyle = array( 'font' => Style_Font(12,'FF0000',true) );          
                $objPHPExcel->getActiveSheet()->getStyle($col_name[0].$row.":".$col_name[count($list_act_report[$sheetIndex][0])-2].$row)->applyFromArray($UNREADFillStyle);
                $objPHPExcel->getActiveSheet()->getStyle($col_name[count($list_act_report[$sheetIndex][0])-1].$row)->applyFromArray($UNREADFontStyle);
                $objPHPExcel->getActiveSheet()->getStyle('K'.$row)->applyFromArray(array('font' => Style_Font(11,'FF0000',true)));
            }
            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);
           // var_dump($val); 
    }
   // var_dump($value); 
  //  exit;
    $row++;
    $objPHPExcel->getActiveSheet()->setAutoFilter($col_name[0]."1:".$col_name[count($list_act_report[$sheetIndex][0])-1]."1");
    $objPHPExcel->getActiveSheet()->freezePane('A2');
   }
   } 
} else {

            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "No data ".$til.".");
            $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true)));
            //echo "Non data."; exit;
}
// $objPHPExcel->getActiveSheet()->setTitle($title);
$ind++;


}
if ($til == 'loss code'){



  
}

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->getStyle('ZZ1')->getNumberFormat()->setFormatCode('_-* #,##0_-;-* #,##0_-;_-* "-"??_-;_-@_-');
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
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


function Style_Fill($color=null) {

    return array( 'type'  => PHPExcel_Style_Fill::FILL_SOLID,                           
                  'color' => array('rgb' => $color)                                    
                );                                   
}

function Style_Font($size=11, $color='FFFFFF', $bol=false) {

    return  array(
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

function set_autosize($colA = 'A' ,$colB = 'Z',  $objPHPExcel = nul, $index = 0)
{
        $objPHPExcel->setActiveSheetIndex($index); 
    foreach(range($colA, $colB) as $columnID) 
    {        
        $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);       
    }                     
}


?>

 
