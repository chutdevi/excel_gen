<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/PHPExcel-1.8.1/Classes/PHPExcel.php';

$freez = 'A4';
$start_col = 3; 
// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$data_col = array();

//var_dump($list_act_report); exit;
//exit;
$ind = 0;
$col_name = array();

foreach ( range('A', 'Z') as $cm ) { array_push($col_name, $cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "A".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "B".$cm); }
$i   = 0;   
$ind = 0;
foreach ($title as $inTil => $til) 
{
             $objPHPExcel->createSheet();
             $objPHPExcel->setActiveSheetIndex($ind);
             $objPHPExcel->getActiveSheet()->setTitle( "$til ( ". date('Y-m-d') . " )" );

    $sheetIndex  =  strtolower(str_replace(' ', '_', $title[$ind])); 
    $count_index = 0;
    $count_data  =  count($list_act_report[$sheetIndex]) + 6;
    if ($count_data - 6  > 0) 
    {      
       // if ($til == 'FA Daily') { 
            $count_index =  count($list_act_report[$sheetIndex][0]) - 1 ;
            $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 40 );   //Row SIZE
            $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 33 );
            $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 33 );
            $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 30 );
            $objPHPExcel->getActiveSheet()->getRowDimension( 5 )->setRowHeight( 30 );
            $objPHPExcel->getActiveSheet()
                ->getStyle('1:5')
                ->getAlignment()
                ->setWrapText(true)
                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);       
                                   


            $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80);    
            
            $objPHPExcel->getActiveSheet()->freezePane('A6'); #-- Freeze
            $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[$count_index].'4')->applyFromArray(array('fill'    => Style_Fill('0a1429')));
            $objPHPExcel->getActiveSheet()->getStyle('K1:'.$col_name[$count_index].'4')->applyFromArray(array('fill'    => Style_Fill('0a1429'))); //4d0000           
            $objPHPExcel->getActiveSheet()->getStyle('A2:'.$col_name[$count_index].'4')->applyFromArray(array('fill'    => Style_Fill('0a1429')));
            $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[$count_index].'5')->applyFromArray(array('fill'    => Style_Fill('E6EDF5')));
            $objPHPExcel->getActiveSheet()->getStyle('A4:'.$col_name[$count_index].'4')->applyFromArray(array('fill'    => Style_Fill('cce6ff')));
            $objPHPExcel->getActiveSheet()->getStyle('A5:'.$col_name[$count_index].'5')->applyFromArray(array('fill'    => Style_Fill('e6f2ff')));
            $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[$count_index].'5')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
            foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
            {
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."4", str_replace("_", " ", strtoupper($key)));       
            }            
    $row = 6; #-- Start data 
            foreach ($list_act_report[$sheetIndex] as $key => $value) 
            {
                
                $col = 0;
                foreach ($value as $body => $val) 
                {
                        $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);


                        if ($val == '')
                         {
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col-1].($row), 0);                        
                         } 
                          $ckminustime = intval ($value['USE_TIME']) - intval ($value['LOSS']); 
                          $EFF = $ckminustime/intval ($value['USE_TIME']);
                         if($EFF*100 < 80) 
                         {
                        //  echo $EFF;exit;
                          $UNREADFillStyle = array( 'fill' => Style_Fill('ffffcc') );   
                          $UNREADFontStyle = array( 'font' => Style_Font(12,'0000FF',true) );          
                          $objPHPExcel->getActiveSheet()->getStyle('A'.$row.":".'J'.$row)->applyFromArray($UNREADFillStyle);
                          $objPHPExcel->getActiveSheet()->getStyle('A'.$row.":".'J'.$row)->applyFromArray($UNREADFontStyle);
                        }

                }
                $row++;
                
            }
                 
          $objPHPExcel->getActiveSheet()->setCellValue('K5', "Model Change.");
          $objPHPExcel->getActiveSheet()->setCellValue('L5', "Mold Change.");
          $objPHPExcel->getActiveSheet()->setCellValue('M5', "Tool Change.");
          $objPHPExcel->getActiveSheet()->setCellValue('N5', "MC Adjust.");
          $objPHPExcel->getActiveSheet()->setCellValue('O5', "DIE Adjust.");
          $objPHPExcel->getActiveSheet()->setCellValue('P5', "MC Breakdown.");
          $objPHPExcel->getActiveSheet()->setCellValue('Q5', "Casting Breakdown.");
          $objPHPExcel->getActiveSheet()->setCellValue('R5', "RM Shortage.");
          $objPHPExcel->getActiveSheet()->setCellValue('S5', "Quality Problem.");
          $objPHPExcel->getActiveSheet()->setCellValue('T5', "Waiting Box.");
          $objPHPExcel->getActiveSheet()->setCellValue('U5', "Manage.");


            foreach(range( 10 ,$count_index) as $columnID)
                $objPHPExcel->getActiveSheet()->setCellValue("$col_name[$columnID]3", "=SUBTOTAL(9,$col_name[$columnID]6:$col_name[$columnID]".$count_data .")");

            foreach(range( 10 ,$count_index) as $columnID)
                $objPHPExcel->getActiveSheet()->setCellValue("$col_name[$columnID]2", "=($col_name[$columnID]3/J3)");

            $objPHPExcel->getActiveSheet()->getStyle('K2:'.$col_name[$count_index].'2')->getNumberFormat()->setFormatCode('_ #,##0%_-;[RED](#,##0%)_-;_ "-"??_-;_-@_-');

            //var_dump($list_act_report); exit;
#-----------------------------------------------------------------------------------------------------------------
            $objPHPExcel->getActiveSheet()->insertNewColumnBefore('K', 1);
            $objPHPExcel->getActiveSheet()->setCellValue('K4', "Man Hr Pcs \n [min] ");

            for($i=6 ; $i < $count_data; $i++)
          {
            $minusTime = '((F'.$i.'+'.'J'.$i.')*E'.$i.')/H'.$i;
            $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, '='.$minusTime);
          }
            $minusTime = '((F3+J3)*E3)/H3';
            $SumTime = '(SUBTOTAL(9,K6:K'. $count_data. '))';
            $objPHPExcel->getActiveSheet()->setCellValue('K3', '='.$minusTime);



#MAN-----------------------------------------------------------------------------------------------------------------          
            $SumTime = '(SUBTOTAL(9,E6:E'. (count( $list_act_report[$sheetIndex] )+5) . '))'; 
            $objPHPExcel->getActiveSheet()->setCellValue('E3', '='.$SumTime);
#USE TIME-------------------------------------------------------------------------------------------------------------
            $SumTime = '(SUBTOTAL(9,F6:F'. (count( $list_act_report[$sheetIndex] )+5) . '))';
            $objPHPExcel->getActiveSheet()->setCellValue('F3', '='.$SumTime);
#PLAN-----------------------------------------------------------------------------------------------------------------
            $SumTime = '(SUBTOTAL(9,G6:G'. (count( $list_act_report[$sheetIndex] )+5) . '))';
            $objPHPExcel->getActiveSheet()->setCellValue('G3', '='.$SumTime);
#ACTUAL----------------------------------------------------------------------------------------------------------------
            $SumTime = '(SUBTOTAL(9,H6:H'. (count( $list_act_report[$sheetIndex] )+5) . '))';
            $objPHPExcel->getActiveSheet()->setCellValue('H3', '='.$SumTime);
#DIFF------------------------------------------------------------------------------------------------------------------
            $SumTime = '(SUBTOTAL(9,I6:I'. (count( $list_act_report[$sheetIndex] )+5) . '))';
            $objPHPExcel->getActiveSheet()->setCellValue('I3', '='.$SumTime);
#LOSS------------------------------------------------------------------------------------------------------------------

            $SumTime = '(SUBTOTAL(9,J6:J'. (count( $list_act_report[$sheetIndex] )+5) . '))';
            $objPHPExcel->getActiveSheet()->setCellValue('J3', '='.$SumTime);
            $objPHPExcel->getActiveSheet()->getStyle('K3:K'.$count_data)->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[RED](#,##0.00)_-;_-* "-"??_-;_-@_-');

#---------------------------------------------------------------------------------------------------------------------- EFF
            $objPHPExcel->getActiveSheet()->insertNewColumnBefore('K', 1);
            $objPHPExcel->getActiveSheet()->setCellValue('K4', "EFF.(%) ");

        for($i=6 ; $i < $count_data; $i++)
          {
            $minusTime = '(F'.$i.'-'.'J'.$i.')';
            $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, '=IF(F'.$i.'<1,0,'.$minusTime.'/'.'F'.$i.')');
          }

          $SumTime = '(SUBTOTAL(9,F6:F'. $count_data. '))';
          $objPHPExcel->getActiveSheet()->setCellValue('K3', '=('.$SumTime.'-J3)/'.$SumTime);
          $objPHPExcel->getActiveSheet()->getStyle('K3:K'.$count_data)->getNumberFormat()->setFormatCode('_-* #,##0%_-;[RED](#,##0%)_-;_-* "-"??_-;_-@_-');
          $objPHPExcel->getActiveSheet()->getStyle('F6:J'.$count_data)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');
          $objPHPExcel->getActiveSheet()->getStyle('M6:W'.$count_data)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); 

          $objPHPExcel->getActiveSheet()->getStyle('M3:W'.$count_data)->getNumberFormat()->setFormatCode('_ #,##0_-;[Red](#,##0)_-;_ "-"??_-;_-@_-'); //TOTAL M3:W  loss _-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-

          $objPHPExcel->getActiveSheet()->getStyle('A3:J'.$count_data)->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-'); //TOTAL A3:J
#-----------------------------------------------------------------------------------------------------------------
#font -----------------------------------------------------------------------------------------------------------------

            $objPHPExcel->getActiveSheet()->setAutoFilter('A5:'.$col_name[11].'5'); #-- fillter
            $objPHPExcel->getActiveSheet()->getStyle('A1:J1')->applyFromArray(array('font' => Style_Font(28,'FFFFFF',true, 'Consolas')));
            $objPHPExcel->getActiveSheet()->getStyle('A2:J2')->applyFromArray(array('font' => Style_Font(18,'FFFFFF',true, 'Calibri')));
            $objPHPExcel->getActiveSheet()->getStyle('K2:W2')->applyFromArray(array('font' => Style_Font(14,'FFFFFF',true, 'Calibri')));

            $objPHPExcel->getActiveSheet()->getStyle('K1:U1')->applyFromArray(array('font' => Style_Font(22,'FFFFFF',true)));
            $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[$count_index].$count_data)->applyFromArray(array('font' => Style_Font(14,'000000',true,'Calibri')));
            $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[$count_index].$count_data)->applyFromArray(array('font' => Style_Font(10,'000000',false,'Calibri')));


        
                            $objPHPExcel->getActiveSheet()->setCellValue('A1', "DAILY FA REPORT OF ".strtoupper(date('F Y')));
                         //   $objPHPExcel->getActiveSheet()->setCellValue('A2', "DATE OF".strtoupper(date('F Y')));
                            $objPHPExcel->getActiveSheet()->setCellValue('A3', "TOTAL");
                         //   $objPHPExcel->getActiveSheet()->setCellValue('A2', "DETAIL OF ".strtoupper(date('d-M-Y',  strtotime((date('d')-1) . "-" . date('M') . "-" . date('Y'))    )));
                            $objPHPExcel->getActiveSheet()->setCellValue('A2', "DETAIL OF: ".strtoupper(date('d-M-Y',  strtotime((date('d')-1) . "-" . date('M') . "-" . date('Y'))    )));
                            $objPHPExcel->getActiveSheet()->setCellValue('M1', "IMPORTANT LOSS TIME CODE [MIN]");
                    
            
            //foreach(range( 0 ,$count_index) as $columnID) 
            //$objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$columnID]) ->setAutoSize(true);
                $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('10');
                $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('12');
                $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('8');
                $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('8');
                $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('14');
                $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('14'); 
                $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('14'); 
                $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('8'); 
                $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('8'); 
                $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth('14'); 
                $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth('14');
                $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth('18');
                $objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth('18');
                $objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth('18');
                $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth('18');
                $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth('18');
                $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth('18');
                $objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth('18');
                $objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth('18');
                $objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth('18');
                $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth('18');
                $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth('18');            
            #--MERT  
                $objPHPExcel->getActiveSheet()->mergeCells('A1:L1');
                $objPHPExcel->getActiveSheet()->mergeCells('M1:W1');
                $objPHPExcel->getActiveSheet()->mergeCells('A2:L2');
                $objPHPExcel->getActiveSheet()->mergeCells('A3:D3');
                //$objPHPExcel->getActiveSheet()->mergeCells('Y1:AF1');


                                            
               foreach (range(12, 22) as $index) Style_group_lv1($col_name, $index, $objPHPExcel); 
               $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('X1', "X");
               $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('Y1', "↢ Unhide to view important loss time code");
               $objPHPExcel->getActiveSheet()->getStyle('X1')->applyFromArray(array('font' => Style_Font(36,'FF0000',true,'Wingdings 3')));
               $objPHPExcel->getActiveSheet()->getStyle('Y1')->applyFromArray(array('font' => Style_Font(18,'FF0000',true,'Franklin Gothic Book')));
               $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('F4', "USE TIME STD \n [min]");
               $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('G4', "PLAN \n [pcs.]");
               $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('H4', "ACTUAL \n [pcs.]");
               $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('J4', "LOSS \n [min]");
              // $objPHPExcel->getActiveSheet()->getStyle('M2:W'.(count( $list_act_report[$sheetIndex] )+5))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[Red](#,##0.00)_-;_-* "-"??_-;_-@_-'); 

               $objPHPExcel->getActiveSheet()->mergeCells('Y1:AF1');
    
   // } elseif ($till = 'Loss Code') {
                    //echo "AAAA";
    } else {

                    $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "No data ".$til.".");
                    $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true,'Franklin Gothic Book')));
                    //echo "Non data."; exit;
    }
// $objPHPExcel->getActiveSheet()->setTitle($title);
$ind++;


}

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);
 $objPHPExcel->getActiveSheet()->getStyle('ZZ1');
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
function Style_group_lv1($cell=null, $index=0, $objPHPExcel=null)
{
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setOutlineLevel(1);
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setVisible(false);
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setCollapsed(true); 
}

?>

 
