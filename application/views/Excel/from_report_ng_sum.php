<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$data_col = array();
//var_dump($list_act_report); exit;

//exit;    
$col_name = array();

    foreach ( range('A', 'Z') as $cm ) { array_push($col_name, $cm); }
    foreach ( range('A', 'Z') as $cm ) { array_push($col_name, 'A'.$cm); }
    foreach ( range('A', 'Z') as $cm ) { array_push($col_name, 'B'.$cm); }
    foreach ( range('A', 'Z') as $cm ) { array_push($col_name, 'C'.$cm); }
    foreach ( range('A', 'Z') as $cm ) { array_push($col_name, 'D'.$cm); }
    foreach ( range('A', 'Z') as $cm ) { array_push($col_name, 'E'.$cm); }
    foreach ( range('A', 'Z') as $cm ) { array_push($col_name, 'F'.$cm); }
$ind = 0;
foreach ($title as $inTil => $til) 
{
         $objPHPExcel->createSheet();
         $objPHPExcel->setActiveSheetIndex($inTil);
         $objPHPExcel->getActiveSheet()->setTitle("$til");

//echo $til; exit;
        $sheetIndex =  strtolower(str_replace(' ', '_', $title[$inTil])); 
        $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 35 );
        $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(65);
        $objPHPExcel->getActiveSheet()
            ->getStyle('1')
            ->getAlignment()
            ->setWrapText(true)
            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);       
        //echo count($list_act_report[$sheetIndex][0])-1; exit;






// var_dump($field_name[50]['field_nm']); 
//var_dump($list_act_report['defect_summary'][0]);
//exit;   

     $i = 0;
     if(count($list_act_report[$sheetIndex]) > 0 )
     {
        //echo count(count($list_act_report[$sheetIndex])); exit;
        foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
        {

            $key = str_replace('CD_', '' , $key);
            $key = str_replace('QC_', '' , $key);
            $key = str_replace('_'  , ' ',  $key);
            if ($i > 9)
            {
                $objPHPExcel->getActiveSheet()->getStyle($col_name[$i]."1")->getNumberFormat()->setFormatCode('000');
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."1", intval($key));                
            }
            else
            {
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."1", strtoupper($key));     
            }

            $objPHPExcel->getActiveSheet()->getStyle($col_name[$i-1].'1')->applyFromArray(array('font'    => Style_Font(10,'FFFFFF',true)));
            $objPHPExcel->getActiveSheet()->getStyle($col_name[$i-1].'1')->applyFromArray(array('fill'    => Style_Fill($colhead)));
            $objPHPExcel->getActiveSheet()->getStyle($col_name[$i-1].'1')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));
            // if (substr($key,0,3) == 'day') 
            // {
            //     $dayCh = substr($key,3,strlen($key)-3);
            //     if ($dayCh == date('d')-1) 
            //     {
            //         //echo $key;
            //         break;
            //     }
            // }
              
        }
//$objPHPExcel->getActiveSheet()->getStyle($col_name[0]."1:".$col_name[count($list_act_report[$sheetIndex][0])-1]."1")->applyFromArray($style);
     }
}   


        
            $row = 2;
            $indSheet = 0;
            foreach ($list_act_report as $key => $value) 
            {
                if(count($list_act_report[$key]) > 0 )
                { 
                    if ($key == 'defect_summary')
                    {

                            //set_autosize('A','CZ', $objPHPExcel, $indSheet);
                            $objPHPExcel->setActiveSheetIndex($indSheet);
                            $objPHPExcel->getActiveSheet()->insertNewRowBefore(1,3); 
                            $objPHPExcel->getActiveSheet()->freezePane('A5');
                            $r = 5;


                                foreach ($value as $nr => $val) 
                                {
                                  //echo count($value); exit;
                                $indCol = 0;
                                    foreach ($val as $rowData => $data) 
                                    {

                                       // var_dump($rowData); exit;                                                                                                                                                                                  


                                        if($data > 0 && $indCol-1 > 9)
                                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$indCol].($r))->applyFromArray(array('font' => Style_Font(14,'0000FF',true)));
                                        if($data > 0 && $rowData == 'ACTUAL_OK')
                                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$indCol].($r))->applyFromArray(array('font' => Style_Font(15,'008000',true)));
                                        if($data > 0 && $rowData == 'RECEIVE')
                                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$indCol].($r))->applyFromArray(array('font' => Style_Font(15,'008000',true)));                                        
                
                                            //echo $r.$data; exit;
                                            //$objPHPExcel->getActiveSheet()->setCellValue($col_name[$indCol].'5', $data);
                                    
                                    
                                           if($r == 5)
                                           {
                                                if ($indCol-1 > 9) 
                                                {
                                                 $objPHPExcel->getActiveSheet()->setCellValue($col_name[$indCol].'1', "=SUBTOTAL(9,".$col_name[$indCol]."5:".$col_name[$indCol].(count($value)+7).")");
                                                }

                                           }
                                            if ($data == 0 && $indCol-1 > 7) 
                                            {
                                                 $objPHPExcel->getActiveSheet()->setCellValue($col_name[$indCol++].($r), '-');
                                            }
                                            else
                                            {
                                                if ($data === '3E00') 
                                                {
                                                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$indCol++].($r), "'".$data);
                                                }
                                                else
                                                {
                                                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$indCol++].($r), $data);
                                                }                                                
                                                 //$objPHPExcel->getActiveSheet()->setCellValue($col_name[$indCol++].($r), $data); 
                                            }
                                    }    // break;
                                               $r++;

                                }
                                $MntHis = ( (date('d')+0) == 1 ) ? date('F-Y', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('F-Y'));
                                $Yeday = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : date('d', strtotime(date('Y')."-".(date('m')+0)."-".(date('d')-1)));


                                $MontCol = ( (date('d')+0) == 1 ) ? date('M', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('M'));
                                $YearCol = ( (date('d')+0) == 1 ) ? date('Y', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('Y'));
                                $objPHPExcel->getActiveSheet()->setCellValue('A1', "SUMMARY DEFECT OF ".strtoupper($MntHis));
                                $objPHPExcel->getActiveSheet()->setCellValue('A3', "ACCUMULATE FROM (  01 ".strtoupper($MontCol)." - ".strtoupper($Yeday." ".$MontCol)."  )");
                                $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'16365C',true)));
                                $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                                $objPHPExcel->getActiveSheet()->getStyle('A3')->applyFromArray(array('font' => Style_Font(28,'16365C',true)));
                                $objPHPExcel->getActiveSheet()->getStyle('A3')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));                                
                                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 40 );
                                $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 40 );
                                $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 36 );
                                
                                $objPHPExcel->getActiveSheet()
                                            ->getStyle('J5:CQ'.(count($value)+5))
                                            ->getAlignment()
                                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                                $objPHPExcel->getActiveSheet()
                                            ->getStyle('I1:K2')
                                            ->getAlignment()
                                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
                                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                                $objPHPExcel->getActiveSheet()
                                            ->getStyle('A1:H3')
                                            ->getAlignment()
                                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                                $objPHPExcel->getActiveSheet()->setCellValue('I1', "TOTAL DEFECT >>");
                                $objPHPExcel->getActiveSheet()->setCellValue('I2', "TOTAL ACTUAL >>");
                                $objPHPExcel->getActiveSheet()->setCellValue('J1', "=SUM(L1:CR1)");
                                $objPHPExcel->getActiveSheet()->setCellValue('J2', "=SUBTOTAL(9,J5:J".(count($value)+7).")");
                                $objPHPExcel->getActiveSheet()->setCellValue('K2', "=SUBTOTAL(9,K5:K".(count($value)+7).")");












                                $objPHPExcel->getActiveSheet()->getStyle('L1:CQ1')->applyFromArray(array('font' => Style_Font(16, 'FF0000',true)));
                                $objPHPExcel->getActiveSheet()->getStyle('I1:I2')->applyFromArray(array( 'font' => Style_Font(16, '632523', TRUE)));
                                
                                $objPHPExcel->getActiveSheet()->getStyle('J1')->applyFromArray(array( 'font' => Style_Font(17, 'FF0000', true)));
                                $objPHPExcel->getActiveSheet()->getStyle('K2')->applyFromArray(array( 'font' => Style_Font(17, '008000', true)));
                                $objPHPExcel->getActiveSheet()->getStyle('J2')->applyFromArray(array( 'font' => Style_Font(17, '008000', true)));
                                $objPHPExcel->getActiveSheet()->getStyle('J1:CQ1')->getNumberFormat()->setFormatCode('_* #,##0_-;[RED](#,##0)_-;_* "-"_-;_-@_-');
                                $objPHPExcel->getActiveSheet()->getStyle('J2:L2')->getNumberFormat()->setFormatCode('_* #,##0_-;[RED](#,##0)_-;_* "-"_-;_-@_-');
                                $objPHPExcel->getActiveSheet()->getStyle('J5:CQ'.(count($value)+5))->getNumberFormat()->setFormatCode('_* #,##0_-;[RED](#,##0)_-;_* "-"_-;_-@_-');

                                $objPHPExcel->getActiveSheet()->getStyle('I1:QC3')->applyFromArray(array(
                                                                                                            'borders' => array( 'allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'B1A0C7'))));                               
                                $objPHPExcel->getActiveSheet()->getStyle('L1:CQ1')->applyFromArray(array(
                                                                                                            'borders' => array(
                                                                                                                                'allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'B1A0C7')))); 
                                                                                                                              
                                $objPHPExcel->getActiveSheet()->getStyle('I1:CQ1')->applyFromArray(array('fill' => Style_Fill('C5D9F1')));
                                $objPHPExcel->getActiveSheet()->getStyle('I2:K2')->applyFromArray(array('fill'  => Style_Fill('C5D9F1')));   



                                $objPHPExcel->getActiveSheet()->mergeCells('A1:H2');
                                $objPHPExcel->getActiveSheet()->mergeCells('A3:H3');
                                $objPHPExcel->getActiveSheet()->mergeCells('I2:I3');  
                                

#================================================================================ group ng ====================================================

                                $objPHPExcel->getActiveSheet()->insertNewColumnBefore('L', 7);

                                $objPHPExcel->getActiveSheet()->setCellValue('J4',  "PROD. ACTUAL");
                                $objPHPExcel->getActiveSheet()->setCellValue('K4',  "RECEIVE");
                                $objPHPExcel->getActiveSheet()->setCellValue('L2', "=SUBTOTAL(9,L5:L".(count($value)+7).")");
                                $objPHPExcel->getActiveSheet()->setCellValue('J3',  "Ok (pcs)");
                                $objPHPExcel->getActiveSheet()->setCellValue('K3',  "Ok (pcs)");
                                $objPHPExcel->getActiveSheet()->setCellValue('L3',  "Defect (pcs)");

                                $objPHPExcel->getActiveSheet()->setCellValue('S2',  "RM");
                                $objPHPExcel->getActiveSheet()->setCellValue('AF2', "MA");
                                $objPHPExcel->getActiveSheet()->setCellValue('BP2', "PD4");
                                $objPHPExcel->getActiveSheet()->setCellValue('CI2', "PE");
                                $objPHPExcel->getActiveSheet()->setCellValue('CN2', "OTHER");
                                $objPHPExcel->getActiveSheet()->getStyle('I1:L3')->applyFromArray(array('fill' => Style_Fill('C5D9F1')));

                                $objPHPExcel->getActiveSheet()->getStyle('S2')->applyFromArray(array( 'fill' => Style_Fill('99FFCC')));
                                $objPHPExcel->getActiveSheet()->getStyle('AF2')->applyFromArray(array('fill' => Style_Fill('FF99CC')));
                                $objPHPExcel->getActiveSheet()->getStyle('BP2')->applyFromArray(array('fill' => Style_Fill('FFFF99')));
                                $objPHPExcel->getActiveSheet()->getStyle('CI2')->applyFromArray(array('fill' => Style_Fill('E0E0E0')));
                                $objPHPExcel->getActiveSheet()->getStyle('CN2')->applyFromArray(array('fill' => Style_Fill('99FF99')));

                                $objPHPExcel->getActiveSheet()->getStyle('J3:L3')->applyFromArray(array('font' => Style_Font(10, 'FF8000',true)));
                                $objPHPExcel->getActiveSheet()->getStyle('S2:CX2')->applyFromArray(array('font' => Style_Font(28, '000066',true)));


                                $objPHPExcel->getActiveSheet()->mergeCells('S2:'.'AE3');
                                $objPHPExcel->getActiveSheet()->mergeCells('AF2:'.'BO3');  
                                $objPHPExcel->getActiveSheet()->mergeCells('BP2:'.'CH3'); 
                                $objPHPExcel->getActiveSheet()->mergeCells('CI2:'.'CM3');
                                $objPHPExcel->getActiveSheet()->mergeCells('CN2:'.'CX3'); 
                                $objPHPExcel->getActiveSheet()->mergeCells('J1:R1');
                                //$objPHPExcel->getActiveSheet()->mergeCells('L2:CX2');                            
                                //$objPHPExcel->getActiveSheet()->mergeCells('J1:R1');     
#================================================================================ sum ng ======================================================

  

                                                     
                                                                                           
                               



                                $objPHPExcel->getActiveSheet()->setCellValue('M2',  "DEFECT PERCENT ( % )");
                                $objPHPExcel->getActiveSheet()->setCellValue('L4',  "TOTAL DEFECT");
                                $objPHPExcel->getActiveSheet()->setCellValue('M4',  "TOTAL");
                                $objPHPExcel->getActiveSheet()->setCellValue('N4',  "RM");
                                $objPHPExcel->getActiveSheet()->setCellValue('O4',  "MA");
                                $objPHPExcel->getActiveSheet()->setCellValue('P4',  "PD4");
                                $objPHPExcel->getActiveSheet()->setCellValue('Q4',  "PE");
                                $objPHPExcel->getActiveSheet()->setCellValue('R4',  "OTHER");
                                $objPHPExcel->getActiveSheet()
                                            ->getStyle('J2:CX3')
                                            ->getAlignment()
                                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

                                $objPHPExcel->getActiveSheet()->getStyle('M2')->applyFromArray(array( 'fill' => Style_Fill('663300')));


                                $objPHPExcel->getActiveSheet()->getStyle('M2')->applyFromArray(array('font' => Style_Font(16, 'FFFFFF',true)));
                                $objPHPExcel->getActiveSheet()->mergeCells('M2:R3');




                                foreach (range('L', 'R') as $colum ) 
                                {
                                    for ($i=5; $i < (count($value)+5) ; $i++) 
                                    { 
                                        if($colum == 'L')
                                        $objPHPExcel->getActiveSheet()->setCellValue($colum . $i ,  '=SUM(S'.$i . ':CX' . $i . ')');
                                        if($colum == 'M')
                                        $objPHPExcel->getActiveSheet()->setCellValue($colum . $i ,  '=IF($J' . $i . '="-",(IF($K'. $i . '="-","-",(L' . $i . '*100)/$K'. $i . ')),(L' . $i .'*100)/$J' . $i . ' )');
                                        if($colum == 'N')
                                        $objPHPExcel->getActiveSheet()->setCellValue($colum . $i ,  '=IF($J' . $i . '="-",(IF($K'. $i . '="-","-",(SUM(S' .  $i . ':AE' . $i . ')*100)/$K'. $i . ')),(SUM(S' . $i . ':AE' . $i . ')*100)/$J'. $i . ')');
                                        if($colum == 'O')
                                        $objPHPExcel->getActiveSheet()->setCellValue($colum . $i ,  '=IF($J' . $i . '="-",(IF($K'. $i . '="-","-",(SUM(AF' . $i . ':BO' . $i . ')*100)/$K'. $i . ')),(SUM(AF' . $i . ':BO' . $i . ')*100)/$J'. $i . ')');
                                        if($colum == 'P')
                                        $objPHPExcel->getActiveSheet()->setCellValue($colum . $i ,  '=IF($J' . $i . '="-",(IF($K'. $i . '="-","-",(SUM(BP' . $i . ':CH' . $i . ')*100)/$K'. $i . ')),(SUM(BP' . $i . ':CH' . $i . ')*100)/$J'. $i . ')');
                                        if($colum == 'Q')
                                        $objPHPExcel->getActiveSheet()->setCellValue($colum . $i ,  '=IF($J' . $i . '="-",(IF($K'. $i . '="-","-",(SUM(CI' . $i . ':CM' . $i . ')*100)/$K'. $i . ')),(SUM(CI' . $i . ':CM' . $i . ')*100)/$J'. $i . ')');
                                        if($colum == 'R')
                                        $objPHPExcel->getActiveSheet()->setCellValue($colum . $i ,  '=IF($J' . $i . '="-",(IF($K'. $i . '="-","-",(SUM(CN' . $i . ':CX' . $i . ')*100)/$K'. $i . ')),(SUM(CN' . $i . ':CX' . $i . ')*100)/$J'. $i . ')');                                    
                                    }
                                    $objPHPExcel->getActiveSheet()->getStyle($colum . '5:' . $colum . (count($value)+5))->applyFromArray(array( 'font' => Style_Font(15, 'FF0000', true)));
                                    if ($colum == 'L') 
                                    $objPHPExcel->getActiveSheet()->getStyle($colum . '5:' . $colum . (count($value)+5))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_* "-"_-;_-@_-');
                                    else
                                    $objPHPExcel->getActiveSheet()->getStyle($colum . '5:' . $colum . (count($value)+5))->getNumberFormat()->setFormatCode('_-* #,##0.00_-;[RED](#,##0)_-;_* "-"_-;_-@_-');
                                }


                    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('8');
                    foreach (range('B', 'E') as $c) 
                            $objPHPExcel->getActiveSheet()->getColumnDimension($c)->setWidth('10');
                    
                    $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('60');
                    $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('20');
                    $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('35');
                    $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('28');
                    foreach (range('J', 'L') as $c)
                            $objPHPExcel->getActiveSheet()->getColumnDimension($c)->setWidth('20');
                    foreach (range('M', 'R') as $c)
                            $objPHPExcel->getActiveSheet()->getColumnDimension($c)->setWidth('12');
                    foreach (range('S', 'Z') as $c)
                            $objPHPExcel->getActiveSheet()->getColumnDimension($c)->setWidth('12');
                    foreach (range('A', 'Z') as $c)
                            $objPHPExcel->getActiveSheet()->getColumnDimension('A'.$c)->setWidth('12');                        
                    foreach (range('A', 'Z') as $c)
                            $objPHPExcel->getActiveSheet()->getColumnDimension('B'.$c)->setWidth('12');
                    foreach (range('A', 'Z') as $c)
                            $objPHPExcel->getActiveSheet()->getColumnDimension('C'.$c)->setWidth('12');   
            $objPHPExcel->getActiveSheet()->getStyle('S4:CX4')->getNumberFormat()->setFormatCode('000');                                                 
                    // $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('65');
                    // $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('65');


                   
          //$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80);  
          $objPHPExcel->getActiveSheet()->setAutoFilter('A4:R4');
          $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setVisible(false);
          $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setVisible(false);
          //$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setVisible(false);

          // $objPHPExcel->getActiveSheet()->getColumnDimension('M')->setVisible(false);
          // $objPHPExcel->getActiveSheet()->getColumnDimension('N')->setVisible(false);
          // $objPHPExcel->getActiveSheet()->getColumnDimension('O')->setVisible(false);
          // $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setVisible(false);
          // $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setVisible(false);
          // $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setVisible(false);
          $objPHPExcel->getActiveSheet()->freezePane('A5');
          //$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setVisible(false);
         // $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setVisible(false);      
          //$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setVisible(false); 
          $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(30,'16365C',true)));
                                //$objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
          $objPHPExcel->getActiveSheet()->getStyle('A3')->applyFromArray(array('font' => Style_Font(22,'16365C',true)));
                                //$objPHPExcel->getActiveSheet()->getStyle('A3')->applyFromArray(array('fill' => Style_Fill('FDE9D9'))); 
                   }
         }   
    $indSheet++;
    //echo $indSheet; exit;
    $row = 2;
    
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

 
