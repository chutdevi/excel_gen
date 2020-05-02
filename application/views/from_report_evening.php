<?php
error_reporting(E_ALL);
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
    $count_data  =  count($list_act_report[$sheetIndex]) + 5;
    if ($count_data - 5  > 0) 
    {      
#========================================================================================================================  Put field ====================================================================================        
            $count_index =  count($list_act_report[$sheetIndex][0]) - 1 ;
            $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 28 );
            $objPHPExcel->getActiveSheet()->getRowDimension( 5 )->setRowHeight( 12 );
            $objPHPExcel->getActiveSheet()
                ->getStyle('1:4')
                ->getAlignment()
                ->setWrapText(true)
                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_BOTTOM)
                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);       
                                   
            $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(75);    
            $objPHPExcel->getActiveSheet()->setAutoFilter('A5:E5');

            $objPHPExcel->getActiveSheet()->getStyle('A2:'.$col_name[$count_index+2]."2")->applyFromArray(array('font' => Style_Font(11,$colhead_font,true,false,'Franklin Gothic Book'))); 
            $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[$count_index+2]."3")->applyFromArray(array('font' => Style_Font(11,$colhead_font,true,false,'Franklin Gothic Book')));          
            $objPHPExcel->getActiveSheet()->getStyle('A4:'.$col_name[$count_index+2]."4")->applyFromArray(array('font' => Style_Font(11,$colhead_font,true,false,'Franklin Gothic Book')));
            $objPHPExcel->getActiveSheet()->getStyle('E4:'.$col_name[$count_index+2]."4")->applyFromArray(array('font' => Style_Font(9,$colhead_font,true,false,'Franklin Gothic Book')));  
            $objPHPExcel->getActiveSheet()->getStyle('A6:'.$col_name[$count_index+2].($count_data+3))->applyFromArray(array('font' => Style_Font(10,'000000',false,false,'Ebrima')));
            $objPHPExcel->getActiveSheet()->getStyle('D6:'.$col_name[$count_index+2].($count_data+3))->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* "-"??_-;_-@_-');            
            //$objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[$count_index].'1')->applyFromArray(array('fill'    => Style_Fill($colhead)));
            //$objPHPExcel->getActiveSheet()->getStyle('A2:'.$col_name[$count_index].'2')->applyFromArray(array('fill'    => Style_Fill($colhead)));
            //$objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[$count_index].'1')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFF99'))));
            foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
            {
                //echo substr($key,0,9)."<hr>"; 
                if(substr($key,0,9) == 'deli_date')
                 {                    
                    if ( substr($key,10,1) == 1) $key = substr($key,9,2)."st";
                elseif ( substr($key,10,1) == 2) $key = substr($key,9,2)."nd";
                elseif ( substr($key,10,1) == 3) $key = substr($key,9,2)."rd";
                else $key = substr($key,9,2)."th";

                //echo $key."<hr>"; 
                    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."4", str_replace("_", " ", $key));
                 }
                else
                 {
                    $key = str_replace("month1_r", "Month_R", $key);
                    $key = str_replace("month2_r", "Month_R", $key);
                    $key = str_replace("month3_r", "Month_R", $key);
                    $key = str_replace("month4_r", "Month_R", $key);
                    $key = str_replace("month5_r", "Month_R", $key);
                    $key = str_replace("stock_qty", "Stock_QTY", $key);
                    //echo $key."<hr>";
                    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."4", str_replace("_", " ", $key));
                 }
                       
            } // exit;     
#========================================================================================================================  Put data ====================================================================================                
    $row = 6;
            foreach ($list_act_report[$sheetIndex] as $key => $value) 
            {               
               $col = 0;
                foreach ($value as $body => $val) 
                {
                        $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);
                        if($val == 3)  $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('###"E00"');
                        if($val != "" && $body == "customer") 
                        {
                            $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->applyFromArray(array('fill' => Style_Fill('c6d7ee')));
                            $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->applyFromArray(array('font' => Style_Font(11,'000000',true,false,'Franklin Gothic Book')));
                        }
                }
                $row++;               
            }
#========================================================================================================================  Put data ==================================================================================== 
           $objPHPExcel->getActiveSheet()->removeColumn("A", 1);
           $objPHPExcel->getActiveSheet()->removeColumn("AJ", 1);
           $objPHPExcel->getActiveSheet()->removeColumn("AY", 3);        
#========================================================================================================================  Group field ====================================================================================
           $month1 = date('y-M');
           $month2 = ((date('m')+1) > 12 ) ? date('y-M',strtotime((date('Y')+1)."-".((date('m')+1)-12)."-".(1))) : date('y-M',strtotime((date('Y')+0)."-".(date('m')+1)."-".(1))) ;
           $month3 = ((date('m')+2) > 12 ) ? date('y-M',strtotime((date('Y')+1)."-".((date('m')+2)-12)."-".(1))) : date('y-M',strtotime((date('Y')+0)."-".(date('m')+2)."-".(1))) ;
           $month4 = ((date('m')+3) > 12 ) ? date('y-M',strtotime((date('Y')+1)."-".((date('m')+3)-12)."-".(1))) : date('y-M',strtotime((date('Y')+0)."-".(date('m')+3)."-".(1))) ;
           $month5 = ((date('m')+4) > 12 ) ? date('y-M',strtotime((date('Y')+1)."-".((date('m')+4)-12)."-".(1))) : date('y-M',strtotime((date('Y')+0)."-".(date('m')+4)."-".(1))) ;
           $month_start = date('F Y',strtotime((date('Y')+0)."-".(date('m')+0)."-".(date('d')+0))) ;
           $month_end   = ((date('m')+4) > 12 ) ? date('F Y',strtotime((date('Y')+1)."-".((date('m')+4)-12)."-".(date('d')+0))) : date('F Y',strtotime((date('Y')+0)."-".(date('m')+4)."-".(date('d')+0))) ;                      
           //echo $month1."<hr>".$month2."<hr>".$month3."<hr>".$month4."<hr>".$month5."<hr>".((date('m')+3)-12)."<hr>".date('y-M-d',strtotime((date('Y')+1)."-".((date('m')+3)-12)."-".(date('d')-1))); exit;
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('E',  1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('AK', 2);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('AN', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('AP', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('AR', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('AT', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('AV', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('AX', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('AZ', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('BB', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('BD', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('BF', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('BH', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('BJ', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('BL', 1);
           $objPHPExcel->getActiveSheet()->insertNewColumnBefore('BN', 1);

           $objPHPExcel->getActiveSheet()->setCellValue('A1',  'CUSTOMER DEMAND INFORMATION');
           $objPHPExcel->getActiveSheet()->setCellValue('F1',  'Daily Delivery Plan');
           $objPHPExcel->getActiveSheet()->setCellValue('F2',   date('F Y'));
           $objPHPExcel->getActiveSheet()->setCellValue('AM1',  'CUSTOMER DEMAND');
           $objPHPExcel->getActiveSheet()->setCellValue('AM2',  $month_start . ' - '. $month_end);

           $objPHPExcel->getActiveSheet()->setCellValue('D2',  'Ex/JA');
           $objPHPExcel->getActiveSheet()->setCellValue('E2',  'Stock');
           $objPHPExcel->getActiveSheet()->setCellValue('E3',  'Level');
           $objPHPExcel->getActiveSheet()->setCellValue('E4',  '[Day]');
           $objPHPExcel->getActiveSheet()->setCellValue('AK4', 'Accum.');
           $objPHPExcel->getActiveSheet()->setCellValue('AL3', '[ % ]');
           $objPHPExcel->getActiveSheet()->setCellValue('AL4', 'Progress');
           $objPHPExcel->getActiveSheet()->setCellValue('AN4', 'Daily');           
           $objPHPExcel->getActiveSheet()->setCellValue('AP4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('AR4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('AT4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('AV4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('AX4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('AZ4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('BB4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('BD4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('BF4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('BH4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('BJ4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('BL4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('BN4', 'Daily');
           $objPHPExcel->getActiveSheet()->setCellValue('BP4', 'Daily');

           $objPHPExcel->getActiveSheet()->setCellValue('AM3', $month1);           
           $objPHPExcel->getActiveSheet()->setCellValue('AO3', $month1);
           $objPHPExcel->getActiveSheet()->setCellValue('AQ3', $month1);
           $objPHPExcel->getActiveSheet()->setCellValue('AS3', $month2);
           $objPHPExcel->getActiveSheet()->setCellValue('AU3', $month2);
           $objPHPExcel->getActiveSheet()->setCellValue('AW3', $month2);
           $objPHPExcel->getActiveSheet()->setCellValue('AY3', $month3);
           $objPHPExcel->getActiveSheet()->setCellValue('BA3', $month3);
           $objPHPExcel->getActiveSheet()->setCellValue('BC3', $month3);
           $objPHPExcel->getActiveSheet()->setCellValue('BE3', $month4);
           $objPHPExcel->getActiveSheet()->setCellValue('BG3', $month4);
           $objPHPExcel->getActiveSheet()->setCellValue('BI3', $month4);
           $objPHPExcel->getActiveSheet()->setCellValue('BK3', $month5);
           $objPHPExcel->getActiveSheet()->setCellValue('BM3', $month5);
           $objPHPExcel->getActiveSheet()->setCellValue('BO3', $month5);

           $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(18,$colhead_font,true,true,'Franklin Gothic Book')));
           $objPHPExcel->getActiveSheet()->getStyle('F1')->applyFromArray(array('font' => Style_Font(14,$colhead_font,true,true,'Franklin Gothic Book')));
           $objPHPExcel->getActiveSheet()->getStyle('F2')->applyFromArray(array('font' => Style_Font(14,$colhead_font,true,true,'Franklin Gothic Book')));

           $objPHPExcel->getActiveSheet()->getStyle('AM1')->applyFromArray(array('font' => Style_Font(16,$colhead_font,true,true,'Franklin Gothic Book')));
           $objPHPExcel->getActiveSheet()->getStyle('AM2')->applyFromArray(array('font' => Style_Font(14,$colhead_font,true,true,'Franklin Gothic Book')));

           $objPHPExcel->getActiveSheet()->setCellValue('A2', 'Customer');
           $objPHPExcel->getActiveSheet()->setCellValue('B2', 'Model');
           $objPHPExcel->getActiveSheet()->setCellValue('C2', 'Ref');           

           $objPHPExcel->getActiveSheet()->mergeCells('A1:'.'E1');
           $objPHPExcel->getActiveSheet()->mergeCells('F1:'.'AL1');
           $objPHPExcel->getActiveSheet()->mergeCells('F2:'.'AL2');
           $objPHPExcel->getActiveSheet()->mergeCells('AM1:'.'BP1');
           $objPHPExcel->getActiveSheet()->mergeCells('AM2:'.'BP2');

           $objPHPExcel->getActiveSheet()->mergeCells('A2:'.'A4');
           $objPHPExcel->getActiveSheet()->mergeCells('B2:'.'B4');
           $objPHPExcel->getActiveSheet()->mergeCells('C2:'.'C4');

           $objPHPExcel->getActiveSheet()->mergeCells('AM3:'.'AN3');
           $objPHPExcel->getActiveSheet()->mergeCells('AO3:'.'AP3');
           $objPHPExcel->getActiveSheet()->mergeCells('AQ3:'.'AR3');

           $objPHPExcel->getActiveSheet()->mergeCells('AS3:'.'AT3');
           $objPHPExcel->getActiveSheet()->mergeCells('AU3:'.'AV3');
           $objPHPExcel->getActiveSheet()->mergeCells('AW3:'.'AX3');

           $objPHPExcel->getActiveSheet()->mergeCells('AY3:'.'AZ3');
           $objPHPExcel->getActiveSheet()->mergeCells('BA3:'.'BB3');
           $objPHPExcel->getActiveSheet()->mergeCells('BC3:'.'BD3');       

           $objPHPExcel->getActiveSheet()->mergeCells('BE3:'.'BF3');
           $objPHPExcel->getActiveSheet()->mergeCells('BG3:'.'BH3');
           $objPHPExcel->getActiveSheet()->mergeCells('BI3:'.'BJ3');      

           $objPHPExcel->getActiveSheet()->mergeCells('BK3:'.'BL3');
           $objPHPExcel->getActiveSheet()->mergeCells('BM3:'.'BN3');
           $objPHPExcel->getActiveSheet()->mergeCells('BO3:'.'BP3');     

           $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 35 ); 
           $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 35 );  

           foreach(range(0,1)  as $columnID) $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$columnID])->setWidth('21');
           foreach(range(2,2)  as $columnID) $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$columnID])->setWidth('36');           
           foreach(range(5,67) as $columnID) $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$columnID])->setWidth('11');                                  
#========================================================================================================================  Formula data ===================================================================================
foreach (range(6, $count_data) as $rw) 
{
           $objPHPExcel->getActiveSheet()->setCellValue('AK'.$rw,'=SUM(F'.$rw.":".'AJ'.$rw.')'); #Accum.
           $objPHPExcel->getActiveSheet()->setCellValue('AL'.$rw,'=IF(AM'.$rw.'=0,0,'.'(AK'.$rw."/".'AM'.$rw.') )' );   #progess.

           $objPHPExcel->getActiveSheet()->setCellValue('AN'.$rw,'=(AM'.$rw.'/'.$wd['m1'].')');           
           $objPHPExcel->getActiveSheet()->setCellValue('AP'.$rw,'=(AO'.$rw.'/'.$wd['m1'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('AR'.$rw,'=(AQ'.$rw.'/'.$wd['m1'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('AT'.$rw,'=(AS'.$rw.'/'.$wd['m2'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('AV'.$rw,'=(AU'.$rw.'/'.$wd['m2'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('AX'.$rw,'=(AW'.$rw.'/'.$wd['m2'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('AZ'.$rw,'=(AY'.$rw.'/'.$wd['m3'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('BB'.$rw,'=(BA'.$rw.'/'.$wd['m3'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('BD'.$rw,'=(BC'.$rw.'/'.$wd['m3'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('BF'.$rw,'=(BE'.$rw.'/'.$wd['m4'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('BH'.$rw,'=(BG'.$rw.'/'.$wd['m4'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('BJ'.$rw,'=(BI'.$rw.'/'.$wd['m4'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('BL'.$rw,'=(BK'.$rw.'/'.$wd['m5'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('BN'.$rw,'=(BM'.$rw.'/'.$wd['m5'].')');
           $objPHPExcel->getActiveSheet()->setCellValue('BP'.$rw,'=(BO'.$rw.'/'.$wd['m5'].')'); 

           $objPHPExcel->getActiveSheet()->setCellValue('E'.$rw,'=IF(AN'.$rw.'=0,0,'.'(D'.$rw."/".'AN'.$rw.') )' );
}
$objPHPExcel->getActiveSheet()->getStyle('AL6:'.'AL'.($count_data+3))->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');
#========================================================================================================================  Group data ====================================================================================
               foreach (range(5, (date('d')+3) ) as $index) Style_group_lv1($col_name, $index, $objPHPExcel);             # code...
               if(date('d') < 31) foreach (range((date('d')+5), 35) as $index) Style_group_lv1($col_name, $index, $objPHPExcel); 
               foreach (range(40, 43) as $index) Style_group_lv1($col_name, $index, $objPHPExcel);
               foreach (range(46, 49) as $index) Style_group_lv1($col_name, $index, $objPHPExcel); 
               foreach (range(52, 55) as $index) Style_group_lv1($col_name, $index, $objPHPExcel);
               foreach (range(58, 61) as $index) Style_group_lv1($col_name, $index, $objPHPExcel);                                             
               foreach (range(64, 67) as $index) Style_group_lv1($col_name, $index, $objPHPExcel);               
#==========================================================================================================================   Group data row ================================================================================
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(6,1); 
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(27,1);
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(41,1);
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(54,1);
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(56,1);
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(58,1); 
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(60,1);
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(62,1);
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(64,1);
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(67,1);
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(80,1);
               $objPHPExcel->getActiveSheet()->insertNewRowBefore(82,1);
       
               $objPHPExcel->getActiveSheet()->setCellValue('A6',  'ISUZU');           
               $objPHPExcel->getActiveSheet()->setCellValue('A27', 'MITSUBISHI');
               $objPHPExcel->getActiveSheet()->setCellValue('A41', 'KUBOTA');
               $objPHPExcel->getActiveSheet()->setCellValue('A54', 'KET');
               $objPHPExcel->getActiveSheet()->setCellValue('A56', 'HCTD');
               $objPHPExcel->getActiveSheet()->setCellValue('A58', 'SMT');
               $objPHPExcel->getActiveSheet()->setCellValue('A60', 'SIM');
               $objPHPExcel->getActiveSheet()->setCellValue('A62', 'PROTON');
               $objPHPExcel->getActiveSheet()->setCellValue('A64', 'PERODUA(DATT)');
               $objPHPExcel->getActiveSheet()->setCellValue('A67', 'TBK');
               $objPHPExcel->getActiveSheet()->setCellValue('A80', 'TBK-C');
               $objPHPExcel->getActiveSheet()->setCellValue('A82', 'MTA');
                
$objPHPExcel->getActiveSheet()->getStyle('A6'.":".$col_name[$count_index+13]. '6')->applyFromArray(array( 'font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A27'.":".$col_name[$count_index+13].'27')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A41'.":".$col_name[$count_index+13].'41')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A54'.":".$col_name[$count_index+13].'54')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A56'.":".$col_name[$count_index+13].'56')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A58'.":".$col_name[$count_index+13].'58')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A60'.":".$col_name[$count_index+13].'60')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A62'.":".$col_name[$count_index+13].'62')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A64'.":".$col_name[$count_index+13].'64')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A67'.":".$col_name[$count_index+13].'67')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A80'.":".$col_name[$count_index+13].'80')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));
$objPHPExcel->getActiveSheet()->getStyle('A82'.":".$col_name[$count_index+13].'82')->applyFromArray(array('font' => Style_Font(14,'000000',true,false,'Franklin Gothic Book')));                                                                                
                $objPHPExcel->getActiveSheet()->getStyle('A6' .":".$col_name[$count_index+13].'6')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A27'.":".$col_name[$count_index+13].'27')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A41'.":".$col_name[$count_index+13].'41')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A54'.":".$col_name[$count_index+13].'54')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A56'.":".$col_name[$count_index+13].'56')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A58'.":".$col_name[$count_index+13].'58')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A60'.":".$col_name[$count_index+13].'60')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A62'.":".$col_name[$count_index+13].'62')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A64'.":".$col_name[$count_index+13].'64')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A67'.":".$col_name[$count_index+13].'67')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A80'.":".$col_name[$count_index+13].'80')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));
                $objPHPExcel->getActiveSheet()->getStyle('A82'.":".$col_name[$count_index+13].'82')->applyFromArray(array('fill' => Style_Fill('FDE9D9')));

                $objPHPExcel->getActiveSheet()->getRowDimension( 6 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 27 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 41 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 54 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 56 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 58 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 60 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 62 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 64 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 67 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 80 )->setRowHeight( 20 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 82 )->setRowHeight( 20 );

            $objPHPExcel->getActiveSheet()
                ->getStyle('A6:C'.($count_data+3))
                ->getAlignment()
                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);                  

                
#===============================================================================================   SUMMARY CUSTOMER DEMAND BY PRODUCTS CATEGORY ================================================================================
                $objPHPExcel->getActiveSheet()->setCellValue('A86',  'SUMMARY CUSTOMER DEMAND BY PRODUCTS CATEGORY');
                $objPHPExcel->getActiveSheet()->getStyle('A86')->applyFromArray(array('font' => Style_Font(12,'FFFFFF',true,true,'Franklin Gothic Book')));
                $objPHPExcel->getActiveSheet()->getStyle('A86' . ":" . $col_name[$count_index+13].'86')->applyFromArray(array('fill'  => Style_Fill('0c393f')));
                $objPHPExcel->getActiveSheet()->getStyle('A94' . ":" . $col_name[$count_index+13].'94')->applyFromArray(array('fill'  => Style_Fill('0c393f')));

                $objPHPExcel->getActiveSheet()->setCellValue('A87',  'î');
                $objPHPExcel->getActiveSheet()->setCellValue('A88',  'î');
                $objPHPExcel->getActiveSheet()->setCellValue('A89',  'î');
                $objPHPExcel->getActiveSheet()->setCellValue('A90',  'î');
                $objPHPExcel->getActiveSheet()->setCellValue('A91',  'î');
                $objPHPExcel->getActiveSheet()->setCellValue('A92',  'î');
                $objPHPExcel->getActiveSheet()->setCellValue('A93',  'î');

                $objPHPExcel->getActiveSheet()->setCellValue('B87',  'WATER PUMP');
                $objPHPExcel->getActiveSheet()->setCellValue('B88',  'OIL PUMP'  );
                $objPHPExcel->getActiveSheet()->setCellValue('B89',  'WHEEL CYT' );
                $objPHPExcel->getActiveSheet()->setCellValue('B90',  'FORK SHIFT');
                $objPHPExcel->getActiveSheet()->setCellValue('B91',  'BRAKE'     );
                $objPHPExcel->getActiveSheet()->setCellValue('B92',  'GEAR'      );
                $objPHPExcel->getActiveSheet()->setCellValue('B93',  'BEARING'   );                                                                                                

                $objPHPExcel->getActiveSheet()->setCellValue('D87',  '=D$18 + D$19 + D$23 + D$24 + D$25 + D$29 + D$38 + D$39 + D$40 + D$55 + D$59 + D$66 + D$76 + D$78');
                $objPHPExcel->getActiveSheet()->setCellValue('D88',  '=D$21 + D$22 + D$28 + D$37 + D$61 + D$63 + D$65 + D$73 + D$20');
                $objPHPExcel->getActiveSheet()->setCellValue('D89',  '=D57');
                $objPHPExcel->getActiveSheet()->setCellValue('D90',  '=D$42 + D$43 + D$44 + D$45 + D$46 + D$47 + D$48 + D$49 + D$50 + D$51 + D$52 + D$53');
                $objPHPExcel->getActiveSheet()->setCellValue('D91',  '=D$7 + D$8 + D$9 + D$10');
                $objPHPExcel->getActiveSheet()->setCellValue('D92',  '=D$14 + D$15 + D$16 + D$17 + D$31 + D$32 + D$33 + D$34 + D$68 + D$69 + D$70 + D$71 + D$72');
                $objPHPExcel->getActiveSheet()->setCellValue('D93',  '=D83'); 

                $objPHPExcel->getActiveSheet()->setCellValue('E87',  '=IF( AN$87=0,0,(D$87/AN$87) )');
                $objPHPExcel->getActiveSheet()->setCellValue('E88',  '=IF( AN$88=0,0,(D$88/AN$88) )');
                $objPHPExcel->getActiveSheet()->setCellValue('E89',  '=IF( AN$89=0,0,(D$89/AN$89) )');
                $objPHPExcel->getActiveSheet()->setCellValue('E90',  '=IF( AN$90=0,0,(D$90/AN$90) )');
                $objPHPExcel->getActiveSheet()->setCellValue('E91',  '=IF( AN$91=0,0,(D$91/AN$91) )');
                $objPHPExcel->getActiveSheet()->setCellValue('E92',  '=IF( AN$92=0,0,(D$92/AN$92) )');
                $objPHPExcel->getActiveSheet()->setCellValue('E93',  '=IF( AN$93=0,0,(D$93/AN$93) )');

                foreach (range(5, 35) as $index) 
                {
                    $catag1  = $col_name[$index].'$18' ."+";
                    $catag1 .= $col_name[$index].'$19' ."+";
                    $catag1 .= $col_name[$index].'$23' ."+";
                    $catag1 .= $col_name[$index].'$24' ."+";
                    $catag1 .= $col_name[$index].'$25' ."+";
                    $catag1 .= $col_name[$index].'$29' ."+";
                    $catag1 .= $col_name[$index].'$38' ."+";
                    $catag1 .= $col_name[$index].'$39' ."+";
                    $catag1 .= $col_name[$index].'$40' ."+";
                    $catag1 .= $col_name[$index].'$55' ."+";
                    $catag1 .= $col_name[$index].'$59' ."+";
                    $catag1 .= $col_name[$index].'$66' ."+";
                    $catag1 .= $col_name[$index].'$76' ."+";
                    $catag1 .= $col_name[$index].'$78'     ;

                    $catag2  = $col_name[$index].'$21' ."+";
                    $catag2 .= $col_name[$index].'$22' ."+";
                    $catag2 .= $col_name[$index].'$28' ."+";
                    $catag2 .= $col_name[$index].'$37' ."+";
                    $catag2 .= $col_name[$index].'$61' ."+";
                    $catag2 .= $col_name[$index].'$63' ."+";
                    $catag2 .= $col_name[$index].'$65' ."+";
                    $catag2 .= $col_name[$index].'$73' ."+";
                    $catag2 .= $col_name[$index].'$20'     ;

                    $catag3  = $col_name[$index].'$57';

                    $catag4  = $col_name[$index].'$42' ."+";
                    $catag4 .= $col_name[$index].'$43' ."+";
                    $catag4 .= $col_name[$index].'$44' ."+";
                    $catag4 .= $col_name[$index].'$45' ."+";
                    $catag4 .= $col_name[$index].'$46' ."+";
                    $catag4 .= $col_name[$index].'$47' ."+";
                    $catag4 .= $col_name[$index].'$48' ."+";
                    $catag4 .= $col_name[$index].'$49' ."+";
                    $catag4 .= $col_name[$index].'$50' ."+";
                    $catag4 .= $col_name[$index].'$51' ."+";
                    $catag4 .= $col_name[$index].'$52' ."+";
                    $catag4 .= $col_name[$index].'$53'     ;

                    $catag5  = $col_name[$index].'$7' ."+";
                    $catag5 .= $col_name[$index].'$8' ."+";
                    $catag5 .= $col_name[$index].'$9' ."+";
                    $catag5 .= $col_name[$index].'$10'    ;

                    $catag6  = $col_name[$index].'$14' ."+";
                    $catag6 .= $col_name[$index].'$15' ."+";
                    $catag6 .= $col_name[$index].'$16' ."+";
                    $catag6 .= $col_name[$index].'$17' ."+";
                    $catag6 .= $col_name[$index].'$31' ."+";
                    $catag6 .= $col_name[$index].'$32' ."+";
                    $catag6 .= $col_name[$index].'$33' ."+";
                    $catag6 .= $col_name[$index].'$34' ."+";
                    $catag6 .= $col_name[$index].'$68' ."+";
                    $catag6 .= $col_name[$index].'$69' ."+";
                    $catag6 .= $col_name[$index].'$70' ."+";
                    $catag6 .= $col_name[$index].'$71' ."+";
                    $catag6 .= $col_name[$index].'$72'     ;

                    $catag7 = $col_name[$index].'$83';                    
                    //echo $catag1 . "<hr>" . $catag2 . "<hr>" . $catag3 . "<hr>" . $catag4 . "<hr>" . $catag5 . "<hr>" . $catag6 . "<hr>" . $catag7; exit;
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'87',  '='.$catag1 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'88',  '='.$catag2 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'89',  '='.$catag3 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'90',  '='.$catag4 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'91',  '='.$catag5 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'92',  '='.$catag6 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'93',  '='.$catag7 );                    
                }

                $objPHPExcel->getActiveSheet()->setCellValue('AK87','=SUM(F87'.":".'AJ87'.')');
                $objPHPExcel->getActiveSheet()->setCellValue('AK88','=SUM(F88'.":".'AJ88'.')');
                $objPHPExcel->getActiveSheet()->setCellValue('AK89','=SUM(F89'.":".'AJ89'.')');
                $objPHPExcel->getActiveSheet()->setCellValue('AK90','=SUM(F90'.":".'AJ90'.')');
                $objPHPExcel->getActiveSheet()->setCellValue('AK91','=SUM(F91'.":".'AJ91'.')');
                $objPHPExcel->getActiveSheet()->setCellValue('AK92','=SUM(F92'.":".'AJ92'.')');
                $objPHPExcel->getActiveSheet()->setCellValue('AK93','=SUM(F93'.":".'AJ93'.')');

                $objPHPExcel->getActiveSheet()->setCellValue('AL87','=IF(AM87'.'=0,0,'.'(AK87'."/".'AM87'.') )' );   #progess.
                $objPHPExcel->getActiveSheet()->setCellValue('AL88','=IF(AM88'.'=0,0,'.'(AK88'."/".'AM88'.') )' );   #progess.
                $objPHPExcel->getActiveSheet()->setCellValue('AL89','=IF(AM89'.'=0,0,'.'(AK89'."/".'AM89'.') )' );   #progess.
                $objPHPExcel->getActiveSheet()->setCellValue('AL90','=IF(AM90'.'=0,0,'.'(AK90'."/".'AM90'.') )' );   #progess.
                $objPHPExcel->getActiveSheet()->setCellValue('AL91','=IF(AM91'.'=0,0,'.'(AK91'."/".'AM91'.') )' );   #progess.
                $objPHPExcel->getActiveSheet()->setCellValue('AL92','=IF(AM92'.'=0,0,'.'(AK92'."/".'AM92'.') )' );   #progess.
                $objPHPExcel->getActiveSheet()->setCellValue('AL93','=IF(AM93'.'=0,0,'.'(AK93'."/".'AM93'.') )' );   #progess.               

$demand_col = array( 38,      44,      50,      56,      62,      64,      66 );

                foreach (range(38, 67) as $index) 
                {
                    $catag1  = $col_name[$index].'$18' ."+";
                    $catag1 .= $col_name[$index].'$19' ."+";
                    $catag1 .= $col_name[$index].'$23' ."+";
                    $catag1 .= $col_name[$index].'$24' ."+";
                    $catag1 .= $col_name[$index].'$25' ."+";
                    $catag1 .= $col_name[$index].'$29' ."+";
                    $catag1 .= $col_name[$index].'$38' ."+";
                    $catag1 .= $col_name[$index].'$39' ."+";
                    $catag1 .= $col_name[$index].'$40' ."+";
                    $catag1 .= $col_name[$index].'$55' ."+";
                    $catag1 .= $col_name[$index].'$59' ."+";
                    $catag1 .= $col_name[$index].'$66' ."+";
                    $catag1 .= $col_name[$index].'$76' ."+";
                    $catag1 .= $col_name[$index].'$78'     ;

                    $catag2  = $col_name[$index].'$21' ."+";
                    $catag2 .= $col_name[$index].'$22' ."+";
                    $catag2 .= $col_name[$index].'$28' ."+";
                    $catag2 .= $col_name[$index].'$37' ."+";
                    $catag2 .= $col_name[$index].'$61' ."+";
                    $catag2 .= $col_name[$index].'$63' ."+";
                    $catag2 .= $col_name[$index].'$65' ."+";
                    $catag2 .= $col_name[$index].'$73' ."+";
                    $catag2 .= $col_name[$index].'$20'     ;

                    $catag3  = $col_name[$index].'$57';

                    $catag4  = $col_name[$index].'$42' ."+";
                    $catag4 .= $col_name[$index].'$43' ."+";
                    $catag4 .= $col_name[$index].'$44' ."+";
                    $catag4 .= $col_name[$index].'$45' ."+";
                    $catag4 .= $col_name[$index].'$46' ."+";
                    $catag4 .= $col_name[$index].'$47' ."+";
                    $catag4 .= $col_name[$index].'$48' ."+";
                    $catag4 .= $col_name[$index].'$49' ."+";
                    $catag4 .= $col_name[$index].'$50' ."+";
                    $catag4 .= $col_name[$index].'$51' ."+";
                    $catag4 .= $col_name[$index].'$52' ."+";
                    $catag4 .= $col_name[$index].'$53'     ;

                    $catag5  = $col_name[$index].'$7' ."+";
                    $catag5 .= $col_name[$index].'$8' ."+";
                    $catag5 .= $col_name[$index].'$9' ."+";
                    $catag5 .= $col_name[$index].'$10'    ;

                    $catag6  = $col_name[$index].'$14' ."+";
                    $catag6 .= $col_name[$index].'$15' ."+";
                    $catag6 .= $col_name[$index].'$16' ."+";
                    $catag6 .= $col_name[$index].'$17' ."+";
                    $catag6 .= $col_name[$index].'$31' ."+";
                    $catag6 .= $col_name[$index].'$32' ."+";
                    $catag6 .= $col_name[$index].'$33' ."+";
                    $catag6 .= $col_name[$index].'$34' ."+";
                    $catag6 .= $col_name[$index].'$68' ."+";
                    $catag6 .= $col_name[$index].'$69' ."+";
                    $catag6 .= $col_name[$index].'$70' ."+";
                    $catag6 .= $col_name[$index].'$71' ."+";
                    $catag6 .= $col_name[$index].'$72'     ;

                    $catag7 = $col_name[$index].'$83';                    
                   //echo $catag1 . "<hr>" . $catag2 . "<hr>" . $catag3 . "<hr>" . $catag4 . "<hr>" . $catag5 . "<hr>" . $catag6 . "<hr>" . $catag7; exit;
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'87',  '='.$catag1 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'88',  '='.$catag2 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'89',  '='.$catag3 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'90',  '='.$catag4 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'91',  '='.$catag5 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'92',  '='.$catag6 );
                    $objPHPExcel->getActiveSheet()->setCellValue($col_name[$index].'93',  '='.$catag7 );                    
                }

                $objPHPExcel->getActiveSheet()->getStyle('A87:A93')->applyFromArray(array('font' => Style_Font(18,'000000',false,true,'Wingdings 3')));
                $objPHPExcel->getActiveSheet()->getStyle('B87:B93')->applyFromArray(array('font' => Style_Font(11,'000000',true,true,'Franklin Gothic Book')));
                $objPHPExcel->getActiveSheet()->getStyle('B87:B93')->applyFromArray(array('font' => Style_Font(11,'000000',true,true,'Ebrima')));

                $objPHPExcel->getActiveSheet()->getStyle('D87:' .'AK93')->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');
                $objPHPExcel->getActiveSheet()->getStyle('AL87:'.'AL93')->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');
                $objPHPExcel->getActiveSheet()->getStyle('AM87:'.'BP93')->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');

                $objPHPExcel->getActiveSheet()->mergeCells('B87:'.'C87');
                $objPHPExcel->getActiveSheet()->mergeCells('B88:'.'C88');
                $objPHPExcel->getActiveSheet()->mergeCells('B89:'.'C89');
                $objPHPExcel->getActiveSheet()->mergeCells('B90:'.'C90');
                $objPHPExcel->getActiveSheet()->mergeCells('B91:'.'C91');
                $objPHPExcel->getActiveSheet()->mergeCells('B92:'.'C92');
                $objPHPExcel->getActiveSheet()->mergeCells('B93:'.'C93');

                $objPHPExcel->getActiveSheet()
                    ->getStyle('A87:A93'.($count_data+3))
                    ->getAlignment()
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
#=================================================================================================================================================================================================================================
    } else {
                    $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A1', "No data ".$til.".");
                    $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true,false,'Franklin Gothic Book')));
    }
$ind++;

}

$objPHPExcel->setActiveSheetIndex(0);

$objPHPExcel->removeSheetByIndex(count($title));

$objPHPExcel->getActiveSheet()->getCell('A95')->setValue('Issued by PCS '.date('d-M-Y'));;
$objPHPExcel->getActiveSheet()->getStyle('A95')->applyFromArray(array('font' => Style_Font(11,'000000',false,true,'Comic Sans MS')));    
$objPHPExcel->getActiveSheet()->getStyle('A95')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setSelectedCell('A95');                              
$objPHPExcel->getActiveSheet()->freezePane('A6');                              
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

function Style_Font($size=11, $color='FFFFFF', $bol=false, $ita=false, $fname='Consolas') {

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
function Style_group_lv1($cell=null, $index=0, $objPHPExcel=null)
{
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setOutlineLevel(1);
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setVisible(false);
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setCollapsed(true); 
}
?>