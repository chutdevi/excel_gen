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
$lastmount = substr(date('Y/m/t',strtotime($yearA."/".$monthA."/".$dayA)),8, 2);
//var_dump($list_act_report); exit;
$col_name = array();
$subplan     = array();
$subactual   = array();
$subdiff     = array();
$subacc_diff = array();
$subng       = array();
$stop = 0;
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, $cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "A".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "B".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "C".$cm); }

//$ct = array( "", "", "", "", "", "", "", "", "" );
$ct = array( array(), array(), array(), array(), array(), array(), array(), array(), array() ); 
//var_dump($list_act_report); exit;

$ind = 0;
$i=2;
foreach ($title as $inTil => $til) 
{
             $objPHPExcel->createSheet();
             $objPHPExcel->setActiveSheetIndex($inTil);
             //$objPHPExcel->setActiveSheetIndex(0);

            $sheetIndex  =  strtolower(str_replace(' ', '_', $title[$inTil])); 
            $count_index = 0;
            $i = 2;   
            // $ind = 0;
            $count_data  =  count($list_act_report[$sheetIndex]);

    if ($count_data > 0) 
    {      
#========================================================================================================================  Put field ====================================================================================    
          if( $sheetIndex == 'fluctuation')
          {    

                $objPHPExcel->getActiveSheet()->setTitle( "$til"  );
                $objPHPExcel->getActiveSheet()->setShowGridlines(False);
                $st_col = 5;
                $st_dat = 8;
                $sub = 9;
                $count_index =  count($list_act_report[$sheetIndex][0])-1 ;
                $row = $st_dat;
                $count_data  =  count($list_act_report[$sheetIndex]) + $row-1;
                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 10 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 8 );
                foreach(range(3, 5) as $r)
                $objPHPExcel->getActiveSheet()->getRowDimension( $r )->setRowHeight( 26 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 6  )->setRowHeight( 5 ); 
                $objPHPExcel->getActiveSheet()->getRowDimension( 7  )->setRowHeight( 10 ); 
                foreach(range(8, $count_data) as $r)               
                $objPHPExcel->getActiveSheet()->getRowDimension( $r )->setRowHeight( 20 );
                // $objPHPExcel->getActiveSheet()->getRowDimension( 12 )->setRowHeight( 10 ); 
                // $objPHPExcel->getActiveSheet()->getRowDimension( 13 )->setRowHeight( 10 ); 

                $objPHPExcel->getActiveSheet()->freezePane('J'.$row);   
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(90);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('D'.($st_dat-1) . ":" . 'E'.($st_dat-1) );   

				$objPHPExcel->getActiveSheet()->setCellValue('C3',  "ORDER FLUCTUATE" );  

				$objPHPExcel->getActiveSheet()->setCellValue('K3',  date('d F Y', strtotime( date('y') ."-". date('m') . "-" . (date('d')-1) ) ) );  
				$objPHPExcel->getActiveSheet()->setCellValue('K4',  "Previous Daily" );                         

				$objPHPExcel->getActiveSheet()->setCellValue('N3',  date('d F Y', strtotime( date('y') ."-". date('m') . "-" . (date('d')-0) ) ) );  
				$objPHPExcel->getActiveSheet()->setCellValue('N4',  "Current Daily" );      

				$objPHPExcel->getActiveSheet()->setCellValue('Q3',  "Fluctuation" );  
				$objPHPExcel->getActiveSheet()->setCellValue('Q4',  "DIFF" );

				$objPHPExcel->getActiveSheet()->setCellValue('T3',  "Previous Month vs Current Month" );  
				$objPHPExcel->getActiveSheet()->setCellValue('T4',  "DIFF" );
				//$objPHPExcel->getActiveSheet()->setCellValue('U4',  "DIFF" );
                   foreach(array('X', 'AC', 'AH', 'AM') as $c)
                   {
                        
                        $objPHPExcel->getActiveSheet()->setCellValue( $c . '3' ,  'Forecast' );
                        $objPHPExcel->getActiveSheet()->setCellValue( $c . '4' ,  "Forecast in ". date('M', strtotime( date('y') ."-". (date('m')-1) . "-" . (date('d')-0) ) ) ); 
                        $objPHPExcel->getActiveSheet()->mergeCells( $c . '4'.':'. $c . '5');
                   }  
  


				$objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray(array('font' => Style_Font(22,"FFFFFF",false,false)));

				$objPHPExcel->getActiveSheet()->getStyle('C'.$st_dat.':'.'C'.$count_data)->applyFromArray(array('font' => Style_Font(14,"FFFFFF",true,false)));


				$objPHPExcel->getActiveSheet()->getStyle('K3'.':'.$col_name[$count_index].'3')->applyFromArray(array('font' => Style_Font(11,"FFFFFF",true,false)));
				$objPHPExcel->getActiveSheet()->getStyle('K'.$st_col.':'.$col_name[$count_index].$st_col)->applyFromArray(array('font' => Style_Font(10,"000000",false,false)));


				$objPHPExcel->getActiveSheet()->getStyle('C'.($st_dat-1).':'.'H'.($st_dat-1))->applyFromArray(array('fill' => Style_Fill('1a3365')));   
				$objPHPExcel->getActiveSheet()->getStyle('K'.($st_dat-1).':'.'L'.($st_dat-1))->applyFromArray(array('fill' => Style_Fill('284d00')));
				$objPHPExcel->getActiveSheet()->getStyle('N'.($st_dat-1).':'.'O'.($st_dat-1))->applyFromArray(array('fill' => Style_Fill('000099')));
				$objPHPExcel->getActiveSheet()->getStyle('Q'.($st_dat-1).':'.'R'.($st_dat-1))->applyFromArray(array('fill' => Style_Fill('b30000')));
				$objPHPExcel->getActiveSheet()->getStyle('T'.($st_dat-1).':'.'V'.($st_dat-1))->applyFromArray(array('fill' => Style_Fill('009933')));    

				$objPHPExcel->getActiveSheet()->getStyle('X' .($st_dat-1).':'.'AA'.($st_dat-1))->applyFromArray(array('fill' => Style_Fill('0099ff'))); 
				$objPHPExcel->getActiveSheet()->getStyle('AC'.($st_dat-1).':'.'AF'.($st_dat-1))->applyFromArray(array('fill' => Style_Fill('ff3300'))); 
				$objPHPExcel->getActiveSheet()->getStyle('AH'.($st_dat-1).':'.'AK'.($st_dat-1))->applyFromArray(array('fill' => Style_Fill('333300'))); 
				$objPHPExcel->getActiveSheet()->getStyle('AM'.($st_dat-1).':'.'AP'.($st_dat-1))->applyFromArray(array('fill' => Style_Fill('990033'))); 												

				$objPHPExcel->getActiveSheet()->getStyle('C3'.':'.'H'.'4')->applyFromArray(array('fill' => Style_Fill('1a3365')));
				$objPHPExcel->getActiveSheet()->getStyle('K3'.':'.'L'.'3')->applyFromArray(array('fill' => Style_Fill('284d00')));
				$objPHPExcel->getActiveSheet()->getStyle('N3'.':'.'O'.'3')->applyFromArray(array('fill' => Style_Fill('000099')));
				$objPHPExcel->getActiveSheet()->getStyle('Q3'.':'.'R'.'3')->applyFromArray(array('fill' => Style_Fill('b30000'))); 
				$objPHPExcel->getActiveSheet()->getStyle('T3'.':'.'V'.'3')->applyFromArray(array('fill' => Style_Fill('009933')));

				$objPHPExcel->getActiveSheet()->getStyle('X3' .':'.'AA'.'3')->applyFromArray(array('fill' => Style_Fill('0099ff')));
				$objPHPExcel->getActiveSheet()->getStyle('AC3'.':'.'AF'.'3')->applyFromArray(array('fill' => Style_Fill('ff3300')));
				$objPHPExcel->getActiveSheet()->getStyle('AH3'.':'.'AK'.'3')->applyFromArray(array('fill' => Style_Fill('333300'))); 
				$objPHPExcel->getActiveSheet()->getStyle('AM3'.':'.'AP'.'3')->applyFromArray(array('fill' => Style_Fill('990033')));

        $objPHPExcel->getActiveSheet()->mergeCells('C3:'.'H4');
        $objPHPExcel->getActiveSheet()->mergeCells('K3:'.'L3');
        $objPHPExcel->getActiveSheet()->mergeCells('K4:'.'L4');

        $objPHPExcel->getActiveSheet()->mergeCells('N3:'.'O3');
        $objPHPExcel->getActiveSheet()->mergeCells('N4:'.'O4');

        $objPHPExcel->getActiveSheet()->mergeCells('Q3:'.'R3');
        $objPHPExcel->getActiveSheet()->mergeCells('Q4:'.'R4');       

        $objPHPExcel->getActiveSheet()->mergeCells('T3:'.'V3');
        $objPHPExcel->getActiveSheet()->mergeCells('T4:'.'V4');

        $objPHPExcel->getActiveSheet()->mergeCells('X3:' .'AA3');
        $objPHPExcel->getActiveSheet()->mergeCells('AC3:'.'AF3');
        $objPHPExcel->getActiveSheet()->mergeCells('AH3:'.'AK3');
        $objPHPExcel->getActiveSheet()->mergeCells('AM3:'.'AP3');

        $objPHPExcel->getActiveSheet()->mergeCells('Y4:' .'AA4');
        $objPHPExcel->getActiveSheet()->mergeCells('AD4:'.'AF4');
        $objPHPExcel->getActiveSheet()->mergeCells('AI4:'.'AK4');
        $objPHPExcel->getActiveSheet()->mergeCells('AN4:'.'AP4');

                $objPHPExcel->getActiveSheet()->getStyle('B2:'.$col_name[$count_index+2].($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000023')))); 


                $objPHPExcel->getActiveSheet()->getStyle('C3'.':'.'H'.$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'1a3365'))));

                $objPHPExcel->getActiveSheet()->getStyle('C'.$st_col.':'.'H'.$st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'1a3365'))));

                $objPHPExcel->getActiveSheet()->getStyle('D'.$st_dat.':'.'H'.$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

                $objPHPExcel->getActiveSheet()->getStyle('E'.$st_dat.':'.'G'.$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('vertical'    => Style_border(PHPExcel_Style_Border::BORDER_THIN,'1a3365'))));    

set_head($objPHPExcel,'K', 'L', $st_col, $st_dat, $count_data, '284d00');
set_head($objPHPExcel,'N', 'O', $st_col, $st_dat, $count_data, '000099');
set_head($objPHPExcel,'Q', 'R', $st_col, $st_dat, $count_data, 'b30000');                                           
set_head($objPHPExcel,'T', 'V', $st_col, $st_dat, $count_data, '009933');

$colors = array('0099ff', 'ff3300', '333300', '990033');

set_head1($objPHPExcel,'X',  'AA', $st_col, $st_dat, $count_data, $colors[0]);
set_head1($objPHPExcel,'AC', 'AF', $st_col, $st_dat, $count_data, $colors[1]);
set_head1($objPHPExcel,'AH', 'AK', $st_col, $st_dat, $count_data, $colors[2]);
set_head1($objPHPExcel,'AM', 'AP', $st_col, $st_dat, $count_data, $colors[3]);




 // #=====================================================================================================================================================================             
                $cat = ($count_data+6);   
                $col = 2;
                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {


                    if ($key != "ID" && $key != "IG")
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++].$st_col, str_replace("_", " ", $key));

                    if (substr($key, 0, (strlen($key)-1) ) == 'BANK')
                        { 
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[($i-1)].$st_col, "" );
                        $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($i-1)])->setWidth('2.71');

                        }
                    if ($key == "DAY_1" || $key == "DAY_2")
                    { 
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat),   "=IFERROR(" . $col_name[($i-2)] . ($cat)   ."/" . $wd['m1'] .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+1), "=IFERROR(" . $col_name[($i-2)] . ($cat+1) ."/" . $wd['m1'] .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+2), "=IFERROR(" . $col_name[($i-2)] . ($cat+2) ."/" . $wd['m1'] .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+3), "=IFERROR(" . $col_name[($i-2)] . ($cat+3) ."/" . $wd['m1'] .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+4), "=IFERROR(" . $col_name[($i-2)] . ($cat+4) ."/" . $wd['m1'] .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+5), "=IFERROR(" . $col_name[($i-2)] . ($cat+5) ."/" . $wd['m1'] .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+6), "=IFERROR(" . $col_name[($i-2)] . ($cat+6) ."/" . $wd['m1'] .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+7), "=IFERROR(" . $col_name[($i-2)] . ($cat+7) ."/" . $wd['m1'] .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+8), "=IFERROR(" . $col_name[($i-2)] . ($cat+8) ."/" . $wd['m1'] .",0)");

                      //secho $col_name[ $col ].($cat) ; exit;
                    }
                    elseif ( $key == 'DIFF_2' )
                    {

                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat),   "=IFERROR(" . $col_name[($i-2)] . ($cat)     ."/" . $col_name[($i-8)] . ($cat)   .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+1), "=IFERROR(" . $col_name[($i-2)] . ($cat+1)   ."/" . $col_name[($i-8)] . ($cat+1) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+2), "=IFERROR(" . $col_name[($i-2)] . ($cat+2)   ."/" . $col_name[($i-8)] . ($cat+2) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+3), "=IFERROR(" . $col_name[($i-2)] . ($cat+3)   ."/" . $col_name[($i-8)] . ($cat+3) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+4), "=IFERROR(" . $col_name[($i-2)] . ($cat+4)   ."/" . $col_name[($i-8)] . ($cat+4) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+5), "=IFERROR(" . $col_name[($i-2)] . ($cat+5)   ."/" . $col_name[($i-8)] . ($cat+5) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+6), "=IFERROR(" . $col_name[($i-2)] . ($cat+6)   ."/" . $col_name[($i-8)] . ($cat+6) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+7), "=IFERROR(" . $col_name[($i-2)] . ($cat+7)   ."/" . $col_name[($i-8)] . ($cat+7) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+8), "=IFERROR(" . $col_name[($i-2)] . ($cat+8)   ."/" . $col_name[($i-8)] . ($cat+8) .",0)");

                        $objConditional = new PHPExcel_Style_Conditional();
                        $objConditional->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
                                        ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_GREATERTHAN)
                                        ->addCondition('0.2')
                                        ->getStyle()
                                        ->applyFromArray(
                         array(
                          'font'=>array(
                           'color'=>array('argb'=>'ff0000'),
                           'bold'  => true
                          ),
                          'fill'=>array(
                           'type' =>PHPExcel_Style_Fill::FILL_SOLID,
                           'startcolor' =>array('argb' => 'ffffcc'),
                           'endcolor' =>array('argb' => 'ffff00')
                          )
                         )
                        );
                        $conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle($col_name[ ($i-1) ].($cat).':'.$col_name[ ($i-1) ].($cat+7))->getConditionalStyles();
                        array_push($conditionalStyles,$objConditional);
                        $objPHPExcel->getActiveSheet()->getStyle($col_name[ ($i-1) ].($cat).':'.$col_name[ ($i-1) ].($cat+7))->setConditionalStyles($conditionalStyles);
                    }
                    elseif ( $key == 'DIFF_P' )
                    {

                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat),   "=IFERROR(" . $col_name[($i-2)] . ($cat)     ."/" . $col_name[($i-3)] . ($cat)   .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+1), "=IFERROR(" . $col_name[($i-2)] . ($cat+1)   ."/" . $col_name[($i-3)] . ($cat+1) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+2), "=IFERROR(" . $col_name[($i-2)] . ($cat+2)   ."/" . $col_name[($i-3)] . ($cat+2) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+3), "=IFERROR(" . $col_name[($i-2)] . ($cat+3)   ."/" . $col_name[($i-3)] . ($cat+3) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+4), "=IFERROR(" . $col_name[($i-2)] . ($cat+4)   ."/" . $col_name[($i-3)] . ($cat+4) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+5), "=IFERROR(" . $col_name[($i-2)] . ($cat+5)   ."/" . $col_name[($i-3)] . ($cat+5) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+6), "=IFERROR(" . $col_name[($i-2)] . ($cat+6)   ."/" . $col_name[($i-3)] . ($cat+6) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+7), "=IFERROR(" . $col_name[($i-2)] . ($cat+7)   ."/" . $col_name[($i-3)] . ($cat+7) .",0)");
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[ ($i-1) ].($cat+8), "=IFERROR(" . $col_name[($i-2)] . ($cat+8)   ."/" . $col_name[($i-3)] . ($cat+8) .",0)");
                        $objConditional = new PHPExcel_Style_Conditional();
                        $objConditional->setConditionType(PHPExcel_Style_Conditional::CONDITION_CELLIS)
                                        ->setOperatorType(PHPExcel_Style_Conditional::OPERATOR_LESSTHAN)#OPERATOR_LESSTHAN OPERATOR_GREATERTHAN
                                        ->addCondition('-0.19')
                                        ->getStyle()
                                        ->applyFromArray(
                         array(
                          'font'=>array(
                           'color'=>array('argb'=>'ff0000'),
                           'bold'  => true
                          ),
                          'fill'=>array(
                           'type' =>PHPExcel_Style_Fill::FILL_SOLID,
                           'startcolor' =>array('argb' => 'ffffcc'),
                           'endcolor' =>array('argb' => 'ffff00')
                          )
                         )
                        );
                        $conditionalStyles = $objPHPExcel->getActiveSheet()->getStyle($col_name[ ($i-1) ].($cat).':'.$col_name[ ($i-1) ].($cat+7))->getConditionalStyles();
                        array_push($conditionalStyles,$objConditional);
                        $objPHPExcel->getActiveSheet()->getStyle($col_name[ ($i-1) ].($cat).':'.$col_name[ ($i-1) ].($cat+7))->setConditionalStyles($conditionalStyles);
                    }                                         

                } // exit; 

#========================================================================================================================  Put data ====================================================================================                
                $st=$row;
                $gt=$row;
                $mt=$row;
                $en=0;
                $cat = ($count_data+6);
                $ccp = $list_act_report[$sheetIndex][0]['CUSTOMER'];
                $gcp = $list_act_report[$sheetIndex][0]['GROUP_PART'];
                $mcp = $list_act_report[$sheetIndex][0]['MODEL'];
                //echo $com; exit;
                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {               
                   $col = 2;
                  if( $value['IG'] == 1 ) array_push( $ct[ ($value['IG']-1) ], $row );
                  elseif( $value['IG'] == 2 ) array_push( $ct[ ($value['IG']-1) ], $row );
                  elseif( $value['IG'] == 3 ) array_push( $ct[ ($value['IG']-1) ], $row );
                  elseif( $value['IG'] == 4 ) array_push( $ct[ ($value['IG']-1) ], $row );
                  elseif( $value['IG'] == 5 ) array_push( $ct[ ($value['IG']-1) ], $row );
                  elseif( $value['IG'] == 6 ) array_push( $ct[ ($value['IG']-1) ], $row );
                  elseif( $value['IG'] == 7 ) array_push( $ct[ ($value['IG']-1) ], $row );
                  elseif( $value['IG'] == 8 ) array_push( $ct[ ($value['IG']-1) ], $row );
                  elseif( $value['IG'] == 9 ) array_push( $ct[ ($value['IG']-1) ], $row );



                    foreach ($value as $body => $val) 
                    {

                        if ($body != "ID" && $body != "IG")
                            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$col++].($row), $val);

                          //var_dump($wd); exit;
                        if( $body == 'DAY_1' || $body == 'DAY_2')
                        {
                            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[($col-1)].($row), "=(" . $val ."/" . $wd['m1'] .")");

                            //$objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[($col-1)].($cat), "=(" . $col_name[($col-2)] . ($cat) ."/" . $wd['m1'] .")");
                        }


                        if($val == 3 && $body == 'MODEL')  $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('###"E00"');

                        if($body == 'CUSTOMER')
                        {
                        	if($val != $ccp  || $row  == $count_data  )
                        	{

                        		if( ($row - $st) > 0 ) $objPHPExcel->getActiveSheet()->getStyle('C' . $st . ':' . 'C' . ($row-1))->applyFromArray(array('fill' => Style_Fill('1a3365')));  


                                if( ($row  == $count_data) )
                                {
                                    $objPHPExcel->getActiveSheet()->mergeCells( 'A' . $st . ':' . 'A' . ($row) );
                                    $objPHPExcel->getActiveSheet()->mergeCells( 'C' . $st . ':' . 'C' . ($row) );    

                                    $objPHPExcel->getActiveSheet()->getStyle( 'C'. ($row) )
                                                              ->applyFromArray(array(
                                                                'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'c2c2a3'))));  

                                    $objPHPExcel->getActiveSheet()->getStyle('D'. ($row) .':'. $col_name[7] . ($row) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));    

                                    $objPHPExcel->getActiveSheet()->getStyle('K'. ($row) .':'. 'L' . ($row) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));  

                                    $objPHPExcel->getActiveSheet()->getStyle('N'. ($row) .':'. 'O' . ($row) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                                    $objPHPExcel->getActiveSheet()->getStyle('Q'. ($row) .':'. 'R' . ($row) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                    $objPHPExcel->getActiveSheet()->getStyle('T'. ($row) .':'. 'V' . ($row) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                                    $objPHPExcel->getActiveSheet()->getStyle('X'. ($row) .':'. 'AA' . ($row) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                    $objPHPExcel->getActiveSheet()->getStyle('AC'. ($row) .':'. 'AF' . ($row) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                    $objPHPExcel->getActiveSheet()->getStyle('AH'. ($row) .':'. 'AK' . ($row) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                    $objPHPExcel->getActiveSheet()->getStyle('AM'. ($row) .':'. 'AP' . ($row) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));      

                                }else
                                {
                                    $objPHPExcel->getActiveSheet()->mergeCells( 'A' . $st . ':' . 'A' . ($row-1) );
                                    $objPHPExcel->getActiveSheet()->mergeCells( 'C' . $st . ':' . 'C' . ($row-1) );                                    

                                    $objPHPExcel->getActiveSheet()->getStyle( 'C'. ($row-1) )
                                                              ->applyFromArray(array(
                                                                'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'c2c2a3'))));  

                                    $objPHPExcel->getActiveSheet()->getStyle('D'. ($row-1) .':'. $col_name[7] . ($row-1) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));    

                                    $objPHPExcel->getActiveSheet()->getStyle('K'. ($row-1) .':'. 'L' . ($row-1) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));  

                                    $objPHPExcel->getActiveSheet()->getStyle('N'. ($row-1) .':'. 'O' . ($row-1) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                                    $objPHPExcel->getActiveSheet()->getStyle('Q'. ($row-1) .':'. 'R' . ($row-1) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                    $objPHPExcel->getActiveSheet()->getStyle('T'. ($row-1) .':'. 'V' . ($row-1) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                                    $objPHPExcel->getActiveSheet()->getStyle('X'. ($row-1) .':'. 'AA' . ($row-1) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                    $objPHPExcel->getActiveSheet()->getStyle('AC'. ($row-1) .':'. 'AF' . ($row-1) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                    $objPHPExcel->getActiveSheet()->getStyle('AH'. ($row-1) .':'. 'AK' . ($row-1) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                    $objPHPExcel->getActiveSheet()->getStyle('AM'. ($row-1) .':'. 'AP' . ($row-1) )
                                                                  ->applyFromArray(array(
                                                                    'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));                       
                                }

                        		$ccp = $val;
                                if( !($row == $count_data) ) $st  = $row;


                        	}

                        }


                        if($body == 'GROUP_PART')
                        {
                        	if( $val != $gcp || $row == $count_data)
                        	{


                                    if( $row == $count_data)     $objPHPExcel->getActiveSheet()->mergeCells( 'F' . $gt . ':' . 'F' . ($row) ); 
                                    elseif( ( $row - $gt ) > 0 ) $objPHPExcel->getActiveSheet()->mergeCells( 'F' . $gt . ':' . 'F' . ($row-1) );
                                        
                                    //if( $row == $count_data)
                                        $gcp = $val;     
                                        $gt  = $row;                  

                        	}          

                        }
                        if($body == 'MODEL')
                        {
                        	if( $val != $mcp || $row == $count_data)
                        	{


                                    if( $row == $count_data)     $objPHPExcel->getActiveSheet()->mergeCells( 'G' . $mt . ':' . 'G' . ($row) ); 
                                    elseif( ( $row - $mt ) > 0 ) $objPHPExcel->getActiveSheet()->mergeCells( 'G' . $mt . ':' . 'G' . ($row-1) );
                                        
                                    //if( $row == $count_data)
                                        $mcp = $val;     
                                        $mt  = $row;                  

                        	}          

                        }                        
                        if($body == 'DIFF_1' && ( $val < 0 ) )
                        {
                          if( $value['DIFF_2'] < 0.20  )
                          {
                              $colo = 'ffffcc';
                                $objPHPExcel->getActiveSheet()->getStyle('K'. $row .':'.'L'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('N'. $row .':'.'O'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('Q'. $row .':'.'R'. $row)->applyFromArray(array('fill' => Style_Fill($colo))); 
                                $objPHPExcel->getActiveSheet()->getStyle('K'.$row.':'.'R'.$row)->applyFromArray(array('font' => Style_Font(11,"ff0000",false,false)));
                          }
                          else
                          {
                              $colo = 'ffff00';
                                $objPHPExcel->getActiveSheet()->getStyle('K'. $row .':'.'L'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('N'. $row .':'.'O'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('Q'. $row .':'.'R'. $row)->applyFromArray(array('fill' => Style_Fill($colo))); 
                                $objPHPExcel->getActiveSheet()->getStyle('K'.$row.':'.'R'.$row)->applyFromArray(array('font' => Style_Font(11,"ff0000",true,false)));



                          }

                        }
                        elseif($body == 'DIFF_1' && ( $val > 0 ) )
                        {
                          if( $value['DIFF_2'] < 0.20  )
                          {
                              $colo = 'b3ffb3';
                                $objPHPExcel->getActiveSheet()->getStyle('K'. $row .':'.'L'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('N'. $row .':'.'O'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('Q'. $row .':'.'R'. $row)->applyFromArray(array('fill' => Style_Fill($colo))); 
                                $objPHPExcel->getActiveSheet()->getStyle('K'.$row.':'.'R'.$row)->applyFromArray(array('font' => Style_Font(11,"002db3",false,false)));
                          }
                           else
                          {
                              $colo = '8cff66';
                                $objPHPExcel->getActiveSheet()->getStyle('K'. $row .':'.'L'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('N'. $row .':'.'O'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('Q'. $row .':'.'R'. $row)->applyFromArray(array('fill' => Style_Fill($colo))); 
                                $objPHPExcel->getActiveSheet()->getStyle('K'.$row.':'.'R'.$row)->applyFromArray(array('font' => Style_Font(11,"002db3",true,false)));
                          }  
                        }     



                        if($body == 'DIFF_L' && ( $val < 0 ) )
                        {
                          if( $value['DIFF_P'] < 0.20  )
                          {
                              $colo = 'ffffcc';
                                $objPHPExcel->getActiveSheet()->getStyle('T'. $row .':'.'V'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('T'.$row.':'.'V'.$row)->applyFromArray(array('font' => Style_Font(11,"ff0000",false,false)));
                          }
                          else
                          {
                              $colo = 'ffff00';
                                $objPHPExcel->getActiveSheet()->getStyle('T'. $row .':'.'V'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('T'.$row.':'.'V'.$row)->applyFromArray(array('font' => Style_Font(11,"ff0000",true,false)));
                          }
                        }
                        elseif($body == 'DIFF_L' && ( $val > 0 ) )
                        {
                          if( $value['DIFF_P'] < 0.20  )
                          {
                              $colo = 'b3ffb3';
                               $objPHPExcel->getActiveSheet()->getStyle('T'. $row .':'.'V'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                               $objPHPExcel->getActiveSheet()->getStyle('T'.$row.':'.'V'.$row)->applyFromArray(array('font' => Style_Font(11,"002db3",false,false)));
                          }
                           else
                          {
                              $colo = '8cff66';
                                $objPHPExcel->getActiveSheet()->getStyle('T'. $row .':'.'V'. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle('T'.$row.':'.'V'.$row)->applyFromArray(array('font' => Style_Font(11,"002db3",true,false)));
                          }                         
                        } 

                    }

                    $row++;  

                }
                //cho $st . " " . $row . " " . $count_data; exit;
                // foreach (range(low, high) as $key => $value) {
                //   # code...
                // }
                foreach(array('K', 'N', 'Q', 'T', 'U', 'X' ,'Y', 'Z', 'AA', 'AC', 'AD', 'AE', 'AF', 'AH', 'AI', 'AJ', 'AK', 'AM', 'AN', 'AO', 'AP') as $cel )
                  put_data($objPHPExcel, $ct, $cel, ($count_data+6));
               // var_dump($ct); exit;
                $objPHPExcel->getActiveSheet()->getStyle('C'.($row-1))->applyFromArray(array('fill' => Style_Fill('1a3365')));



// #=====================================================================================================================================================================  

//echo $col; exit;





                //$objPHPExcel->getActiveSheet()->duplicateStyle( $objPHPExcel->getActiveSheet()->getStyle( 'B8' ), ('C8') );
 //echo $til . " = " . $subplan[0] . "<hr>" ;

                   $objPHPExcel->getActiveSheet()->setCellValue('K5',  "Order" );
                   $objPHPExcel->getActiveSheet()->setCellValue('L5',  "Order Use (Day)" );
                   $objPHPExcel->getActiveSheet()->setCellValue('N5',  "Order" );
                   $objPHPExcel->getActiveSheet()->setCellValue('O5',  "Order Use (Day)" );
                   $objPHPExcel->getActiveSheet()->setCellValue('Q5',  "Diff To Day" );
                   $objPHPExcel->getActiveSheet()->setCellValue('R5',  "Diff (%)" );   
                   $objPHPExcel->getActiveSheet()->setCellValue('T5',  "Order ". date('M', strtotime( date('y') ."-". (date('m')-1) . "-" . (date('d')-0) ) ));
                   $objPHPExcel->getActiveSheet()->setCellValue('U5',  "Diff" );      
                   $objPHPExcel->getActiveSheet()->setCellValue('V5',  "Diff (%)" );      
                   $re_mon = 1;
                   foreach(array('Y', 'AD', 'AI', 'AN') as $mon)
                    {
                    	$mn = ($re_mon++);
                        $his_month = date('F-Y', strtotime( "+ $mn month" , strtotime( date('Y') . "-" . date('m') . "-" . '01' ) ) )   ;
                        //echo $mon . '4'; exit;
                        $objPHPExcel->getActiveSheet()->setCellValue( $mon . '4' , $his_month);
                        $objPHPExcel->getActiveSheet()->setCellValue( $mon . '5' , runInsPD( date('d', strtotime( date('y') ."-". (date('m')-0) . "-" . (date('d')-1) ) ) ));
                       // $objPHPExcel->getActiveSheet()->setCellValue( $mon . '5' , "Previous Day");
                    }    
                   foreach(array('Z', 'AE', 'AJ', 'AO') as $c)
                     //   $objPHPExcel->getActiveSheet()->setCellValue( $c . '5' ,   "Current Day");
                        $objPHPExcel->getActiveSheet()->setCellValue( $c . '5' ,   runInsPD( date('d', strtotime( date('y') ."-". (date('m')-0) . "-" . (date('d')-0) ) ) ));

                   foreach(array('AA', 'AF', 'AK', 'AP') as $c)
                        $objPHPExcel->getActiveSheet()->setCellValue( $c . '5' ,   'Diff' );   
                                 

                
                $objPHPExcel->getActiveSheet()->getStyle('C'.$count_data.':'.$col_name[$count_index].$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_NONE,'FFFFFF'))));


                $objPHPExcel->getActiveSheet()->getStyle('K'.$st_dat.':'.$col_name[$count_index].($count_data+20))
                                              ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');


                $objPHPExcel->getActiveSheet()->getStyle('R'.$st_dat.':'.'R'.($count_data+20))
                                              ->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');
                $objPHPExcel->getActiveSheet()->getStyle('V'.$st_dat.':'.'V'.($count_data+20))
                                              ->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');

                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[0])->setWidth('1.71');     #A
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[1])->setWidth('4.71');     #B no
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[2])->setWidth('20.71');
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[3])->setWidth('16.71'); 
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[4])->setWidth('15.71');    #D plnt
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[5])->setWidth('19.71');     #C pd                
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[6])->setWidth('19.71');    #H it_nm
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[7])->setWidth('31.71');    #I model
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[8])->setWidth('1.21');
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[9])->setWidth('1.21');
                $data_row = array(10,11,13,14,16,17,19,20,21,23, 28, 33, 38);

                foreach($data_row as $in)
                	$objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$in])->setWidth('14.71');
                foreach(array(12,15,18) as $in)
                	$objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$in])->setWidth('1.71');    
                foreach(array(42,43) as $in)
                  $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$in])->setWidth('2.71'); 

                $data_row = array(24,25,26,29,30,31,34,35,36,39,40,41);

                foreach($data_row as $in)
                	$objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$in])->setWidth('9.71');

                //echo count($ig); exit;
                category_item( $objPHPExcel, $ig, $col_name, ($count_data+6), (($count_data+5)+count($ig)), $count_index );

                Style_Alignment(('B1'.':'.$col_name[$count_index+6].$st_col), 3, false, $objPHPExcel);
                // Style_Alignment(('J'.'7'.':'.$col_name[$count_index+6].$count_data), 6, false, $objPHPExcel);
                Style_Alignment(('B2'.':'.'H'.$count_data), 3, false, $objPHPExcel);
          
                // }               
				Style_group_Col( $col_name, 3, $objPHPExcel );
				Style_group_Col( $col_name, 4, $objPHPExcel );

    foreach(range($st+1, $count_data) as $ind_r) Style_group_Row( $ind_r, $objPHPExcel);
                
      foreach(range(23, 42) as $lv)
        Style_group_Col( $col_name, $lv, $objPHPExcel,1 , true );



       foreach(range(23, 26) as $lv)
         Style_group_Col( $col_name, $lv, $objPHPExcel ,2, true );
 
       foreach(range(28, 31) as $lv)
         Style_group_Col( $col_name, $lv, $objPHPExcel ,2, true  );

       foreach(range(33, 36) as $lv)
         Style_group_Col( $col_name, $lv, $objPHPExcel ,2, true  );

       foreach(range(38, 41) as $lv)
         Style_group_Col( $col_name, $lv, $objPHPExcel ,2 , true );
    }
elseif ($sheetIndex == 'fluctuation_history') 
  {
    
                //echo "6666"; exit;
                $objPHPExcel->getActiveSheet()->setTitle( "$til"  );
                $objPHPExcel->getActiveSheet()->setShowGridlines(False);
                $st_col = 5;
                $st_dat = 8;
                $sub = 9;
                $count_index =  count($list_act_report[$sheetIndex][0])  - ( (31 - $lastmount) * 2 ) ;
                $row = $st_dat;
                $count_data  =  count($list_act_report[$sheetIndex]) + $row-1;
                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 10 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 10 );
                foreach(range(3, 5) as $r)
                $objPHPExcel->getActiveSheet()->getRowDimension( $r )->setRowHeight( 26 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 6  )->setRowHeight( 5 ); 
                $objPHPExcel->getActiveSheet()->getRowDimension( 7  )->setRowHeight( 10 ); 
                foreach(range(8, $count_data) as $r)               
                $objPHPExcel->getActiveSheet()->getRowDimension( $r )->setRowHeight( 20 );
                // $objPHPExcel->getActiveSheet()->getRowDimension( 12 )->setRowHeight( 10 ); 
                // $objPHPExcel->getActiveSheet()->getRowDimension( 13 )->setRowHeight( 10 ); 

                $objPHPExcel->getActiveSheet()->freezePane('J'.$row);   
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(90);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('D'.($st_dat-1) . ":" . 'E'.($st_dat-1) );  

                $objPHPExcel->getActiveSheet()->setCellValue('C3',  "ORDER FLUCTUATE" );
                $objPHPExcel->getActiveSheet()->setCellValue('K3',  "HISTORY" );
                $objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray(array('font' => Style_Font(22,"FFFFFF",false,false)));
                $objPHPExcel->getActiveSheet()->getStyle('K3')->applyFromArray(array('font' => Style_Font(22,"FFFFFF",false,false)));

                $objPHPExcel->getActiveSheet()->getStyle('C'.$st_dat.':'.'C'.$count_data)->applyFromArray(array('font' => Style_Font(14,"FFFFFF",true,false)));
                $objPHPExcel->getActiveSheet()->getStyle('C'.($st_dat-1).':'.'H'.($st_dat-1))->applyFromArray(array('fill' => Style_Fill('1a3365')));

                $objPHPExcel->getActiveSheet()->getStyle('C3'.':'.'H'.'4')->applyFromArray(array('fill' => Style_Fill('1a3365')));
                $objPHPExcel->getActiveSheet()->getStyle('K3'.':'.$col_name[$count_index].'4')->applyFromArray(array('fill' => Style_Fill('ff4000')));
                $objPHPExcel->getActiveSheet()->getStyle('K'.($st_dat-1).':'.$col_name[$count_index].($st_dat-1))->applyFromArray(array('fill' => Style_Fill('ff4000')));
                $objPHPExcel->getActiveSheet()->mergeCells('C3:'.'H4');
                $objPHPExcel->getActiveSheet()->mergeCells('K3:'.$col_name[$count_index].'4');

                $objPHPExcel->getActiveSheet()->getStyle('B2:'.$col_name[$count_index+2].($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000023')))); 
                $objPHPExcel->getActiveSheet()->getStyle('C3'.':'.'H'.$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'1a3365'))));

                $objPHPExcel->getActiveSheet()->getStyle('C'.$st_col.':'.'H'.$st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'1a3365'))));

                $objPHPExcel->getActiveSheet()->getStyle('D'.$st_dat.':'.'H'.$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

                $objPHPExcel->getActiveSheet()->getStyle('E'.$st_dat.':'.'G'.$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('vertical'    => Style_border(PHPExcel_Style_Border::BORDER_THIN,'1a3365'))));
 #===================================================================================================================================================================== 

                $objPHPExcel->getActiveSheet()->getStyle('K4'.':'.$col_name[$count_index].'4')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'ff4000'))));

                $objPHPExcel->getActiveSheet()->getStyle('K'.$st_col.':'.$col_name[$count_index].$st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'ff4000'))));

                $objPHPExcel->getActiveSheet()->getStyle('K'.$st_dat.':'.$col_name[$count_index].$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

                $objPHPExcel->getActiveSheet()->getStyle('K'.$st_dat.':'.$col_name[$count_index].$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('vertical' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'ff4000'))));

                $objPHPExcel->getActiveSheet()->getStyle('K3'.':'.$col_name[$count_index].$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('outline'  => Style_border(PHPExcel_Style_Border::BORDER_THICK,'ff4000'))));                                              

 #===================================================================================================================================================================== 





                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {
                  //echo substr($key, (strlen($key)-2), 2)  . "<hr>"; 
                  if ( is_numeric(substr($key, (strlen($key)-2), 2)) && substr($key, (strlen($key)-2), 2) > $lastmount ) { $stop = 99; break; }
                  //if ( $stop > 0 ) break;

                    if ($key != "NO")
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++].$st_col, str_replace("_", " ", $key));
                    if ( substr($key, 0, (strlen($key)-2) ) == 'STATUS')
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[($i-1)].$st_col,  runInsPD( substr($key, (strlen($key)-2), 2) ) );
                    if ( substr($key, 0, (strlen($key)-3) ) == 'Reason')
                    {
                        $objPHPExcel->getActiveSheet()->getStyle( $col_name[($i-1)].$st_col )->applyFromArray(array('fill' => Style_Fill('ccddff')));
                        $objPHPExcel->getActiveSheet()->getStyle( $col_name[($i-1)].$st_dat . ":" . $col_name[($i-1)].$count_data )->applyFromArray(array('fill' => Style_Fill('ccddff')));
                    }
                         
                    if (substr($key, 0, (strlen($key)-1) ) == 'B')
                        { 
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[($i-1)].$st_col, "" );
                        $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($i-1)])->setWidth('1.21');

                        }                   
                }  //exit;



#========================================================================================================================  Put data ====================================================================================                
                $st=$row;
                $gt=$row;
                $en=0;
                $stop = 0;
                $ccp = $list_act_report[$sheetIndex][0]['CUSTOMER'];
                $gcp = $list_act_report[$sheetIndex][0]['GROUP_PART'];


                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {               
                   $col = 2;
                   //$st = $row;
                   //echo "33333" . "<hr>";  exit;
                  // $st=2;
                  // $en=0;
                    foreach ($value as $body => $val) 
                    {
   
                        //if ( is_numeric(substr($body, (strlen($body)-2), 2  ) ) && substr( $body, (strlen($body)-3), 3 ) > $lastmount ) break;
                  if ( is_numeric(substr($body, (strlen($body)-2), 2)) && substr($body, (strlen($body)-2), 2) > $lastmount ) { $stop = 99; break; }
                  //if ( $stop > 0 ) break;

                        if ($body != "NO")
                            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$col++].($row), $val);



                        if($val == 3 && $body == 'MODEL')  $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('###"E00"');




                        if($body == 'CUSTOMER')
                        {
                          if( $val != $ccp || $row + 1 >= $count_data )
                          {

                            if( ($row - $st) > 0 )
                            $objPHPExcel->getActiveSheet()->getStyle('C' . $st . ':' . 'C' . ($row-1))->applyFromArray(array('fill' => Style_Fill('1a3365')));  
                            $objPHPExcel->getActiveSheet()->mergeCells( 'A' . $st . ':' . 'A' . ($row-1) );
                            $objPHPExcel->getActiveSheet()->mergeCells( 'C' . $st . ':' . 'C' . ($row-1) );

                              $objPHPExcel->getActiveSheet()->getStyle( 'C'. ($row-1) )
                                                            ->applyFromArray(array(
                                                              'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'c2c2a3'))));

                                $objPHPExcel->getActiveSheet()->getStyle('D'. ($row-1) .':'. $col_name[7] . ($row-1) )
                                                              ->applyFromArray(array(
                                                                'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));    
                            $ccp = $val;
                            $st  = $row;

                          }

                         }
                        if($body == 'GROUP_PART')
                        {
                          if($val != $gcp  )
                          {
                            if( ($row - $gt) > 0 )
                            $objPHPExcel->getActiveSheet()->mergeCells( 'F' . $gt . ':' . 'F' . ($row-1) );
                            $gcp = $val;
                            $gt  = $row;

                          }

                         }
                        if(substr($body, 0, (strlen($body)-2) ) == 'STATUS' && ( $val < 0 ) )
                        {
                          //$objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$st_dat . ":" . $col_name[$col-1].$count_data)->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');
                                      
                          if( $val > -0.20  )
                          {
                              $colo = 'ffffcc';
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1]. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1]. $row.':'.$col_name[$col].$row)->applyFromArray(array('font' => Style_Font(11,"ff0000",false,false)));
                          }
                          else
                          {
                              $colo = 'ffff00';
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1]. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1]. $row.':'.$col_name[$col].$row)->applyFromArray(array('font' => Style_Font(11,"ff0000",true,false)));
                          }

                        }
                        elseif(substr($body, 0, (strlen($body)-2) ) == 'STATUS'  && ( $val > 0 ) )
                        {
                          //$objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$st_dat . ":" . $col_name[$col-1].$count_data)->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');

                          if( $val < 0.20  )
                          {
                              $colo = 'b3ffb3';
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1]. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1]. $row.':'.$col_name[$col].$row)->applyFromArray(array('font' => Style_Font(11,"002db3",false,false)));
                          }
                           else
                          {
                              $colo = '8cff66';
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1]. $row)->applyFromArray(array('fill' => Style_Fill($colo)));
                                $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1]. $row.':'.$col_name[$col].$row)->applyFromArray(array('font' => Style_Font(11,"002db3",true,false)));
                          }       
                        }   
                      //   elseif( substr($body, 0, (strlen($body)-2) ) == 'STATUS'  && ( $val == 0 ) ) 
                      //        $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].$st_dat . ":" . $col_name[$col-1].$count_data)->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-'); 
                       }
                    $row++;               
                  }
                  //exit;
                $objPHPExcel->getActiveSheet()->getStyle('C'.($row-1))->applyFromArray(array('fill' => Style_Fill('1a3365')));

                $objPHPExcel->getActiveSheet()->getStyle( $col_name[ ( (date('d')*2)+8) ].$st_col . ":" . $col_name[ ((date('d')*2)+9) ].$st_col )->applyFromArray(array('fill' => Style_Fill('cc99ff')));

                $objPHPExcel->getActiveSheet()->getStyle($col_name[10].$st_dat . ":" . $col_name[$count_index].$count_data)->getNumberFormat()->setFormatCode('_-* #,##0%_-;[Red](#,##0%)_-;_-* "-"??_-;_-@_-');
                //echo $inTil; exit;
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[0])->setWidth('1.71');     #A
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[1])->setWidth('4.71');     #B no
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[2])->setWidth('20.71');
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[3])->setWidth('16.71'); 
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[4])->setWidth('15.71');    #D plnt
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[5])->setWidth('19.71');     #C pd                
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[6])->setWidth('19.71');    #H it_nm
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[7])->setWidth('31.71');    #I model

                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+1)])->setWidth('2.71');    #H it_nm
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+2)])->setWidth('2.71');    #I model
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+3)])->setWidth('2.71');
                foreach(range(10, $count_index) as $ind)
                    $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$ind])->setWidth('13.71');

                $objPHPExcel->getActiveSheet()->getStyle('C'.$count_data.':'.$col_name[$count_index].$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_NONE,'FFFFFF'))));



                Style_Alignment(('B1'.':'.$col_name[$count_index+6].$st_col), 3, false, $objPHPExcel);
                // Style_Alignment(('J'.'7'.':'.$col_name[$count_index+6].$count_data), 6, false, $objPHPExcel);
                Style_Alignment(('B2'.':'.'H'.$count_data), 3, false, $objPHPExcel);    


        Style_group_Col( $col_name, 3, $objPHPExcel );
        Style_group_Col( $col_name, 4, $objPHPExcel );

        $da = ( (date('d')+0) == 1 ) ? 1 :  (date('d')-0); 



        if ($da == 1 )
         foreach(range(12, $count_index) as $id ) 
             Style_group_Col( $col_name, $id, $objPHPExcel ); 
        elseif( $da == ($lastmount) )
         foreach( range(10, ($count_index-2) ) as $id ) 
             Style_group_Col( $col_name, $id, $objPHPExcel );            
        else
        {
         foreach( range(10, ($da*2+7) ) as $id ) 
             Style_group_Col( $col_name, $id, $objPHPExcel );     
         $da = $da+2;   
         foreach( range( ($da*2+6), $count_index ) as $id ) 
             Style_group_Col( $col_name, $id, $objPHPExcel );                    
        }

    // # code...
       }
     
//echo "123123"; exit;




#========================================================================================================================  Put data ====================================================================================         
    } else {
                    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('A1', "No data ".$til.".");
                    $objPHPExcel->getActiveSheet()->setTitle( "$til"  );
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

function Style_group_Col($cell=null, $index=0, $objPHPExcel=null, $level=1, $vi=false, $co=true)
{
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setOutlineLevel($level);
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setVisible($vi);
    $objPHPExcel->getActiveSheet()->getColumnDimension ($cell[$index])->setCollapsed($co); 
}
function Style_group_Row($index=0, $objPHPExcel=null, $vi=false, $co=true)
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

function runInsPD($dat = 1){
  //$d = ( $dat < 10  ) ? "0" . $dat : $dat;
  $d = $dat;
    if ( substr($d, 1, strlen($d)-1) == "1" && $d != "11" ) $d =  $d . "st";
    else if  ( substr($d, 1, strlen($d)-1) == "2" && $d != "12" ) $d =  $d . "nd";
    else if  ( substr($d, 1, strlen($d)-1) == "3" && $d != "13" ) $d =  $d . "rd";
    else $d =  $d . "th";
   return $d;
}


function set_head($objPHPExcel=null, $ind_st, $ind_en, $st_col, $st_dat, $count_data, $cl)
{
                  $objPHPExcel->getActiveSheet()->getStyle( $ind_st . '4'.':'. $ind_en . '4')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK, $cl ))));

                $objPHPExcel->getActiveSheet()->getStyle( $ind_st .$st_col . ':' . $ind_en .$st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN, $cl ))));

                $objPHPExcel->getActiveSheet()->getStyle( $ind_st . $st_dat . ':' . $ind_en .$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN, 'a6a6a6' ))));

                $objPHPExcel->getActiveSheet()->getStyle( $ind_st . $st_dat . ':' . $ind_en .$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('vertical'    => Style_border(PHPExcel_Style_Border::BORDER_THIN, $cl ))));

                $objPHPExcel->getActiveSheet()->getStyle( $ind_st . '3'. ':' . $ind_en .$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('outline'  => Style_border(PHPExcel_Style_Border::BORDER_THICK, $cl ))));

}

function set_head1($objPHPExcel=null, $ind_st, $ind_en, $st_col, $st_dat, $count_data, $cl)
{
                $objPHPExcel->getActiveSheet()->getStyle( $ind_st . ($st_col-1).':'. $ind_en . ($st_col))
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN, $cl))));

                $objPHPExcel->getActiveSheet()->getStyle( $ind_st . $st_dat.':'. $ind_en . $count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN, 'a6a6a6'))));

                $objPHPExcel->getActiveSheet()->getStyle( $ind_st . $st_dat.':'. $ind_en . $count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('vertical' => Style_border(PHPExcel_Style_Border::BORDER_THIN, $cl))));

                $objPHPExcel->getActiveSheet()->getStyle( $ind_st . '3' . ':'. $ind_en . $count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('outline'  => Style_border(PHPExcel_Style_Border::BORDER_THICK, $cl))));                                              
                $objPHPExcel->getActiveSheet()->getStyle( $ind_st . '4' . ':'. $ind_en . '4')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK, $cl))));

}

function set_head2($objPHPExcel=null, $ind_st, $ind_en, $st_dat, $count_data, $cl)
{

                $objPHPExcel->getActiveSheet()->getStyle($ind_st.($st_dat-2).':'.$ind_en.($st_dat-2))->applyFromArray(array('fill' => Style_Fill($cl)));


                $objPHPExcel->getActiveSheet()->getStyle( $ind_st. $st_dat .':'.$ind_en. $count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

                $objPHPExcel->getActiveSheet()->getStyle( $ind_st. $st_dat .':'.$ind_en. $st_dat)
                                              ->applyFromArray(array(
                                                'borders' => array('top'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

                $objPHPExcel->getActiveSheet()->getStyle( $ind_st . $st_dat.':'. $ind_en . $count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('vertical' => Style_border(PHPExcel_Style_Border::BORDER_THIN, $cl))));
                $objPHPExcel->getActiveSheet()->getStyle( $ind_st.($st_dat-2).':'.$ind_en. $count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,$cl))));
}

function category_item( $objPHPExcel=null, $dat, $col_name, $st_data, $count_data, $count_index )
{
               // echo ($st_data-4) . " = " . $col_name[$count_index+2] . " -> " . ( $count_data+1 ); exit;
                $row = $st_data;
                $objPHPExcel->getActiveSheet()->getStyle(  'B' . ($st_data-3) . ':' . $col_name[$count_index+2] . ( $count_data+1) )
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000023'))));
                $objPHPExcel->getActiveSheet()->getRowDimension( ($st_data-1) )->setRowHeight( 5 );

                $objPHPExcel->getActiveSheet()->setCellValue( 'C' . ($st_data-2 ),  "SUMMARY CUSTOMER DEMAND BY PRODUCTS CATEGORY" );
                $objPHPExcel->getActiveSheet()->getStyle( 'C' . ($st_data-2) )->applyFromArray(array('font' => Style_Font(12,"FFFFFF",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('C'.($st_data-2).':'.'H'.($st_data-2))->applyFromArray(array('fill' => Style_Fill('1a3365')));

                $objPHPExcel->getActiveSheet()->mergeCells('C'.($st_data-2).':'.'H'.($st_data-2));

                set_head2($objPHPExcel,'K', 'L',  $st_data, $count_data, '284d00');
                set_head2($objPHPExcel,'N', 'O',  $st_data, $count_data, '000099');
                set_head2($objPHPExcel,'Q', 'R',  $st_data, $count_data, 'b30000');                                           
                set_head2($objPHPExcel,'T', 'V',  $st_data, $count_data, '009933');

                $colors = array('0099ff', 'ff3300', '333300', '990033');

                set_head2($objPHPExcel,'X',  'AA',  $st_data, $count_data, $colors[0]);
                set_head2($objPHPExcel,'AC', 'AF',  $st_data, $count_data, $colors[1]);
                set_head2($objPHPExcel,'AH', 'AK',  $st_data, $count_data, $colors[2]);
                set_head2($objPHPExcel,'AM', 'AP',  $st_data, $count_data, $colors[3]);


                $objPHPExcel->getActiveSheet()->getStyle('C'.($st_data-2).':'.'H'. $count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'1a3365'))));

                $objPHPExcel->getActiveSheet()->getStyle('C'. $st_data .':'.'H'. $count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));
                $objPHPExcel->getActiveSheet()->getStyle('C'. $st_data .':'.'H'. $st_data)
                                              ->applyFromArray(array(
                                                'borders' => array('top'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));
                $objPHPExcel->getActiveSheet()->getStyle('C'. $count_data . ':'. $col_name[$count_index].$count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6')))); 

               foreach ($dat as $value) 
                {
                  //var_dump($value); exit;
                  $objPHPExcel->getActiveSheet()->setCellValue( 'D' . ($row++),  $value['GP'] );
                  $objPHPExcel->getActiveSheet()->setCellValue( 'C' . ($row-1),  'c' );
                  $objPHPExcel->getActiveSheet()->mergeCells('D'. ($row-1) . ':' . 'H' . ($row-1) );
                } 


                Style_Alignment( 'C' . ($st_data-2) . ':' . 'D' . ($st_data-2), 3, false, $objPHPExcel);
                //Style_Alignment('D' . $st_data . ':' . 'D' . $count_data), 9, false, $objPHPExcel);
                Style_Alignment( 'C' . $st_data . ':' . 'C' . $count_data, 3, false, $objPHPExcel);

                $objPHPExcel->getActiveSheet()->getStyle( 'D' . $st_data . ':' . 'D' . $count_data  )->applyFromArray(array('font' => Style_Font(11,"1a3365",true,true)));
                $objPHPExcel->getActiveSheet()->getStyle( 'C' . $st_data . ':' . 'C' . $count_data  )->applyFromArray(array('font' => Style_Font(12,"1a3365",true,false,'Wingdings 3')));




}

function put_data($objPHPExcel, $dat, $cell, $row)
{

  $str = "=SUBTOTAL(109,";
  foreach ( $dat as $key => $value ) 
  {
    $str = "=SUBTOTAL(109,";
      foreach( $value as $ro => $val)
      {
        $str .= $cell.$val.",";
      }
             
    $objPHPExcel->getActiveSheet()->setCellValue($cell.($row+$key), substr($str, 0, strlen($str)-1) . ")" );       
  }
  //exit;
}











?>