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

//echo $lastmount; exit;
//var_dump($list_act_report); exit;
$col_name = array();
$subplan     = array();
$subactual   = array();
$subdiff     = array();
$subacc_diff = array();
$subng       = array();
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, $cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "A".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "B".$cm); }
foreach ( range('A', 'Z') as $cm ) { array_push($col_name, "C".$cm); }

 
foreach ( range(0, 31) as $cm ) { array_push($subplan,     "=SUBTOTAL(109"); }
foreach ( range(0, 31) as $cm ) { array_push($subactual,   "=SUBTOTAL(109"); }
foreach ( range(0, 31) as $cm ) { array_push($subacc_diff, "=SUBTOTAL(109"); }
foreach ( range(0, 31) as $cm ) { array_push($subdiff,     "=SUBTOTAL(109"); }
foreach ( range(0, 31) as $cm ) { array_push($subng,       "=SUBTOTAL(109"); }
//var_dump($subacc_diff); exit;

$ind = 0;
$i=2;
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
            $i = 2;   
            // $ind = 0;
            $count_data  =  count($list_act_report[$sheetIndex]);

            foreach ( range(0, 31) as $cm ) 
            { 
             $subplan[$cm]     = "=SUBTOTAL(109"; 
             $subactual[$cm]   = "=SUBTOTAL(109"; 
             $subacc_diff[$cm] = "=SUBTOTAL(109"; 
             $subdiff[$cm]     = "=SUBTOTAL(109"; 
             $subng[$cm]       = "=SUBTOTAL(109"; 
            }

    if ($count_data > 0) 
    {      
#========================================================================================================================  Put field ====================================================================================        
                $objPHPExcel->getActiveSheet()->setTitle( "$til"  );
                $objPHPExcel->getActiveSheet()->setShowGridlines(False);
                $st_col = 5;
                $st_dat = 14;
                $sub = 9;
                $count_index =  count($list_act_report[$sheetIndex][0])  - ( 31 - $lastmount ) ;
                $row = $st_dat;
                $count_data  =  count($list_act_report[$sheetIndex]) + $row-1;
                $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 10 );
                foreach(range(2, 5) as $r)
                $objPHPExcel->getActiveSheet()->getRowDimension( $r )->setRowHeight( 15 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 6  )->setRowHeight( 7 ); 
                foreach(range(8, 11) as $r)               
                $objPHPExcel->getActiveSheet()->getRowDimension( $r )->setRowHeight( 18 );
                $objPHPExcel->getActiveSheet()->getRowDimension( 12 )->setRowHeight( 10 ); 
                $objPHPExcel->getActiveSheet()->getRowDimension( 13 )->setRowHeight( 10 ); 

                $objPHPExcel->getActiveSheet()->freezePane('J'.$row);   
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(90);    
                $objPHPExcel->getActiveSheet()->setAutoFilter('C'.($st_dat-1).':'.$col_name[$count_index].($st_dat-1));                


                $objPHPExcel->getActiveSheet()->getStyle('B2:'.$col_name[$count_index+1].($count_data+1))
                                              ->applyFromArray(array(
                                                'borders' => array('outline'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000023')))); 

                $objPHPExcel->getActiveSheet()->getStyle('C'.$st_dat.':'.$col_name[$count_index].$count_data)
                                              ->applyFromArray(array(
                                                'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));

                                              

                $objPHPExcel->getActiveSheet()->getStyle('C'.$st_col.':'.$col_name[$count_index].$st_col)->applyFromArray(array('fill' => Style_Fill('305496')));

                $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].$st_col.':'.$col_name[($count_index+4)].$st_col)->applyFromArray(array('fill' => Style_Fill('305496')));

                $objPHPExcel->getActiveSheet()->getStyle('C'.($st_dat-1).':'.$col_name[$count_index].($st_dat-1))->applyFromArray(array('fill' => Style_Fill('e0ebeb')));







                $objPHPExcel->getActiveSheet()->getStyle('C'.$st_col.':'.'H'.$st_col)->applyFromArray(array('font' => Style_Font(11,"FFFFFF",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('I'.$st_col.':'.$col_name[$count_index+6].($st_col))->applyFromArray(array('font' => Style_Font(10,"FFFFFF",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('I'.$st_dat.':'.$col_name[$count_index].$count_data)->applyFromArray(array('font' => Style_Font(10,"000000",false,true)));
                //echo $row; exit;                
                   
                foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
                {
                	if ( is_numeric(substr($key,0,2)) && substr($key,0,2) > $lastmount ) break;   
                    if ($key != "NO")
                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++].$st_col, str_replace("_", " ", $key));

                    if ( holiday(substr($key,0,2), $holiday) )
                     {
                        //echo substr($key,0,2);  exit;
                        $objPHPExcel->getActiveSheet()->getStyle($col_name[($i-1)]. $st_dat . ':' . $col_name[($i-1)].$count_data)->applyFromArray( array( 'fill' => Style_Fill('B9FDDE') ) );
                        $objPHPExcel->getActiveSheet()->getStyle($col_name[($i-1)]. '7' . ':' . $col_name[($i-1)].'11')->applyFromArray( array( 'fill' => Style_Fill('B9FDDE') ) );                                                           
                     }  
                   
                } // exit;     
#========================================================================================================================  Put data ====================================================================================                

                foreach ($list_act_report[$sheetIndex] as $key => $value) 
                {               
                   $col = 2;
                    foreach ($value as $body => $val) 
                    {


                    	 if ( is_numeric(substr($body,0,2)) && substr($body,0,2) > $lastmount ) break;

                        if ($body != "NO")
                            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$col++].($row), $val);
                        
                        if ( $col == 10 &&  $value['H'] == '4' )
                            $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$col-1].($row) , "=".$col_name[$col-1].($row-1));
                        
                        if ( $col > 10 && $value['H'] == '4' )
                            $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$col-1].($row) , "=".$col_name[$col-1].($row-1) . "+" .  $col_name[$col-2].($row) );

                        if($val == 3 && $body == 'MODEL')  $objPHPExcel->getActiveSheet()->getStyle($col_name[$col-1].($row))->getNumberFormat()->setFormatCode('###"E00"');

                        if ($val == "1" && $body == "H" )
                        {
                            $objPHPExcel->getActiveSheet()->mergeCells('C'.$row.':'.'C'.($row+4));
                            $objPHPExcel->getActiveSheet()->mergeCells('D'.$row.':'.'D'.($row+4));
                            $objPHPExcel->getActiveSheet()->mergeCells('E'.$row.':'.'E'.($row+4));
                            $objPHPExcel->getActiveSheet()->mergeCells('F'.$row.':'.'F'.($row+4));
                            $objPHPExcel->getActiveSheet()->mergeCells('G'.$row.':'.'G'.($row+4));
                            $objPHPExcel->getActiveSheet()->mergeCells('H'.$row.':'.'H'.($row+4));

                            $objPHPExcel->getActiveSheet()->setCellValue('I'.$row,  '  Plan');
                        }
                        elseif($val == "2" && $body == "H")
                        {

                            $objPHPExcel->getActiveSheet()->setCellValue('I'.$row,  '  Actual');

                            $objPHPExcel->getActiveSheet()->getStyle('I'. $row .':'. $col_name[$count_index] . $row)->applyFromArray(array('fill' => Style_Fill('d9ffb3')));
                            $objPHPExcel->getActiveSheet()->getStyle('I'. $row .':'. $col_name[$count_index] . $row)->applyFromArray(array('font' => Style_Font(10,"0000ff",true,true)));


                        }  
                        elseif($val == "3" && $body == "H")
                        {

                            $objPHPExcel->getActiveSheet()->setCellValue('I'.$row,  '  Diff');


                        }
                        elseif($val == "4" && $body == "H")
                        {

                            $objPHPExcel->getActiveSheet()->setCellValue('I'.$row,  '  Acc. Diff');


                        }
                        elseif($val == "5" && $body == "H")
                        {

                            $objPHPExcel->getActiveSheet()->setCellValue('I'.$row,  '  Ng');

                            $objPHPExcel->getActiveSheet()->getStyle('C'. $row .':'. $col_name[$count_index] . $row)
                                                          ->applyFromArray(array(
                                                            'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));

                            $objPHPExcel->getActiveSheet()->getStyle('I'. $row .':'. $col_name[$count_index] . $row)->applyFromArray(array('fill' => Style_Fill('eaeae1')));
                            $objPHPExcel->getActiveSheet()->getStyle('I'. $row .':'. $col_name[$count_index] . $row)->applyFromArray(array('font' => Style_Font(10,"ff0000",true,true)));
                            //$objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_name[$col++].($row), $val);

                            
                        }  

                        //echo 
                        if( $col > 8 && strlen($body) == 4 )
                        {
                            $in = intval(substr($body, 0,2)) - 1;

                            if ($value['H'] == '1')
                            {
                                $subplan[$in] .= "," . $col_name[($in+9)].($row);
                                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($in+9)] . '7', $subplan[$in] . ")");
                            }

                            elseif ($value['H'] == '2')
                            {
                                $subactual[$in] .= "," . $col_name[($in+9)].($row);
                                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($in+9)] . '8', $subactual[$in] . ")");
                            }

                            elseif ($value['H'] == '3')
                            {
                                $subdiff[$in] .= "," . $col_name[($in+9)].($row);
                                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($in+9)] . '9', $subdiff[$in] . ")");
                            }


                            elseif ($value['H'] == '4')
                            {
                                $subacc_diff[$in] .= "," . $col_name[($in+9)].($row);
                                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($in+9)] . '10', $subacc_diff[$in] . ")");
                            }


                            elseif ($value['H'] == '5')
                            {
                                $subng[$in] .= "," . $col_name[($in+9)].($row);
                                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($in+9)] . '11', $subng[$in] . ")");
                            }                                                                                    



                        }

                    }

                    $row++;         

                }
 //echo $til . " = " . $subplan[0] . "<hr>" ;

                //$objPHPExcel->getActiveSheet()->setCellValue('J7',  $subplan[0] . ")" );
                $objPHPExcel->getActiveSheet()->setCellValue('D2', 'Production of '.date('F Y',strtotime($en)) );
                $objPHPExcel->getActiveSheet()->setCellValue('C7', 'TOTAL');
                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+3)].$st_col, 'Summary');

                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+4)].($st_col+2), '=SUM(' .$col_name[9]."7".":".$col_name[(8+$lastmount)]."7".")");
                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+4)].($st_col+3), '=SUM(' .$col_name[9]."8".":".$col_name[(8+$lastmount)]."8".")");
                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+4)].($st_col+4), '=SUM(' .$col_name[9]."9".":".$col_name[(8+$lastmount)]."9".")");
                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+4)].($st_col+5), '=' .$col_name[(8+$lastmount)]."10");
                $objPHPExcel->getActiveSheet()->setCellValue( $col_name[($count_index+4)].($st_col+6), '=SUM(' .$col_name[9]."11".":".$col_name[(8+$lastmount)]."11".")");

                $objPHPExcel->getActiveSheet()->setCellValue('I5',  '#');

                $objPHPExcel->getActiveSheet()->setCellValue('I7',  '  Plan');
                $objPHPExcel->getActiveSheet()->setCellValue('I8',  '  Actual');
                $objPHPExcel->getActiveSheet()->setCellValue('I9',  '  Diff');
                $objPHPExcel->getActiveSheet()->setCellValue('I10', '  Acc. Diff');
                $objPHPExcel->getActiveSheet()->setCellValue('I11', '  Ng');

                $objPHPExcel->getActiveSheet()->setCellValue($col_name[($count_index+3)].'7',  '  Plan');
                $objPHPExcel->getActiveSheet()->setCellValue($col_name[($count_index+3)].'8',  '  Actual');
                $objPHPExcel->getActiveSheet()->setCellValue($col_name[($count_index+3)].'9',  '  Diff');
                $objPHPExcel->getActiveSheet()->setCellValue($col_name[($count_index+3)].'10', '  Acc. Diff');
                $objPHPExcel->getActiveSheet()->setCellValue($col_name[($count_index+3)].'11', '  Ng');

                $objPHPExcel->getActiveSheet()->setCellValue('C' . ($count_data+3),  'TBKK (Thailand) Co.,Ltd. by Pcsystem.');

                $objPHPExcel->getActiveSheet()->getStyle('D2')->applyFromArray(array('font' => Style_Font(26,"305496",true,true,'Franklin Gothic Heavy')));

                $objPHPExcel->getActiveSheet()->getStyle('C7')->applyFromArray(array('font' => Style_Font(29,"000000",true,true )));

                $objPHPExcel->getActiveSheet()->getStyle('I7:I11')->applyFromArray(array('font' => Style_Font(10,"000000",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('J'.'7'.':'.$col_name[$count_index+6].'11')->applyFromArray(array('font' => Style_Font(10,"000000",true,true)));              

                $objPHPExcel->getActiveSheet()->getStyle('C'.$st_dat.':'.'H'.$count_data)->applyFromArray(array('font' => Style_Font(10,"000000",false,true)));

                $objPHPExcel->getActiveSheet()->getStyle('C' . ($count_data+3) .':'. 'D' .($count_data+3))->applyFromArray(array('font' => Style_Font(9,"000000",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('I8' .':'. $col_name[$count_index] . '8')->applyFromArray(array('fill' => Style_Fill('d9ffb3')));
                $objPHPExcel->getActiveSheet()->getStyle('I8' .':'. $col_name[$count_index] . '8')->applyFromArray(array('font' => Style_Font(10,"0000ff",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle('I11' .':'. $col_name[$count_index] . '11')->applyFromArray(array('fill' => Style_Fill('eaeae1')));
                $objPHPExcel->getActiveSheet()->getStyle('I11' .':'. $col_name[$count_index] . '11')->applyFromArray(array('font' => Style_Font(10,"ff0000",true,true)));


                $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'8'.':'. $col_name[($count_index+4)] . '8')->applyFromArray(array('fill' => Style_Fill('d9ffb3')));
                $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'8'.':'. $col_name[($count_index+4)] . '8')->applyFromArray(array('font' => Style_Font(10,"0000ff",true,true)));

                $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'11'.':'. $col_name[($count_index+4)] . '11')->applyFromArray(array('fill' => Style_Fill('eaeae1')));
                $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'11'.':'. $col_name[($count_index+4)] . '11')->applyFromArray(array('font' => Style_Font(10,"ff0000",true,true)));
                
                $objPHPExcel->getActiveSheet()->getStyle('D2:'.'H'.'3')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'305496'))));

                $objPHPExcel->getActiveSheet()->getStyle('C6:'.$col_name[$count_index].'6')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));

                $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'6'.':'.$col_name[($count_index+4)].'6')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));                                              

                $objPHPExcel->getActiveSheet()->getStyle('C11:'.$col_name[$count_index].'11')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));

                $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'11'.':'.$col_name[($count_index+4)].'11')
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000')))); 

                $objPHPExcel->getActiveSheet()->getStyle('I7:'.$col_name[$count_index].'11')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'a6a6a6'))));

                $objPHPExcel->getActiveSheet()->getStyle($col_name[($count_index+3)].'7'.':'.$col_name[($count_index+4)].'11')
                                              ->applyFromArray(array(
                                                'borders' => array('inside'   => Style_border(PHPExcel_Style_Border::BORDER_DOTTED,'a6a6a6'))));

                $objPHPExcel->getActiveSheet()->getStyle('I7:'.$col_name[$count_index].'11')
                                              ->applyFromArray(array(
                                                'borders' => array('left'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));

                $objPHPExcel->getActiveSheet()->getStyle('C'.$st_col.':'.$col_name[$count_index].$st_col)
                                              ->applyFromArray(array(
                                                'borders' => array('inside' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'bebebe'))));

                $objPHPExcel->getActiveSheet()->getStyle( 'I' . $st_dat .':'. 'I' . $count_data )
                                              ->applyFromArray(array(
                                                'borders' => array('left'   => Style_border(PHPExcel_Style_Border::BORDER_THICK,'000000'))));


                $objPHPExcel->getActiveSheet()->getStyle( 'C' . ($count_data+3) .':'. 'F' .($count_data+3) )
                                              ->applyFromArray(array(
                                                'borders' => array('bottom'   => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));    


                                                              
//echo $col_name[42]; exit;

                $objPHPExcel->getActiveSheet()->getStyle('I'.$st_dat.':'.$col_name[$count_index].$count_data)
                                              ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');


                $objPHPExcel->getActiveSheet()->getStyle('J'.'7'.':'.$col_name[$count_index+6].'11')
                                              ->getNumberFormat()->setFormatCode('_-* #,##0_-;[Red](#,##0)_-;_-* "-"??_-;_-@_-');

                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[0])->setWidth('2.71');     #A
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[1])->setWidth('3.71');     #B no
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[2])->setWidth('7.71');     #D plnt
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[3])->setWidth('7.71');     #C pd                
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[4])->setWidth('30.71');    #H it_nm
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[5])->setWidth('16.71');    #I model
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[6])->setWidth('26.71');    #J
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[7])->setWidth('20.71');    #K
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[8])->setWidth('9.71');     #L
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+1)])->setWidth('3.71');    #M
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+2)])->setWidth('3.71');    #N
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+3)])->setWidth('10.71');   #M
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+4)])->setWidth('16.71');   #M
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[($count_index+5)])->setWidth('3.71');    #N                
                foreach(range(9, $count_index) as $r)
                $objPHPExcel->getActiveSheet()->getColumnDimension($col_name[$r])->setWidth('10.71');   #M                 


               // Style_Alignment('C2:C5',3, false, $objPHPExcel);
               // Style_Alignment('J3',3, false, $objPHPExcel);
                Style_Alignment(('B'.$st_col.':'.$col_name[$count_index+6].$st_col), 3, false, $objPHPExcel);
                Style_Alignment(('J'.'7'.':'.$col_name[$count_index+6].$count_data), 6, false, $objPHPExcel);
                Style_Alignment(('B2'.':'.'H'.$count_data), 3, false, $objPHPExcel);
                Style_Alignment(($col_name[($count_index+3)].'7'.':'.$col_name[($count_index+3)].'11'), 9, false, $objPHPExcel);
                //Style_Alignment(('B'.$st_dat.':'.'I'.$count_data), 9, false, $objPHPExcel);
                $objPHPExcel->getActiveSheet()->mergeCells('D2:'.'H3');
                $objPHPExcel->getActiveSheet()->mergeCells($col_name[($count_index+3)].$st_col.':'.$col_name[($count_index+4)].$st_col);
                $objPHPExcel->getActiveSheet()->mergeCells($col_name[($count_index+4)].'9'.':'.$col_name[($count_index+4)].'10');
                $objPHPExcel->getActiveSheet()->mergeCells('C7:'.'H11');


                //$objPHPExcel->getActiveSheet()->duplicateStyle( $objPHPExcel->getActiveSheet()->getStyle( $col_name[44].($st_dat+4) ), ('C14:C18') );


                // Style_group_lv1_Col($col_name, 4, $objPHPExcel );


                // if ( date('d') < 3  )
                //     {
                //         foreach (range( 16 , $count_index ) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel );
                //     }
                // elseif ( date('d') > 3 && date('d') < $lastmount ) 
                //     {
                //         foreach (range( 9 , ( (date('d')-4)+8 ) ) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel );

                //         foreach (range( ((date('d')+4)+8) , $count_index ) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel );
                                                                         
                //     }
                // else
                //     {
                //         foreach (range( 9 , ($count_index-2) ) as $index) Style_group_lv1_Col($col_name, $index, $objPHPExcel );
                //     }

                //     if( (date('d')-0) != 1 )
                //     {
                //         $objPHPExcel->getActiveSheet()->getStyle( $col_name[( (date('d')-1)+8 )].($st_col-1).':'. $col_name[( (date('d')-1)+8 )].$count_data )
                //                                       ->applyFromArray(array(
                //                                             'borders' => array('outline' => Style_border(PHPExcel_Style_Border::BORDER_THICK,'00cc00'))));   #GREEN IS YESTERDAY

                //         $objPHPExcel->getActiveSheet()->setCellValue( $col_name[( (date('d')-1)+8 )].($st_col-1), 'Yesterday');
                //         $objPHPExcel->getActiveSheet()->getStyle($col_name[( (date('d')-1)+8 )].($st_col-1))->applyFromArray(array('fill' => Style_Fill('00cc00')));

                //         $objPHPExcel->getActiveSheet()->getStyle( $col_name[( (date('d')-1)+8 )].($st_col-1) )
                //                                       ->applyFromArray(array(
                //                                             'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'ffff4d'))));                           
                //     }

                //         $objPHPExcel->getActiveSheet()->getStyle( $col_name[( (date('d')-0)+8 )].($st_col-1).':'. $col_name[( (date('d')-0)+8 )].$count_data )
                //                                       ->applyFromArray(array(
                //                                             'borders' => array('outline' => Style_border(PHPExcel_Style_Border::BORDER_THICK,'ff0000'))));   #RED IS TODAY

                //         $objPHPExcel->getActiveSheet()->getStyle( $col_name[( (date('d')-0)+8 )].($st_col-1) )
                //                                       ->applyFromArray(array(
                //                                             'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'ffff4d'))));   
                                                            

                //         $objPHPExcel->getActiveSheet()->setCellValue( $col_name[( (date('d')-0)+8 )].($st_col-1), 'Today'); 


                //         $objPHPExcel->getActiveSheet()->getStyle($col_name[( (date('d')-0)+8 )].($st_col-1))->applyFromArray(array('fill' => Style_Fill('ff0000')));

                //         $objPHPExcel->getActiveSheet()->getStyle($col_name[( (date('d')-1)+8 )].($st_col-1))->applyFromArray(array('font' => Style_Font(9,"ffff4d",true,true)));
                //         $objPHPExcel->getActiveSheet()->getStyle($col_name[( (date('d')-0)+8 )].($st_col-1))->applyFromArray(array('font' => Style_Font(9,"ffff4d",true,true)));

                //         Style_Alignment( ($col_name[( (date('d')-1)+8 )].($st_col-1) . ':' . $col_name[( (date('d')-0)+8 )].($st_col-1) ), 3, false, $objPHPExcel);


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
                  //  echo $count_data;                                      
#========================================================================================================================  Put data ====================================================================================         
    } else {
                    $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue('A1', "No data ".$til.".");
                    $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true,false,'Franklin Gothic Book')));
    }
$ind++;


//echo $til; exit;

}

//exit;
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