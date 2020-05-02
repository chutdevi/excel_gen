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
$col_index = array();
foreach ( range('A', 'Z') as $cm ) {
    array_push($col_index, $cm);
}
foreach ( range('A', 'Z') as $cm ) {
    array_push($col_index, "A$cm");
}
//echo count($list_act_report); exit;
 $head1 = array( 
         'No.'
        , 'PLANT'
        , 'PD'
        , 'LINE CD'
        , 'SEC NAME'
        , 'PLAN'
        , 'ACTUAL'
        , 'DIFF.'
        , 'NG.'
        , 'PLAN'
        , 'ACTUAL'
        , 'DIFF.'        
        , 'NG.'
        , 'OPRT TIME.'
        , 'LOSS TIME.' 
        , 'PLAN'
        , 'PLAN'
        , 'PLAN'            
        , 'Accum. PLAN'   
        , 'Accum. ACTUAL'
        , 'Accum. DIFF.'
        , 'Accum. NG.'
        , 'DEFECT PERCENT'
        , '(%)'
        , 'PLAN THIS MONTH' );
  $head2 = array( 
         'No.'
        ,'PLANT CD'
        ,'PD'
        ,'LINE CD'
        ,'SEC NM'
        ,'ITEM CD'
        ,'ITEM NAME'
        ,'MONTHLY PLAN'
        ,'ACTUAL SUMMARY'
        ,'DIFF.' 
    );
$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1 : (date('d'));
$MontCol = ( (date('d')+0) == 1 ) ? date('M', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('M'));
//echo $dateCol . '/' . $MontCol; exit;
for ($i=1; $i < $dateCol ; $i++) { 
    $dayT = $dateCol-($dateCol-$i);
    $dayT = ($dayT < 10 ) ? "0".$dayT."/".$MontCol : $dayT."/".$MontCol;
    array_push($head2, $dayT);
}

// var_dump($holiday); exit;
//var_dump($head2); exit;
$colume_head = array( $head1, $head1, $head1, $head1, $head1, $head1, $head1, $head1, $head2 );

//=======================================================================================  config Style ================================================================================

foreach ($title as $inTil => $til) {

        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex($inTil);
        $objPHPExcel->getActiveSheet()->setTitle("$til");
        $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 30 );
        $objPHPExcel->getActiveSheet()
                    ->getStyle('1:3')
                    ->getAlignment()
                    ->setWrapText(true)
                    ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

        $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('5.5');
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('6.5');
        $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('8');         
        $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('10');
        $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('65');
       
         $i=0;
            if ($inTil < count($title)-1) 
            {
                $til = ( $til == "ALL SECTION" ) ? "Production report $todayA" : $til; 
                $objPHPExcel->getActiveSheet()->mergeCells('A1:E2');
                $objPHPExcel->getActiveSheet()->setAutoFilter('A3:'.$col_index[count($head1)-1].'3');
                $objPHPExcel->getActiveSheet()->freezePane('F4');
                style_his_r3('A3:'.$col_index[count($head1)-1].'3', $objPHPExcel, $inTil, 'E0E0E0', '000000', 11, true);
                //style_his_r3('A3:'.$col_index[count($head2)-1].'3', $objPHPExcel, $inTil, '0066CC', 'EFFF00', 10, TRUE);
                style_his_r3('A1',$objPHPExcel, $inTil,'FFFFFF', '000000', 20, true);                     
                // HEAD
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue("A1", $til                           ) ;
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue("F1", "Production 2 days ago."       ) 
                                                         ->setCellValue("F2", $yesterdayB                    )
                                                         ->setCellValue("J2", $yesterdayA                    ) ;
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue("P1", "Production Plan"              ) 
                                                         ->setCellValue("P2", $todayA                        )
                                                         ->setCellValue("Q2", $tomorrowA                     ) 
                                                         ->setCellValue("R2", $tomorrowB                     );
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue("S1", "Monthly Production"           ) 
                                                         ->setCellValue("S2",  date('F').' Production'       ) ;
                $objPHPExcel->getActiveSheet()->mergeCells('F1:O1');
                $objPHPExcel->getActiveSheet()->mergeCells('F2:I2');
                $objPHPExcel->getActiveSheet()->mergeCells('J2:O2');
                style_his_r3('A1',   $objPHPExcel, $inTil, 'FFFFFF', '000000', 20, true); 
                style_his_r3('F1:O1',$objPHPExcel, $inTil, 'd1fcf8', 'ef5d1a', 18, true);
                style_his_r3('F2:I2',$objPHPExcel, $inTil, 'f8fdff', 'ec6eb9', 13, true);
                style_his_r3('J2:O2',$objPHPExcel, $inTil, 'f8fdff', '805ffb', 13, true); 
           
            //80a9ed                                             
    //------------------------------------------------------------------------------------------------------------------------------    
                                            
                $objPHPExcel->getActiveSheet()->mergeCells('P1:R1');
                $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth('15');
                $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth('15');
                $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth('15');
                
                style_his_r3('P1:Q1',$objPHPExcel, $inTil, 'd9d9d9', 'ef5d1a', 18, true);
                style_his_r3('P2',$objPHPExcel, $inTil, 'f8fdff', '805ffb', 13, true);
                style_his_r3('Q2',$objPHPExcel, $inTil, 'f8fdff', 'ec6eb9', 13, true);
                style_his_r3('R2',$objPHPExcel, $inTil, 'f8fdff', 'ec6eb9', 13, true);                        
    //------------------------------------------------------------------------------------------------------------------------------    
                  
                $objPHPExcel->getActiveSheet()->mergeCells('S1:'.$col_index[count($head1)-1].'1');
                $objPHPExcel->getActiveSheet()->mergeCells('S2:'.$col_index[count($head1)-1].'2');

                style_his_r3('S1:'.$col_index[count($head1)-1].'1',$objPHPExcel, $inTil, 'f8cdad', 'ef5d1a', 18, true);
                style_his_r3('S2:'.$col_index[count($head1)-1].'2',$objPHPExcel, $inTil, 'f8fdff', 'ec6eb9', 13, true);                    
    //------------------------------------------------------------------------------------------------------------------------------     
                foreach(range('F','O') as $columnID) {  $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setWidth('10');   } 
                foreach(range('P','R') as $columnID) {  $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setWidth('17');   }
                foreach(range('S',$col_index[count($head1)-1]) as $columnID) {  $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setWidth('13'); } 
                $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth('5');    
                                 
                $til = ( $til == "Production report $todayA" ) ? "ALL SECTION" : $til;                                                 
            }
            else 
            {
                // HEAD
                $MntHis = ( (date('d')+0) == 1 ) ? date('F-Y', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('F-Y'));
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue("A1", "Production report ".$MntHis   ) ;
                $objPHPExcel->getActiveSheet()->setAutoFilter('A3:J3');
                $objPHPExcel->getActiveSheet()->freezePane('K4');
                $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80);
                $objPHPExcel->getActiveSheet()->mergeCells('A1:'.$col_index[count($head2)-1].'2');
                foreach (range('A','D') as $value) { $objPHPExcel->getActiveSheet()->getColumnDimension($value)->setWidth('9');}
                //$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('65');                      
                $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('21');           
                $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('45');                        
                $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('15');       
                $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('15');       
                $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('15');

                style_his_r3('A1:'.$col_index[count($head2)-1].'1',$objPHPExcel, $inTil, 'c6dc2c', '001e2c',16, TRUE);
                style_his_r3('A3:'.$col_index[count($head2)-1].'3',$objPHPExcel, $inTil, '0066CC', 'EFFF00',10, TRUE);
                    
                foreach ($head2 as $k => $va) 
                    {
                        if ( holiday($va, $holiday) )
                            $objPHPExcel->getActiveSheet()->getStyle($col_index[$k].'4:'.$col_index[$k].(count($list_act_report['production_actual_history'])+3))->applyFromArray( array( 'fill' => fill_color('E6B8B7') ) );   

                    }                
            }
                foreach ( $colume_head[$inTil] as $hn)
                {

                        $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_index[$i++]."3", strtoupper($hn)) ; 
                } 
                if ( ($til != "ALL SECTION") AND ($til != "Production actual history") ) 
                {
                   // $objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth('5');
                    $objPHPExcel->getActiveSheet()->removeColumn('B')->removeColumn('B');
                }
                 

}

//=======================================================================================  Input data ================================================================================
//var_dump($list_act_report['production_actual_history'][0]['date5']); exit;
 //echo count($list_act_report['all_section']); exit;
// echo count($list_act_report) ."<hr>";
    $r=4;
    $i=0;
    foreach ($title as $ind => $va_nm) 
    {
            $ind_ar = strtolower(str_replace(" ", "_", $va_nm));
            $styleSum  =  array('bottom' => border_color(PHPExcel_Style_Border::BORDER_DOUBLE,'000000'), 'top' => border_color(PHPExcel_Style_Border::BORDER_THIN,'000000') );
            $r=4;
            if ($ind_ar != "k2pd06") 
            {
                foreach ($list_act_report[$ind_ar] as $in_row => $value) 
                {
                    $i=0;           
                        foreach ($value as $row_val => $data) 
                        {
                            $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_index[$i++].$r , $data) ;
                            if ($va_nm == "Production actual history") 
                            {
                                switch ($data) 
                                    {
                                        case 'K1PD01':
                                                     $objPHPExcel->getActiveSheet()->getStyle('A'.$r.':'.'J'.$r)->applyFromArray( array( 'fill' => fill_color('C5D9F1') ) );
                                            break;
                                        case 'K1PD02':
                                                     $objPHPExcel->getActiveSheet()->getStyle('A'.$r.':'.'J'.$r)->applyFromArray( array( 'fill' => fill_color('DCE6F1') ) );
                                            break;
                                        case 'K1PD03':
                                                     $objPHPExcel->getActiveSheet()->getStyle('A'.$r.':'.'J'.$r)->applyFromArray( array( 'fill' => fill_color('F2DCDB') ) );
                                            break;
                                        case 'K1PD04':
                                                     $objPHPExcel->getActiveSheet()->getStyle('A'.$r.':'.'J'.$r)->applyFromArray( array( 'fill' => fill_color('EBF1DE') ) );
                                            break;
                                        case 'K1PD05':
                                                     $objPHPExcel->getActiveSheet()->getStyle('A'.$r.':'.'J'.$r)->applyFromArray( array( 'fill' => fill_color('E4DFEC') ) );
                                            break;
                                        case 'K2PD06':
                                                     $objPHPExcel->getActiveSheet()->getStyle('A'.$r.':'.'J'.$r)->applyFromArray( array( 'fill' => fill_color('DAEEF3') ) );
                                            break; 
                                        case 'K1PL00':
                                                     $objPHPExcel->getActiveSheet()->getStyle('A'.$r.':'.'J'.$r)->applyFromArray( array( 'fill' => fill_color('FDE9D9') ) );
                                            break;                                                                           
                                    }

                                if ( ( $va_nm == "Production actual history" ) AND ( $i == count($head2) )  ) 
                                {
                                   break;
                                }       
                            }


                        }
                   
                    $r++;



                    if (($r >= count($list_act_report[$ind_ar])+4) AND ($va_nm != "Production actual history") ) 
                    {
                      $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A'.($r-1) , "TOTAL");
                      $objPHPExcel->getActiveSheet()->getRowDimension( ($r-1) )->setRowHeight( 35 );
                      
                    
                      $objPHPExcel->getActiveSheet()->getStyle('A'.($r-1))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);                                 
                      if ($ind_ar == 'all_section')
                      {
                        $objPHPExcel->getActiveSheet()->mergeCells('A'.($r-1).':'.'E'.($r-1));
                        $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_index[count($head1)-1].($r-1) )->getNumberFormat()->setFormatCode('#,##0_);[Red](#,##0)');
                        $objPHPExcel->getActiveSheet()->getStyle('W4:'.'W'.($r-1))->getNumberFormat()->setFormatCode('#,##0.00');
                        $objPHPExcel->getActiveSheet()->getStyle('N4:'.'O'.($r-1))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                        $objPHPExcel->getActiveSheet()->getStyle('A'.($r-1).':'.$col_index[count($head1)-1].($r-1))->applyFromArray( array( 'font' => font_color(12,true), 'borders' => $styleSum ) );
                        $objPHPExcel->getActiveSheet()->getStyle('I4:I'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                        $objPHPExcel->getActiveSheet()->getStyle('O4:O'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                        $objPHPExcel->getActiveSheet()->getStyle('R4:R'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                        $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setVisible(false);
                        $objPHPExcel->getActiveSheet()->getStyle('Q4:Q'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                      } 
                      else 
                      {
                        $objPHPExcel->getActiveSheet()->mergeCells('A'.($r-1).':'.'C'.($r-1));
                        $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_index[count($head1)-1].($r-1) )->getNumberFormat()->setFormatCode('#,##0_);[Red](#,##0)');
                        $objPHPExcel->getActiveSheet()->getStyle('U4:'.'U'.($r-1))->getNumberFormat()->setFormatCode('#,##0.00');
                        $objPHPExcel->getActiveSheet()->getStyle('L4:'.'M'.($r-1))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                        $objPHPExcel->getActiveSheet()->getStyle('A'.($r-1).':'.$col_index[count($head1)-3].($r-1))->applyFromArray( array( 'font' => font_color(12,true), 'borders' => $styleSum ) );
                        $objPHPExcel->getActiveSheet()->getStyle('G4:G'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                        $objPHPExcel->getActiveSheet()->getStyle('M4:M'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                        $objPHPExcel->getActiveSheet()->getStyle('P4:P'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );    
                        $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setVisible(false);   
                        $objPHPExcel->getActiveSheet()->getStyle('O4:O'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );                 
                      }
                        
                    }
                    if (($r >= count($list_act_report[$ind_ar])+4) AND ($va_nm == "Production actual history")) 
                    {
                       $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_index[count($head1)-1].($r-1) )->getNumberFormat()->setFormatCode('#,##0_);[Red](#,##0)');
                    }        

                }
                $objPHPExcel->getActiveSheet()->getStyle('FF1');
            }
            else
            {
                foreach ($list_act_report[$ind_ar] as $in_row => $value) 
                {

                        $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A'.$r , strtoupper( str_replace("_", " ", $in_row) )) ;
                        $objPHPExcel->getActiveSheet()->mergeCells('A'.$r.':'.$col_index[count($head1)-3].$r);
                        $objPHPExcel->getActiveSheet()->getStyle('A'.$r.':'.$col_index[count($head1)-3].$r)
                                    ->applyFromArray(array('fill' => fill_color('FFCCE5')));
                        $objPHPExcel->getActiveSheet()->getStyle('G'.$r)->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                        $objPHPExcel->getActiveSheet()->getStyle('M'.$r)->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                        $objPHPExcel->getActiveSheet()->getStyle('P'.$r)->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );                                    
                        $r++;       
                        foreach ($value as $row_val => $data_index) 
                        {
                            $i=0;
                            foreach ($data_index as $row => $data) 
                            {
                                $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_index[$i++].$r , $data) ;
                            }
                            $r++;
                             
                        }   
                          if (($r >= count($value)+5)) 
                            {
                              $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A'.($r-1) , "TOTAL");
                              $objPHPExcel->getActiveSheet()->getRowDimension( ($r-1) )->setRowHeight( 35 );
                              $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_index[count($head1)-1].($r-1) )->getNumberFormat()->setFormatCode('#,##0_);[Red](#,##0)');
                              $objPHPExcel->getActiveSheet()->getStyle('U5:'.'U'.($r-1))->getNumberFormat()->setFormatCode('#,##0.00');
                              $objPHPExcel->getActiveSheet()->mergeCells('A'.($r-1).':'.'C'.($r-1));
                              $objPHPExcel->getActiveSheet()->getStyle('L5:'.'M'.($r-1))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                              $objPHPExcel->getActiveSheet()
                                          ->getStyle('A'.($r-1))
                                          ->getAlignment()                                      
                                          ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 
                              $objPHPExcel->getActiveSheet()->getStyle('A'.($r-1).':'.$col_index[count($head1)-3].($r-1))->applyFromArray( array( 'font' => font_color(12,true), 'borders' => $styleSum ) );
                        $objPHPExcel->getActiveSheet()->getStyle('G4:G'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                        $objPHPExcel->getActiveSheet()->getStyle('M4:M'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                        $objPHPExcel->getActiveSheet()->getStyle('P4:P'.($r-1))->applyFromArray( array( 'borders' => array ( 'right' => border_color(PHPExcel_Style_Border::BORDER_THICK,'E26E0A') ) ) );
                            } 
                }
                //$objPHPExcel->getActiveSheet()->removeColumn('B');
                $objPHPExcel->getActiveSheet()->getStyle('FF1'); 
            }



       
    } 

                    
//=======================================================================================  Style data ================================================================================
//FFCCE5
    // $r=4;
    // $i=0;
    // foreach ($title as $ind => $va_nm) 
    // {
    //         $ind_ar = strtolower(str_replace(" ", "_", $va_nm));
    //         $r=4;
    //         if ($ind_ar != "k2pd06") 
    //         {
    //             foreach ($list_act_report[$ind_ar] as $in_row => $value) 
    //             {

    //                   $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A'.($r-1) , "TOTAL") ; 
                        
    //             }
    //         }
    //         else
    //         {
    //             foreach ($list_act_report[$ind_ar] as $in_row => $value) 
    //             {
    //                //echo count($list_act_report[$ind_ar]). "<hr>" . count($in_row). "<hr>" . count($value); exit;

    //                     $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A'.$r , strtoupper( str_replace("_", " ", $in_row) )) ;
    //                     $r++;       
    //                     foreach ($value as $row_val => $data_index) 
    //                     {
    //                         $i=0;
    //                         foreach ($data_index as $row => $data) 
    //                         {
    //                             $objPHPExcel->setActiveSheetIndex($ind)->setCellValue($col_index[$i++].$r , $data) ;
    //                         }
    //                         $r++;
                             
    //                     }   
    //                       if (($r >= count($value)+5)) 
    //                         {
    //                           $objPHPExcel->setActiveSheetIndex($ind)->setCellValue('A'.($r-1) , "TOTAL") ;

    //                         }    
    //             }
    //         }
        
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







function style_his_r3($rane, $objPHPExcel, $Index, $cFill='FFFFCC', $cFont='000000', $fSize=16, $bol=true){
        $objPHPExcel->setActiveSheetIndex($Index);
                $objPHPExcel->getActiveSheet()
                  //  ->mergeCells($rane)
                    ->getStyle($rane)
                    ->applyFromArray(
                                    array(
                                            'fill' => array(
                                                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                                'color' => array('rgb' =>  $cFill)
                                            ),
                                            'font' => array(
                                                'size' => $fSize,
                                                'bold' => $bol,
                                                'color' => array('rgb' => $cFont)
                                            )
                                        )
                                    )   
                    ->getAlignment()
                    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
}
function style_his_r2($rane, $objPHPExcel){
                    $objPHPExcel->setActiveSheetIndex(9);
                            $objPHPExcel->getActiveSheet()
                                ->getStyle($rane)
                                ->applyFromArray(
                                            array(
                                                'fill' => array(
                                                    'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                                    'color' => array('rgb' => 'FFFFCC')
                                                    //'font' => array('size' => 11,'bold' => true)
                                                )
                                            )
                                        )
                                ->getAlignment()
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
}



function fill_color($color='FFFFFF')
{
    return array('type' => PHPExcel_Style_Fill::FILL_SOLID,'color' => array('rgb' =>  $color));
}

function font_color($fSize=11, $bol=false, $cFont='000000')
{
    return  array( 'size' => $fSize, 'bold' => $bol, 'color' => array('rgb' => $cFont) );                                                                                 
}

function border_color($line='BORDER_THICK', $color='000000')
{
    return array( 'style' => $line, 'color' => array('rgb' => $color)) ;
}


function holiday($dat, $hol)
{

    foreach ($hol as $ld) 
        if ( substr( $ld['d_t'], 8,2 ) == substr( $dat, 0,2 ) ) 
            return true;
}


 ?>
