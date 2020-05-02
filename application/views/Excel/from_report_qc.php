<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
ini_set('max_execution_time', 0); 
ini_set('memory_limit','2048M');
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$data_col = array();
//$Today = (date('d')-1)."-".date('M')."-".date('Y');
$Today = ( (date('d')+0) == 1 ) ? date('d-M-Y', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : date('d-M-Y', strtotime(date('Y')."-".(date('m')+0)."-".(date('d')-1)));

$Onday = ( (date('d')+0) == 1 ) ? 32   : date('d', strtotime(date('Y')."-".(date('m')+0)."-".(date('d'))));
$Yeday = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : date('d', strtotime(date('Y')."-".(date('m')+0)."-".(date('d')-1)));


$MontCol = ( (date('d')+0) == 1 ) ? date('M', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('M'));
$YearCol = ( (date('d')+0) == 1 ) ? date('Y', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('Y'));
//echo $Today; exit;
//var_dump($list_act_report); exit;
//echo $Onday . $Yeday .$Today ; exit;
//echo $Onday . "-" .  $MontCol . "-" . $YearCol ; exit;
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

        $sheetIndex =  strtolower(str_replace(' ', '_', $title[$inTil]));

     $i = 0;
     $day = 1;
     if(count($list_act_report[$sheetIndex]) > 0 )
     {
        foreach ( $list_act_report[$sheetIndex][0] as $key => $val ) 
        {
            $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i++]."1", str_replace("_", " ", strtoupper($key)));     
            if ( substr($key,0,4) == 'date' ) 
            {
                $objPHPExcel->setActiveSheetIndex($inTil)->setCellValue($col_name[$i-1]."1", str_replace("_", " ", date('d M', strtotime( ($day++) . "-" .$MontCol . "-" . $YearCol ))));
                $dayCh = substr($key,4,strlen($key)-3);
                if ($dayCh == $Yeday) 
                {
                    break;
                }
            }
        }
         //foreach(range('A','Z') as $columnID) { $objPHPExcel->getActiveSheet()->getColumnDimension('B'.$columnID)->setAutoSize(true); }         
     }



}       

$row = 2;
$indSheet = 0;
foreach ($list_act_report as $key => $value) 
    {
                //echo substr('DATE1',4,2); exit;
                //var_dump($key); exit;
     if(count($list_act_report[$key]) > 0 )
         { 
                   if ($key == 'defect_daily') 
                   {
                    $objPHPExcel->setActiveSheetIndex($indSheet);
                    $objPHPExcel->getActiveSheet()->insertNewRowBefore(1,2);
                    $objPHPExcel->getActiveSheet()->freezePane('M4');
                    $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 35 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 2 )->setRowHeight( 35 );
                    $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 35 );
                    $objPHPExcel->getActiveSheet()
                                ->getStyle(('A3:'.$col_name[$Onday+10].'3'))
                                ->getAlignment()
                                ->setWrapText(true)
                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 

                    $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(60);
                    $objPHPExcel->getActiveSheet()->setAutoFilter('A3:'.$col_name[11].'3');                    
                    $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[$Onday+10].'3')->applyFromArray(array('font'    => Style_Font(10,'FFFFFF',true)));
                    $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[$Onday+10].'3')->applyFromArray(array('fill'    => Style_Fill($colhead)));
                    $objPHPExcel->getActiveSheet()->getStyle('A3:'.$col_name[$Onday+10].'3')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF')))); 

                    $startData = 4;
                    $r = 4;
                            foreach ($value as $nr => $val) 
                            {
                                $indCol = 0;
                                        foreach ($val as $rowData => $data) 
                                        {
                                           
                                           if($indCol > 11 && $r == 4)
                                           {
                                            $objPHPExcel->getActiveSheet()->setCellValue( $col_name[$indCol].'2', '=SUBTOTAL(9,' . $col_name[$indCol] . $r . ":" . $col_name[$indCol] . (count( $list_act_report[$key] )+5) . ")" );
                                           } 

                                            if ($rowData == 'model') 
                                            {
                                                if ($data == '3E00') 
                                                {
                                                   $objPHPExcel->setActiveSheetIndex($indSheet)->setCellValue($col_name[$indCol++].($r), "'".$data);
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

                                                    if ($rowData == 'ng_cd')
                                                    {
                                                            $objPHPExcel->getActiveSheet()->getStyle($col_name[$indCol-1].($r))->getNumberFormat()->setFormatCode('000');
                                                            $objPHPExcel->setActiveSheetIndex($indSheet)->setCellValue($col_name[$indCol-1].($r), intval($data));
                                                    }
                                                    if ( substr($rowData,0,4) == 'date' ) 
                                                    {
                                                        if (substr($rowData,4,2) == $Yeday) 
                                                        {
                                                            break;
                                                        }
                                                        else
                                                        {
                                                         if ( holiday(substr($rowData,4,2), $holiday) )
                                                            $objPHPExcel->getActiveSheet()
                                                                        ->getStyle($col_name[$indCol-1]. '4:' . $col_name[$indCol-1].(count( $list_act_report[$key] )+3))
                                                                        ->applyFromArray( array( 'fill' => Style_Fill('B9FDDE') ) );                                                           
                                                        }
                                                    }                                        

                                        }
                                $r++;
                            }
                    		sunday($Today, 1, 4, (count( $list_act_report[$key] )+3), $col_name, $indSheet, $objPHPExcel);       
                            $objPHPExcel->getActiveSheet()->setCellValue('A1', "DAILY DEFECT  OF ".strtoupper(date('F Y')));
                            $objPHPExcel->getActiveSheet()->setCellValue('K1', "TOTAL WEEK");
                            $objPHPExcel->getActiveSheet()->setCellValue('K2', "TOTAL");
                            $objPHPExcel->getActiveSheet()->setCellValue('L1', '=SUBTOTAL(9,L4:L' . (count( $list_act_report[$key] )+5) . ")" ); 


                            $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(48,'000000',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('K1:K2')->applyFromArray(array('fill'    => Style_Fill('003319')));
                            $objPHPExcel->getActiveSheet()->getStyle('K1:K2')->applyFromArray(array('font' => Style_Font(16,'FDE9D9',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('L1')->applyFromArray(array('fill'    => Style_Fill('202020')));

                            //sunday($Today, 1, 4, (count( $list_act_report[$key] )+3), $col_name, $indSheet, $objPHPExcel);  #Function sunday -----------------------------------------------

                            $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[$Onday+10].'2')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));

                            $objPHPExcel->getActiveSheet()
                                        ->getStyle('A1:'.$col_name[$Onday+10].'2')
                                        ->getAlignment()
                                        ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
                                        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);                            






                            $objPHPExcel->getActiveSheet()->mergeCells('A1:J2');
                            $objPHPExcel->getActiveSheet()->mergeCells('L1:L2');

                            $objPHPExcel->getActiveSheet()->getStyle('L1:'.$col_name[$Onday+10].'1' )
                                                          ->getNumberFormat()->setFormatCode('_* #,##0_-;[RED](#,##0)_-');

                            $objPHPExcel->getActiveSheet()->getStyle('L2:'.$col_name[$Onday+10].(count( $list_act_report[$key] )+5) )
                                                          ->getNumberFormat()->setFormatCode('_-* #,##0_-;[RED](#,##0)_-;_-* [BLACK]"-"??_-;_-@_-');

                            $objPHPExcel->getActiveSheet()->getStyle('M4:'.$col_name[$Onday+10].(count( $list_act_report[$key] )+5))->applyFromArray(array('font' => Style_Font(10,'FF0000',true)));

                            $objPHPExcel->getActiveSheet()->getStyle('L1:'.$col_name[$Onday+10].'1')->applyFromArray(array('font' => Style_Font(16,'FFFFFF',true)));
                            $objPHPExcel->getActiveSheet()->getStyle('L1:'.$col_name[$Onday+10].'1')->getFont()->setUnderline(true);

                            $objPHPExcel->getActiveSheet()->getStyle('L2:'.'L'.$col_name[$Onday+10].'2')->applyFromArray(array('font' => Style_Font(11,'FF0000',true)));


                            $objPHPExcel->getActiveSheet()->getStyle('L4:'.'L'.(count( $list_act_report[$key] )+5))->applyFromArray(array('font' => Style_Font(11,'000066',true)));

                            foreach (range('B', 'E') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('66');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('19');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('31');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('31');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('10');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth('31');
                            $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth('22');
                            foreach (range('M', 'Z') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setWidth('12');
                            foreach (range('A', 'Q') as $key) $objPHPExcel->getActiveSheet()->getColumnDimension('A'.$key)->setWidth('12');

  
					$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setVisible(false);
					$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setVisible(false);
					$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setVisible(false);
					$objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray(array('font' => Style_Font(28,'000000',true)));
	                         
                    }
                   elseif ($key == 'code_detail') 
                   {
                        $objPHPExcel->setActiveSheetIndex($indSheet);
                        $objPHPExcel->getActiveSheet()->getRowDimension( 1 )->setRowHeight( 35 );
                        $objPHPExcel->getActiveSheet()
                            ->getStyle('A1:'.$col_name[2].'1')
                            ->getAlignment()
                            ->setWrapText(true)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 

                        $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(90);
                        $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[2].'1')->applyFromArray(array('font'    => Style_Font(10,'FFFFFF',true)));
                        $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[2].'1')->applyFromArray(array('fill'    => Style_Fill($colhead)));
                        $objPHPExcel->getActiveSheet()->getStyle('A1:'.$col_name[2].'1')->applyFromArray(array('borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'FFFFFF'))));

                        $r = 2;
                            foreach ($value as $nr => $val) 
                            {
                            $indCol = 0;
                                foreach ($val as $rowData => $data) 
                                {
                                    $objPHPExcel->setActiveSheetIndex($indSheet)->setCellValue($col_name[$indCol++].($r), $data);
                                                    if ($rowData == 'CODE')
                                                    {
                                                            $objPHPExcel->getActiveSheet()->getStyle($col_name[$indCol-1].($r))->getNumberFormat()->setFormatCode('000');
                                                            $objPHPExcel->setActiveSheetIndex($indSheet)->setCellValue($col_name[$indCol-1].($r), intval($data));
                                                    }                                    
                                }
                            $r++;
                            }
                           set_autosize('A','C', $objPHPExcel, $indSheet);


                    }

         }   
$indSheet++;
    //echo $indSheet; exit;
$row = 2;
    
} 



// Set active sheet index to the first sheet, so Excel opens this as the first sheet

$objPHPExcel->setActiveSheetIndex(0);
//$objPHPExcel->getActiveSheet()->getStyle('ZF9999')->applyFromArray(array('fill' => Style_Fill($colhead)));
$objPHPExcel->getActiveSheet()->setCellValue('ZF9999', '');
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

function sunday($dat = '01-01-2018',  $focus = 1, $start_row = 4, $end_row = 5, $col = null, $ind = 0,  $objPHPExcel = nul)
{
    $objPHPExcel->setActiveSheetIndex($ind);
    $d = date('d', strtotime($dat));
    //echo $d; exit;
    $fillSum   = array();
    $fillTotal = array();
    $indexSun = 12;
    $kla = 0;
    $merge_nosun = '';
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
                $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun-6].$focus, '=SUM(' . $col[$indexSun-6] . '2' . ":" . $col[$indexSun] . '2' . ")" );
                $objPHPExcel->getActiveSheet()->mergeCells($col[$indexSun-6] . $focus .":" . $col[$indexSun] . $focus); 
                array_push($fillSum, $col[$indexSun-6] . $focus);
                array_push($fillTotal, $col[$indexSun-6] . '2' . ":" . $col[$indexSun] . '2' );
                if( $indexSun + 6 > $d+10 && $valDay < $d-1)
                {
                        $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun+1].$focus, '=SUM(' . $col[$indexSun+1] . '2' . ":" . $col[$d+11] . '2' . ")" );
                        $objPHPExcel->getActiveSheet()->mergeCells($col[$indexSun+1] . $focus .":" . $col[$d+11] . $focus);
                        array_push($fillSum, $col[$indexSun+1] . $focus);
                        array_push($fillTotal, $col[$indexSun+1] . '2' . ":" . $col[$d+11] . '2' );
                        //break;
                }

               // echo $Sunday; exit;
            }
            //FFC2C2

            elseif($indexSun > 12 && $indexSun-6 < 12)
            {
                $objPHPExcel->getActiveSheet()->setCellValue( $col[12].$focus, '=SUM(' . $col[12] . '2' . ":" . $col[$indexSun] . '2' . ")" );
                $objPHPExcel->getActiveSheet()->mergeCells($col[12] . $focus .":" . $col[$indexSun] . $focus);
                array_push($fillSum, $col[12] . $focus);
                array_push($fillTotal, $col[12] . '2' . ":" . $col[$indexSun] . '2' );
            }
        

            // echo date('d', strtotime($Tday)) . "-" . $Nday . $col[$indexSun] . "<hr>";
            //echo date('d', strtotime($Tday)) . "<hr>";

            $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun].$start_row . ":" . $col[$indexSun].$end_row)->applyFromArray(array('fill'    => Style_Fill('FFC2C2'))); 
            $kla = 99;       
        }
        else
        {            
			if($indexSun == 12 && $indexSun+6 > $d)
            {
                $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun].$focus, '=SUM(' . $col[$indexSun] . '2)' );
                $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun].$focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun].'2')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) ); 
  	           	$kla = 99;
            } 
            elseif($indexSun > 12 &&  $indexSun-1 == $Sunday)
            {
                $objPHPExcel->getActiveSheet()->setCellValue( $col[$indexSun].$focus, '=SUM(' . $col[$indexSun] . '2)' );
                $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun].$focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun].'2')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) ); 
  	           	$kla = 99;
            }             
            elseif($indexSun > 12 && $indexSun-7 < $Sunday)
            {
                $objPHPExcel->getActiveSheet()->setCellValue( $col[$Sunday+1].$focus, '=SUM(' . $col[$Sunday+1] . '2' . ":" . $col[$indexSun] . '2' . ")" );
                $objPHPExcel->getActiveSheet()->getStyle($col[$Sunday+1].$focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                $objPHPExcel->getActiveSheet()->getStyle($col[$Sunday+1] . '2' . ":" . $col[$indexSun] . '2')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) );
                $merge_nosun = $col[$Sunday+1] . $focus . ":" . $col[$indexSun] . $focus;
                $kla = 0;
            }   
            elseif($indexSun > 12 && $indexSun-6 < 12)
            {
                $objPHPExcel->getActiveSheet()->setCellValue( $col[12].$focus, '=SUM(' . $col[$Sunday+1] . '2' . ":" . $col[$indexSun] . '2' . ")" );
                $objPHPExcel->getActiveSheet()->getStyle($col[12].$focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                $objPHPExcel->getActiveSheet()->getStyle($col[12] . '2' . ":" . $col[$indexSun] . '2')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) );
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
            //    		    $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun+1] . $focus .":" . $col[$d+11] . $focus)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
            //    		    $objPHPExcel->getActiveSheet()->getStyle($col[$indexSun+1] . '2' .":" . $col[$d+11] . '2')->applyFromArray( array( 'fill' => Style_Fill('FFCCE5') ) ); 
            //    		    $merge_nosun = $col[$indexSun+1] . $focus .":" . $col[$d+11] . $focus;
            //    		    $kla = 0;
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
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('4F6228') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('D8E4BC') ) );//D8E4BC
                                }
                                elseif($key == 1)
                                {
                                                    $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('0F243E') ) );
                                                    $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('B8CCE4') ) );
                                }
                                elseif($key == 2)
                                {
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('512603') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('FCD5B4') ) );
                                }
                                elseif($key == 3)
                                {
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('4B4B4B') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('D9D9D9') ) );
                                }
                                elseif($key == 4)
                                {
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('193300') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('66CC00') ) );
                                }
                                else
                                {
                                                     $objPHPExcel->getActiveSheet()->getStyle($value)->applyFromArray( array( 'fill' => Style_Fill('330019') ) );
                                                     $objPHPExcel->getActiveSheet()->getStyle($fillTotal[$key])->applyFromArray( array( 'fill' => Style_Fill('FF99CC') ) );
                                }

     
    }

}


function input_fill($fillSum=null, $fillTotal=null, $objPHPExcel=null)
{




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

 
