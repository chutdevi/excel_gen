<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');
if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
//$rendererName = PHPExcel_Settings::PDF_RENDERER_TCPDF;
$rendererName = PHPExcel_Settings::PDF_RENDERER_MPDF;
//$rendererName = PHPExcel_Settings::PDF_RENDERER_DOMPDF;
//$rendererLibrary = 'tcPDF5.9';
$rendererLibrary = 'mPDF5.4';
//$rendererLibrary = 'domPDF0.6.0beta3';
$rendererLibraryPath = dirname(__FILE__).'/Classes/libraries/PDF/' . $rendererLibrary;  

//============================================================================================= date =================================================================================

$lastmount = substr(date('Y-m-t',strtotime('today')),8, 2);

$myfile = fopen("G:/taks_request_sheet/bin/".'vend.txt', 'w') or die("Unable to open file!");
$str_vend = '';
//echo $yesterdayA."<br>".$yesterdayB."<br>".$yesterdayC; exit;
//var_dump($kla);exit();

foreach ($kla as $ido => $vend) 
{

//============================================================================================= date =================================================================================
    

    ///file_get_contents('http://192.168.161.102/dep_trainer/Api_tool/api_test');
    //echo $str; exit;

$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
                             ->setLastModifiedBy("Maarten Balliauw")
                             ->setTitle("PDF Test Document")
                             ->setSubject("PDF Test Document")
                             ->setDescription("Test document for PDF, generated using PHP classes.")
                             ->setKeywords("pdf php")
                             ->setCategory("Test result file");


//$objPHPExcel->setActiveSheetIndex(0);
$col_name = array();
foreach ( range('A', 'Z') as $cm ) {array_push($col_name, $cm);}
    

foreach ( range('A', 'Z') as $cm ) {array_push($col_name, "A$cm");}
    
foreach ( range('A', 'Z') as $cm ) {array_push($col_name, "B$cm");}

//echo $dateCol . '/' . $MontCol; exit;


//var_dump($kla); exit;
//var_dump($head2); exit;


                // $objPHPExcel->getActiveSheet()
                //     ->getPageSetup()
                //     ->setRowsToRepeatAtTopByStartAndEnd(1, 5);
                // $objPHPExcel->getActiveSheet()
                //     ->setShowGridlines(False);

//=======================================================================================  config Style ================================================================================





            $indSheet = 0;
            foreach ($title as $inTil => $til) 
            {
                    $sheetIndex =  strtolower(str_replace(' ', '_', $title[$inTil]));
                    $objPHPExcel->createSheet();
                    $objPHPExcel->setActiveSheetIndex($inTil);
                    $objPHPExcel->getActiveSheet()->setTitle("$inTil");
                    $objPHPExcel->getActiveSheet()->setShowGridlines(False);
                    $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 30 );
                      $objPHPExcel->getActiveSheet()
                                                ->getStyle(('1:4'))
                                                ->getAlignment()
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 


                    // $objPHPExcel->getActiveSheet()
                    //     ->getPageSetup()
                    //     ->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                    // $objPHPExcel->getActiveSheet()
                    //     ->getPageSetup()
                    //     ->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                    // $objPHPExcel->getActiveSheet()
                    //     ->getPageMargins()->setTop(0.53);
                    // $objPHPExcel->getActiveSheet()
                    //     ->getPageMargins()->setRight(0.39);
                    // $objPHPExcel->getActiveSheet()
                    //     ->getPageMargins()->setLeft(0.39);
                    // $objPHPExcel->getActiveSheet()
                    //     ->getPageMargins()->setBottom(0.53);

                    // $objPHPExcel->getActiveSheet()->getPageSetup()->setFitToPage(true);
                    // $objPHPExcel->getActiveSheet()->getPageSetup()->setFitToWidth(true);
                    // $objPHPExcel->getActiveSheet()->getPageSetup()->setFitToHeight(true);
                    // $objPHPExcel->getActiveSheet()->getPageSetup()->setHorizontalCentered(true);                    


                 $i = 3;
                 $day = 1;


                if ($indSheet < 1) 
                {

                     if(count($list_act_report['report_request_sheet']) > 0 )
                         { 
                                   // if ($key == 'fa_supply_list') 
                                   // {
                                   $objPHPExcel->setActiveSheetIndex($indSheet);
                                    $st_cal = 11 ;                                
                                    $st_dat = 12 ;
                                    $cu_dat = count($list_act_report['report_request_sheet']) ;
                                    //$objPHPExcel->getActiveSheet()->insertNewRowBefore(1,2);
                                    //$objPHPExcel->getActiveSheet()->freezePane('M4');
                                  
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 50 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 50 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 5 )->setRowHeight( 24 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 6 )->setRowHeight( 45 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 7 )->setRowHeight( 24 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 8 )->setRowHeight( 40 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 9 )->setRowHeight( 40 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 10 )->setRowHeight( 40 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 11 )->setRowHeight( 15 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 12 )->setRowHeight( 72 );


                                    //$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80); 
                                     
                                    $r = 13;
                                    $vnd_po = 'temp';
                                    $indCol = 1;
                                    //var_dump($list_act_report); exit;
                                             $objPHPExcel->getActiveSheet()
                                                                        ->getStyle('13:'.($cu_dat+13))
                                                                        ->getAlignment()
                                                                        ->setWrapText(false)
                                                                        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

                                             $objPHPExcel->getActiveSheet()
                                                                        ->getStyle('B12'.':D'.($cu_dat+13))
                                                                        ->getAlignment()
                                                                        ->setWrapText(false)
                                                                        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                                         ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

                                            foreach ($list_act_report['report_request_sheet'] as $nr => $val) 
                                             {

                                                   

                                                                if( $val['ORDER_NO'] != $vnd_po && $val['VEND_CD'] == $vend['VEND_CD'] )
                                                                {
                                                                  
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[1].($r),   $val['ORDER_NO'] );
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[2].($r),   $val['ITEM_CD'] );
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[2].($r+1), $val['ITEM_NAME'] );
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[4].($r),   $val['QTY'] );
                                                                  //$objPHPExcel->getActiveSheet()->setCellValue($col_name[6].($r),   $val['PUCH_ODR_DLV_DATE']  );
                                                                  //$objPHPExcel->getActiveSheet()->setCellValue($col_name[8].($r), 'REMARK' );
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[1].($r).':'.$col_name[1].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[2].($r).':'.$col_name[3].($r)  );
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[2].($r+1).':'.$col_name[3].($r+1) );
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[4].($r).':'.$col_name[5].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[6].($r).':'.$col_name[7].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[8].($r).':'.$col_name[9].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->setCellValue('B6', $val['VEND_NAME']);
                                                                  $objPHPExcel->getActiveSheet()->setCellValue('B8', $val['VEND_CD']);
                                                                  $objPHPExcel->getActiveSheet()->setCellValue('E10', $val['USER_NAME']);

                                                                   $vnd_po = $val['ORDER_NO'];

                                                                   $r = $r+2;
                                                                }
               
                                             }
                                            $r = $r-1 ;
                                         
                                            $objPHPExcel->getActiveSheet()->getPageSetup()->setPrintArea('B2:J'.$r);
                                            $objPHPExcel->getActiveSheet()->setCellValue('I3', 'Created date');
                                            $objPHPExcel->getActiveSheet()->setCellValue('J3', 'page');
                                            $objPHPExcel->getActiveSheet()->setCellValue('B10', 'We sent request as follows.');
                                            $objPHPExcel->getActiveSheet()->setCellValue('B12', 'ORDER NO');
                                            $objPHPExcel->getActiveSheet()->setCellValue('C12', 'ITEM ');
                                            $objPHPExcel->getActiveSheet()->setCellValue('E12', 'QTY');
                                            $objPHPExcel->getActiveSheet()->setCellValue('G12', 'PKG CODE');
                                            $objPHPExcel->getActiveSheet()->setCellValue('I12', 'REMARK');

                                            $objPHPExcel->getActiveSheet()->setCellValue('I4', date('Y/m/d') );
                                            $objPHPExcel->getActiveSheet()->setCellValue('J4', 1 );


                                            $objPHPExcel->getActiveSheet()
                                                ->getStyle(('B12:I12'))
                                                ->getAlignment()
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 


                                            $req = ($indSheet == 1 ) ? "Request sheet(COPY)" : "Request sheet";
                                            $objPHPExcel->getActiveSheet()->setCellValue('C3', $req); 
                                            $objPHPExcel->getActiveSheet()
                                                ->getStyle(('D3'.':'.$col_name[11].'3'))
                                                ->getAlignment()
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 

                                            $objPHPExcel->getActiveSheet()
                                                ->getStyle(('D4'.':'.$col_name[11].'4'))
                                                ->getAlignment()
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // example title 

                                            $objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray(array('font'    => Style_Font(52, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B6')->applyFromArray(array('font'    => Style_Font(36, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B8')->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('E10')->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('I4:J4')->applyFromArray(array('font'    => Style_Font(22, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B12:I12')->applyFromArray(array('font'    => Style_Font(30, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B10:D10')->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('I3:J3')->applyFromArray(array('font'    => Style_Font(24, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B13:'.'J'.($r))->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            
                                           
                                             $objPHPExcel->getActiveSheet()->getStyle('B6:E6')->applyFromArray(array(
                                                                                      'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                             $objPHPExcel->getActiveSheet()->getStyle('B12:'.'J'.($r))->applyFromArray(array(
                                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                             $objPHPExcel->getActiveSheet()->getStyle('I3:'.'J4')->applyFromArray(array(
                                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                                           $objPHPExcel->getActiveSheet()->mergeCells('C3:H4');
                                           $objPHPExcel->getActiveSheet()->mergeCells('B10:D10');
                                           $objPHPExcel->getActiveSheet()->mergeCells('E10:H10');
                                           $objPHPExcel->getActiveSheet()->mergeCells('B6:E6');
                                           $objPHPExcel->getActiveSheet()->mergeCells('C12:D12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('E12:F12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('G12:H12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('I12:J12');

                             
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('8.43');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('50');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('45');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('45');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('18');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('18');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('15');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('15');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('40');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('20');

                                            
                                            

                   
                                 # code...
                }                           


                #-------------------------------------------------------------------------------------------------------------------------------------------


                                    
                              

                 // echo $indSheet; exit;
                        
                $row = 5;
                 } 
                 if ($indSheet == 2) 
                {

                     if(count($list_act_report['report_request_sheet']) > 0 )
                         { 
                                   // if ($key == 'fa_supply_list') 
                                   // {
                                   $objPHPExcel->setActiveSheetIndex($indSheet);
                                    $st_cal = 11 ;                                
                                    $st_dat = 12 ;
                                    $cu_dat = count($list_act_report['report_request_sheet']) ;
                                    //$objPHPExcel->getActiveSheet()->insertNewRowBefore(1,2);
                                    //$objPHPExcel->getActiveSheet()->freezePane('M4');
                                  
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 50 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 50 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 5 )->setRowHeight( 24 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 6 )->setRowHeight( 45 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 7 )->setRowHeight( 24 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 8 )->setRowHeight( 40 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 9 )->setRowHeight( 40 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 10 )->setRowHeight( 40 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 11 )->setRowHeight( 15 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 12 )->setRowHeight( 72 );


                                   // $objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80); 
                                     
                                    $r = 13;
                                    $vnd_po = 'temp';
                                    $indCol = 1;
                                    //var_dump($list_act_report); exit;
                                             $objPHPExcel->getActiveSheet()
                                                                        ->getStyle('13:'.($cu_dat+13))
                                                                        ->getAlignment()
                                                                        ->setWrapText(false)
                                                                        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

                                             $objPHPExcel->getActiveSheet()
                                                                        ->getStyle('B12'.':D'.($cu_dat+13))
                                                                        ->getAlignment()
                                                                        ->setWrapText(false)
                                                                        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                                         ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

                                            foreach ($list_act_report['report_request_sheet'] as $nr => $val) 
                                             {

                                                   

                                                                if( $val['ORDER_NO'] != $vnd_po && $val['VEND_CD'] == $vend['VEND_CD'] )
                                                                {
                                                                  
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[1].($r),   $val['ORDER_NO'] );
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[2].($r),   $val['ITEM_CD'] );
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[2].($r+1), $val['ITEM_NAME'] );
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[4].($r),   $val['QTY'] );
                                                                  //$objPHPExcel->getActiveSheet()->setCellValue($col_name[6].($r),   $val['PUCH_ODR_DLV_DATE']  );
                                                                  //$objPHPExcel->getActiveSheet()->setCellValue($col_name[8].($r), 'REMARK' );
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[1].($r).':'.$col_name[1].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[2].($r).':'.$col_name[3].($r)  );
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[2].($r+1).':'.$col_name[3].($r+1) );
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[4].($r).':'.$col_name[5].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[6].($r).':'.$col_name[7].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[8].($r).':'.$col_name[9].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->setCellValue('B6', $val['VEND_NAME']);
                                                                  $objPHPExcel->getActiveSheet()->setCellValue('B8', $val['VEND_CD']);
                                                                  

                                                                   $vnd_po = $val['ORDER_NO'];

                                                                   $r = $r+2;
                                                                }
               
                                             }
                                            $r = $r-1 ;
                                         
                                            $objPHPExcel->getActiveSheet()->getPageSetup()->setPrintArea('B2:J'.$r);
                                            $objPHPExcel->getActiveSheet()->setCellValue('I3', 'Created date');
                                            $objPHPExcel->getActiveSheet()->setCellValue('J3', 'page');
                                            $objPHPExcel->getActiveSheet()->setCellValue('B10', 'We received request as follows.');
                                            $objPHPExcel->getActiveSheet()->setCellValue('B12', 'ORDER NO');
                                            $objPHPExcel->getActiveSheet()->setCellValue('C12', 'ITEM ');
                                            $objPHPExcel->getActiveSheet()->setCellValue('E12', 'QTY');
                                            $objPHPExcel->getActiveSheet()->setCellValue('G12', 'PKG CODE');
                                            $objPHPExcel->getActiveSheet()->setCellValue('I12', 'REMARK');
                                            $objPHPExcel->getActiveSheet()->setCellValue('I6', 'RECEIVED');

                                            $objPHPExcel->getActiveSheet()->setCellValue('I4', date('Y/m/d') );
                                            $objPHPExcel->getActiveSheet()->setCellValue('J4', 1 );


                                            $objPHPExcel->getActiveSheet()
                                                ->getStyle(('B12:I12'))
                                                ->getAlignment()
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 


                                            $req ="Receiving sheet";
                                            $objPHPExcel->getActiveSheet()->setCellValue('C3', $req); 
                                            $objPHPExcel->getActiveSheet()
                                                ->getStyle(('D3'.':'.$col_name[11].'3'))
                                                ->getAlignment()
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 

                                            $objPHPExcel->getActiveSheet()
                                                ->getStyle(('D4'.':'.$col_name[11].'4'))
                                                ->getAlignment()
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // example title 

                                            $objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray(array('font'    => Style_Font(52, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B6')->applyFromArray(array('font'    => Style_Font(36, '002d4d', false, 'Courier New')));
                                             $objPHPExcel->getActiveSheet()->getStyle('I6')->applyFromArray(array('font'    => Style_Font(36, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B8')->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('E10')->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('I4:J4')->applyFromArray(array('font'    => Style_Font(22, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B12:I12')->applyFromArray(array('font'    => Style_Font(30, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B10:D10')->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('I3:J3')->applyFromArray(array('font'    => Style_Font(24, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B13:'.'J'.($r))->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            
                                           
                                             $objPHPExcel->getActiveSheet()->getStyle('B6:E6')->applyFromArray(array(
                                                                                      'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                             $objPHPExcel->getActiveSheet()->getStyle('B12:'.'J'.($r))->applyFromArray(array(
                                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                             $objPHPExcel->getActiveSheet()->getStyle('I3:'.'J4')->applyFromArray(array(
                                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                             $objPHPExcel->getActiveSheet()->getStyle('I6:'.'I11')->applyFromArray(array(
                                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                                           $objPHPExcel->getActiveSheet()->mergeCells('C3:H4');
                                           $objPHPExcel->getActiveSheet()->mergeCells('B10:D10');
                                           $objPHPExcel->getActiveSheet()->mergeCells('B6:E6');
                                           $objPHPExcel->getActiveSheet()->mergeCells('C12:D12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('E12:F12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('G12:H12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('I12:J12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('I7:I11');

                             
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('8.43');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('50');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('45');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('45');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('18');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('18');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('15');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('15');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('40');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('20');

                                            
                                            

                   
                                 # code...
                }                           


                #-------------------------------------------------------------------------------------------------------------------------------------------


                                    
                              

                 // echo $indSheet; exit;
                        
                $row = 5;
                 } 
                 if ($indSheet == 3) 
                {

                     if(count($list_act_report['report_request_sheet']) > 0 )
                         { 
                                   // if ($key == 'fa_supply_list') 
                                   // {
                                   $objPHPExcel->setActiveSheetIndex($indSheet);
                                    $st_cal = 11 ;                                
                                    $st_dat = 12 ;
                                    $cu_dat = count($list_act_report['report_request_sheet']) ;
                                    //$objPHPExcel->getActiveSheet()->insertNewRowBefore(1,2);
                                    //$objPHPExcel->getActiveSheet()->freezePane('M4');
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 3 )->setRowHeight( 50 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 4 )->setRowHeight( 50 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 5 )->setRowHeight( 24 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 6 )->setRowHeight( 45 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 7 )->setRowHeight( 45 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 8 )->setRowHeight( 40 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 9 )->setRowHeight( 45 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 10 )->setRowHeight( 40 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 11 )->setRowHeight( 15 );
                                     $objPHPExcel->getActiveSheet()->getRowDimension( 12 )->setRowHeight( 72 );


                                    //$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(80); 
                                     
                                    $r = 13;
                                    $vnd_po = 'temp';
                                    $indCol = 1;
                                    //var_dump($list_act_report); exit;
                                             $objPHPExcel->getActiveSheet()
                                                                        ->getStyle('13:'.($cu_dat+13))
                                                                        ->getAlignment()
                                                                        ->setWrapText(false)
                                                                        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

                                             $objPHPExcel->getActiveSheet()
                                                                        ->getStyle('B12'.':D'.($cu_dat+13))
                                                                        ->getAlignment()
                                                                        ->setWrapText(false)
                                                                        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                                         ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

                                            foreach ($list_act_report['report_request_sheet'] as $nr => $val) 
                                             {

                                                   

                                                                if( $val['ORDER_NO'] != $vnd_po && $val['VEND_CD'] == $vend['VEND_CD'] )
                                                                {
                                                                  
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[1].($r),   $val['ORDER_NO'] );
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[2].($r),   $val['ITEM_CD'] );
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[2].($r+1), $val['ITEM_NAME'] );
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[4].($r),   $val['QTY'] );
                                                                  $objPHPExcel->getActiveSheet()->setCellValue($col_name[8].($r),   $val['WH_CD'] );

                                                                  //$objPHPExcel->getActiveSheet()->setCellValue($col_name[6].($r),   $val['PUCH_ODR_DLV_DATE']  );
                                                                  //$objPHPExcel->getActiveSheet()->setCellValue($col_name[8].($r), 'REMARK' );
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[1].($r).':'.$col_name[1].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[2].($r).':'.$col_name[3].($r)  );
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[2].($r+1).':'.$col_name[3].($r+1) );
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[4].($r).':'.$col_name[5].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[6].($r).':'.$col_name[7].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[8].($r).':'.$col_name[9].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->mergeCells($col_name[10].($r).':'.$col_name[11].($r+1));
                                                                  $objPHPExcel->getActiveSheet()->setCellValue('B7', $val['VEND_NAME']);
                                                                  $objPHPExcel->getActiveSheet()->setCellValue('B8', $val['VEND_CD']);
                                                                  $objPHPExcel->getActiveSheet()->setCellValue('G10', $val['USER_NAME']);

                                                                   $vnd_po = $val['ORDER_NO'];

                                                                   $r = $r+2;
                                                                }
               
                                             }
                                            $r = $r-1 ;
                                         
                                            $objPHPExcel->getActiveSheet()->getPageSetup()->setPrintArea('B2:L'.$r);
                                            $objPHPExcel->getActiveSheet()->setCellValue('J3', 'Created date');
                                            $objPHPExcel->getActiveSheet()->setCellValue('L3', 'page');
                                            $objPHPExcel->getActiveSheet()->setCellValue('B6', 'ISSUE FOR');
                                            $objPHPExcel->getActiveSheet()->setCellValue('B9', 'ISSUE DATE');
                                            $objPHPExcel->getActiveSheet()->setCellValue('B12', 'ORDER NO');
                                            $objPHPExcel->getActiveSheet()->setCellValue('C12', 'ITEM');
                                            $objPHPExcel->getActiveSheet()->setCellValue('E12', 'QTY');
                                            $objPHPExcel->getActiveSheet()->setCellValue('G12', 'PKG CODE');
                                            $objPHPExcel->getActiveSheet()->setCellValue('I12', 'ISSUE SS CODE');
                                            $objPHPExcel->getActiveSheet()->setCellValue('K12', 'CHECK');

                                            $objPHPExcel->getActiveSheet()->setCellValue('J4', date('Y/m/d') );
                                            $objPHPExcel->getActiveSheet()->setCellValue('L4', 1 );


                                            $objPHPExcel->getActiveSheet()
                                                ->getStyle(('B12:L12'))
                                                ->getAlignment()
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 


                                            $req = "Issue sheet(logistics) ";
                                            $objPHPExcel->getActiveSheet()->setCellValue('C3', $req); 
                                            $objPHPExcel->getActiveSheet()
                                                ->getStyle(('D3'.':'.$col_name[11].'3'))
                                                ->getAlignment()
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 

                                            $objPHPExcel->getActiveSheet()
                                                ->getStyle(('D4'.':'.$col_name[11].'4'))
                                                ->getAlignment()
                                                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                                                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // example title 

                                            $objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray(array('font'    => Style_Font(52, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B6')->applyFromArray(array('font'    => Style_Font(38, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B7')->applyFromArray(array('font'    => Style_Font(36, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B8')->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B9')->applyFromArray(array('font'    => Style_Font(38, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('G10')->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('J4:L4')->applyFromArray(array('font'    => Style_Font(22, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B12:L12')->applyFromArray(array('font'    => Style_Font(30, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B11:D11')->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('J3:L3')->applyFromArray(array('font'    => Style_Font(24, '002d4d', true, 'Courier New')));
                                            $objPHPExcel->getActiveSheet()->getStyle('B13:'.'L'.($r))->applyFromArray(array('font'    => Style_Font(28, '002d4d', false, 'Courier New')));
                                            
                                           
                                             $objPHPExcel->getActiveSheet()->getStyle('B7:E7')->applyFromArray(array(
                                                                                      'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                              $objPHPExcel->getActiveSheet()->getStyle('B10:E10')->applyFromArray(array(
                                                                                      'borders' => array('bottom' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                             $objPHPExcel->getActiveSheet()->getStyle('B12:'.'L'.($r))->applyFromArray(array(
                                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));
                                             $objPHPExcel->getActiveSheet()->getStyle('J3:'.'L4')->applyFromArray(array(
                                                                                      'borders' => array('allborders' => Style_border(PHPExcel_Style_Border::BORDER_THIN,'000000'))));

                                            $objPHPExcel->getActiveSheet()->mergeCells('J3:K3');
                                            $objPHPExcel->getActiveSheet()->mergeCells('J4:K4');
                                           $objPHPExcel->getActiveSheet()->mergeCells('C3:I4');
                                           $objPHPExcel->getActiveSheet()->mergeCells('G10:J10');
                                           $objPHPExcel->getActiveSheet()->mergeCells('B6:E6');
                                           $objPHPExcel->getActiveSheet()->mergeCells('C12:D12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('E12:F12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('G12:H12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('I12:J12');
                                           $objPHPExcel->getActiveSheet()->mergeCells('K12:L12');

                             
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('8.43');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('50');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('45');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('45');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth('18');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('18');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth('15');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth('15');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth('15');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth('30');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth('15');
                                            $objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth('20');

                                            
                                            

                   
                                 # code...
                }                           


                #-------------------------------------------------------------------------------------------------------------------------------------------


                                    
                              

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


            //$objPHPExcel->setActiveSheetIndex(0);
              
            //$objPHPExcel->removeSheetByIndex(count($title));
            if (!PHPExcel_Settings::setPdfRenderer(
                    $rendererName,
                    $rendererLibraryPath
                )) {
                die(
                    'NOTICE: Please set the $rendererName and $rendererLibraryPath values' .
                    '<br />' .
                    'at the top of this script as appropriate for your directory structure'
                );
            }
            $today = date("My");
            //Redirect output to a client’s web browser (Excel2007)
            // header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            // $con = 'Content-Disposition: attachment;filename='.$filename.date('d').'.xlsx';
            // //echo $con; exit;
            // header($con);
            // header('Cache-Control: max-age=0');
            // // If you're serving to IE 9, then the following may be needed
            // header('Cache-Control: max-age=1');

            // // If you're serving to IE over SSL, then the following may be needed
            // header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            // header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
            // header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            // header ('Pragma: public'); // HTTP/1.0
            $str_vend .= $vend['VEND_CD'].'-'.date('dmy') . "," ;

           // $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            //$objWriter->save('G:/file_export/request_sheet/'.$vend['VEND_CD'].'-'.date('dmy').'.xlsx');
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'PDF');
            //$objWriter->selectSheetsByIndex(0);
            //$objWriter->setSheetIndex(0,1,2,3);
            $objWriter->writeAllSheets();
            $objWriter->save('G:/file_export/request_sheet/'.$vend['VEND_CD'].'-'.date('dmy').'.pdf');
    # code...
}

    $str_vend = substr($str_vend, 0, (strlen($str_vend)-1) ) ;
    fwrite($myfile,$str_vend);
    fclose($myfile);


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






 ?>
