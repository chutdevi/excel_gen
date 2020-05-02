<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli') die('This example should only be run from a Web Browser');
 
require_once dirname(__FILE__, 2) . '/PHPExcel-1.8.1/Classes/PHPExcel.php';
require_once dirname(__FILE__, 2) . '/PHPExcel-1.8.1/function_from_report.php';
$gdImage =   dirname(__FILE__, 2) . '/img/NEW-TBKK-LOGO_0.png';

// var $FJ;
// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$FJ = new FUNCTION_GENERATE( $objPHPExcel ); 

	$formula['FML1'] = "=SUM(%s:%s)";
	$formula['FML2'] = "=%s-COUNTIF(%s:%s,0)";
	$formula['FML3'] = "=SUBTOTAL(9,%s:%s)";
	$formula['FML6'] = "=IFERROR(%s/((%s*60)+(%s*60)), 0)"; 
	//$formula['FML6'] = "=IFERROR(%s/((%s*60)+(%s*60)), 0)";

	$formula['FML5'] = "=IFERROR(%s/%s, 0)";
	$formula['FML4'] = "=IFERROR(%s/(%s+%s), 0)";

	$format['FRM1']  = '_-* #,##0_-;_-* [Red](#,##0)_-;_-* "-"_-;_-@_-';
	$format['FRM2']  = '_-* #,##0.00_-;_-* [Red](#,##0.00)_-;_-* "-"_-;_-@_-';
	$format['FRM3']  = '_-* #,##0.00%_-;_-* [Red](#,##0.00%)_-;_-* "-"_-;_-@_-';
	$format['FRM4']  = '_-* #,##0.00 "hr."_-;_-* [Red](#,##0.00 "hr.")_-;_-* "-"_-;_-@_-';
	$format['FRM5']  = '_-* #,##0 "min."_-;_-* [Red](#,##0 "min.")_-;_-* "-"_-;_-@_-';
	$cn = $FJ->OUTPUT_COLUMNNAMES(55);
	$cx = $FJ->OUTPUT_COLUMNINDEX(55);

#STYLE ACTUAL dataORY
	$rc = $FJ->row_column   = 8;
	$rd = $FJ->row_content  = 10;
	$cs = $FJ->column_start = ( $cn["B"]+1 ); 
	$FJ->column_index  = $cx;
 
	$FJ->amount_column = $ac = count ( $data[0] )+( $cs-1 );
	$cval =  $data[0];
	$FJ->amount_row = $ad = count ( $data )+( $rd-1 );
	$FJ->inx = 0; $FJ->CREATE_SHEET( "ALL SECTION" );
	$FJ->IND(0);
	$FJ->CREATE_FREEZE('N'.$rd );
	$FJ->CREATE_COLORTAB( "c0504d" );
	$FJ->CREATE_ZOOMSCALE(62); 
	$FJ->STYLE_GRIDLINES(False);
	$FJ->CREATE_FILTER( $cx[$cs-1] . ( $rc+1 ), $cx[$ac] . ( $rc+1 ) );

	$FJ->CREATE_HEAD( $cval, true );
	$FJ->CREATE_BODY( $data, true );	


	$FJ->STYLE_GROUP_COLUMN($cn["C"]);
	$FJ->STYLE_GROUP_COLUMN($cn["F"]);
	$FJ->STYLE_GROUP_COLUMN($cn["H"]); 

	foreach( range($cn["K"], $ac) as $c){
		$FJ->CREATE_TEXT( $cx[$c] . '5' , sprintf( $formula["FML3"],  $cx[$c].$rd,  $cx[$c].$ad ) );
	} 
	foreach( range($rd, $ad) as $l){   
		$FJ->CREATE_TEXT( $cx[$cn["L"]]  . $l , sprintf( $formula["FML1"], $cx[$cn["N"]] .$l, $cx[$ac] 		. $l ) ); 
		$FJ->CREATE_TEXT( $cx[$cn["M"]]  . $l , sprintf( "=%s - %s"	     , $cx[$cn["L"]] .$l, $cx[$cn["K"]] . $l ) );  
	}	
	$FJ->STYLE_ALIGNMENT( 'K'. '5' , $cx[$ac]. '5' );
	$FJ->STYLE_ALIGNMENT( 'B'. $rd , $cx[$cn["J"]]. $ad );

	$FJ->CREATE_FORMAT( 'K'. '5' , $cx[$ac]. '5' , $format['FRM1'] );
	$FJ->CREATE_FORMAT( 'K'. $rd , $cx[$ac]. $ad , $format['FRM1'] );
	
	// foreach($hol1 as $i => $v)
	// {
	// 	foreach( range($cn["N"], $ac)  as $c)
	// 		{
	// 			$t = $FJ->GETDATA_ONCELL( $cx[$c] . $rc );
	// 			if( $t == $v["DD"])
	// 				{
	// 					$FJ->CREATE_FILL($cx[$c].'5', $cx[$c].'5', 'f2dcdb' );
	// 					$FJ->CREATE_FILL($cx[$c].$rc, $cx[$c].$rc, 'f2dcdb' );
	// 					$FJ->CREATE_FILL($cx[$c].$rd, $cx[$c].$ad, 'f2dcdb' );
	// 					$FJ->CREATE_TEXT($cx[$c].($rc+1), $v['FND']);
	// 				}
	// 		}
	// }
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->removeSheetByIndex( $FJ->OUTPUT_INDEXSHEET()-1 );

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save($fln);

echo $fln;
 		