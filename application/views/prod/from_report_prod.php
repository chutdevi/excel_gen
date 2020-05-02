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
#STYLE ALL SECTION AND PD
	$rc = $FJ->row_column   = 9;
	$rd = $FJ->row_content  = 11;
	$cs = $FJ->column_start = ( $cn["B"]+1 ); 
	$FJ->column_index  = $cx;
	$FJ->amount_column = $ac = count ( $data["PD01"][0] )+( $cs-1 );


	$cval =  $data["PD01"][0];
	//$FJ->amount_row    = $ad = count ( $Dcon    )+( $rd-1 );


	$FJ->inx = 0; $FJ->CREATE_SHEET( "ALL SECTION" );
	$FJ->inx = 1; $FJ->CREATE_SHEET( "PD01" );
	$FJ->inx = 2; $FJ->CREATE_SHEET( "PD02" );
	$FJ->inx = 3; $FJ->CREATE_SHEET( "PD03" );
	$FJ->inx = 4; $FJ->CREATE_SHEET( "PD04" );
	$FJ->inx = 5; $FJ->CREATE_SHEET( "PD05" );
	$FJ->inx = 6; $FJ->CREATE_SHEET( "PD06" );
	$FJ->inx = 7; $FJ->CREATE_SHEET( "LG00" );
	$FJ->inx = 8; $FJ->CREATE_SHEET( "ACTUAL HISTORY" );  

	foreach( range(0, 7) as $inx){
 
		$FJ->amount_row = $ad = count ( $data["PD01"] )+( $rd-1 );
		$FJ->IND($inx);
		$FJ->CREATE_FREEZE('F'.$rd );
		$FJ->CREATE_COLORTAB( "4f81bd" );
		$FJ->CREATE_ZOOMSCALE(62); 
		$FJ->STYLE_GRIDLINES(False);
		$FJ->CREATE_FILTER( $cx[$cs-1] . ( $rc+1 ), $cx[$ac] . ( $rc+1 ) );

		$FJ->CREATE_IMAMG('B2', $gdImage, 5, 5, 210, 180);


		$FJ->CREATE_TEXT( $cx[$cn["E"]] . "2" , sprintf( "PRODUCTION REPORT OF %s", strtoupper( date('Y F d', strtotime($days) ))));
		$FJ->CREATE_TEXT( $cx[$cn["E"]] . "3" , sprintf( "TBKK [ Thailand %s ] Pcsystem vol 2.0", date('Y') ));

		$FJ->CREATE_TEXT( $cx[$cn["B"]]  . "5" , sprintf( "%s", $FJ->GETDATA_SHEETNAME() ) );
		$FJ->CREATE_TEXT( $cx[$cn["F"]]  . "5" , sprintf( "PRODUCTION DATA LAST 2 DAYS"));
		$FJ->CREATE_TEXT( $cx[$cn["Z"]]  . "5" , sprintf( "PRODUCTION PLAN"));
		$FJ->CREATE_TEXT( $cx[$cn["AC"]] . "5" , sprintf( "PRODUCTION DATA OF MONTH"));

		$FJ->CREATE_TEXT( $cx[$cn["B"]]  . "6" , sprintf( "SUMMARY DATA", date('Y') ) );
		$FJ->CREATE_TEXT( $cx[$cn["F"]]  . "6" , sprintf( "%s", date('Y-m-d', strtotime("- 2 day", strtotime($days) ) ) ) );
		$FJ->CREATE_TEXT( $cx[$cn["P"]]  . "6" , sprintf( "%s", date('Y-m-d', strtotime("- 1 day", strtotime($days) ) ) ) );
		$FJ->CREATE_TEXT( $cx[$cn["Z"]]  . "6" , sprintf( "%s", date('Y-m-d', strtotime("+ 0 day", strtotime($days) ) ) ) );
		$FJ->CREATE_TEXT( $cx[$cn["AA"]] . "6" , sprintf( "%s", date('Y-m-d', strtotime("+ 1 day", strtotime($days) ) ) ) );
		$FJ->CREATE_TEXT( $cx[$cn["AB"]] . "6" , sprintf( "%s", date('Y-m-d', strtotime("+ 2 day", strtotime($days) ) ) ) );
		$FJ->CREATE_TEXT( $cx[$cn["AC"]] . "6" , sprintf( "ACCUMULATE DATA FROM %s TO %s", date('Y-m-01', strtotime("+ 0 day", strtotime($days))), date('Y-m-d', strtotime("- 0 day", strtotime($days)))  ));  

		$FJ->CREATE_FONT( $cx[$cn["E"]] . "2" ,$cx[$cn["E"]] . "2", 58, 'FFFFFF', true );
		$FJ->CREATE_FONT( $cx[$cn["E"]] . "3" ,$cx[$cn["E"]] . "3", 10, 'FFFFFF', true );

		$FJ->CREATE_FONT( $cx[$cn["B"]] . "5" ,$cx[$cn["B"]] . "5", 28, '404040', true );
		$FJ->CREATE_FONT( $cx[$cn["B"]] . "6" ,$cx[$cn["B"]] . "6", 20, '404040', true );

		$FJ->CREATE_FONT( $cx[$cn["F"]] . "5" ,$cx[$ac] . "5", 18, 'FFFFFF', true );
		$FJ->CREATE_FONT( $cx[$cn["F"]] . "6" ,$cx[$ac] . "6", 16, 'FFFFFF', true );
		$FJ->CREATE_FONT( $cx[$cn["F"]] . "7" ,$cx[$ac] . "7", 16, '36608D', true );
		$FJ->CREATE_FONT( $cx[$cn["B"]] . $rc ,$cx[$ac] . $rc, 14, 'FFFFFF', true );



		$FJ->CREATE_FILL( $cx[ $cn["B"] ] . '2', $cx[ $cn["D"] ] . '3', 'bfbfbf' );
		$FJ->CREATE_FILL( $cx[ $cn["E"] ] . '2', $cx[ $ac ]      . '3', '36608d' );
		$FJ->CREATE_FILL( $cx[ $cn["B"] ] . '5', $cx[ $cn["E"] ] . '7', 'bfbfbf' );
		$FJ->CREATE_FILL( $cx[$cs-1] . ( $rc+1 ) , $cx[ $ac ] .( $rc+1 ), '808080' );
		
		$FJ->CREATE_FILL( $cx[$cn["F"]] . '5' , $cx[$cn["F"]] . '5', '31869b' );
		$FJ->CREATE_FILL( $cx[$cn["Z"]] . '5' , $cx[$cn["Z"]] . '5', '36608d' );
		$FJ->CREATE_FILL( $cx[$cn["AC"]] .'5' , $cx[$cn["AC"]]. '5', '31869b' );

		$FJ->CREATE_FILL( $cx[$cn["F"]] . '6' , $cx[$cn["F"]] . '6', '538dca' );
		$FJ->CREATE_FILL( $cx[$cn["P"]] . '6' , $cx[$cn["P"]] . '6', '31869b' );
		$FJ->CREATE_FILL( $cx[$cn["Z"]] . '6' , $cx[$cn["AB"]]. '6' , '36608d' );
		$FJ->CREATE_FILL( $cx[$cn["AC"]] . '6' , $cx[$cn["AC"]] . '6' , '538dca' );


		$FJ->CREATE_FILL( $cx[$cn["B"]] . $rc , $cx[$cn["E"]] . $rc, '36608d' );

		$FJ->CREATE_FILL( $cx[$cn["F"]] . $rc , $cx[$cn["O"]] . $rc, '538dca' );
		$FJ->CREATE_FILL( $cx[$cn["P"]] . $rc , $cx[$cn["Y"]] . $rc, '31869b' );
		$FJ->CREATE_FILL( $cx[$cn["Z"]] . $rc  , $cx[$cn["AB"]] . $rc, '538dca' );
		$FJ->CREATE_FILL( $cx[$cn["AC"]] . $rc , $cx[$cn["AO"]] . $rc, '31869b' );

		$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'2', $cx[ $ac ].'3', 'BFBFBF', "outline", 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'5', $cx[ $ac ].'7', 'BFBFBF', "outline", 'BORDER_THICK' );

		$FJ->CREATE_BORDER( $cx[ $cn["Z"] ] .'6', $cx[ $cn["AB"] ].'6', 'BFBFBF', "inside" , 'BORDER_THIN' );
		$FJ->CREATE_BORDER( $cx[ $cn["F"] ] .'7', $cx[ $cn["AG"] ].'7', 'BFBFBF', "inside" , 'BORDER_THIN' );
		$FJ->CREATE_BORDER( $cx[ $cn["AH"] ].'7', $cx[ $ac ]      .'7', 'BFBFBF', "inside" , 'BORDER_THIN' );
		$FJ->CREATE_BORDER( $cx[ $cn["B"] ].$rc,  $cx[ $cn["AG"] ].$rc, 'FFFFFF', "inside" , 'BORDER_THIN' );
		$FJ->CREATE_BORDER( $cx[ $cn["AH"] ].$rc, $cx[ $ac ]      .$rc, 'FFFFFF', "inside" , 'BORDER_THIN' );		

		$FJ->CREATE_BORDER( $cx[ $cn["F"] ].'5', $cx[ $cn["Y"] ].'5'  , 'BFBFBF', "outline", 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["Z"] ].'5', $cx[ $cn["AB"] ].'5' , 'BFBFBF', "outline", 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["AC"] ].'5', $cx[ $cn["AO"] ].'5', 'BFBFBF', "outline", 'BORDER_THICK' );

		$FJ->CREATE_BORDER( $cx[ $cn["F"] ].'6', $cx[ $cn["O"] ] .'7', 'BFBFBF', "outline", 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["P"] ].'6', $cx[ $cn["Y"] ] .'7', 'BFBFBF', "outline", 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["Z"] ].'6', $cx[ $cn["AB"] ].'7', 'BFBFBF', "outline", 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["AC"] ].'6', $cx[ $cn["AO"] ].'7','BFBFBF', "outline", 'BORDER_THICK' );

		$FJ->CREATE_BORDER( $cx[ $cn["AC"] ].'5', $cx[ $cn["AO"] ].'5','BFBFBF', "outline", 'BORDER_THICK' );

		$FJ->CREATE_BORDER( $cx[ $cn["F"]  ].'7', $cx[ $ac ] .'7', 'BFBFBF', "top" , 'BORDER_THIN' );	
		
		$FJ->CREATE_BORDER( $cx[ $cn["E"] ].$rc , $cx[ $cn["E"] ].$rc, 'BFBFBF', "right", 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["O"] ].$rc , $cx[ $cn["O"] ].$rc, 'BFBFBF', "right", 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["Z"] ].$rc , $cx[ $cn["Z"] ].$rc, 'BFBFBF', "right", 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["AB"] ].$rc, $cx[ $cn["AB"] ].$rc,'BFBFBF', "right", 'BORDER_THICK' );

		//$FJ->STYLE_WIDTH_RANGE( $cn["F"], $ac, 10);
		$FJ->STYLE_WIDTH("A",  4);
		$FJ->STYLE_WIDTH("B",  9);
		$FJ->STYLE_WIDTH("C", 12);
		$FJ->STYLE_WIDTH("D", 22);
		$FJ->STYLE_WIDTH("E", 55.5);
		$FJ->STYLE_WIDTH_RANGE( $cn["F"],  $cn["Y"] ,  13);
		$FJ->STYLE_WIDTH_RANGE( $cn["Z"],  $cn["AB"],  19);
		$FJ->STYLE_WIDTH_RANGE( $cn["AC"], $cn["AG"],  15);
		$FJ->STYLE_WIDTH("AH", 5);
		$FJ->STYLE_WIDTH_RANGE( $cn["AI"], $cn["AO"],  15);
		$FJ->STYLE_WIDTH_RANGE( $cn["J"],  $cn["N"] ,  20);
		$FJ->STYLE_WIDTH_RANGE( $cn["T"],  $cn["X"] ,  20);
		$FJ->STYLE_WIDTH_RANGE( $cn["AI"],  $cn["AM"] ,  20);
		$FJ->STYLE_HEIGHT( 1, 20 );	
		$FJ->STYLE_HEIGHT( 2, 128);
		$FJ->STYLE_HEIGHT( 3, 18 );
		$FJ->STYLE_HEIGHT( 4, 10 );
		$FJ->STYLE_HEIGHT( 5, 35 );
		$FJ->STYLE_HEIGHT( 6, 38 );
		$FJ->STYLE_HEIGHT( 7, 35 );
		$FJ->STYLE_HEIGHT( 8, 10 );
		$FJ->STYLE_HEIGHT( 9, 48 );
		$FJ->STYLE_HEIGHT( $rc+1, 16.5); 		

		$FJ->CREATE_HEAD( $cval, true );

		$FJ->CREATE_TEXT( $cx[$cn["C"]]  . $rc , "PD" );
		$FJ->CREATE_TEXT( $cx[$cn["D"]]  . $rc , "LINE CD" );
		$FJ->CREATE_TEXT( $cx[$cn["E"]]  . $rc , "LINE NAME" );
		$FJ->CREATE_TEXT( $cx[$cn["F"]]  . $rc , "PLAN" );
		$FJ->CREATE_TEXT( $cx[$cn["G"]]  . $rc , "ACTUAL" );
		$FJ->CREATE_TEXT( $cx[$cn["H"]]  . $rc , "DIFF." );
		$FJ->CREATE_TEXT( $cx[$cn["I"]]  . $rc , "NG." );
		$FJ->CREATE_TEXT( $cx[$cn["J"]]  . $rc , "TOTAL TIME" );
		$FJ->CREATE_TEXT( $cx[$cn["K"]]  . $rc , "BREAK TIME" );
		$FJ->CREATE_TEXT( $cx[$cn["L"]]  . $rc , "LOSS" );
		$FJ->CREATE_TEXT( $cx[$cn["M"]]  . $rc , "WORK TIME" );
		$FJ->CREATE_TEXT( $cx[$cn["N"]]  . $rc , "CYCLE TIME" );
		$FJ->CREATE_TEXT( $cx[$cn["O"]]  . $rc , "EFF." );
		$FJ->CREATE_TEXT( $cx[$cn["P"]]  . $rc , "PLAN" );
		$FJ->CREATE_TEXT( $cx[$cn["Q"]]  . $rc , "ACTUAL" );
		$FJ->CREATE_TEXT( $cx[$cn["R"]]  . $rc , "DIFF." );
		$FJ->CREATE_TEXT( $cx[$cn["S"]]  . $rc , "NG." );
		$FJ->CREATE_TEXT( $cx[$cn["T"]]  . $rc , "TOTAL TIME" );
		$FJ->CREATE_TEXT( $cx[$cn["U"]]  . $rc , "BREAK TIME" );
		$FJ->CREATE_TEXT( $cx[$cn["V"]]  . $rc , "LOSS" );
		$FJ->CREATE_TEXT( $cx[$cn["W"]]  . $rc , "WORK TIME" );
		$FJ->CREATE_TEXT( $cx[$cn["X"]]  . $rc , "CYCLE TIME" );
		$FJ->CREATE_TEXT( $cx[$cn["Y"]]  . $rc , "EFF." );
		$FJ->CREATE_TEXT( $cx[$cn["Z"]]  . $rc , "PLAN" );
		$FJ->CREATE_TEXT( $cx[$cn["AA"]] . $rc , "PLAN" );
		$FJ->CREATE_TEXT( $cx[$cn["AB"]] . $rc , "PLAN" );
		$FJ->CREATE_TEXT( $cx[$cn["AC"]] . $rc , "ACCUM. PLAN" );
		$FJ->CREATE_TEXT( $cx[$cn["AD"]] . $rc , "ACCUM. ACTUAL" );
		$FJ->CREATE_TEXT( $cx[$cn["AE"]] . $rc , "ACCUM. DIFF." );
		$FJ->CREATE_TEXT( $cx[$cn["AF"]] . $rc , "ACCUM. NG." );
		$FJ->CREATE_TEXT( $cx[$cn["AG"]] . $rc , "DEFECT PERCENT" );
		$FJ->CREATE_TEXT( $cx[$cn["AH"]] . $rc , "(%)" );
		$FJ->CREATE_TEXT( $cx[$cn["AI"]] . $rc , "ACCUM. TOTAL TIME" );
		$FJ->CREATE_TEXT( $cx[$cn["AJ"]] . $rc , "ACCUM. BREAK TIME" );
		$FJ->CREATE_TEXT( $cx[$cn["AK"]] . $rc , "ACCUM. LOSS" );
		$FJ->CREATE_TEXT( $cx[$cn["AL"]] . $rc , "ACCUM. WORK TM" );
		$FJ->CREATE_TEXT( $cx[$cn["AM"]] . $rc , "CYCEL TIME  MONTH" );
		$FJ->CREATE_TEXT( $cx[$cn["AN"]] . $rc , "EFF.  MONTH" );
		$FJ->CREATE_TEXT( $cx[$cn["AO"]] . $rc , "PLAN THIS MONTH" );

 
		$FJ->CREATE_COMMENT('F1' ,sprintf("Click button on top to unhide \"LINE NAME\"" ) );
		$FJ->CREATE_COMMENT('P1' ,sprintf("Click button on top to unhide \"Total detail of %s\""      ,date('Y-m-d', strtotime("- 2 day", strtotime($days) ) ) ) );
		$FJ->CREATE_COMMENT('Z1' ,sprintf("Click button on top to unhide \"Total detail of %s\""      ,date('Y-m-d', strtotime("- 1 day", strtotime($days) ) ) ) );
		$FJ->CREATE_COMMENT('AC1',sprintf("Click button on top to unhide \"Plan of %s\""              ,date('Y-m-d', strtotime("+ 2 day", strtotime($days) ) ) ) );
		$FJ->CREATE_COMMENT('AO1',sprintf("Click button on top to unhide \"Total detail accum of %s\"",date('Y-m', strtotime("- 0 day"  , strtotime($days) ) ) ), 100, 350 );
		// date('Y-m-d', strtotime("- 2 day", strtotime($days) ) )
		// date('Y-m-d', strtotime("- 1 day", strtotime($days) ) )
		// date('Y-m-d', strtotime("+ 0 day", strtotime($days) ) )
		// date('Y-m-d', strtotime("+ 1 day", strtotime($days) ) )
		// date('Y-m-d', strtotime("+ 2 day", strtotime($days) ) )
		$FJ->CREATE_MERGECELL('B' . '2' ,'D' . '3');
		$FJ->CREATE_MERGECELL('E' . '2' , $cx[$ac] . '2');
		$FJ->CREATE_MERGECELL('E' . '3' , $cx[$ac] . '3');
		$FJ->CREATE_MERGECELL('B' . '5' ,'E' . '5');
		$FJ->CREATE_MERGECELL('B' . '6' ,'E' . '7');
		$FJ->CREATE_MERGECELL('F' . '5' ,'Y' . '5');
		$FJ->CREATE_MERGECELL('Z' . '5' ,'AB'. '5');
		$FJ->CREATE_MERGECELL('AC'. '5' ,'AO'. '5');
		$FJ->CREATE_MERGECELL('F' . '6' ,'O' . '6');
		$FJ->CREATE_MERGECELL('P' . '6' ,'Y' . '6');
		$FJ->CREATE_MERGECELL('AC'. '6' ,'AO'. '6'); 

		$FJ->STYLE_ALIGNMENT( 'E2' , 'E3', 'CL' ); 
		$FJ->STYLE_ALIGNMENT( 'B5' , $cx[$ac] . '7');
		$FJ->STYLE_ALIGNMENT( 'B'.$rc , $cx[$ac] . $rc, 'CC', true );

		
		$FJ->STYLE_GROUP_COLUMN( $cn["E"] );
		$FJ->STYLE_GROUP_COLUMN( $cn["AB"]);

		$FJ->STYLE_GROUP_COLUMNS($cn["J"] , $cn["O"]);
		$FJ->STYLE_GROUP_COLUMNS($cn["T"] , $cn["Y"]);
		$FJ->STYLE_GROUP_COLUMNS($cn["AI"], $cn["AN"]);
		$FJ->STYLE_GROUP_COLUMN( $cn["N"], 2);
		$FJ->STYLE_GROUP_COLUMN( $cn["X"], 2);
		$FJ->STYLE_GROUP_COLUMN( $cn["AM"],2);

	}

#SHEET ALL SECTION
		$FJ->IND(0);
		$FJ->CREATE_COLORTAB( "8064a2" );
		$FJ->amount_row  = $ad = count ( $data["PD01"] ) + $rd;
		$FJ->CREATE_BODY( $data["PD01"] );
		
		$FJ->row_content = $ad;
		$FJ->amount_row  = $ad += count ( $data["PD02"] ); 
		$FJ->CREATE_BODY( $data["PD02"] );
		
		$FJ->row_content = $ad;
		$FJ->amount_row  = $ad += count ( $data["PD03"] );
		$FJ->CREATE_BODY( $data["PD03"] );

		$FJ->row_content = $ad;
		$FJ->amount_row  = $ad += count ( $data["PD04"] );
		$FJ->CREATE_BODY( $data["PD04"] );

		$FJ->row_content = $ad;
		$FJ->amount_row  = $ad += count ( $data["PD05"] );
		$FJ->CREATE_BODY( $data["PD05"] );

		$FJ->row_content = $ad;
		$FJ->amount_row  = $ad += count ( $data["PD06"] );
		$FJ->CREATE_BODY( $data["PD06"] );

		$FJ->row_content = $ad;
		$FJ->amount_row  = $ad += count ( $data["LG00"] );
		$FJ->CREATE_BODY( $data["LG00"] );


		$FJ->row_content = $ad;
		$FJ->amount_row  = $ad += count ( $data["PL00"] );
		$FJ->CREATE_BODY( $data["PL00"] );						
		// foreach( range($rd, $ad) as $l){
		// 	$FJ->CREATE_TEXT( $cx[$cn["N"]] . $l , sprintf( "=%s + %s", $cx[$cn["J"]].$l, $cx[$cn["K"]].$l ) );
		// }		
	//$dm = date('F Y', strtotime($days) );
	//$lm = date('t',strtotime($days) ); 
 
 
#SHEET PD
	$rd = $FJ->row_content = 11;

	foreach(range(0,7) as $h){
		$FJ->IND($h);
		if( $h > 0){
			$FJ->amount_row  = $ad = count ( $data[$FJ->GETDATA_SHEETNAME()] ) + $rd-1;
			$FJ->CREATE_BODY( $data[$FJ->GETDATA_SHEETNAME()], true );
			$FJ->STYLE_GROUP_COLUMN( $cn["C"] ); 
		}else{ 
			$ad = $ad-1;
			foreach( range($rd, $ad) as $l){
				$FJ->CREATE_TEXT( $cx[$cn["B"]] . $l  , sprintf( "=SUBTOTAL(3,$%s$%s:$%s$%s)", $cx[$cn["C"]], $rd, $cx[$cn["C"]], $l ) ); 
			}	
		}

		$FJ->CREATE_BORDER( $cx[ $cn["B"] ] . $rd , $cx[ $ac ] . $ad, 'BFBFBF', "inside", 'BORDER_THIN' ); 
		$FJ->CREATE_BORDER( $cx[ $cn["B"] ]  . ($rd-1) , $cx[ $cn["E"] ] . $ad, '808080', "outline" , 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["F"] ]  . ($rd-1) , $cx[ $cn["O"] ] . $ad, '808080', "outline" , 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["P"] ]  . ($rd-1) , $cx[ $cn["Y"] ] . $ad, '808080', "outline" , 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["Z"] ]  . ($rd-1) , $cx[ $cn["AB"] ]. $ad, '808080', "outline" , 'BORDER_THICK' );
		$FJ->CREATE_BORDER( $cx[ $cn["AC"] ] . ($rd-1) , $cx[ $cn["AO"] ]. $ad, '808080', "outline" , 'BORDER_THICK' );

		$FJ->CREATE_FONT(  $cx[ $cn["B"] ]  . $rd , $cx[ $ac ] . $ad, 14 );

		$FJ->STYLE_HEIGHT_RANGE( $rd, $ad,  22.5); 
		foreach( range($cn["F"], $cn["AO"]) as $c){
			$FJ->CREATE_TEXT( $cx[$c] . '7' , sprintf( $formula["FML3"],  $cx[$c].$rd,  $cx[$c].$ad ) );
		}
		foreach( range($cn["J"], $cn["M"]) as $c){
			$FJ->CREATE_TEXT( $cx[$c] . '7' , sprintf( $formula["FML3"]."/60",  $cx[$c].$rd,  $cx[$c].$ad ) );
		}
		foreach( range($cn["T"], $cn["W"]) as $c){
			$FJ->CREATE_TEXT( $cx[$c] . '7' , sprintf( $formula["FML3"]."/60",  $cx[$c].$rd,  $cx[$c].$ad ) );
		}  
		foreach( range($cn["AI"], $cn["AL"]) as $c){
			$FJ->CREATE_TEXT( $cx[$c] . '7' , sprintf( $formula["FML3"]."/60",  $cx[$c].$rd,  $cx[$c].$ad ) );
		}     
		foreach( range($rd, $ad) as $l){  
			$FJ->CREATE_TEXT( $cx[$cn["M"]]  . $l , sprintf( "=%s - %s - %s", $cx[$cn["J"]] .$l, $cx[$cn["K"]] .$l, $cx[$cn["L"]] .$l ) );
			$FJ->CREATE_TEXT( $cx[$cn["W"]]  . $l , sprintf( "=%s - %s - %s", $cx[$cn["T"]] .$l, $cx[$cn["U"]] .$l, $cx[$cn["V"]] .$l ) );
			$FJ->CREATE_TEXT( $cx[$cn["AL"]] . $l , sprintf( "=%s - %s - %s", $cx[$cn["AI"]].$l, $cx[$cn["AJ"]].$l, $cx[$cn["AK"]] .$l ) );
			$FJ->CREATE_TEXT( $cx[$cn["AG"]] . $l , sprintf( $formula["FML5"]." * 100", $cx[$cn["AF"]].$l, $cx[$cn["AD"]].$l ) );

			$FJ->CREATE_TEXT( $cx[$cn["O"]]  . $l , sprintf( $formula["FML4"], $cx[$cn["N"]]  .$l, $cx[$cn["L"]] .$l, $cx[$cn["M"]] .$l ) );
			$FJ->CREATE_TEXT( $cx[$cn["Y"]]  . $l , sprintf( $formula["FML4"], $cx[$cn["X"]]  .$l, $cx[$cn["V"]] .$l, $cx[$cn["W"]] .$l ) );
			$FJ->CREATE_TEXT( $cx[$cn["AN"]] . $l , sprintf( $formula["FML4"], $cx[$cn["AM"]] .$l, $cx[$cn["AK"]].$l, $cx[$cn["AL"]].$l ) ); 


			$str_comt = " ข้อมูลที่แสดง เป็นข้อมูลที่เกิดจากการ \n นำ จำนวนการผลิต( %s ) x ( ผลรวม cycletime part ที่ทำการผลิต )( %s ) \n ข้อมุล cycle per line = %s";
			if( $FJ->GETDATA_ONCELL(  $cx[$cn["G"]].$l ) > 0 )
				$FJ->CREATE_COMMENT('N' . $l,  sprintf(  $str_comt , $FJ->GETDATA_ONCELL(  $cx[$cn["G"]].$l ), $FJ->GETDATA_ONCELL(  $cx[$cn["N"]].$l ), ( $FJ->GETDATA_ONCELL( $cx[$cn["N"]].$l ) / $FJ->GETDATA_ONCELL ($cx[$cn["G"]].$l ) ) )  ,130, 360);
			
			if( $FJ->GETDATA_ONCELL(  $cx[$cn["Q"]].$l ) > 0 )
				$FJ->CREATE_COMMENT('X' . $l,  sprintf(  $str_comt , $FJ->GETDATA_ONCELL(  $cx[$cn["Q"]].$l ), $FJ->GETDATA_ONCELL(  $cx[$cn["X"]].$l ), ( $FJ->GETDATA_ONCELL( $cx[$cn["X"]].$l ) / $FJ->GETDATA_ONCELL( $cx[$cn["Q"]].$l ) ) )  ,130, 360);
			
			if( $FJ->GETDATA_ONCELL(  $cx[$cn["AC"]].$l ) > 0 )
				$FJ->CREATE_COMMENT('AM' . $l, sprintf(  $str_comt , $FJ->GETDATA_ONCELL(  $cx[$cn["AC"]].$l ), $FJ->GETDATA_ONCELL( $cx[$cn["AM"]].$l ),( $FJ->GETDATA_ONCELL( $cx[$cn["AM"]].$l) / $FJ->GETDATA_ONCELL( $cx[$cn["AC"]].$l ) ) ) ,130, 360);						
		}

			$FJ->CREATE_COMMENT('N'.'7', sprintf("ข้อมูลที่แสดง เป็นข้อมูลที่เกิดจากการ \n นำ จำนวนการผลิต x ( ผลรวม cycletime part ที่ทำการผลิต )") );

			$FJ->CREATE_COMMENT('X'.'7', sprintf("ข้อมูลที่แสดง เป็นข้อมูลที่เกิดจากการ \n นำ จำนวนการผลิต x ( ผลรวม cycletime part ที่ทำการผลิต )") );

			$FJ->CREATE_COMMENT('AM'.'7',sprintf("ข้อมูลที่แสดง เป็นข้อมูลที่เกิดจากการ \n นำ จำนวนการผลิต x ( ผลรวม cycletime part ที่ทำการผลิต)" ) );	

		$FJ->CREATE_TEXT( $cx[$cn["O"]]  . '7' , sprintf( $formula["FML6"], $cx[$cn["N"]]  .'7', $cx[$cn["L"]] .'7', $cx[$cn["M"]] .'7' ) );
		$FJ->CREATE_TEXT( $cx[$cn["Y"]]  . '7' , sprintf( $formula["FML6"], $cx[$cn["X"]]  .'7', $cx[$cn["V"]] .'7', $cx[$cn["W"]] .'7' ) );
		$FJ->CREATE_TEXT( $cx[$cn["AN"]] . '7' , sprintf( $formula["FML6"], $cx[$cn["AM"]] .'7', $cx[$cn["AK"]].'7', $cx[$cn["AL"]].'7' ) ); 

		$FJ->CREATE_TEXT( $cx[$cn["AG"]] . '7' , sprintf( $formula["FML5"]." * 100", $cx[$cn["AF"]] .'7', $cx[$cn["AD"]].'7' ) );
		$FJ->CREATE_TEXT( $cx[$cn["AH"]] . '7' , sprintf( "%%" ) ); 

		$FJ->STYLE_ALIGNMENT( 'B'. $rd , 'E' . $ad );

		$FJ->CREATE_FORMAT(   'F' . '7' , $cx[$ac]. '7' , $format['FRM1'] ); 
		$FJ->CREATE_FORMAT(   'N' . '7' , 'N' . '7' , $format['FRM2'] );	
		$FJ->CREATE_FORMAT(   'O' . '7' , 'O' . '7' , $format['FRM3'] );
		$FJ->CREATE_FORMAT(   'X' . '7' , 'X' . '7' , $format['FRM2'] );	
		$FJ->CREATE_FORMAT(   'Y' . '7' , 'Y' . '7' , $format['FRM3'] );
		$FJ->CREATE_FORMAT(   'AG'. '7' , 'AG'. '7' , $format['FRM2'] );
		$FJ->CREATE_FORMAT(   'AM'. '7' , 'AM'. '7' , $format['FRM2'] );
		$FJ->CREATE_FORMAT(   'AN'. '7' , 'AN'. '7' , $format['FRM3'] );
		
		$FJ->CREATE_FORMAT(   'B' . $rd , $cx[$ac]. $ad , $format['FRM1'] );
		$FJ->CREATE_FORMAT(   'N' . $rd , 'N' . $ad , $format['FRM2'] );	
		$FJ->CREATE_FORMAT(   'O' . $rd , 'O' . $ad , $format['FRM3'] );
		$FJ->CREATE_FORMAT(   'X' . $rd , 'X' . $ad , $format['FRM2'] );	
		$FJ->CREATE_FORMAT(   'Y' . $rd , 'Y' . $ad , $format['FRM3'] );
		$FJ->CREATE_FORMAT(   'AG'. $rd , 'AG'. $ad , $format['FRM2'] );
		$FJ->CREATE_FORMAT(   'AM'. $rd , 'AM'. $ad , $format['FRM2'] );
		$FJ->CREATE_FORMAT(   'AN'. $rd , 'AN'. $ad , $format['FRM3'] );
		
		$FJ->CREATE_FORMAT(   'J'. $rd , 'M'. $ad , $format['FRM5'] );
		$FJ->CREATE_FORMAT(   'T'. $rd , 'W'. $ad , $format['FRM5'] );
		$FJ->CREATE_FORMAT(   'AI'.$rd , 'AL'.$ad , $format['FRM5'] );

		$FJ->CREATE_FORMAT(   'J'. '7' , 'M'. '7' , $format['FRM4'] );
		$FJ->CREATE_FORMAT(   'T'. '7' , 'W'. '7' , $format['FRM4'] );
		$FJ->CREATE_FORMAT(   'AI'.'7' , 'AL'.'7' , $format['FRM4'] );
	}  

#STYLE ACTUAL HISTORY
	$rc = $FJ->row_column   = 8;
	$rd = $FJ->row_content  = 10;
	$cs = $FJ->column_start = ( $cn["B"]+1 ); 
	$FJ->column_index  = $cx;
 
	$FJ->amount_column = $ac = count ( $hist[0] )+( $cs-1 ); 
	$cval =  $hist[0];
	$FJ->amount_row = $ad = count ( $hist )+( $rd-1 );

	$FJ->IND(8);
	$FJ->CREATE_FREEZE('N'.$rd );
	$FJ->CREATE_COLORTAB( "c0504d" );
	$FJ->CREATE_ZOOMSCALE(62); 
	$FJ->STYLE_GRIDLINES(False);
	$FJ->CREATE_FILTER( $cx[$cs-1] . ( $rc+1 ), $cx[$ac] . ( $rc+1 ) );

	$FJ->CREATE_TEXT( $cx[$cn["B"]] . "2" , sprintf( "TBKK [ Thailand ]" ));
	$FJ->CREATE_TEXT( $cx[$cn["F"]] . "2" , sprintf( "HISTORY PRODUCTION ACTUAL  OF %s",  strtoupper( date('Y F', strtotime($days) )))); 
	$FJ->CREATE_TEXT( $cx[$cn["B"]] . "4" , sprintf( "SUMMARY DATA"  ) );

	$FJ->CREATE_TEXT( $cx[$cn["F"]]  . "5" , sprintf( "PRODUCTION DATA LAST 2 DAYS"));
	$FJ->CREATE_TEXT( $cx[$cn["Z"]]  . "5" , sprintf( "PRODUCTION PLAN"));
	$FJ->CREATE_TEXT( $cx[$cn["AC"]] . "5" , sprintf( "PRODUCTION DATA OF MONTH"));

  

	$FJ->CREATE_FONT( $cx[$cn["B"]] . "2" ,$cx[$cn["B"]] . "2", 28, 'FFFFFF', true, false, 'Arial Black' );
	$FJ->CREATE_FONT( $cx[$cn["F"]] . "2" ,$cx[$cn["F"]] . "2", 34, '36608d', true, false, 'Arial Narrow' );

	$FJ->CREATE_FONT( $cx[$cn["B"]] . "4" ,$cx[$cn["B"]] . "4", 36, 'FFFFFF', true, false, 'Arial Black');
	$FJ->CREATE_FONT( $cx[$cn["K"]] . "5" ,$cx[$ac] . "5", 16, '36608d', true );

	$FJ->CREATE_FONT( $cx[$cn["B"]] . $rc ,$cx[$cn["M"]] . $rc, 16, 'FFFFFF', true );
	$FJ->CREATE_FONT( $cx[$cn["K"]] . $rc ,$cx[$ac] . $rc, 16, '36608d', true );	

	$FJ->CREATE_FONT( $cx[$cn["B"]] . $rc ,$cx[$cn["M"]] . $rc, 16, 'FFFFFF', true );
	$FJ->CREATE_FONT( $cx[$cn["N"]] . $rc ,$cx[$ac] . $rc, 16, '36608d', true );
	$FJ->CREATE_FONT(  $cx[ $cn["B"] ]  . ($rc+1) , $cx[ $ac ] . ($rc+1), 12,'ffff33', true );
	$FJ->CREATE_FONT(  $cx[ $cn["B"] ]  . $rd , $cx[ $ac ] . $ad, 14 );


	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . '2', $cx[ $cn["B"] ] . '2', '36608d' );
	$FJ->CREATE_FILL( $cx[ $cn["F"] ] . '2', $cx[ $cn["F"] ] . '2', 'f2f2f2' );
	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . '4', $cx[ $cn["B"] ] . '4', '36608d' );
	$FJ->CREATE_FILL( $cx[ $cn["K"] ] . '4', $cx[ $ac ] . '4', '36608d' ); 
	$FJ->CREATE_FILL( $cx[ $cn["K"] ] . '5', $cx[ $ac ] . '5', 'f2f2f2' ); 
	$FJ->CREATE_FILL( $cx[ $cn["K"] ] . '6', $cx[ $ac ] . '6', '36608d' ); 



	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . $rc, $cx[ $cn["M"] ] . $rc, '36608d' );
	$FJ->CREATE_FILL( $cx[ $cn["N"] ] . $rc, $cx[ $ac ] . $rc, 'f2f2f2' );   
	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . ($rc+1),  $cx[ $ac ] . ($rc+1), '808080' ); 
	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . $rd, $cx[ $ac ] . $ad, 'f2f2f2' ); 

	$FJ->CREATE_BORDER( $cx[ $cn["K"] ].'5', $cx[ $ac ].'5', 'BFBFBF', "inside", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].$rc, $cx[ $ac ].$rc, 'FFFFFF', "inside", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( $cx[ $cn["N"] ].$rc, $cx[ $ac ].$rc, 'BFBFBF', "inside", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].$rd, $cx[ $ac ].$ad, 'BFBFBF', "inside", 'BORDER_THIN' );

	$FJ->CREATE_BORDER( $cx[ $cn["K"] ].'5', $cx[ $ac ].'5', '36608d', "top", 	 'BORDER_THIN' );
	$FJ->CREATE_BORDER( $cx[ $cn["K"] ].'5', $cx[ $ac ].'5', '36608d', "bottom", 'BORDER_THIN' );

	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'2', $cx[ $ac ].'2'     , '36608d', "top", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'2', $cx[ $ac ].'2'     , '36608d', "bottom", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'2', $cx[ $cn["E"] ].'2', '36608d', "outline", 'BORDER_THICK' );

	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'4', $cx[ $ac ].'6'     , '36608d', "top", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'4', $cx[ $ac ].'6'     , '36608d', "bottom", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'4', $cx[ $cn["J"] ].'6', '36608d', "outline", 'BORDER_THICK' );

	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].$rc, $cx[ $ac ].$rc     , '36608d', "top",   'BORDER_THICK' ); 
	// $FJ->CREATE_BORDER( $cx[ $cn["B"] ].$rc, $cx[ $cn["M"] ].$rc, '36608d', "left", 'BORDER_THICK' );

	$FJ->CREATE_BORDER( $cx[ $cn["B"] ] .($rc+1), $cx[ $cn["J"] ] . $ad, '808080', "outline" , 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["K"] ] .($rc+1), $cx[ $cn["M"] ] . $ad, '808080', "outline" , 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["N"] ] .($rc+1), $cx[ $ac ]      . $ad, '808080', "outline" , 'BORDER_THICK' );  
	//$FJ->STYLE_WIDTH_RANGE( $cn["F"], $ac, 10);
	$FJ->STYLE_WIDTH("A",  4);
	$FJ->STYLE_WIDTH("B",  10);
	$FJ->STYLE_WIDTH("C",  12);
	$FJ->STYLE_WIDTH("D",  17);
	$FJ->STYLE_WIDTH("E",  22);
	$FJ->STYLE_WIDTH("F", 55.5);
	$FJ->STYLE_WIDTH("G", 26.5);
	$FJ->STYLE_WIDTH("H", 35.5);
	$FJ->STYLE_WIDTH("I", 26.5);
	$FJ->STYLE_WIDTH("J", 13.5);
	$FJ->STYLE_WIDTH_RANGE( $cn["K"],  $cn["M"] ,  20);
	$FJ->STYLE_WIDTH_RANGE( $cn["N"],  $ac  ,   17.75); 
	$FJ->STYLE_WIDTH($cx[($ac+1)], 4); 

	$FJ->STYLE_HEIGHT( 1, 10.5 );	
	$FJ->STYLE_HEIGHT( 2, 100);
	$FJ->STYLE_HEIGHT( 3, 17 );
	$FJ->STYLE_HEIGHT( 4, 10 );
	$FJ->STYLE_HEIGHT( 5, 60 );
	$FJ->STYLE_HEIGHT( 6, 10 );
	$FJ->STYLE_HEIGHT( 7, 10 );
	$FJ->STYLE_HEIGHT( 8, 46 ); 
	$FJ->STYLE_HEIGHT( ($rc+1), 16.5); 
	$FJ->STYLE_HEIGHT_RANGE( $rd, $ad,  22.5); 

	$FJ->CREATE_HEAD( $cval, true );
	$FJ->CREATE_BODY( $hist, true );	

	$FJ->CREATE_TEXT( $cx[$cn["C"]]  . $rc , "PLANT" );
	$FJ->CREATE_TEXT( $cx[$cn["D"]]  . $rc , "PD" );
	$FJ->CREATE_TEXT( $cx[$cn["E"]]  . $rc , "LINE CD" );
	$FJ->CREATE_TEXT( $cx[$cn["F"]]  . $rc , "LINE NAME" );
	$FJ->CREATE_TEXT( $cx[$cn["G"]]  . $rc , "ITEM CD" );
	$FJ->CREATE_TEXT( $cx[$cn["H"]]  . $rc , "ITEM NAME" );
	$FJ->CREATE_TEXT( $cx[$cn["I"]]  . $rc , "MODEL" );
	$FJ->CREATE_TEXT( $cx[$cn["J"]]  . $rc , "PRODUCT TYPE" );
	$FJ->CREATE_TEXT( $cx[$cn["K"]]  . $rc , "PLAN" );
	$FJ->CREATE_TEXT( $cx[$cn["L"]]  . $rc , "ACTUAL" );
	$FJ->CREATE_TEXT( $cx[$cn["M"]]  . $rc , "DIFF" );


	$FJ->CREATE_MERGECELL('B' . '2' ,'E' . '2');
	$FJ->CREATE_MERGECELL('F' . '2' , $cx[$ac] . '2'); 
	$FJ->CREATE_MERGECELL('B' . '4' ,'J' . '6'); 

	$FJ->STYLE_ALIGNMENT( 'B2' , 'B2' ); 
	$FJ->STYLE_ALIGNMENT( 'F2' , 'F2', 'CL' );
	$FJ->STYLE_ALIGNMENT( 'B4' , 'B4' );
	$FJ->STYLE_ALIGNMENT( $cx[ $cn["B"] ]  . ($rc+1) , $cx[ $ac ] . ($rc+1) );
	$FJ->STYLE_ALIGNMENT( 'B'.$rc , $cx[$ac] . $rc, 'CC', true );	

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

	foreach($hol1 as $i => $v)
	{
		foreach( range($cn["N"], $ac)  as $c)
			{
				$t = $FJ->GETDATA_ONCELL( $cx[$c] . $rc );
				if( $t == $v["DD"])
					{
						$FJ->CREATE_FILL($cx[$c].'5', $cx[$c].'5', 'f2dcdb' );
						$FJ->CREATE_FILL($cx[$c].$rc, $cx[$c].$rc, 'f2dcdb' );
						$FJ->CREATE_FILL($cx[$c].$rd, $cx[$c].$ad, 'f2dcdb' );
						$FJ->CREATE_TEXT($cx[$c].($rc+1), $v['FND']);
					}
			}
	}
	foreach($sat1 as $i => $v)
	{
		foreach( range($cn["N"], $ac)  as $c)
			{
				$t = $FJ->GETDATA_ONCELL( $cx[$c] . $rc );
				if( $t == $v["DD"])
					{
						$FJ->CREATE_FILL($cx[$c].'5', $cx[$c].'5', 'daeed5' );
						$FJ->CREATE_FILL($cx[$c].$rc, $cx[$c].$rc, 'daeed5' );
						$FJ->CREATE_FILL($cx[$c].$rd, $cx[$c].$ad, 'daeed5' );
						$FJ->CREATE_TEXT($cx[$c].($rc+1), $v['FND']);
					}
			}
	}

// Set active sheet index to the first sheet, so Excel opens this as the first sheet 7030a0
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->removeSheetByIndex( $FJ->OUTPUT_INDEXSHEET()-1 );
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save($fln);

echo $fln;
//output_file($filename);
//exit;
 
#function @ use
	function RANCOLOR_HOLIDAY_CELL($f, $h, $r, $e)
		{
			var_dump($h); exit;
			echo $f->GETDATA_ONCELL("B2"); 
		} 
 		