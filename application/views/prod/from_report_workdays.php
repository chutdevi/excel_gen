<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli') die('This example should only be run from a Web Browser');
   
require_once dirname(__FILE__) . '/PHPExcel-1.8.1/Classes/PHPExcel.php';
require_once dirname(__FILE__) . '/PHPExcel-1.8.1/function_from_report.php';

// var $FJ;
// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$FJ = new FUNCTION_GENERATE( $objPHPExcel );

	$formula['E1'] = "=SUM(%s:%s)";
	$formula['F1'] = "=%s-COUNTIF(%s:%s,0)";
	$formula['G1'] = "=SUBTOTAL(9,%s:%s)";
	$formula['G3'] = "=SUBTOTAL(4,%s:%s)";

	$format['F1']  = '_-* #,##0_-"Days";-* #,##0_-;_-* "-"_-;_-@_-';
	$format['E1']  = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-';

$cn = $FJ->OUTPUT_COLUMNNAMES(50);
$cx = $FJ->OUTPUT_COLUMNINDEX(50);
 
#@SHEET 1
	$Dcon  = $cont["this"];
	$rc = $FJ->row_column   = 2;
	$rd = $FJ->row_content  = 4;
	$cs = $FJ->column_start = 1;
	$FJ->inx = 0;
	$FJ->column_index = $cx;
	$FJ->amount_column = $ac = count ( $Dcon[0] )+( $cs-1 );
	$FJ->amount_row    = $ad = count ( $Dcon    )+( $rd-1 );

	$dm = date('F Y', strtotime($days) );
	$lm = date('t',strtotime($days) );


	$FJ->CREATE_SHEET( "Production-" .  date('F Y', strtotime($days) )  );  

	$FJ->CREATE_COLORTAB( "00004d" );

	$FJ->STYLE_GRIDLINES(False);

	$FJ->CREATE_HEAD( $Dcon[0] );
	$FJ->CREATE_BODY( $Dcon );
	$FJ->CREATE_FILTER( $cx[$cs] . ( $rc+1 ), $cx[$ac] . ( $rc+1 ) );
//RANCOLOR_HOLIDAY_CELL($FJ, $sat1, $rc, $ac);
//exit;
	$FJ->CREATE_FREEZE('G4');
	$FJ->CREATE_TEXT( 'A1', "Plan " .  $dm );
	$FJ->CREATE_TEXT( 'E1', sprintf( $formula['E1'], "G1", $cx[$ac] . "1"  ) );
	$FJ->CREATE_TEXT( 'F1', sprintf( $formula['F1'], $lm , "G1", $cx[$ac] . "1"  ) );

	$FJ->STYLE_ALIGNMENT( 'A1' , $cx[$ac] . '1' );
	$FJ->STYLE_ALIGNMENT( 'A3' , $cx[$ac] . '3' );
	$FJ->STYLE_ALIGNMENT( 'A2' , $cx[$ac] . '2' );

	$FJ->CREATE_FORMAT('F1','F1', $format['F1'] );
	$FJ->CREATE_FORMAT('E1','E1', $format['E1'] );
	$FJ->CREATE_FORMAT('G1', $cx[$ac] . '1', $format['E1'] );
	$FJ->CREATE_FORMAT('G4', $cx[$ac] . $ad, $format['E1'] );

	$FJ->CREATE_FONT('A1' ,'A1', 12, 'FF0000', true );

	$FJ->CREATE_FONT('E1' ,'E1', 12, '000099', true );
	$FJ->CREATE_FONT('F1' ,'F1', 12, '990000', true );
	$FJ->CREATE_FONT('A2' , $cx[$ac] . '2', 12,'FFFFF0', true );
	$FJ->CREATE_FONT('A3' , $cx[$ac] . '3', 8 ,'FF0000', true );

	$FJ->CREATE_FILL('A1' , 'A1', 'e0ebeb' );
	$FJ->CREATE_FILL('A2' , $cx[$ac] . '2', '29293d' );
	$FJ->CREATE_FILL('A3' , $cx[$ac] . '3', 'd1e0e0' );

	foreach( range($cn["G"], $ac) as $c )
		$FJ->CREATE_TEXT( $cx[$c] . "1" , sprintf( $formula['G1'], $cx[$c].$rd, $cx[$c].$ad ) );

	$FJ->CREATE_MERGECELL('A1','D1');

	$FJ->STYLE_WIDTH_RANGE( $cn["A"], $cn["B"], 8.5);
	$FJ->STYLE_WIDTH_RANGE( $cn["G"], $ac, 10);
	$FJ->STYLE_WIDTH("C", 60);
	$FJ->STYLE_WIDTH("D", 19);
	$FJ->STYLE_WIDTH("E", 19);
	$FJ->STYLE_WIDTH("F", 26);

	$FJ->STYLE_HEIGHT( ( $rc+1 ) , 12);

	$FJ->CREATE_BORDER( 'A2', $cx[$ac].'2', 'FFFFFC', "outline", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( 'A2', $cx[$ac].'2', 'FFFFFC', "inside" , 'BORDER_THIN' );
	$FJ->CREATE_BORDER( 'A1', $cx[$ac].'1', '9494b8', "inside" , 'BORDER_THIN' );

	$FJ->CREATE_BORDER( 'A1', $cx[$ac].$ad, '000000', "outline", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( 'A4', $cx[$ac].$ad, '9494b8', "inside" , 'BORDER_THIN' );

	$FJ->STYLE_GROUP_COLUMN($cn["C"]);

	foreach($hol1 as $i => $v)
		{
			foreach( range($cn["G"], $ac)  as $c) 
				{
					$t = $FJ->GETDATA_ONCELL( $cx[$c] . '2' );
					if( $t == $v["DD"])
						{
							$FJ->CREATE_FILL($cx[$c].'1', $cx[$c].'1', 'ffbf80' );
							$FJ->CREATE_FILL($cx[$c].$rd, $cx[$c].$ad, 'ffbf80' );
							$FJ->CREATE_TEXT($cx[$c].'3', $v['FND']);
						}
				}
		}
	foreach($sat1 as $i => $v)
		{
			foreach( range($cn["G"], $ac)  as $c)
				{
					$t = $FJ->GETDATA_ONCELL( $cx[$c] . '2' );
					if( $t == $v["DD"])
						{
							$FJ->CREATE_FILL($cx[$c].'1', $cx[$c].'1', '80ff80' );
							$FJ->CREATE_FILL($cx[$c].$rd, $cx[$c].$ad, '80ff80' );
							$FJ->CREATE_TEXT($cx[$c].'3', $v['FND']);
						}
				}
		}


#@SHEET 2
	$Dcon  = $cont["online_this"];
	$rc = $FJ->row_column   = 3;
	$rd = $FJ->row_content  = 5;
	$cs = $FJ->column_start = 2; 
	$FJ->inx = 1;

	$FJ->amount_column = $ac = count ( $Dcon[0] )+( $cs-1 );
	$FJ->amount_row    = $ad = count ( $Dcon    )+( $rd-1 );;

	$FJ->CREATE_SHEET( "Line work-" .  date('F Y', strtotime($days) )  );  
	$FJ->CREATE_COLORTAB( "0099cc" );
	$FJ->STYLE_GRIDLINES(False);
	$FJ->CREATE_HEAD( $Dcon[0] );
	$FJ->CREATE_BODY( $Dcon );
	$FJ->CREATE_FILTER( $cx[$cs]. ( $rc+1 ), $cx[$ac] . ( $rc+1 ) );
	$FJ->CREATE_FREEZE('E5');	

	$FJ->STYLE_ALIGNMENT( 'A1' , $cx[$ac] . $ad );
	$FJ->STYLE_ALIGNMENT( $cx[ $cn["F"] ] . $rc, $cx[ $ac ] . $rc );

	$FJ->CREATE_MERGECELL('B2','C2');	
	
	$FJ->STYLE_WIDTH_RANGE( $cn["B"], $cn["C"], 8.5);
	$FJ->STYLE_WIDTH_RANGE( $cn["D"], $cn["E"], 0.5);
	$FJ->STYLE_WIDTH_RANGE( $cn["F"], $ac, 4.5);

	$FJ->STYLE_WIDTH( $cx[ $cn["A" ] ],1.3);
	$FJ->STYLE_WIDTH( $cx[  $ac+1  ]  ,1.3);

	$FJ->STYLE_HEIGHT_RANGE( $rd , $ad ,20);
	
	$FJ->STYLE_HEIGHT( 1, 10 );	
	$FJ->STYLE_HEIGHT( 2, 25);
	$FJ->STYLE_HEIGHT( 3, 35);
	$FJ->STYLE_HEIGHT( ( $rc+1 ), 10.5);
	$FJ->STYLE_HEIGHT( $ad+1, 10);

	$FJ->STYLE_TEXTROTATION_RANGE( $cn["F"], $ac, 3, -90 );

	$FJ->CREATE_TEXT( $cx[ $cn["B"] ] . '2', "LINE WORK" );

	$FJ->CREATE_FILL('A1' , $cx[$ac+1].($ad+1), '16365c' );
	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . $rc, $cx[ $cn["C"] ] . $rc, '808080' );	
	$FJ->CREATE_FILL( $cx[ $cn["F"] ] . $rc, $cx[ $ac ]      . $rc, '808080' );
	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . $rd, $cx[ $cn["C"] ] . $ad, '808080' );


	$FJ->CREATE_FILL( $cx[ $cn["F"] ] . $rd, $cx[ $cn["C"] ] . $rd, '808080' );

	$FJ->CREATE_FILL( $cx[ $cn["F"] ] . ( $rc-1 ), $cx[ $ac ] . ( $rc-1 ), '538dd5' );	

	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . ( $rc+1 ), $cx[ $cn["C"] ] . ( $rc+1 ), '990000' );
	$FJ->CREATE_FILL( $cx[ $cn["F"] ] . ( $rc+1 ), $cx[ $ac ]      . ( $rc+1 ), '990000' );

	$FJ->CREATE_FONT( $cx[ $cn["B"] ] . '1',       $cx[ $cn["B"] ] . '1'    , 12, 'FFFFFF', true );
	$FJ->CREATE_FONT( $cx[ $cs ]      . ( $rc-1 ), $cx[ $ac ]      . ( $rc ), 12, 'FFFFFF', true );
	$FJ->CREATE_FONT( $cx[ $cn["B"] ] . ( $rd )  , $cx[ $cn["C"] ] . ( $ad ), 11, 'FFFFFF', true );
	$FJ->CREATE_FONT( $cx[ $cn["F"] ] . ( $rd )  , $cx[ $ac ]      . ( $ad ), 11, 'FFFFFF', true );

	$FJ->CREATE_BORDER( $cx[ $cn["F"] ].'2', $cx[ $ac ].'2', '990000', "outline", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["F"] ].'3', $cx[ $ac ].$ad, '990000', "outline", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'3', $cx[ $cn["C"] ] . '3', '990000', "outline", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'5', $cx[ $cn["C"] ] . $ad, '990000', "outline", 'BORDER_THICK' );

	$FJ->CREATE_BORDER( $cx[ $cn["F"] ].'2', $cx[ $ac ].'2',      '990000', "inside", 'BORDER_THIN' );	
	$FJ->CREATE_BORDER( $cx[ $cn["F"] ].'3', $cx[ $ac ].'3',      '990000', "inside", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'3', $cx[ $cn["C"] ].'3', '990000', "inside", 'BORDER_THIN' );

	$FJ->CREATE_BORDER( $cx[ $cn["F"] ]. ( $rd ), $cx[ $ac ]      . $ad, '990000', "inside", 'BORDER_THIN' );	
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ]. ( $rd ), $cx[ $cn["C"] ] . $ad, '990000', "inside", 'BORDER_THIN' );

	foreach( range($cn["F"], $ac) as $c )
		$FJ->CREATE_TEXT( $cx[$c] . "2" , sprintf( $formula['G1'], $cx[$c].$rd, $cx[$c].$ad ) );

	$DT1 = $FJ->OUTPUT_CONDITIONFORMAT("=", 0);
	$FJ->STYLE_CONDITIONFORMAT_NUMBER($DT1, '"OFF"' );
	$FJ->STYLE_CONDITIONFORMAT_FONT($DT1, 11,'FF00000',"Calibri", true);
	$FJ->STYLE_CONDITIONFORMAT_FILL($DT1, 'ffff00', 'ffff00');
		

	$DT2 = $FJ->OUTPUT_CONDITIONFORMAT("=", 1);
	$FJ->STYLE_CONDITIONFORMAT_NUMBER( $DT2, '"ON"' );
	$FJ->STYLE_CONDITIONFORMAT_FONT( $DT2, 11,'000080',"Calibri", true);
	$FJ->STYLE_CONDITIONFORMAT_FILL( $DT2, '00cc00', '00cc00');
	//$FJ->APPEND_CONDITIONFORMAT( $DT2, $cx[ $cn["F"] ].( $rd ), $cx[ $ac ]. $ad );	


	$FJ->CREATE_CONDITIONFORMAT($cx[ $cn["F"] ].( $rd ), $cx[ $ac ]. $ad, array( $DT1, $DT2 ));

	foreach($hol1 as $i => $v)
		{
			foreach( range($cn["F"], $ac)  as $c)
				{
					$t = $FJ->GETDATA_ONCELL( $cx[$c] . '3' );
					if( $t == $v["DD"])
						{
							$FJ->CREATE_FILL($cx[$c].'3', $cx[$c].'3', '4f6228' );
						}
				}
		}
	foreach($sat1 as $i => $v)
		{
			foreach( range($cn["G"], $ac)  as $c)
				{
					$t = $FJ->GETDATA_ONCELL( $cx[$c] . '3' );
					if( $t == $v["DD"])
						{
							$FJ->CREATE_FILL($cx[$c].'3', $cx[$c].'3', '0f243e' );
						}
				}
		}


#@SHEET 3
	$Dcon  = $cont["next"];
	$rc = $FJ->row_column  = 2;
	$rd = $FJ->row_content = 4;
	$cs = $FJ->column_start = 1; 
	$FJ->inx = 2;
	$FJ->column_index = $cx;

	$FJ->amount_column = $ac = count ( $Dcon[0] )+( $cs-1 );
	$FJ->amount_row    = $ad = count ( $Dcon    )+( $rd-1 );

	$FJ->CREATE_SHEET( "Production-" . date('F Y', strtotime("+ 1 month", strtotime($days) ) )  );  

	$FJ->CREATE_COLORTAB( "26004d" );

	$FJ->STYLE_GRIDLINES(False);

	$FJ->CREATE_HEAD( $Dcon[0] );
	$FJ->CREATE_BODY( $Dcon );
	$FJ->CREATE_FILTER( $cx[$cs]. ( $rc+1 ), $cx[$ac] . ( $rc+1 ) );

	$FJ->CREATE_FREEZE('G4');
	$FJ->CREATE_TEXT( 'A1', "Plan " .  date('F Y', strtotime("+ 1 month", strtotime($days) ) ) );
	$FJ->CREATE_TEXT( 'E1', sprintf("=SUM(%s:%s)"		   , "G1", $cx[$ac] . "1"  ) );
	$FJ->CREATE_TEXT( 'F1', sprintf("=%s-COUNTIF(%s:%s,0)", date('t',strtotime("+ 1 month", strtotime($days) ) ), "G1", $cx[$ac] . "1"  ) );

	$FJ->STYLE_ALIGNMENT( 'A1' , $cx[$ac] . '1' );
	$FJ->STYLE_ALIGNMENT( 'A2' , $cx[$ac] . '2' );
	$FJ->STYLE_ALIGNMENT( 'A3' , $cx[$ac] . '3' );

	$FJ->CREATE_FORMAT('F1','F1', '_-* #,##0_-"Days";-* #,##0_-;_-* "-"_-;_-@_-');
	$FJ->CREATE_FORMAT('E1','E1', '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-');
	$FJ->CREATE_FORMAT('G1', $cx[$ac] . '1', '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-');
	$FJ->CREATE_FORMAT('G4', $cx[$ac] . $ad, '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-');

	$FJ->CREATE_FONT('A1' ,'A1', 12, 'FF0000', true );
	$FJ->CREATE_FONT('E1' ,'E1', 12, '000099', true );
	$FJ->CREATE_FONT('F1' ,'F1', 12, '990000', true );
	$FJ->CREATE_FONT('A2' , $cx[$ac] . '2', 12,'FFFFF0', true );
	$FJ->CREATE_FONT('A3' , $cx[$ac] . '3', 8 ,'FF0000', true );

	$FJ->CREATE_FILL('A1' , 'A1', 'e0ebeb' );
	$FJ->CREATE_FILL('A2' , $cx[$ac] . '2', '29293d' );
	$FJ->CREATE_FILL('A3' , $cx[$ac] . '3', 'd1e0e0' );

	foreach( range($cn["G"], $ac) as $cind )
		$FJ->CREATE_TEXT( $cx[$cind] . "1" , sprintf("=SUBTOTAL(9,%s:%s)", $cx[$cind].$rd, $cx[$cind].$ad ) );

	$FJ->CREATE_MERGECELL('A1','D1');


	$FJ->STYLE_WIDTH_RANGE( $cn["A"], $cn["B"], 8.5);
	$FJ->STYLE_WIDTH_RANGE( $cn["G"], $ac, 10);
	$FJ->STYLE_WIDTH("C", 60);
	$FJ->STYLE_WIDTH("D", 19);
	$FJ->STYLE_WIDTH("E", 19);
	$FJ->STYLE_WIDTH("F", 26);

	$FJ->STYLE_HEIGHT( ( $rc+1 ) , 12);

	$FJ->CREATE_BORDER( 'A2', $cx[$ac].'2', 'FFFFFC', "outline", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( 'A2', $cx[$ac].'2', 'FFFFFC', "inside" , 'BORDER_THIN' );
	$FJ->CREATE_BORDER( 'A1', $cx[$ac].'1', '9494b8', "inside" , 'BORDER_THIN' );

	$FJ->CREATE_BORDER( 'A1', $cx[$ac].$ad, '000000', "outline", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( 'A4', $cx[$ac].$ad, '9494b8', "inside", 'BORDER_THIN' );

	$FJ->STYLE_GROUP_COLUMN($cn["C"]);

	foreach($hol2 as $i => $v)
		{
			foreach( range($cn["G"], $ac)  as $c)
				{
					$t = $FJ->GETDATA_ONCELL( $cx[$c] . '2' );
					if( $t == $v["DD"])
						{
							$FJ->CREATE_FILL($cx[$c].'1', $cx[$c].'1', 'ffbf80' );
							$FJ->CREATE_FILL($cx[$c].$rd, $cx[$c].$ad, 'ffbf80' );
							$FJ->CREATE_TEXT($cx[$c].'3', $v['FND']);
						}
				}
		}
	foreach($sat2 as $i => $v)
		{
			foreach( range($cn["G"], $ac)  as $c)
				{
					$t = $FJ->GETDATA_ONCELL( $cx[$c] . '2' );
					if( $t == $v["DD"])
						{
							$FJ->CREATE_FILL($cx[$c].'1', $cx[$c].'1', '80ff80' );
							$FJ->CREATE_FILL($cx[$c].$rd, $cx[$c].$ad, '80ff80' );
							$FJ->CREATE_TEXT($cx[$c].'3', $v['FND']);
						}
				}
		}


#@SHEET 4
	$Dcon  = $cont["online_next"];
	$rc = $FJ->row_column   = 3;
	$rd = $FJ->row_content  = 5;
	$cs = $FJ->column_start = 2;
	$FJ->inx = 3;
	$FJ->column_index = $cx;

	$FJ->amount_column = $ac = count ( $Dcon[0] )+( $cs-1 );
	$FJ->amount_row    = $ad = count ( $Dcon    )+( $rd-1 );;

	$FJ->CREATE_SHEET( "Line work-" .  date('F Y', strtotime("+ 1 month", strtotime($days) ) )  );  

	$FJ->CREATE_COLORTAB( "0099cc" );//666699

	$FJ->STYLE_GRIDLINES(False);


	$FJ->CREATE_HEAD( $Dcon[0] );
	$FJ->CREATE_BODY( $Dcon );
	$FJ->CREATE_FILTER( $cx[$cs]. ( $rc+1 ), $cx[$ac] . ( $rc+1 ) );
	$FJ->CREATE_FREEZE('E5');	

	$FJ->STYLE_ALIGNMENT( 'A1' , $cx[$ac] . $ad );
	$FJ->STYLE_ALIGNMENT( $cx[ $cn["F"] ] . $rc, $cx[ $ac ] . $rc);

	$FJ->CREATE_MERGECELL('B2','C2');	
	
	$FJ->STYLE_WIDTH_RANGE( $cn["B"], $cn["C"], 8.5);
	$FJ->STYLE_WIDTH_RANGE( $cn["D"], $cn["E"], 0.5);
	$FJ->STYLE_WIDTH_RANGE( $cn["F"], $ac, 4.5);

	$FJ->STYLE_WIDTH( $cx[ $cn["A" ] ],1.3);
	$FJ->STYLE_WIDTH( $cx[  $ac+1  ]  ,1.3);

	$FJ->STYLE_HEIGHT_RANGE( $rd , $ad ,20);
	
	$FJ->STYLE_HEIGHT( 1, 10 );	
	$FJ->STYLE_HEIGHT( 2, 25);
	$FJ->STYLE_HEIGHT( 3, 35);
	$FJ->STYLE_HEIGHT( ( $rc+1 ), 10.5);
	$FJ->STYLE_HEIGHT( $ad+1, 10);

	$FJ->STYLE_TEXTROTATION_RANGE( $cn["F"], $ac, 3, -90 );

	$FJ->CREATE_TEXT( $cx[ $cn["B"] ] . '2', "LINE WORK" );

	$FJ->CREATE_FILL('A1' , $cx[$ac+1].($ad+1), '16365c' );
	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . $rc, $cx[ $cn["C"] ] . $rc, '808080' );	
	$FJ->CREATE_FILL( $cx[ $cn["F"] ] . $rc, $cx[ $ac ]      . $rc, '808080' );
	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . $rd, $cx[ $cn["C"] ] . $ad, '808080' );


	$FJ->CREATE_FILL( $cx[ $cn["F"] ] . $rd, $cx[ $cn["C"] ] . $rd, '808080' );

	$FJ->CREATE_FILL( $cx[ $cn["F"] ] . ( $rc-1 ), $cx[ $ac ] . ( $rc-1 ), '538dd5' );	

	$FJ->CREATE_FILL( $cx[ $cn["B"] ] . ( $rc+1 ), $cx[ $cn["C"] ] . ( $rc+1 ), '990000' );
	$FJ->CREATE_FILL( $cx[ $cn["F"] ] . ( $rc+1 ), $cx[ $ac ]      . ( $rc+1 ), '990000' );

	$FJ->CREATE_FONT( $cx[ $cn["B"] ] . '1',       $cx[ $cn["B"] ] . '1'    , 12, 'FFFFFF', true );
	$FJ->CREATE_FONT( $cx[ $cs ]      . ( $rc-1 ), $cx[ $ac ]      . ( $rc ), 12, 'FFFFFF', true );
	$FJ->CREATE_FONT( $cx[ $cn["B"] ] . ( $rd )  , $cx[ $cn["C"] ] . ( $ad ), 11, 'FFFFFF', true );
	$FJ->CREATE_FONT( $cx[ $cn["F"] ] . ( $rd )  , $cx[ $ac ]      . ( $ad ), 11, 'FFFFFF', true );

	$FJ->CREATE_BORDER( $cx[ $cn["F"] ].'2', $cx[ $ac ].'2', '990000', "outline", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["F"] ].'3', $cx[ $ac ].$ad, '990000', "outline", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'3', $cx[ $cn["C"] ] . '3', '990000', "outline", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'5', $cx[ $cn["C"] ] . $ad, '990000', "outline", 'BORDER_THICK' );

	$FJ->CREATE_BORDER( $cx[ $cn["F"] ].'2', $cx[ $ac ].'2',      '990000', "inside", 'BORDER_THIN' );	
	$FJ->CREATE_BORDER( $cx[ $cn["F"] ].'3', $cx[ $ac ].'3',      '990000', "inside", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ].'3', $cx[ $cn["C"] ].'3', '990000', "inside", 'BORDER_THIN' );

	$FJ->CREATE_BORDER( $cx[ $cn["F"] ]. ( $rd ), $cx[ $ac ]      . $ad, '990000', "inside", 'BORDER_THIN' );	
	$FJ->CREATE_BORDER( $cx[ $cn["B"] ]. ( $rd ), $cx[ $cn["C"] ] . $ad, '990000', "inside", 'BORDER_THIN' );

	foreach( range($cn["F"], $ac) as $c )
		$FJ->CREATE_TEXT( $cx[$c] . "2" , sprintf( $formula['G1'], $cx[$c].$rd, $cx[$c].$ad ) );

	$DT1 = $FJ->OUTPUT_CONDITIONFORMAT("=", 0);
	$FJ->STYLE_CONDITIONFORMAT_NUMBER($DT1, '"OFF"' );
	$FJ->STYLE_CONDITIONFORMAT_FONT($DT1, 11,'FF00000',"Calibri", true);
	$FJ->STYLE_CONDITIONFORMAT_FILL($DT1, 'ffff00', 'ffff00');
		

	$DT2 = $FJ->OUTPUT_CONDITIONFORMAT("=", 1);
	$FJ->STYLE_CONDITIONFORMAT_NUMBER( $DT2, '"ON"' );
	$FJ->STYLE_CONDITIONFORMAT_FONT( $DT2, 11,'000080',"Calibri", true);
	$FJ->STYLE_CONDITIONFORMAT_FILL( $DT2, '00cc00', '00cc00');
	//$FJ->APPEND_CONDITIONFORMAT( $DT2, $cx[ $cn["F"] ].( $rd ), $cx[ $ac ]. $ad );	


	$FJ->CREATE_CONDITIONFORMAT($cx[ $cn["F"] ].( $rd ), $cx[ $ac ]. $ad, array( $DT1, $DT2 ));

	foreach($hol2 as $i => $v)
		{
			foreach( range($cn["F"], $ac)  as $c)
				{
					$t = $FJ->GETDATA_ONCELL( $cx[$c] . '3' );
					if( $t == $v["DD"])
						{
							$FJ->CREATE_FILL($cx[$c].'3', $cx[$c].'3', '4f6228' );
						}
				}
		}
	foreach($sat2 as $i => $v)
		{
			foreach( range($cn["G"], $ac)  as $c)
				{
					$t = $FJ->GETDATA_ONCELL( $cx[$c] . '3' );
					if( $t == $v["DD"])
						{
							$FJ->CREATE_FILL($cx[$c].'3', $cx[$c].'3', '0f243e' );
						}
				}
		}


#@SHEET 5
	$Dcon  = $cont["summary_line"];
	$rc = $FJ->row_column   = 11;
	$rd = $FJ->row_content  = 13;
	$cs = $FJ->column_start = 3; 
	$FJ->inx = 4;

	$FJ->amount_column = $ac = count ( $Dcon[0] )+( $cs-1 );
	$FJ->amount_row    = $ad = count ( $Dcon    )+( $rd-1 );

	$FJ->CREATE_SHEET( "Summary" );  
	$FJ->CREATE_COLORTAB( "FF0000" );

	$FJ->STYLE_GRIDLINES(False);


	$FJ->CREATE_HEAD( $Dcon[0] );
	$FJ->CREATE_BODY( $Dcon );
	$FJ->CREATE_FILTER( $cx[$cs] . ( $rc+1 ), $cx[$ac] . ( $rc+1 ) );

	$FJ->CREATE_FREEZE('F13');	
	$FJ->STYLE_ALIGNMENT( 'A1' , $cx[$ac] . $ad );



	$FJ->STYLE_WIDTH_RANGE( $cn["A"], $cn["B"], 2 );
	$FJ->STYLE_WIDTH_RANGE( $cn["C"], $cn["D"], 11  );
	$FJ->STYLE_WIDTH_RANGE( $cn["E"], $cn["F"], 0.5 );
	$FJ->STYLE_WIDTH_RANGE( $cn["N"], $cn["O"], 2 );
	$FJ->STYLE_WIDTH_RANGE( $cn["G"], $ac, 12 );

	$FJ->STYLE_WIDTH( $cx[ $cn["A" ] ],1.3);
	$FJ->STYLE_WIDTH( $cx[  $ac+1  ]  ,1.3);

	$FJ->STYLE_HEIGHT_RANGE( 1 , 2 ,10.5);

	$FJ->STYLE_HEIGHT( 3, 18 );
	$FJ->STYLE_HEIGHT( 5, 18 );
	$FJ->STYLE_HEIGHT( 7, 18 );
	$FJ->STYLE_HEIGHT( 4, 8 );	
	$FJ->STYLE_HEIGHT( 6, 8 );
	$FJ->STYLE_HEIGHT( 8, 8 );
	$FJ->STYLE_HEIGHT( 11, 25);

	$FJ->CREATE_FILL('B2' , $cx[$ac+1].($ad+1), '31869b' );
	$FJ->CREATE_FILL('C3' , 'C3', '632523' );
	$FJ->CREATE_FILL('C5' , 'C5', '366092' );
	$FJ->CREATE_FILL('C7' , 'C7', 'ffc000' );
	$FJ->CREATE_FILL('D9' , 'D9', '215967' );

	$FJ->CREATE_FILL('C12' , 'D12', '366092' );
	$FJ->CREATE_FILL('G12' , 'M12', '7030a0' );

	$FJ->CREATE_FILL('G3' , $cx[$cn["M"]] . "3", '538dd5' );
	$FJ->CREATE_FILL('G5' , $cx[$cn["M"]] . "5", '76933c' );
	$FJ->CREATE_FILL('G7' , $cx[$cn["M"]] . "7", 'da9694' );

	$FJ->CREATE_FONT( $cx[ $cn["B"] ] . '2', $cx[ $cn["N"] ] . $ad    , 11, 'fde9d9', true );
	$FJ->CREATE_FONT( $cx[ $cn["C"] ] . '3', $cx[ $cn["C"] ] . '3'    , 11, 'FFFF00', true );
	$FJ->CREATE_FONT( $cx[ $cn["C"] ] . '5', $cx[ $cn["C"] ] . '5'    , 11, '92D050', true );
	$FJ->CREATE_FONT( $cx[ $cn["C"] ] . '7', $cx[ $cn["C"] ] . '7'    , 11, 'FF0000', true );

	//$FJ->CREATE_FORMAT( 'G9', $cx[$ac] . $ad, $format['E1'] );

	$FJ->CREATE_BORDER( $cx[ $cn["G"] ].'3', $cx[ $cn["M"] ] .'3', '632523', "outline", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["G"] ].'5', $cx[ $cn["M"] ] .'5', '494529', "outline", 'BORDER_THICK' );
	$FJ->CREATE_BORDER( $cx[ $cn["G"] ].'7', $cx[ $cn["M"] ] .'7', '963634', "outline", 'BORDER_THICK' );

	$FJ->CREATE_BORDER( $cx[ $cn["C"] ].$rc, $cx[ $cn["D"] ] .$ad, '366092', "outline", 'BORDER_THICK' );	
	$FJ->CREATE_BORDER( $cx[ $cn["G"] ].$rc, $cx[ $ac ]      .$ad, '7030a0', "outline", 'BORDER_THICK' );

	$FJ->CREATE_BORDER( $cx[ $cn["G"] ].'3', $cx[ $cn["M"] ] .'3', '632523', "inside", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( $cx[ $cn["G"] ].'5', $cx[ $cn["M"] ] .'5', '494529', "inside", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( $cx[ $cn["G"] ].'7', $cx[ $cn["M"] ] .'7', '963634', "inside", 'BORDER_THIN' );

	$FJ->CREATE_BORDER( $cx[ $cn["C"] ].$rc, $cx[ $cn["D"] ] .$rc, '366092', "inside", 'BORDER_THIN' );	
	$FJ->CREATE_BORDER( $cx[ $cn["C"] ].$rd, $cx[ $cn["D"] ] .$ad, '366092', "inside", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( $cx[ $cn["G"] ].$rc, $cx[ $cn["M"] ] .$rc, '7030a0', "inside", 'BORDER_THIN' );
	$FJ->CREATE_BORDER( $cx[ $cn["G"] ].$rd, $cx[ $cn["M"] ] .$ad, '7030a0', "inside", 'BORDER_THIN' );




	$FJ->CREATE_TEXT( $cx[ $cn["C"] ] . '3', "Actual Max day" );
	$FJ->CREATE_TEXT( $cx[ $cn["C"] ] . '5', "Calendar workdays" );
	$FJ->CREATE_TEXT( $cx[ $cn["C"] ] . '7', "Over workdays" );
	$FJ->CREATE_TEXT( $cx[ $cn["D"] ] . '9', "Years" );

	$FJ->CREATE_RUNFOEMULA_ROWDATA( $cn["G"], $ac, 3, $rd, $ad, $formula["G3"] );

	$ty  = $cn["G"];
	$y   = (int)date('Y' , strtotime( "-3 month", strtotime($days) ));
	//$cor = array('4f6228','244062' );
	foreach( range(-3,3) as $x => $v )
		{
				$thisMonth  = (int)date('ym');
				$runsMonth  = (int)date('ym', strtotime( $v . " month", strtotime($days) ));
				$runsYears  = (int)date('Y' , strtotime( $v . " month", strtotime($days) ));
				$FJ->CREATE_TEXT( $cx[ $cn["G"]+$x ] . '9', $runsYears );
				$FJ->CREATE_TEXT( $cx[ $cn["G"]+$x ] . '5', $cals[0]["M".($x+1)] );
				$FJ->CREATE_TEXT( $cx[ $cn["G"]+$x ] . '7', "=" . $cx[ $cn["G"]+$x ] . '3' . "-" . $cx[ $cn["G"]+$x ] . '5'  );
				if( $thisMonth >= $runsMonth)
					{
						$FJ->CREATE_TEXT( $cx[ $cn["G"]+$x ] . '12', "(actual)" );
						$FJ->CREATE_FONT( $cx[ $cn["G"]+$x ] . '12', $cx[ $cn["G"]+$x ] . '12' , 11, 'FFFF00', true );			
					} 
				else
					{
						$FJ->CREATE_TEXT( $cx[ $cn["G"]+$x ] . '12', "(plan)" );
						$FJ->CREATE_FONT( $cx[ $cn["G"]+$x ] . '12', $cx[ $cn["G"]+$x ] . '12' , 11, '66FF99', true );	
					}
				if($y != $runsYears)
					{
						$FJ->CREATE_FILL($cx[ $cn["G"] ] . '9',$cx[ $cn["G"]+($x-1) ] . '9', '4f6228' );
						$FJ->CREATE_MERGECELL( $cx[ $cn["G"] ] . '9',$cx[ $cn["G"]+($x-1) ] . '9');	
						$y  = $runsYears;
						$ty = $cn["G"] + $x ; 	
					}	
		}

	$FJ->CREATE_FILL( $cx[ $ty ] . '9', $cx[ $ac ] . '9', '244062' );
	$FJ->CREATE_MERGECELL( $cx[ $ty ] . '9', $cx[ $ac ] . '9' );

	$FJ->CREATE_MERGECELL( 'C' . '3', 'D' . '3' );
	$FJ->CREATE_MERGECELL( 'C' . '5', 'D' . '5' );
	$FJ->CREATE_MERGECELL( 'C' . '7', 'D' . '7' );	

// Set active sheet index to the first sheet, so Excel opens this as the first sheet 7030a0
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->removeSheetByIndex( $FJ->OUTPUT_INDEXSHEET()-1 );
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save($fln);

echo $fln;
//output_file($filename);
exit;






#function @ use
	function RANCOLOR_HOLIDAY_CELL($f, $h, $r, $e)
		{
			var_dump($h); exit;
			echo $f->GETDATA_ONCELL("B2"); 
		}