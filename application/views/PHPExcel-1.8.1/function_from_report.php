<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
#======================================= BORDER OVERVIEW ==================================#
// Array key | Maps to property
// ----------|------------------
// left      | getLeft()
// right     | getRight()
// top       | getTop()
// bottom    | getBottom()
// diagonal  | getDiagonal()
// Additional shortcut borders come in handy like in the example above. These are the shortcut borders available:

// Array key  | Maps to property
// -----------|------------------ 
// allborders | getAllBorders()
// outline    | getOutline()
// inside     | getInside()
// vertical   | getVertical()
// horizontal | getHorizontal()
#======================================= BORDER STYLE  ==================================#
// BORDER_NONE
// BORDER_DASHDOT
// BORDER_DASHDOTDOT
// BORDER_DASHED
// BORDER_DOTTED
// BORDER_DOUBLE
// BORDER_HAIR
// BORDER_MEDIUM
// BORDER_MEDIUMDASHDOT
// BORDER_MEDIUMDASHDOTDOT
// BORDER_MEDIUMDASHED
// BORDER_SLANTDASHDOT
// BORDER_THICK
// BORDER_THIN
#======================================= CONDITION FORMAT CELL  ==================================#
// const CONDITION_NONE                    = 'none';
// const CONDITION_CELLIS                  = 'cellIs';
// const CONDITION_CONTAINSTEXT            = 'containsText';
// const CONDITION_EXPRESSION              = 'expression';
//
// /* Operator types */
// const OPERATOR_NONE                     = '';
// const OPERATOR_BEGINSWITH               = 'beginsWith';
// const OPERATOR_ENDSWITH                 = 'endsWith';
// const OPERATOR_EQUAL                    = 'equal';
// const OPERATOR_GREATERTHAN              = 'greaterThan';
// const OPERATOR_GREATERTHANOREQUAL       = 'greaterThanOrEqual';
// const OPERATOR_LESSTHAN                 = 'lessThan';
// const OPERATOR_LESSTHANOREQUAL          = 'lessThanOrEqual';
// const OPERATOR_NOTEQUAL                 = 'notEqual';
// const OPERATOR_CONTAINSTEXT             = 'containsText';
// const OPERATOR_NOTCONTAINS              = 'notContains';
// const OPERATOR_BETWEEN                  = 'between';


Class FUNCTION_GENERATE
 {
    public $objPHPExcel;
    public $inx;
    public $row_column;
    public $row_content;
    public $column_start;
    public $column_index;

    public $offsetX = 10;
    public $offsetY = 10;
    public $width = 20;
    public $height= 20;
    public $img;

    public $amount_column;
    public $amount_row;
   
   	private $objWorksheet;
   	private $gcell;
    
    function __construct($obj) 
        {
            $this->objPHPExcel = $obj;
        }
    
    function IND($ind)
        {
            $this->inx = $ind;
            $this->objPHPExcel->setActiveSheetIndex($this->inx);
        }
    
    function ROW_INDEX($col, $cnt)
        {
            $this->row_column  = $col;
            $this->row_content = $cnt;
        }
    
    function CELL($c)
        {
            $this->gcell = $c;
        }                
    
    function CREATE_SHEET( $til="" )
        {
            $this->objWorksheet = $this->objPHPExcel->createSheet();
            $this->objPHPExcel->setActiveSheetIndex($this->inx);
            $this->objPHPExcel->getActiveSheet()->setTitle($til);
            $this->objPHPExcel->getActiveSheet()->getTabColor()->setRGB('FFF000');            
        }

    function CREATE_TEXT( $cell="A1", $cont="")
        {
            $this->objPHPExcel->setActiveSheetIndex($this->inx)->setCellValue($cell, $cont);
        }
    
    function CREATE_HEAD( $data=array() , $rn = false)
        {
            $col =   $this->column_start;
            if($rn)  $this->CREATE_TEXT( $this->column_index[$col-1].($this->row_column) , "NO." );
            foreach ( $data as $body => $val)
                {
                    $this->CREATE_TEXT($this->column_index[$col++].($this->row_column), trim( $body ) );
                }
        }
    
    function CREATE_BODY( $data=array(), $rn = false )
        {
            $rw = $this->row_content; 
            foreach ( $data as $body => $value)
                {
                   
                    $col = $this->column_start; 
                    if($rn)  $this->CREATE_TEXT($this->column_index[$col-1].$rw , sprintf( "=SUBTOTAL(3,$%s$%s:$%s$%s)",($this->column_index[$col]), $this->row_content,($this->column_index[$col]), $body+$this->row_content )  );
                    foreach($value as $key => $val)
                        {
                            $this->CREATE_TEXT($this->column_index[$col++].$rw , $val );
                        }
                    $rw++;
                } 
                $this->CREATE_TEXT($this->column_index[$col-1].$rw , sprintf(  "  " ) );
        }
    
    function CREATE_FILTER($cellST="A1", $cellEN="B1")
        {
        	$cell = ( $cellST == $cellEN ) ? $cellST : $cellST . ":" . $cellEN;
            $this->objPHPExcel->getActiveSheet()->setAutoFilter($cell);
        }
    
    function CREATE_ZOOMSCALE($z)
        {
            $this->objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale($z);
        }

    function CREATE_MERGECELL($cellST='A1', $cellEN='A1')
    	{
    		//$cell = ( $cellST == $cellEN ) ? $cellST : $cellST . ":" . $cellEN;
    		$this->objPHPExcel->getActiveSheet()->mergeCells( $cellST .":" . $cellEN );
    	}
    
    function REMOVE_MERGECELL($cellST='A1', $cellEN='A1')
    	{
    		//$cell = ( $cellST == $cellEN ) ? $cellST : $cellST . ":" . $cellEN;
    		$this->objPHPExcel->getActiveSheet()->unmergeCells( $cellST .":" . $cellEN );
    	}
    
    function CREATE_ROW($rw = 1, $ins = 1)
    	{
    		$this->objPHPExcel->getActiveSheet()->insertNewRowBefore($rw, $ins);
        }
        
    function CREATE_COL($cm = "A", $ins = 1)
    	{ 
            $this->objPHPExcel->getActiveSheet()->insertNewColumnBefore($cm, $ins);
    	}    
    function CREATE_FREEZE($cel)
    	{
    		$this->objPHPExcel->getActiveSheet()->freezePane($cel);
    	}
    
    function CREATE_COLORTAB( $col = "FFFFFF" )
    	{
    		$this->objPHPExcel->getActiveSheet()->getTabColor()->setRGB($col);
    	}
    
    function CREATE_BORDER( $cellST='A1', $cellEN='A1', $color='000000', $ove="allborders", $line='BORDER_NONE' )
        { 
        	$cell = ( $cellST == $cellEN ) ? $cellST : $cellST . ":" . $cellEN;
        	$bordetLine = PHPExcel_Style_Border::BORDER_NONE;
            switch ($line) {
            	case 'BORDER_NONE':
            		$bordetLine = PHPExcel_Style_Border::BORDER_NONE;
            		break;
            	case 'BORDER_DASHDOT':
            		$bordetLine = PHPExcel_Style_Border::BORDER_DASHDOT;
            		break;
            	case 'BORDER_DASHED':
            		$bordetLine = PHPExcel_Style_Border::BORDER_DASHED;# code...
            		break;
            	case 'BORDER_DOTTED':
            		$bordetLine = PHPExcel_Style_Border::BORDER_DOTTED;# code...
            		break;
            	case 'BORDER_DOUBLE':
            		$bordetLine = PHPExcel_Style_Border::BORDER_DOUBLE;# code...
            		break;
            	case 'BORDER_HAIR':
            		$bordetLine = PHPExcel_Style_Border::BORDER_NONE;# code...
            		break;
            	case 'BORDER_MEDIUM':
            		$bordetLine = PHPExcel_Style_Border::BORDER_MEDIUM;# code...
            		break;
            	case 'BORDER_MEDIUMDASHDOT':
            		$bordetLine = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT;# code...
            		break;
            	case 'BORDER_MEDIUMDASHDOTDOT':
            		$bordetLine = PHPExcel_Style_Border::BORDER_NONE;# code...
            		break;
            	case 'BORDER_MEDIUMDASHED':
            		$bordetLine = PHPExcel_Style_Border::BORDER_MEDIUMDASHED;# code...
            		break;
            	case 'BORDER_SLANTDASHDOT':
            		$bordetLine = PHPExcel_Style_Border::BORDER_SLANTDASHDOT;# code...
            		break;
            	case 'BORDER_THICK':
            		$bordetLine = PHPExcel_Style_Border::BORDER_THICK;# code...
            		break;
            	case 'BORDER_THIN':
            		$bordetLine = PHPExcel_Style_Border::BORDER_THIN;# code...
            		break;            		            		            		            		            		            		            		            		            		            		            		            	
            	default:
            		$bordetLine = PHPExcel_Style_Border::BORDER_NONE;# code...
            		break;            		
            }
           //var_dump($bordetLine); exit;
           $this->objPHPExcel->getActiveSheet()->getStyle( $cell )->applyFromArray( array('borders' => array( $ove => $this->STYLE_BORDER( $bordetLine, $color))));   
                                                                                              
        }

    function CREATE_FILL($cellST='A1', $cellEN='A1', $color='FFFFFF')
    	{
    		$cell = ( $cellST == $cellEN ) ? $cellST : $cellST . ":" . $cellEN;
    		$this->objPHPExcel->getActiveSheet()->getStyle( $cell )->applyFromArray(array('fill' => $this->STYLE_FILL($color)));
    	}

    function CREATE_FONT($cellST='A1', $cellEN='A1',  $size=12, $color='000000', $bol=false, $ita=false, $fname='Calibri')
    	{
    		$cell = ( $cellST == $cellEN ) ? $cellST : $cellST . ":" . $cellEN;
    		$this->objPHPExcel->getActiveSheet()->getStyle( $cell )->applyFromArray(array('font' => $this->STYLE_FONT($size, $color, $fname, $bol, $ita)));
    	}
    
    function CREATE_FORMAT($cellST='A1', $cellEN='A1', $format="#,##0")
    	{
    		$cell = ( $cellST == $cellEN ) ? $cellST : $cellST . ":" . $cellEN;
    		$this->objPHPExcel->getActiveSheet()->getStyle( $cell )->getNumberFormat()->setFormatCode($format);
    	}
	
	function STYLE_GRIDLINES($sh = true)
		{
			$this->objPHPExcel->getActiveSheet()->setShowGridlines( $sh );
		}    	

    function STYLE_TEXTROTATION($cell = 'A1', $n=0)
        {
            $this->objPHPExcel->getActiveSheet()->getStyle( $cell )->getAlignment()->setTextRotation($n);     
        }
    
    function STYLE_TEXTROTATION_RANGE($numST=0, $numEN=0, $r=1, $n=0)
        {
            foreach (range($numST, $numEN) as $id )
                $this->objPHPExcel->getActiveSheet()->getStyle( $this->column_index[$id] . $r  )->getAlignment()->setTextRotation($n);     
        }
    
    function STYLE_FILL($color=null) 
        {

            return array( 'type'  => PHPExcel_Style_Fill::FILL_SOLID,                           
                          'color' => array('rgb' => $color)                                    
                    );                                   
        }

    function STYLE_FILL_CONDITION($color_st=null, $color_en=null)
        {

            return array( 'type'   => PHPExcel_Style_Fill::FILL_SOLID,                           
                      'startcolor' =>array('argb' => $color_st),
                      'endcolor'   =>array('argb' => $color_en)                             
                    );                                   
        }
    
    function STYLE_FONT($size=11, $color='FFFFFF', $fname='Consolas', $bol=false, $ita=false, $unl=false,$sup=false, $sub=false, $ske=false ) 
        {

            return  array(
	                        'name' 		 => $fname,
	                        'size'       => $size,
	                        'bold'       => $bol,
	                        'italic'     => $ita,
	                        //'underline'  => $unl,
	                        //'superScript'=> $sup,
	                        //'subScript'  => $sub,
	                        //'strike'     => $ske,
	                        'color'      => array('rgb' => $color)
                     	);                               
        }
    
    function STYLE_FONT_CONDITION($size=11, $color='FFFFFF', $fname='Consolas', $bol=false, $ita=false, $unl=false,$sup=false, $sub=false, $ske=false ) 
        {

            return  array(
                            'name'       => $fname,
                            'size'       => $size,
                            'bold'       => $bol,
                            'italic'     => $ita,
                            //'underline'  => $unl,
                            //'superScript'=> $sup,
                            //'subScript'  => $sub,
                            //'strike'     => $ske,
                            'color'      => array('argb' => $color)
                        );                               
        }
    
    function STYLE_BORDER(  $line=PHPExcel_Style_Border::BORDER_THIN, $color='000000' )
        { 
        	//echo $line; exit;
        	return array( 'style' => $line, 'color' => array('rgb' => $color)) ;
        }        
    
    function STYLE_ALIGNMENT($cellST='A1', $cellEN='A1', $sty='CC', $swt=false)
        {
        	$cell = ( $cellST == $cellEN ) ? $cellST : $cellST . ":" . $cellEN;
            switch ($sty) {
                case "BC": #bottom->center
                        $this->objPHPExcel->getActiveSheet()
                            ->getStyle($cell)
                            ->getAlignment()
                            ->setWrapText($swt)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_BOTTOM)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                    break;
                case "TC": #top->center
                        $this->objPHPExcel->getActiveSheet()
                            ->getStyle($cell)
                            ->getAlignment()
                            ->setWrapText($swt)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                    break;
                case "CC": #center->center
                        $this->objPHPExcel->getActiveSheet()
                            ->getStyle($cell)
                            ->getAlignment()
                            ->setWrapText($swt)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                    break;
                case "BR": #bottom->right
                        $this->objPHPExcel->getActiveSheet()
                            ->getStyle($cell)
                            ->getAlignment()
                            ->setWrapText($swt)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_BOTTOM)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                           // echo $sty; exit;
                    break;
                case "TR": #top->right
                        $this->objPHPExcel->getActiveSheet()
                            ->getStyle($cell)
                            ->getAlignment()
                            ->setWrapText($swt)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    break;
                case "CR": #center->right
                        $this->objPHPExcel->getActiveSheet()
                            ->getStyle($cell)
                            ->getAlignment()
                            ->setWrapText($swt)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                    break;
                case "BL": #bottom->left
                        $this->objPHPExcel->getActiveSheet()
                            ->getStyle($cell)
                            ->getAlignment()
                            ->setWrapText($swt)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_BOTTOM)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                           // echo $cell; exit;
                    break;
                case "TL": #top->left
                        $this->objPHPExcel->getActiveSheet()
                            ->getStyle($cell)
                            ->getAlignment()
                            ->setWrapText($swt)
                            ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
                            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                    break;
                case "CL": #center->left
                        $this->objPHPExcel->getActiveSheet()
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

    function STYLE_GROUP_COLUMN($index=0, $level=1, $vi=false, $co=true)
        {
            $this->objPHPExcel->getActiveSheet()->getColumnDimension ($this->column_index[$index])->setOutlineLevel($level);
            $this->objPHPExcel->getActiveSheet()->getColumnDimension ($this->column_index[$index])->setVisible($vi);
            $this->objPHPExcel->getActiveSheet()->getColumnDimension ($this->column_index[$index])->setCollapsed($co); 
        }
    function STYLE_GROUP_COLUMNS($inst=0, $ined=1, $level=1, $vi=false, $co=true)
        {
            foreach(range($inst, $ined) as $inx){
                $this->objPHPExcel->getActiveSheet()->getColumnDimension ($this->column_index[$inx])->setOutlineLevel($level);
                $this->objPHPExcel->getActiveSheet()->getColumnDimension ($this->column_index[$inx])->setVisible($vi);
                $this->objPHPExcel->getActiveSheet()->getColumnDimension ($this->column_index[$inx])->setCollapsed($co); 
            }
        }    
    function STYLE_GROUP_ROW($index=0, $vi=false, $co=true)
        {
            $this->objPHPExcel->getActiveSheet()->getRowDimension($index)->setOutlineLevel(1);
            $this->objPHPExcel->getActiveSheet()->getRowDimension($index)->setVisible($vi);
            $this->objPHPExcel->getActiveSheet()->getRowDimension($index)->setCollapsed($co); 
        }
    function STYLE_GROUP_ROWS($inst=0, $ined=1, $vi=false, $co=true)
        {
            foreach(range($inst, $ined) as $inx){
                $this->objPHPExcel->getActiveSheet()->getRowDimension($inx)->setOutlineLevel(1);
                $this->objPHPExcel->getActiveSheet()->getRowDimension($inx)->setVisible($vi);
                $this->objPHPExcel->getActiveSheet()->getRowDimension($inx)->setCollapsed($co); 
            }
        }
    function STYLE_WIDTH( $dime = "A", $wid = '15' )
    	{
    		$this->objPHPExcel->getActiveSheet()->getColumnDimension($dime)->setWidth( ( (float)$wid + 0.71 ) );
    	}
    
    function STYLE_HEIGHT( $dime = "A", $hei = '15' )
    	{
    		$this->objPHPExcel->getActiveSheet()->getRowDimension( $dime )->setRowHeight( $hei );
    	}
    
    function STYLE_WIDTH_RANGE( $numST=0, $numEN=0, $wid = 15 )
    	{
    		foreach (range($numST, $numEN) as $id ) 
             	$this->objPHPExcel->getActiveSheet()->getColumnDimension( $this->column_index[$id] )->setWidth( ( (float)$wid + 0.71 ) );

    	}
    
    function STYLE_HEIGHT_RANGE( $numST=0, $numEN=0, $hei = 15 )
    	{
    		foreach (range($numST, $numEN) as $id ) 
             	$this->objPHPExcel->getActiveSheet()->getRowDimension( $id )->setRowHeight( $hei );
    	}
    
    function STYLE_ZOOMSCALE($z = 100)
        {
            $this->objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale($z);
        } 

    function OUTPUT_INDEXSHEET()
        {
            return $this->objPHPExcel->getSheetCount();
        }    
    
    function OUTPUT_FILE($namefile)
        {
            //$namefile = "Query_pods_data.sql";
            $file = $namefile; 
            //echo basename($file); exit;
            header("Content-Description: File Transfer"); 
            header("Content-Type: application/octet-stream"); 
            header("Content-Disposition: attachment; filename=".basename($file) ); 
            readfile ($file);
            exit;
        }
	
	function OUTPUT_COLUMNINDEX( $col = 30)
		{
			return $this->CREATE_COLUMNINDEX( $col );
		}
	
	function OUTPUT_COLUMNNAMES( $col = 30)
		{
			return $this->CREATE_COLUMNNAMES( $col );
		} 		          
	
	function CREATE_COLUMNNAMES( $col = 30, $col_array = array() )
		{
			$amColumn = ceil( $col / 26);
			$colind = 1;
			foreach ( range( 0, (int)$amColumn ) as $r ) 
			{
				$colNm = "";
				if (  $r == 0  )
					{
						foreach ( range('A', 'Z') as $cm )$col_array[ $cm ] = $colind++ ;					
					}
				else
					{
						foreach ( range('A', 'Z') as $cm ) 
							{ 
								$colNm = $cm;
								foreach ( range('A', 'Z') as $c )
									{ 
										$col_array[ $colNm.$c ] = $colind++ ;

										if( $colind >= $col+1) return $col_array;
									}							
							}
					}
				
			}
			return $col_array;
		}        
	
	function CREATE_COLUMNINDEX($col = 30, $col_array = array() )
		{
			$amColumn = ceil( $col / 26);
			$colind = 1;
			foreach ( range( 0, (int)$amColumn ) as $r ) 
			{
				$colNm = "";
				if (  $r == 0  )
					{
						foreach ( range('A', 'Z') as $cm )$col_array[ $colind++ ] = $cm ;					
					}
				else
					{
						foreach ( range('A', 'Z') as $cm ) 
							{ 
								$colNm = $cm;
								foreach ( range('A', 'Z') as $c )
									{ 
										$col_array[ $colind++ ] = $colNm.$c ;

										if( $colind >= $col+1) return $col_array;
									}							
							}
					}
				
			}
			return $col_array;
		}     

    function OUTPUT_CONDITIONFORMAT($op="=", $ct='0')
        {
            $p=null;
            $c=null;
            switch ($op) {
                case '=':
                    $p = PHPExcel_Style_Conditional::OPERATOR_EQUAL;
                    $c = PHPExcel_Style_Conditional::CONDITION_CELLIS;
                    break;                    
                case '>':
                    $p = PHPExcel_Style_Conditional::OPERATOR_GREATERTHAN;
                    $c = PHPExcel_Style_Conditional::CONDITION_CELLIS;
                    break;                    
                case '<':
                    $p = PHPExcel_Style_Conditional::OPERATOR_LESSTHAN;
                    $c = PHPExcel_Style_Conditional::CONDITION_CELLIS;
                    break;                    
                case '>=':
                    $p = PHPExcel_Style_Conditional::OPERATOR_GREATERTHANOREQUAL;
                    $c = PHPExcel_Style_Conditional::CONDITION_CELLIS;
                    break;                    
                case '<=':
                    $p = PHPExcel_Style_Conditional::OPERATOR_LESSTHANOREQUAL;
                    $c = PHPExcel_Style_Conditional::CONDITION_CELLIS;
                    break;                    
                case '!=':
                    $p = PHPExcel_Style_Conditional::OPERATOR_NOTEQUAL;
                    $c = PHPExcel_Style_Conditional::CONDITION_CELLIS;
                    break;                    
                case 'YF':
                    $p = PHPExcel_Style_Conditional::OPERATOR_CONTAINSTEXT;
                    $c = PHPExcel_Style_Conditional::CONDITION_CONTAINSTEXT;
                    break;                    
                case 'NF':
                    $p = PHPExcel_Style_Conditional::OPERATOR_NOTCONTAINS;
                    $c = PHPExcel_Style_Conditional::CONDITION_CONTAINSTEXT;
                    break;                    
                case 'BW':
                    $p = PHPExcel_Style_Conditional::OPERATOR_BETWEEN;
                    $c = PHPExcel_Style_Conditional::CONDITION_EXPRESSION;
                    break;                    
                case 'BG':
                    $p = PHPExcel_Style_Conditional::OPERATOR_BEGINSWITH;
                    $c = PHPExcel_Style_Conditional::CONDITION_EXPRESSION;                                                                                                                                                                
                    break;
                case 'ES':
                    $p = PHPExcel_Style_Conditional::OPERATOR_ENDSWITH;
                    $c = PHPExcel_Style_Conditional::CONDITION_EXPRESSION;                                                                                                                                                                
                    break;                
                default:
                    $p = PHPExcel_Style_Conditional::OPERATOR_NONE;
                    $c = PHPExcel_Style_Conditional::CONDITION_CELLIS; 
                    break;
            }


            $objConditional1 = new PHPExcel_Style_Conditional();
            $objConditional1->setConditionType($c)->setOperatorType ($p)->addCondition( $ct );
            return  $objConditional1;            
        }
 
    function CREATE_CONDITIONFORMAT($cellST='A1', $cellEN='A1', $cdt=null)
        {
            $cell = ( $cellST == $cellEN ) ? $cellST : $cellST . ":" . $cellEN;
            $this->objPHPExcel->getActiveSheet()->getStyle( $cell )->setConditionalStyles( $cdt );
        }

    function STYLE_CONDITIONFORMAT_NUMBER($ct=null, $fm="#,##0")
        {
            $ct->getStyle()->getNumberFormat()->setFormatCode($fm);
        }         
    
    function STYLE_CONDITIONFORMAT_FONT($ct=null, $s=12, $c="000000",$f='Calibri', $b=false, $i=false)
        {
            $ct->getStyle()->applyFromArray(array( 'font' => $this->STYLE_FONT_CONDITION($s,$c,$f,$b,$i)));
                                                                          
        }         
    
    function STYLE_CONDITIONFORMAT_FILL($ct=null, $c1="0000FF", $c2="FF0000")
        {
            $ct->getStyle()->applyFromArray(array( 'fill' => $this->STYLE_FILL_CONDITION($c1, $c2)));
        } 

    function APPEND_CONDITIONFORMAT( $ct, $cellST='A1', $cellEN='A1' )
        {
            $cell = ( $cellST == $cellEN ) ? $cellST : $cellST . ":" . $cellEN;
            $conditionalStyles = $this->objPHPExcel->getActiveSheet()->getStyle($cellST)->getConditionalStyles();
            array_push($conditionalStyles, $ct);
            $this->objPHPExcel->getActiveSheet()->getStyle($cell)->setConditionalStyles($conditionalStyles);            
        }
    
    function DUPLICATE_STYLE( $cms='A1', $cmn='A1', $cds='A1', $cdn='A1')
        {
            $cellM = ( $cms == $cmn ) ? $cms : $cms . ":" . $cmn;
            $cellD = ( $cds == $cdn ) ? $cds : $cds . ":" . $cdn;
            $this->objPHPExcel->getActiveSheet()->duplicateStyle( $this->objPHPExcel->getActiveSheet()->getStyle($cellM), $cellD );
        }
    
    function CREATE_RUNFOEMULA_ROWDATA( $numST=1, $numEN=1, $r=1, $rs = 1, $re = 1, $fml="")
        {
            foreach( range($numST, $numEN) as $c )
                $this->CREATE_TEXT( $this->column_index[$c] . $r , sprintf( $fml, $this->column_index[$c].$rs, $this->column_index[$c].$re ) );
        }

    function GETDATA_ONCELL($c)
        {
            $formula = $this->objPHPExcel->getActiveSheet()->getCell($c)->getValue();
            return $formula; 
        }

    function GETDATA_SHEETNAME( )
        {
            $shname = $this->objPHPExcel->getActiveSheet()->getTitle();
            return $shname; 
        }

    function SETOFFSETX($x = 0)
        {
            $this->offsetX = ( $x > 0 ) ? $x : 10;
        } 
    function SETOFFSETY($y = 0)
        {
            $this->offsetY = ( $y > 0 ) ? $y : 10;
        } 
    function SETWIDTH($w = 0)
        {
            $this->width = ( $w > 0 ) ? $w : 20;
        } 
    function SETHEIGHT($h = 0)
        {
            $this->height = ( $h > 0 ) ? $h : 20;
        } 
    function SETIMGAGE($m = "")
        {
            $this->img = ( $m != "" ) ? $m : "none";
        }                                 
    function CREATE_IMAMG($c='A1', $m="", $x=0, $y=0, $w=0, $h=0)
        {
            $this->SETIMGAGE($m);            
            $this->SETOFFSETX($x);
            $this->SETOFFSETY($y);
            $this->SETWIDTH($w);
            $this->SETHEIGHT($h);
            //$gdImage = dirname(__FILE__) . '/img/NEW-TBKK-LOGO_0.png';
            // Add a drawing to the worksheetecho date('H:i:s') . " Add a drawing to the worksheet\n";
            $objDrawing = new PHPExcel_Worksheet_Drawing();
            $objDrawing->setName('Sample image');
            $objDrawing->setDescription('Sample image');
            $objDrawing->setPath($this->img);
            //$objDrawing->setRenderingFunction(PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG);
            //$objDrawing->setMimeType(PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT);
            $objDrawing->setOffsetX($this->offsetX); 
            $objDrawing->setOffsetY($this->offsetY);  
            $objDrawing->setHeight($this->height);
            $objDrawing->setWidth($this->width ); 
            $objDrawing->setCoordinates($c);
            $objDrawing->setWorksheet($this->objPHPExcel->getActiveSheet()); 
        }

    function CREATE_COMMENT($c='A1', $txt="no data!",$h=70, $w=300 ,$s=16, $cr="FF0000", $b=true )
        {
            $this->objPHPExcel->getActiveSheet()->getComment($c)->setAuthor('Efusionsoft');
            $objCommentRichText = $this->objPHPExcel->getActiveSheet()->getComment($c)->getText()->createTextRun('Report:');
 
            $this->objPHPExcel->getActiveSheet()->getComment($c)->setWidth(  $w.'pt' );
            $this->objPHPExcel->getActiveSheet()->getComment($c)->setHeight( $h.'pt' );  
            $this->objPHPExcel->getActiveSheet()->getComment($c)->getText()->createTextRun("\r\n"); 
            $this->objPHPExcel->getActiveSheet()->getComment($c)->getText()->createTextRun($txt);
            

            $objCommentRichText->getFont()->setName('Arial');
            $objCommentRichText->getFont()->setSize($s); 
            $objCommentRichText->getFont()->getColor()->setARGB($cr); 
            $objCommentRichText->getFont()->setBold($b);

        }    
 }




// END Template Class

/* End of file Template.php */
/* Location: ./system/application/libraries/Template.php */