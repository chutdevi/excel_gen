<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Csv extends CI_Controller {

	public function __construct()
	{ 
		parent::__construct();



	}

	public function index()
	{


	}

	public function pu()
	{
		 $tz_object = new DateTimeZone('+0700');
        //date_default_timezone_set('Brazil/East');

	     $datetime = new DateTime();
	     $datetime->setTimezone($tz_object);   

		$cs = $this->input->post('ty_data');

		//echo $cs; exit;

		$las_data =  $this->input->post('lt_data'); 
		$id       =  $this->input->post('nm_data');

		//echo $las_data . "<hr>" . $id . "<hr>" . $cs ; exit;
		$date_start = date('Y/m/d', strtotime($this->input->post('date_start')));
		$date_end   = date('Y/m/d', strtotime($this->input->post('date_end')));

		 $csv  = "Inside Mgt No."                                	   . "," ;
		 $csv .= "Co. code"                                            . "," ;
		 $csv .= "Business pattern code"                               . "," ;
		 $csv .= "Line No."                                            . "," ;
		 $csv .= "Ref. origin type"                                    . "," ;
		 $csv .= "Sales actual Mgt No."                                . "," ;
		 $csv .= "Order Mgt No."                                       . "," ;
		 $csv .= "Sales slip No."                                      . "," ;
		 $csv .= "Ord No."                                             . "," ;
		 $csv .= "Accpt count"                                         . "," ;
		 $csv .= "Receiving check No."                                 . "," ;
		 $csv .= "Chargeable supply No."                               . "," ;
		 $csv .= "Temp. order No."                                     . "," ;
		 $csv .= "Bad disposal history slip No."                       . "," ;
		 $csv .= "Bad disposal history correction count"               . "," ;
		 $csv .= "Bad disposal history correction type"                . "," ;
		 $csv .= "Exception cost history slip No."                     . "," ;
		 $csv .= "Exception cost history correction count"             . "," ;
		 $csv .= "Exception cost history correction type"              . "," ;
		 $csv .= "Generation Proc type"                                . "," ;
		 $csv .= "Pmt recipient code"                                  . "," ;
		 $csv .= "Vendor code"                                         . "," ;
		 $csv .= "Plant code"                                          . "," ;
		 $csv .= "Item No."                                            . "," ;
		 $csv .= "Item name"                                           . "," ;
		 $csv .= "UC"                                                  . "," ;
		 $csv .= "UC type"                                             . "," ;
		 $csv .= "Qt."                                                 . "," ;
		 $csv .= "Measure unit"                                        . "," ;
		 $csv .= "Amt"                                                 . "," ;
		 $csv .= "Discount amount"                                     . "," ;
		 $csv .= "Receipt date"                                        . "," ;
		 $csv .= "Consumption tax code"                                . "," ;
		 $csv .= "Debt Interface flag"                                 . "," ;
		 $csv .= "Debt IF EXEC date"                                   . "," ;
		 $csv .= "Accnt interface flag"                                . "," ;
		 $csv .= "Accnt IF EXEC date"                                  . "," ;
		 $csv .= "AP slip type(black)"                                 . "," ;
		 $csv .= "AP slip type(red)"                                   . "," ;
		 $csv .= "AP stocking Div. code"                               . "," ;
		 $csv .= "AP debt calculation type"                            . "," ;
		 $csv .= "AP dealings type(black)"                             . "," ;
		 $csv .= "AP dealings type(red)"                               . "," ;
		 $csv .= "AP calculation item/auxiliary subject"               . "," ;
		 $csv .= "AP slip code(black)"                                 . "," ;
		 $csv .= "AP slip code(red)"                                   . "," ;
		 $csv .= "AiIF record type column value(debtor)"               . "," ;
		 $csv .= "AiIF record type column value(creditor)"             . "," ;
		 $csv .= "Ai slip type(debtor)"                                . "," ;
		 $csv .= "Ai slip type(creditor)"                              . "," ;
		 $csv .= "Ai issue section(debtor)"                            . "," ;
		 $csv .= "Ai issue section(creditor)"                          . "," ;
		 $csv .= "Ai calculation item(debtor)"                         . "," ;
		 $csv .= "Ai calculation item(creditor)"                       . "," ;
		 $csv .= "Ai auxiliary item(debtor)"                           . "," ;
		 $csv .= "Ai auxiliary item(creditor)"                         . "," ;
		 $csv .= "Ai charge division code(debtor)"                     . "," ;
		 $csv .= "Ai charge division code(creditor)"                   . "," ;
		 $csv .= "Ai detail summary(debtor)"                           . "," ;
		 $csv .= "Ai detail summary(creditor)"                         . "," ;
		 $csv .= "Ai detail outline name type"                         . "," ;
		 $csv .= "Ai consumption tax judgment type(debtor)"            . "," ;
		 $csv .= "Ai consumption tax judgment type(creditor)"          . "," ;
		 $csv .= "Accnt IF EXEC date"                                  . "," ;
		 $csv .= "Including tax or Excluding tax type"                 . "," ;
		 $csv .= "Tax rates 1"                                         . "," ;
		 $csv .= "Tax rates 2"                                         . "," ;
		 $csv .= "Tax rates 3"                                         . "," ;
		 $csv .= "Defect factor class"                                 . "," ;
		 $csv .= "Defect factor code"                                  . "," ;
		 $csv .= "Pre defect discovery origin line code"               . "," ;
		 $csv .= "Pre defect discovery origin vendor code"             . "," ;
		 $csv .= "Progress %"                                          . "," ;
		 $csv .= "Currency code"                                       . "," ;
		 $csv .= "Create slip type"                                    . "," ;
		 $csv .= "Slip Mgt company code"                               . "," ;
		 $csv .= "Ship returned goods flag"                            . "," ;
		 $csv .= "Gr company type"                                     . "," ;
		 $csv .= "Dealings Div."                                       . "," ;
		 $csv .= "Dealings Gr type"                                    . "," ;
		 $csv .= "Bad disposal type"                                   . "," ;
		 $csv .= "Cost processing type"                                . "," ;
		 $csv .= "Partner code"                                        . "," ;
		 $csv .= "UC Mgt company code"                                 . "," ;
		 $csv .= "UC acquisition destination type"                     . "," ;
		 $csv .= "UC acquisition destination traders code"             . "," ;
		 $csv .= "UC basic date"                                       . "," ;
		 $csv .= "UC basic Qt."                                        . "," ;
		 $csv .= "Journalizing judgment type"                          . "," ;
		 $csv .= "Invoice No."                                         . "," ;
		 $csv .= "BOItype"                                             . "," ;
		 $csv .= "ASIA IF EXEC date"                                   . "," ;

		// echo $csv; exit;	
		$cm = '';
		if     ($cs == 3)     $w = "((VEND_CD  LIKE 'T%' OR VEND_CD LIKE 'M%') AND NOT(VEND_CD = 'T10100' OR VEND_CD = 'T11200' OR VEND_CD = 'T11300')) AND  ";
		elseif ($cs == 2)     $w = "(NOT(VEND_CD  LIKE 'T%' OR VEND_CD LIKE 'M%') OR (VEND_CD = 'T10100' OR VEND_CD = 'T11200' OR VEND_CD = 'T11300'))  AND  ";
		elseif ($cs == 33)  { $w = "((VEND_CD  LIKE 'T%' OR VEND_CD LIKE 'M%') AND NOT(VEND_CD = 'T10100' OR VEND_CD = 'T11200' OR VEND_CD = 'T11300')) AND INTERNAL_CTRL_CD > $las_data AND "; $cm = '--'; }
		elseif ($cs == 22)  { $w = "(NOT(VEND_CD  LIKE 'T%' OR VEND_CD LIKE 'M%') OR (VEND_CD = 'T10100' OR VEND_CD = 'T11200' OR VEND_CD = 'T11300'))  AND INTERNAL_CTRL_CD > $las_data AND "; $cm = '--'; }
		elseif ($cs == 11)  { $w = "INTERNAL_CTRL_CD > $las_data AND "; $cm = '--'; }
		else    $w = "----";

		//echo $las_data; exit;
		$data['list_act_report'] = $this->Backreport_model->inf_pu($date_start, $date_end, $w, $cm);

//var_dump($data['list_act_report']); exit;

		$lt_data = 99999999;

		if ( ( count($data['list_act_report'] ) > 1 ) )
			$lt_data = $data['list_act_report'][ (count($data['list_act_report'])-1 ) ]["INTERNAL_CTRL_CD"];
		else
			$cs = 0;



		$data['csv'] = $csv;

		

		//var_dump($data['list_act_report']); exit;
		// $data['title'] = array("Sale report");
		$data['filename'] = "KAIKAKE_". $datetime->format('YmdHis') . ".csv";
		// $data['colhead']  = "CCFFCC";
		// $data['colhead_font']  = "1A1100";		
		$dat = $id . "-" . $lt_data . "-" .  1 . "-" .  $cs . "-" .  $data['filename'];
		
		if ($cs > 10) $this->td($dat);
		
		$data['sta'] = $cs;

		$this->load->view('Export/from_csv',$data);

	}


	public function sa()
	{
		 $tz_object = new DateTimeZone('+0700');
        //date_default_timezone_set('Brazil/East');

	     $datetime = new DateTime();
	     $datetime->setTimezone($tz_object);   

		$cs = $this->input->post('ty_data');

		//echo $cs; exit;

		$las_data =  $this->input->post('lt_data'); 
		$id       =  $this->input->post('nm_data');

		//echo $las_data . "<hr>" . $id . "<hr>" . $cs ; exit;
		$date_start = date('Y/m/d', strtotime($this->input->post('date_start')));
		$date_end   = date('Y/m/d', strtotime($this->input->post('date_end')));

		$csv  = "Inside Mgt No."                                     . "," ;
		$csv .= "Co. code"                                           . "," ;
		$csv .= "Business pattern code"                              . "," ;
		$csv .= "Line No."                                           . "," ;
		$csv .= "Ref. origin type"                                   . "," ;
		$csv .= "Sales actual Mgt No."                               . "," ;
		$csv .= "Order Mgt No."                                      . "," ;
		$csv .= "Sales slip No."                                     . "," ;
		$csv .= "Ord No."                                            . "," ;
		$csv .= "Accpt count"                                        . "," ;
		$csv .= "Receiving check No."                                . "," ;
		$csv .= "Chargeable supply No."                              . "," ;
		$csv .= "Temp. order No."                                    . "," ;
		$csv .= "Bad disposal history slip No."                      . "," ;
		$csv .= "Bad disposal history correction count"              . "," ;
		$csv .= "Bad disposal history correction type"               . "," ;
		$csv .= "Exception cost history slip No."                    . "," ;
		$csv .= "Exception cost history correction count"            . "," ;
		$csv .= "Exception cost history correction type"             . "," ;
		$csv .= "Generation Proc type"                               . "," ;
		$csv .= "Cust code"                                          . "," ;
		$csv .= "Cust item No."                                      . "," ;
		$csv .= "Cust item name"                                     . "," ;
		$csv .= "Last delivery place code"                           . "," ;
		$csv .= "Cust order No."                                     . "," ;
		$csv .= "Plant code"                                         . "," ;
		$csv .= "Item No."                                           . "," ;
		$csv .= "Item name"                                          . "," ;
		$csv .= "UC"                                                 . "," ;
		$csv .= "UC type"                                            . "," ;
		$csv .= "Qt."                                                . "," ;
		$csv .= "Measure unit"                                       . "," ;
		$csv .= "Amt"                                                . "," ;
		$csv .= "Sales date"                                         . "," ;
		$csv .= "Consumption tax code"                               . "," ;
		$csv .= "Accnt interface flag"                               . "," ;
		$csv .= "Accnt IF EXEC date"                                 . "," ;
		$csv .= "AiIF record type column value(debtor)"              . "," ;
		$csv .= "AiIF record type column value(creditor)"            . "," ;
		$csv .= "Ai slip type(debtor)"                               . "," ;
		$csv .= "Ai slip type(creditor)"                             . "," ;
		$csv .= "Ai issue section(debtor)"                           . "," ;
		$csv .= "Ai issue section(creditor)"                         . "," ;
		$csv .= "Ai calculation item(debtor)"                        . "," ;
		$csv .= "Ai calculation item(creditor)"                      . "," ;
		$csv .= "Ai auxiliary item(debtor)"                          . "," ;
		$csv .= "Ai auxiliary item(creditor)"                        . "," ;
		$csv .= "Ai charge division code(debtor)"                    . "," ;
		$csv .= "Ai charge division code(creditor)"                  . "," ;
		$csv .= "Ai detail summary(debtor)"                          . "," ;
		$csv .= "Ai detail summary(creditor)"                        . "," ;
		$csv .= "Ai detail outline name type"                        . "," ;
		$csv .= "Ai consumption tax judgment type(debtor)"           . "," ;
		$csv .= "Ai consumption tax judgment type(creditor)"         . "," ;
		$csv .= "Accnt IF EXEC date"                                 . "," ;
		$csv .= "Including tax or Excluding tax type"                . "," ;
		$csv .= "Tax rates 1"                                        . "," ;
		$csv .= "Tax rates 2"                                        . "," ;
		$csv .= "Tax rates 3"                                        . "," ;
		$csv .= "Currency code"                                      . "," ;
		$csv .= "Create slip type"                                   . "," ;
		$csv .= "Slip Mgt company code"                              . "," ;
		$csv .= "Ship returned goods flag"                           . "," ;
		$csv .= "Gr company type"                                    . "," ;
		$csv .= "Dealings Div."                                      . "," ;
		$csv .= "Dealings Gr type"                                   . "," ;
		$csv .= "Bad disposal type"                                  . "," ;
		$csv .= "Cost processing type"                               . "," ;
		$csv .= "Partner code"                                       . "," ;
		$csv .= "UC Mgt company code"                                . "," ;
		$csv .= "UC acquisition destination type"                    . "," ;
		$csv .= "UC acquisition destination traders code"            . "," ;
		$csv .= "UC basic date"                                      . "," ;
		$csv .= "UC basic Qt."                                       . "," ;
		$csv .= "Journalizing judgment type"                         . "," ;
		$csv .= "Invoice No."                                        . "," ;
		$csv .= "BOItype"                                            . "," ;
		$csv .= "ASIA IF EXEC date"                                  ;

		// echo $csv; exit;	
		// $cm = '';
		// if     ($cs == 3)     $w = "((CUST_CD  LIKE 'T%' OR CUST_CD LIKE 'F%') AND NOT(CUST_CD = 'T10100' OR CUST_CD = 'T11200' OR CUST_CD = 'T11300'))  AND";
		// elseif ($cs == 2)     $w = "((NOT CUST_CD  LIKE 'T%' OR CUST_CD LIKE 'F%') OR (CUST_CD = 'T10100' OR CUST_CD = 'T11200' OR CUST_CD = 'T11300'))  AND";
		// else   { $w = "---"; $cm = '--'; }

		$cm = '';
		if     ($cs == 3)     $w = "AND ((CUST_CD  LIKE 'T%' OR CUST_CD LIKE 'F%') AND NOT(CUST_CD = 'T10100' OR CUST_CD = 'T11200' OR CUST_CD = 'T11300'))   ";
		elseif ($cs == 2)     $w = "AND ((NOT CUST_CD  LIKE 'T%' OR CUST_CD LIKE 'F%') OR (CUST_CD = 'T10100' OR CUST_CD = 'T11200' OR CUST_CD = 'T11300'))   ";
		elseif ($cs == 33)  { $w = "AND ((CUST_CD  LIKE 'T%' OR CUST_CD LIKE 'F%') AND NOT(CUST_CD = 'T10100' OR CUST_CD = 'T11200' OR CUST_CD = 'T11300'))  AND INTERNAL_CTRL_CD > $las_data  "; $cm = '--'; }
		elseif ($cs == 22)  { $w = "AND ((NOT CUST_CD  LIKE 'T%' OR CUST_CD LIKE 'F%') OR (CUST_CD = 'T10100' OR CUST_CD = 'T11200' OR CUST_CD = 'T11300'))  AND INTERNAL_CTRL_CD > $las_data  "; $cm = '--'; }
		elseif ($cs == 11)  { $w = "AND INTERNAL_CTRL_CD > $las_data "; $cm = '--'; }
		else   { $w = "----";  }

		$data['list_act_report'] = $this->Backreport_model->inf_sa($date_start, $date_end, $w, $cm);

		$data['csv'] = $csv;

		

		//var_dump($data['list_act_report']); exit;
		// $data['title'] = array("Sale report");
		$lt_data = 99999999;

		if ( (count($data['list_act_report']) > 0 ) )
			$lt_data = $data['list_act_report'][ (count($data['list_act_report'])-1 ) ]["INTERNAL_CTRL_CD"];
		else
			$cs = 0;



		$data['filename'] = "URIKAKE_". $datetime->format('YmdHis') . ".csv";
		// $data['colhead']  = "CCFFCC";
		// $data['colhead_font']  = "1A1100";

		$dat = $id . "-" . $lt_data . "-" .  2 . "-" .  $cs . "-" .  $data['filename'];

		//echo $dat; exit;
		if ($cs > 10) $this->td($dat);		

		$data['sta'] = $cs;
		
		$this->load->view('Export/from_csv_1',$data);

	}

		public function raw_data_csv()
		{

			//s$this->load->view('Export/export_csv',$data);

			echo "string";

		}


	public function td( $dat )
	{
					$content = file_get_contents('http://192.168.161.102/report_access/Api_tool/api_test/'. $dat  );
					// แปลงข้อมูลที่รับมาในรูป json มาเป็น array จะได้ใช้ง่าย ๆ
					$DATA = json_decode($content, true);

					// //dump ข้อมูลออกมาดู
					//print_r($DATA);
					//print_r($DATA);
					// ลองดึงออกทีล่ะค่า
					//echo "<hr>";


      				//exit;


	}

}

?>