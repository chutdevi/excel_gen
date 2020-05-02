<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Report_demand extends CI_Controller {

	public function __construct()
	{ 
		parent::__construct();



	}

	public function index()
	{


	}

	public function report()
	{
		// $data['list_act_report'] = array( 'defect'      => $this->Backreport_model->mysql_report_service('DEFECT_REPORT'),
										  // 'code_detail' => $this->Backreport_model->mysql_report_service('QC_CODE'));
		$data['list_act_report'] = array( 'pd01' => $this->Backreport2_model->demand_data("'K1PD01'"),
										  'pd02' => $this->Backreport2_model->demand_data("'K1PD02'"),
										  'pd03' => $this->Backreport2_model->demand_data("'K1PD03'"),
										  'pd04' => $this->Backreport2_model->demand_data("'K1PD04'"),
										  'pd05' => $this->Backreport2_model->demand_data("'K1PD05'"),
										  'pd06_machining' => $this->Backreport2_model->demand_data("'K2PD06'".' AND FG.SOURCE_NAME = ' . "'MACHINING H/BEARING'"),
										  'pd06_washing'   => $this->Backreport2_model->demand_data("'K2PD06'".' AND FG.SOURCE_NAME = ' . "'WASHING H/BEARING'") );
		$data['title'] = array('PD01', 'PD02', 'PD03', 'PD04', 'PD05', 'PD06 MACHINING', 'PD06 WASHING');//WASHING H/BEARING
		$data['filename'] = "G:/vbs_demand/bin/" ."production_demand";

		$Monthol = date('Y/m');
		$data['look_month'] =  31 - date('t');
		$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );	
		$data['limit_dat']  =  date('t', strtotime("+ 0 month" ,strtotime(date('Y-m-01')))) + 0;
		$data['del']  		=  0;
		$data['reof']  = date('Y F d');


		//var_dump($data); exit;
		$this->load->view('from_report_demand', $data);
	}
	public function report_1()
	{
		// $data['list_act_report'] = array( 'defect'      => $this->Backreport_model->mysql_report_service('DEFECT_REPORT'),
										  // 'code_detail' => $this->Backreport_model->mysql_report_service('QC_CODE'));
		$data['list_act_report'] = array( 'pd01' => $this->Backreport2_model->demand_data_1("'K1PD01'"),
										  'pd02' => $this->Backreport2_model->demand_data_1("'K1PD02'"),
										  'pd03' => $this->Backreport2_model->demand_data_1("'K1PD03'"),
										  'pd04' => $this->Backreport2_model->demand_data_1("'K1PD04'"),
										  'pd05' => $this->Backreport2_model->demand_data_1("'K1PD05'"),
										  'lg00' => $this->Backreport2_model->demand_data_1("'K1LG00'"),
										  'pd06_machining' => $this->Backreport2_model->demand_data_1("'K2PD06'".' AND FG.SOURCE_NAME = ' . "'MACHINING H/BEARING'"),
										  'pd06_washing'   => $this->Backreport2_model->demand_data_1("'K2PD06'".' AND FG.SOURCE_NAME = ' . "'WASHING H/BEARING'") );
		$data['title'] = array('PD01', 'PD02', 'PD03', 'PD04', 'PD05', 'LG00', 'PD06 MACHINING', 'PD06 WASHING');//WASHING H/BEARING
		$data['filename'] = "G:/vbs_demand_extend/bin/" ."production_demand";


		$data['look_month'] =  ( 31 - date('t', strtotime("+ 0 month" ,strtotime(date('Y-m-01')))) ) +  ( 31 - date('t', strtotime("+ 1 month" ,strtotime(date('Y-m-01')))) ) +  ( 31 - date('t', strtotime("+ 2 month" ,strtotime(date('Y-m-01')))) );
		$data['limit_dat']  =  (int)( date('t', strtotime("+ 0 month" ,strtotime(date('Y-m-01')))) ) + (int)( date('t', strtotime("+ 1 month" ,strtotime(date('Y-m-01')))) ) + (int)( date('t', strtotime("+ 2 month" ,strtotime(date('Y-m-01')))) ) ;
		$data['reof']  		=  date('Y F', strtotime("+ 0 month" ,strtotime(date('Y-m-01')))) . " To " . date('Y F', strtotime("+ 2 month" ,strtotime(date('Y-m-01')))) ;

		$data['del']  		=  0;

		$mst = date('Y/m/d', strtotime("+ 0 month" ,strtotime(date('Y-m-01'))));
		$men = date('Y/m/t', strtotime("+ 2 month" ,strtotime(date('Y-m-01'))));

		$data['holiday']  = $this->Backreport2_model->get_hol( $mst, $men );

		//print_r($data['holiday']); exit;
		$this->load->view('from_report_demand_extend', $data);
	}
	public function report_2()
	{
		// $data['list_act_report'] = array( 'defect'      => $this->Backreport_model->mysql_report_service('DEFECT_REPORT'),
										  // 'code_detail' => $this->Backreport_model->mysql_report_service('QC_CODE'));
		$data['list_act_report'] = array( 'pd01' => $this->Backreport2_model->demand_data("'K1PD01'"),
										  'pd02' => $this->Backreport2_model->demand_data("'K1PD02'"),
										  'pd03' => $this->Backreport2_model->demand_data("'K1PD03'"),
										  'pd04' => $this->Backreport2_model->demand_data("'K1PD04'"),
										  'pd05' => $this->Backreport2_model->demand_data("'K1PD05'"),
										  'pd06_machining' => $this->Backreport2_model->demand_data("'K2PD06'".' AND FG.SOURCE_NAME = ' . "'MACHINING H/BEARING'"),
										  'pd06_washing'   => $this->Backreport2_model->demand_data("'K2PD06'".' AND FG.SOURCE_NAME = ' . "'WASHING H/BEARING'") );
		$data['title'] = array('PD01', 'PD02', 'PD03', 'PD04', 'PD05', 'PD06 MACHINING', 'PD06 WASHING');//WASHING H/BEARING
		$data['filename'] = "G:/vbs_demand_extend/bin/" ."production_demand";

		$data['look_month'] =  31 - date('t', strtotime("+ 2 month" ,strtotime(date('Y-m-d'))));
		$data['limit_dat']  =  date('t', strtotime("+ 2 month" ,strtotime(date('Y-m-d')))) + 0;
		$data['reof']  = date('Y F', strtotime("+ 2 month" ,strtotime(date('Y-m-d'))));
		$data['del']  		=  1;


		$this->load->view('from_report_demand', $data);
	}
}
