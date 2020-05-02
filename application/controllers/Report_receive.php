<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Report_receive extends CI_Controller {

	public function __construct()
	{ 
		parent::__construct();



	}

	public function index()
	{


	}
	// public function Receive_month()
	// {
	// 	$data['list_act_report'] = array('receive_monthly' => $this->Backreport_model->mysql_report_service('RECEIVE_MON'),	
	// 									 'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
	// 									);
	// 	$data['title'] = array( "Receive monthly", "Receive history" );
	// 	$data['filename'] = "receive_month";
	// 	$data['colhead']  = "375623";	
	// 	$data['rate'] =  $this->Backreport_model->mysql_report_service('EXC_RATE');	
	// 	$this->load->view('from_report_receive_monthly_wk',$data);		

	// }
	public function Receive_month()
	{
		$this->load->library('sql_query');

		$data['list_act_report'] = array('receive_monthly' => $this->Backreport2_model->get_exec( $this->sql_query->GETDATA_RECEIVEFORCAST() ),	
										 'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
										);
		$data['title'] = array( "Receive monthly", "Receive history" );
		$data['filename'] = "receive_month";
		$data['colhead']  = "375623";	
		$data['rate'] =  $this->Backreport_model->mysql_report_service('EXC_RATE');	
		$this->load->view('from_report_receive_monthly_wk',$data);		

	}
	public function Receive_month_nm()
	{
		$data['list_act_report'] = array('receive_monthly' => $this->Backreport_model->mysql_report_service('NM_RECEIVE_MON'),	
										 'receive_history' => $this->Backreport_model->mysql_report_service('NM_RECEIVE_MON_HIS'),
										);
		$data['title'] = array( "Receive monthly", "Receive history" );
		$data['filename'] = "receive_month";
		$data['colhead']  = "375623";	
		//$data['rate'] =  $this->Backreport_model->mysql_report_service('EXC_RATE');	
		$this->load->view('from_report_receive_monthly_nm',$data);		

	}
    public function Receive_this()
	{
		$data['list_act_report'] = array('receive' => $this->Backreport_model->mysql_report_service('RECEIVE_MON'),	
										 'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
										);
		$data['title'] = array( "Receive", "Receive history" );
		$data['filename'] = "receive_month";
		$data['colhead']  = "375623";	
		$data['rate'] =  $this->Backreport_model->mysql_report_service('EXC_RATE');	
		$this->load->view('from_report_receive_tm',$data);		

	}

	public function Receive_temp()
	{
		$this->load->library('sql_query');

		$data['list_act_report'] = array('receive_monthly' => $this->Backreport2_model->get_exec( $this->sql_query->GETDATA_RECEIVEFORCAST() ),	
										 'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
										);
		$data['title'] = array( "Receive monthly", "Receive history" );
		$data['filename'] = "receive_month";
		$data['colhead']  = "375623";	
		$data['rate'] =  $this->Backreport_model->mysql_report_service('EXC_RATE');	
		$this->load->view('from_report_receive_monthly_tm',$data);		

	}


	//$this->load->library('session');
}
