<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Report_fasys extends CI_Controller {

	public function __construct()
		{ 
			parent::__construct();


			$this->load->model("Fasys_Model", "mf");
		}

	public function index()
		{


		}
	public function Fasys_history()
		{	
			$data['list_act_report'] = array( 'fa_actual_history' => array( $this->mf->model_fahistory(10)
																		  , $this->mf->model_fahistory(8)));

			// var_dump($data['list_act_report']); exit;
			$data['title'] = array( "Fa actual history");
			$data['filename'] = "fa_actual_history";
			$data['colhead']  = "375623";	
			$data['rate'] =  $this->Backreport_model->mysql_report_service('EXC_RATE');	
			$this->load->view('from_fa_actual_history',$data);	
	
		}

	//$this->load->library('session');
}
