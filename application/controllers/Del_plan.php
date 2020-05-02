<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Del_plan extends CI_Controller {

	var $ph = "excel";
	public function __construct()
	{ 
		parent::__construct(); 

		$this->ph = $this->pdsys_ispath( $this->ph );
		$this->load->model("Del_plan_model", "mp");

	}

	public function index()
		{
 

		}
	public function del_plan_report()
		{	
            // $d = $this->mp->del_plan_report();
			
			$dir = $this->pdsys_ispath( sprintf("%s/Delivery-Plan-PH8",$this->ph ) ); 
			$dir = $this->pdsys_ispath( sprintf("%s/%s",$dir, date('Ym') ) );  
			
			$filename = sprintf("%s/%s.xlsx", $dir, "delivery-plan-ph8-". date('Ymd') );

			
			$data["data"] = $this->mp->del_plan_report();

			$data['fln']    = $filename; 
			$this->load->view('delivery/from_report_del_plan',$data);
			
			echo $filename; exit;
		}
		
	private function pdsys_ispath($dir)
		{
			if( is_dir($dir) === false ) mkdir($dir); 
			return $dir;
		}
	//$this->load->library('session');
}

?> 