<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Report_prodsys extends CI_Controller {

	var $ph = "G:/excel";
	public function __construct()
		{ 
			parent::__construct(); 

			$this->ph = $this->pdsys_ispath( $this->ph );
			$this->load->model("Prdsys_Model", "mp");

		}

	public function index()
		{
 

		}
	public function pdsys_report()
		{	
			$d = $this->mp->getdate_req();
			
			$dir = $this->pdsys_ispath( sprintf("%s/Production-report",$this->ph ) ); 
			$dir = $this->pdsys_ispath( sprintf("%s/%s",$dir, date('Ym', strtotime($d) ) ) );  
			
			$filename = sprintf("%s/%s.xlsx", $dir, $this->mp->getfilename_req(). date('Ymd', strtotime($d) ) );

			
			$seq = $this->mp->model_create_file_seq( $filename );
			//echo $filename;exit;
			$filename = ( $seq == "001" ) ? $filename : sprintf("%s/%s.xlsx", $dir, $this->mp->getfilename_req() . date('Ymd', strtotime($d)). $seq ); 
			
			$data["data"] = $this->mp->model_datareport()["data"];
			$data["hist"] = $this->mp->model_datareport()["hist"];

			$data["days"]   = $d;
			$data["hol1"]  = $this->mp->model_holiday();
			$data["sat1"]  = $this->mp->model_saturdaty();
			$data['fln']    = $filename; 
			$this->load->view('prod/from_report_prod',$data);
			
			echo $this->mp->model_insert_file_list( $filename )[0]["NSEQ"]; exit;
		}
		
		



	private function pdsys_ispath($dir)
		{
			if( is_dir($dir) === false ) mkdir($dir); 
			return $dir;
		}
	//$this->load->library('session');
}

?> 