<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Report_defect extends CI_Controller {

	public function __construct()
	{ 
		parent::__construct();



	}

	public function index()
	{


	}

	public function Ng_weekly()
	{
		$data['list_act_report'] = array( 'defect'      => $this->Backreport_model->mysql_report_service('DEFECT_REPORT'),
										  'code_detail' => $this->Backreport_model->mysql_report_service('QC_CODE'));

		$data['title'] = array('DEFECT', 'CODE DETAIL') ;//WASHING H/BEARING
		$data['filename'] = "gc_weekly";
		$data['colhead']  = "5B90A4";	
		//$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1 : (date('d'));
		//$Monthol = ( (date('d')+0) == 1 ) ? date('Y/m', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('Y/m'));
		//echo $Monthol; exit;
		//$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );
		//$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . date('Y/m') . "'" );
		$this->load->view('from_report_defect_weekly', $data);			
	}

	public function Ng_monthly($mnt = -1)
	{
		$data['list_act_report'] = array( 'defect'      => $this->Backreport_model->mysql_report_service('DEFECT_REPORT_LASTMONTH'),
										  'code_detail' => $this->Backreport_model->mysql_report_service('QC_CODE'));

		$data['title'] = array('DEFECT', 'CODE DETAIL') ;//WASHING H/BEARING
		$data['filename'] = "gc_monthly";
		$data['colhead']  = "5B90A4";
		$data['mnt']  = $mnt;
		//$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1 : (date('d'));
		//$Monthol = ( (date('d')+0) == 1 ) ? date('Y/m', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('Y/m'));
		//echo $Monthol; exit;
		//$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );
		//$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . date('Y/m') . "'" );
		$this->load->view('from_report_defect_monthly', $data);	
	}

}
