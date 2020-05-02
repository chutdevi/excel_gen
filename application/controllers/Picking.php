<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Picking extends CI_Controller {

	public function __construct()
	{ 
		parent::__construct();



	}

	public function index()
	{


	}

	public function pk($data_trn)
	{

		$ar_data = explode("--",$data_trn);
		//var_dump($ar_data)	; exit;
		$delivery_date = date('Y/m/d',strtotime($ar_data[0]));


										 //'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
										
		$data['title'] = array( "PICKING" );
		$data['filename'] = "Picking";
		$data['colhead']  = "305496";
		$data['de_date']  = $delivery_date;	
		$data['de_cust']  = $ar_data[1];	

		//var_dump($data); exit;
		//$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1 : (date('d'));
		$Monthol = date('Y/m');
		//echo $Monthol; exit;
		//$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );
		if( $ar_data[1] == 'IGCE' || $ar_data[1] == 'IGCE-BIZ' || $ar_data[1] == 'IGCE-KD' || $ar_data[1] == 'IMCT' || $ar_data[1] == 'IGCE-SERVICE' || $ar_data[1] == 'MMCH' || $ar_data[1] == 'IMCT-IGCE'  || $ar_data[1] == 'HMMT'){
			$data['list_act_report'] = array( 'picking' => $this->Backreport_model->picking_list( $delivery_date, $ar_data[1], '' ) );
			$this->load->view('from_report_pickking_bake',$data);	

		}elseif(  $ar_data[1] == 'IEMT-SUM' || $ar_data[1] == 'IEMT-BKT'){
			$data['list_act_report'] = array( 'picking' => $this->Backreport_model->picking_list_iemt( $delivery_date, $ar_data[1] ) );
			$this->load->view('from_report_pickking_iemt',$data);	
		}elseif(  $ar_data[1] == 'SKC-SUM'){
			$data['list_act_report'] = array( 'picking' => $this->Backreport_model->picking_list_skc( $delivery_date, $ar_data[1] ) );
			$this->load->view('from_report_pickking',$data);				
		}else{	
			$data['list_act_report'] = array( 'picking' => $this->Backreport_model->picking_list( $delivery_date, $ar_data[1] ) );
		    $this->load->view('from_report_pickking',$data);		
		}

	}
}

?>