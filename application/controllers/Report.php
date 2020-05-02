<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Report extends CI_Controller {

	public function __construct()
	{ 
		parent::__construct();



	}

	public function index()
	{


	}

	public function qcd_report()
	{
 			
 		//echo "Export report";exit;
		//var_dump($this->Backreport_model->work_day()); exit;
		$data['list_act_report'] = array('fg_stock'  => $this->Backreport_model->sql_sv());
		$data['title']    = array("FG Stock") ;
		$data['filename'] = "qcd";
		$data['colhead']  = "808080";
		$data['colhead_font']  = "FF4343";	
		$data['wd'] = $this->Backreport_model->work_day();
		$this->load->view('from_report_evening',$data);		

	}

	//############################################################################################################ Web ##########################################

	public function sal($ds, $dt)
	{


		$date1=date_create($ds);
		$date2=date_create($dt);
		$date_start = date_format($date1,"Y/m/d");
		$date_end   = date_format($date2,"Y/m/d");

		//echo $date_start."<hr>".$date_end; exit;
		//$data['list_act_report'] = array( 'sale_report'  => $this->Backsystem_model->sale_report($date_start, $date_end) );
		//$this->god->welcome();

		$data_sale = $this->Backreport_model->sale_report($date_start, $date_end);
		$data_boi  = $this->Backreport_model->boi();

		//$this->god->boi_rel($data_sale, $data_boi);
		//var_dump( $this->Backreport_model->boi() ); exit;

		$data['list_act_report'] =  array( 'sale_report' => $this->Backreport_model->boi_rel($data_sale, $data_boi) );
		$data['title'] = array("Sale report");
		$data['filename'] = "sale_report";
		$data['colhead']  = "CCFFCC";
		$data['colhead_font']  = "1A1100";		
		$this->load->view('from_report',$data);

	}
	public function sale()
	{


		$date_start = date('Y/m/d', strtotime($this->input->post('date_start')));
		$date_end   = date('Y/m/d', strtotime($this->input->post('date_end')));

		//echo $date_start."<hr>".$date_end; exit;
		//$data['list_act_report'] = array( 'sale_report'  => $this->Backsystem_model->sale_report($date_start, $date_end) );
		//$this->god->welcome();

		$data_sale = $this->Backreport_model->sale_report($date_start, $date_end);
		$data_boi  = $this->Backreport_model->boi();

		//$this->god->boi_rel($data_sale, $data_boi);
		//var_dump( $this->Backreport_model->boi() ); exit;

		$data['list_act_report'] =  array( 'sale_report' => $this->Backreport_model->boi_rel($data_sale, $data_boi) );
		$data['title'] = array("Sale report");
		$data['filename'] = "sale_report";
		$data['colhead']  = "CCFFCC";
		$data['colhead_font']  = "1A1100";		
		$this->load->view('from_report',$data);

	}	
	public function purchase()
	{
		//$date1=date_create($ds);
		//$date2=date_create($dt);
		//$date_start = date_format($date1,"Y/m/d");
		//$date_end   = date_format($date2,"Y/m/d");
		$date_start = date('Y/m/d', strtotime($this->input->post('date_start')));
		$date_end   = date('Y/m/d', strtotime($this->input->post('date_end')));
		//echo $date_start."<hr>".$date_end; exit;
		//$data['list_act_report'] = array( 'sale_report'  => $this->Backsystem_model->sale_report($date_start, $date_end) );
		//$this->god->welcome();

		//$data_sale = $this->Backreport_model->pur_report($date_start, $date_end);

		//echo $this->input->post('ty_data'); exit;

		//$data_boi  = $this->Backreport_model->boi();

		//$this->god->boi_rel($data_sale, $data_boi);
		//var_dump( $this->Backreport_model->boi() ); exit;
		if ( $this->input->post('ty_data') == "2" ) 
			$data['list_act_report'] =  array( 'purchase_report' => $this->Backreport_model->d_pur_report($date_start, $date_end) );
		else
		    $data['list_act_report'] =  array( 'purchase_report' => $this->Backreport_model->o_pur_report($date_start, $date_end) );

		$data['title'] = array("Purchase report");
		$data['filename'] = "purchase_report";
		$data['colhead']  = "00FFCC";
		$data['colhead_font']  = "1A1100";		

		//var_dump($data); exit;
		$this->load->view('from_report_pur',$data);


	}
	//############################################################################################################     Fa Mind    ##########################################
 
	public function Prod_fa_m()
	{
		$data['list_act_report'] = array('fa_daily' => $this->Backreport_model->mysql_report_service('FA_DAILY_REPORT_ACTU_M'));
		$data['title'] = array("FA Daily") ;
		$data['filename'] = "fa_daily";
		$data['colhead']  = "808080";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('from_report_fa_m',$data);	

	}
		public function Fa_loss_ce()
	{
		$data['list_act_report'] = array('fa_daily_loss' => $this->Backreport_model->mysql_report_service('DAILY_REPORT_FOR_CE'));
																		
		$data['title'] = array("FA Daily Loss");
		$data['filename'] = "fa_daily_loss";
		$data['colhead']  = "808080";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('from_report_fa_loss',$data);	

	}
		public function Fa_loss()
	{
		$data['list_act_report'] = array('fa_loss' => $this->Backreport_model->mysql_report_service('LOSS_MANUAL'));
										
										

		$data['title'] = array("Fa Loss");
		$data['filename'] = "fa_daily_loss";
		$data['colhead']  = "808080";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('from_report_fa_loss1',$data);	

	}

	public function Fa_supply()
	{
		$data['list_act_report'] = array('part_supply_pd1' => $this->Backreport_model->mysql_report_service_sup('FA_SUP_LIST'),
										 'part_supply_pd2' => $this->Backreport_model->mysql_report_service_sup('FA_SUP_LIST_PD2'),
										 'part_supply_pd3' => $this->Backreport_model->mysql_report_service_sup('FA_SUP_LIST_PD3'),
										 'part_supply_pd4' => $this->Backreport_model->mysql_report_service_sup('FA_SUP_LIST_PD4'),
										 'part_supply_pd5' => $this->Backreport_model->mysql_report_service_sup('FA_SUP_LIST_PD5'),
										 'part_supply_pd6' => $this->Backreport_model->mysql_report_service_sup('FA_SUP_LIST_PD6'),
										 'part_supply_pcl1' => $this->Backreport_model->mysql_report_service_sup('FA_SUP_LIST_PCL1')
									    );

		$data['title'] = array("Part Supply PD1","Part Supply PD2","Part Supply PD3","Part Supply PD4","Part Supply PD5","Part Supply PD6","Part Supply PCL1") ;
		$data['filename'] = "fa_supply_list";
		$data['colhead']  = "006666";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('from_report_fa_sup_list',$data);	

	}

	public function Fa_remain()
	{
		$data['list_act_report'] = array('fa_remain' => $this->Backreport_model->mysql_fa_remain('FA_REMAIN'));

		$data['title'] = array("Fa remain") ;
		$data['filename'] = "fa_remain";
		$data['colhead']  = "dce6f1";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('from_report_fa_remain',$data);	

	}

	// public function Fa_report_daily()
	// {
	// 	$data['list_act_report'] = array('fa_report_daily' => $this->Backreport_model->mysql_fa_report('OEE_REPORT'),
	// 									  'loss_code' => $this->Backreport_model->mysql_loss_code('LOSS_CODE'));	

	// 	$data['title'] = array("Fa report daily","Loss code") ;
	// 	$data['filename'] = "fa_daily_report";
	// 	$data['colhead']  = "79a6d2";
	// 	$data['colhead_font']  = "FF4343";	
	// 	// var_dump($data)	;
	// 	// exit;	
	// 	$this->load->view('from_report_fa',$data);	

	// }
	public function Fa_report_daily_new()
	{
		$data['list_act_report'] = array('fa_report' => $this->Backreport_model->mysql_fa_report_new('OEE_REPORT'),
										  'loss_code' => $this->Backreport_model->mysql_loss_code('LOSS_CODE'),
										  'shift_code' => $this->Backreport_model->mysql_shift_code('SHIFT_MASTER'));	

		$data['title'] = array("Fa report","Loss code","Shift code") ;
		$data['filename'] = "fa_daily";
		$data['colhead']  = "79a6d2";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('from_report_fa_daily',$data);	// from_report_fa_newversion

	}
	public function Fa_report_daily_new2()
	{
		$data['list_act_report'] = array('fa_report' => $this->Backreport_model->mysql_fa_report_new('OEE_REPORT'),
										  'loss_code' => $this->Backreport_model->mysql_loss_code('LOSS_CODE'));	

		$data['title'] = array("Fa report","Loss code") ;
		$data['filename'] = "fa_daily";
		$data['colhead']  = "79a6d2";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('from_report_fa_newversion2',$data);	

	}
		public function Fa_accum_acc()
	{
		$data['list_act_report'] = array('fa_summary' => $this->Backreport_model->mysql_fa_mon_ac('FA_SUMMARY'));
										 

		$data['title'] = array("Fa summary") ;
		$data['filename'] = "fa_summary";
		$data['colhead']  = "99ccff";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('from_report_fa_mon_ac',$data);	

	}

	public function Loss_code()
	{
		$data['list_act_report'] = array('loss_code' => $this->Backreport_model->mysql_loss_report('IMPOR_DAILY_CODE'));

		$data['title'] = array("loss code") ;
		$data['filename'] = "loss_report";
		$data['colhead']  = "dce6f1";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('from_report_fa',$data);	

	}

	public function Fa_product_cost()
	{
		$data['list_act_report'] = array('production_cost_report' => $this->Backreport_model->mysql_product_cost('OEE_WORK_MONTH'));
										
		$data['title'] = array("Production cost report") ;
		$data['filename'] = "production_cost_report";
		$data['colhead']  = "366092";
		$data['colhead_font']  = "FFFFFF";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('from_report_fa_cost',$data);	

	}
	public function Fa_cost()
	{
		$data['list_act_report'] = array('cost_report' => $this->Backreport_model->mysql_fa_cost('OEE_WORK_MONTH'));
										
		$data['title'] = array("Cost report");
		$data['filename'] = "cost_report";
		$data['colhead']  = "366092";
		$data['colhead_font']  = "FFFFFF";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('Report_cost',$data);	

	}

	public function Fa_monthly()
	{
		$data['list_act_report'] = array('fa_monthly_report' => $this->Backreport_model->mysql_fa_monthly('OEE_WORK_MONTH'));
										// 'loss_code' => $this->Backreport_model->mysql_loss_code('LOSS_CODE')
										
		$data['title'] = array("Fa monthly report"); //,"Loss code"
		$data['filename'] = "Fa_monthly_report";
		$data['colhead']  = "006666";
		$data['colhead_font']  = "000000";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('Report_fa_monthly',$data);	

	}
	public function Fa_monthly_1()
	{
		$data['list_act_report'] = array('fa_monthly_report' => $this->Backreport_model->mysql_fa_monthly('OEE_WORK_MONTH'));
										// 'loss_code' => $this->Backreport_model->mysql_loss_code('LOSS_CODE')
										
		$data['title'] = array("Fa monthly report"); //,"Loss code"
		$data['filename'] = "Fa_monthly_report";
		$data['colhead']  = "009999";
		$data['colhead_font']  = "000000";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('Report_fa_monthly',$data);	

	}
	public function Fa_weekly()
	{
		$data['list_act_report'] = array('fa_weekly_report' => $this->Backreport_model->mysql_fa_weekly('WK_OEE_WORK_MONTH'));
										
										
		$data['title'] = array("Fa weekly report"); 
		$data['filename'] = "Fa_weekly_report";
		$data['colhead']  = "008080";
		$data['colhead_font']  = "FFFFFF";	
		// var_dump($data)	;
		// exit;	
		$this->load->view('fa_weekly',$data);	

	}




	public function Report_request_sheet()
	{
		$data['kla'] = $this->Backreport_model->oracle_request_vend();
		//var_dump($kla); exit;
		$data['list_act_report'] = array('report_request_sheet' => $this->Backreport_model->oracle_request_report());

		$data['title'] = array("Report request sheet","Report request sheet(copy)","Report receiving sheet ","Report Issue sheet ") ;
		$data['filename'] = "request_sheet";
		$data['colhead']  = "dce6f1";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		//echo 'Oor Za bra na'; exit;
		$this->load->view('from_report_request_sheet',$data);	

	}
	public function Report_request_sheet_Oor()
	{
		$data['kla'] = $this->Backreport_model->oracle_request_vend();
		//var_dump($kla); exit;
		$data['list_act_report'] = array('report_request_sheet' => $this->Backreport_model->oracle_request_report());

		$data['title'] = array("Report request sheet","Report request sheet(copy)","Report receiving sheet ","Report Issue sheet ") ;
		$data['filename'] = "request_sheet";
		$data['colhead']  = "dce6f1";
		$data['colhead_font']  = "FF4343";	
		// var_dump($data)	;
		// exit;	
		//echo 'Oor Za bra na'; exit;
		$this->load->view('from_report_request_sheet_Oor',$data);	

	}

	//############################################################################################################   Receive Dew   ##########################################

	public function Receive_month()
	{
		$data['list_act_report'] = array('receive_monthly' => $this->Backreport_model->mysql_report_service('RECEIVE_MON'),	
										 'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
										);
		$data['title'] = array( "Receive monthly", "Receive history" );
		$data['filename'] = "receive_month";
		$data['colhead']  = "375623";	
		$data['rate'] =  $this->Backreport_model->mysql_report_service('EXC_RATE');	
		$this->load->view('from_report_receive_monthly',$data);		

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

	//############################################################################################################ Prod Model Dew  ##########################################

	public function Prod_model()
	{
		$data['list_act_report'] = array( 'pd01' => $this->Backreport_model->mysql_report_service('PROD_PD01'),
										  'pd02' => $this->Backreport_model->mysql_report_service('PROD_PD02'),
										  'pd03' => $this->Backreport_model->mysql_report_service('PROD_PD03'),
										  'pd04' => $this->Backreport_model->mysql_report_service('PROD_PD04'),
										  'pd05' => $this->Backreport_model->mysql_report_service('PROD_PD05'),
										  'pl00' => $this->Backreport_model->mysql_report_service('PROD_PL00'),
										  'pd06' => $this->Backreport_model->mysql_report_service('PROD_PD06')

										);	
										 //'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
										
		$data['title'] = array( "PD01", "PD02", "PD03", "PD04", "PD05", "PL00", "PD06" );
		$data['filename'] = "prod_model";
		$data['colhead']  = "305496";	

		//var_dump($data); exit;
		//$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1 : (date('d'));
		$Monthol = date('Y/m');
		//echo $Monthol; exit;
		$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );
		$this->load->view('from_report_prod_model',$data);		

	}
	public function Prod_model_sum()
	{
		$data['list_act_report'] = array( 'pd01' => $this->Backreport_model->mysql_report_service('PROD_PD01'),
										  'pd02' => $this->Backreport_model->mysql_report_service('PROD_PD02'),
										  'pd03' => $this->Backreport_model->mysql_report_service('PROD_PD03'),
										  'pd04' => $this->Backreport_model->mysql_report_service('PROD_PD04'),
										  'pd05' => $this->Backreport_model->mysql_report_service('PROD_PD05'),
										  'pl00' => $this->Backreport_model->mysql_report_service('PROD_PL00'),
										  'pd06' => $this->Backreport_model->mysql_report_service('PROD_PD06')

										);	
										 //'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
										
		$data['title'] = array( "PD01", "PD02", "PD03", "PD04", "PD05", "PL00", "PD06" );
		$data['filename'] = "prod_model_sum";
		$data['colhead']  = "305496";	

		//var_dump($data); exit;
		//$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1 : (date('d'));
		$Monthol = date('Y/m');
		//echo $Monthol; exit;
		$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );
		$this->load->view('from_report_prod_model_summary',$data);		

	}

	public function pd06_model()
	{
		$data['list_act_report'] = array( 'pd06_mc' => $this->Backreport_model->mysql_report_service('PD06_MC'),
										  'pd06_ws' => $this->Backreport_model->mysql_report_service('PD06_WS')
										);	
										 //'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
										
		$data['title'] = array( "PD06 MC", "PD06 WS" );
		$data['filename'] = "model_pd6";
		$data['colhead']  = "305496";	

		//var_dump($data); exit;
		//$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1 : (date('d'));
		$Monthol = date('Y/m');
		//echo $Monthol; exit;
		$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );
		$this->load->view('from_report_pd6',$data);		

	}

	public function pd06_model_sum()
	{
		$data['list_act_report'] = array( 'pd06_mc' => $this->Backreport_model->mysql_report_service('PD06_MC'),
										  'pd06_ws' => $this->Backreport_model->mysql_report_service('PD06_WS')
										);	
										 //'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
										
		$data['title'] = array( "PD06 MC", "PD06 WS" );
		$data['filename'] = "model_pd6";
		$data['colhead']  = "305496";	

		//var_dump($data); exit;
		//$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1 : (date('d'));
		$Monthol = date('Y/m');
		//echo $Monthol; exit;
		$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );
		$this->load->view('from_report_pd6_summary',$data);		

	}
	//################################################################################################ flutuate order Dew  ##########################################

	public function fluctuate()
	{
		$data['list_act_report'] = array('fluctuation' => $this->Backreport_model->mysql_report_service('FLUTUATE_REPORT'),
										 'fluctuation_history' => $this->Backreport_model->mysql_report_service('FLUTUATE_HIS') );
		$data['title'] = array("FLUCTUATION", "FLUCTUATION HISTORY") ;
		$data['filename'] = "fluctuate";
		$data['colhead']  = "808080" ;		
        $data['wd'] = $this->Backreport_model->work_day();		
        $data['ig'] = $this->Backreport_model->mysql_item_group();
		$this->load->view('from_report_fluctuate',$data);		

	}	
	public function Fg_report()
	{
		$data['list_act_report'] = array('fg_report'       => $this->Backreport_model->FG_report(),	
										 'shipping_plan'   => $this->Backreport_model->FG_ship_report()
										);
		$data['title'] = array("FG report", "Shipping Plan") ;
		$data['filename'] = "fg_report";
		$data['colhead']  = "808080";		
		$this->load->view('Excel/from_report_sale',$data);		

	}

	public function Rm_report()
	{
		$data['list_act_report'] = array('rm_part' => $this->Backreport_model->RM_part());
		$data['title'] = array("RM Part") ;
		$data['filename'] = "rm_part";
		$data['colhead']  = "808080";		
		$this->load->view('Excel/from_report_sale',$data);			

	}

	public function Pods_report()
	{
		$data['list_act_report'] = array('pods_remain' => $this->Backreport_model->Pods_remain_mor());
		$data['title'] = array("PODS Remain") ;
		$data['filename'] = "pods_remain";
		$data['colhead']  = "808080";		
		$this->load->view('Excel/from_report_sale',$data);			

	}		

	public function Ship_report()
	{
		$data['list_act_report'] = array('shipping_remain' => $this->Backreport_model->Shipping_remain_mor());
		$data['title'] = array("Shipping Remain") ;
		$data['filename'] = "ship_remain";
		$data['colhead']  = "808080";		
		$this->load->view('Excel/from_report_sale',$data);			

	}		

	public function Daily_ship()
	{
 		$dayA = date('d');
		$monthA = date('M');
		$yearA = date('Y');
		$day1 = date('d-m-Y', strtotime($yearA."-".$monthA."-".$dayA));
		$day2 = date('d-m-Y', mktime(0, 0, 0, date("m"), date("d")+1, date("Y")));
		$day3 = date('d-m-Y', mktime(0, 0, 0, date("m"), date("d")+2, date("Y")));

		//echo $day1 . "<hr>" . $day2 . "<hr>" . $day3 ;  
		//exit;
		$data['list_act_report'] = array($day1 => $this->Backreport_model->Daily_ship1(), 
										 $day2 => $this->Backreport_model->Daily_ship2(),
										 $day3 => $this->Backreport_model->Daily_ship3());
		$data['title'] = array($day1, $day2, $day3) ;
		$data['filename'] = "daily_shipment_report";
		$data['colhead']  = "808080";		
		$this->load->view('Excel/from_report_sale',$data);			
	}		


	public function Prod_report()
	{
		$data['list_act_report'] = array('all_section' => $this->Backreport_model->Prod_report('RE_ALL'),
										 'k1pd01' => $this->Backreport_model->Prod_report('RE_PD1'), 
										 'k1pd02' => $this->Backreport_model->Prod_report('RE_PD2'),
										 'k1pd03' => $this->Backreport_model->Prod_report('RE_PD3'),
										 'k1pd04' => $this->Backreport_model->Prod_report('RE_PD4'),
									     'k1pd05' => $this->Backreport_model->Prod_report('RE_PD5'),
									     'k2pd06' => array ('machining_h/bearing' => $this->Backreport_model->Prod_report('RE_PD6M'), 'washing_h/bearing' => $this->Backreport_model->Prod_report('RE_PD6W')),
								         'k1pl00' => $this->Backreport_model->Prod_report('RE_PL0'),
								         'production_actual_history' => $this->Backreport_model->Prod_report('Production_history_TEST'),
							             );
		$data['title'] = array('ALL SECTION', 'K1PD01', 'K1PD02', 'K1PD03', 'K1PD04', 'K1PD05', 'K2PD06', 'K1PL00', 'Production actual history') ;//WASHING H/BEARING
		$data['filename'] = "pods_report";
		$data['colhead']  = "808080";	
		$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1 : (date('d'));
		$Monthol = ( (date('d')+0) == 1 ) ? date('Y/m', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('Y/m'));
		//echo $Monthol; exit;
		$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );
	//var_dump($data['holiday']); exit;
		$this->load->view('Excel/from_report_prod_his',$data);			
	}

	public function QC_daily()
	{
		$data['list_act_report'] = array( 'defect_daily'   => $this->Backreport_model->Prod_report('Ng_daily_TEST'),
										  'code_detail' => $this->Backreport_model->Prod_report('QC_CODE'));

		$data['title'] = array('DEFECT DAILY', 'CODE DETAIL') ;//WASHING H/BEARING
		$data['filename'] = "gc_daily";
		$data['colhead']  = "5B90A4";	
		$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1 : (date('d'));
		$Monthol = ( (date('d')+0) == 1 ) ? date('Y/m', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('Y/m'));
		//echo $Monthol; exit;
		$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );
		//$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . date('Y/m') . "'" );
		$this->load->view('Excel/from_report_qc',$data);			
	}
	public function QC_sum()
	{
		$data['list_act_report'] = array( 'defect_summary' => $this->Backreport_model->Prod_report('QC_SUMMARY_v3'));										  

		$data['title'] = array('DEFECT SUMMARY') ;//WASHING H/BEARING
		$data['filename'] = "gc_daily";
		$data['colhead']  = "5B90A4";	
		$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . date('Y/m') . "'" );
		$this->load->view('Excel/from_report_ng_sum',$data);			
	}
	public function Test()
	{
		$data['list_act_report'] = $this->Backreport_model->Test();	

	}
	public function Receive_weekly()
	{
		$orderBy = "ORDER BY NO_RM,  PARENT_SEC_CD, SOURCE_CD, ITEM_CD ASC";
		$data['list_act_report'] = array('receiving_weekly' => $this->Backreport_model->Prod_report('RECEIVE_ACCUM'));
		$data['title'] = array("Receiving Weekly");
		$data['filename'] = "receive_weekly";
		$data['colhead']  = "800000";
		$dateCol = ( (date('d')+0) == 1 ) ? date('d', strtotime(date('Y')."-".(date('m')+0)."-".'0'))+1   : (date('d'));
		$Monthol = ( (date('d')+0) == 1 ) ? date('Y/m', strtotime(date('Y')."-".(date('m')+0)."-".'0'))   : (date('Y/m'));

		$data['holiday']  = $this->Backreport_model->Prod_report( 'date_ho', "WHERE  LEFT(d_t,7) = '" . $Monthol . "'" );
		$this->load->view('Excel/from_report_receive_accum',$data);		

	}

	// QCD OLD VERSION
	// public function QCD_PROD()
	// {
	// 	$data['list_act_report'] = array( 'qcd_production_report' => $this->Backreport_model->Prod_report('QCD_PROD_REPORT'));										  
	// 	$data['title'] = array('QCD Production Report') ;
	// 	$data['filename'] = "qcd_production_report";
	// 	$data['colhead']  = "5B90A4";	
	// 	$this->load->view('Excel/from_report_qcd_prod',$data);			
	// }

	public function QCD_PROD()
	{
		$data['list_act_report'] = array( 'qcd_production_report' => $this->Backreport_model->Prod_report('QCD_PROD_REPORT'),
										  'pd6_qcd_production_report' => $this->Backreport_model->Prod_report('QCD_PROD_REPORT_PD6'));										  
		$data['title'] = array("QCD Production Report" , "PD6 QCD Production Report") ;
		$data['filename'] = "qcd_production_report";
		$data['colhead']  = "5B90A4";	
		$this->load->view('Excel/from_report_qcd_prod',$data);			
	}

	public function QCD_PCL()
	{
		//$data['list_act_report'] = array('pcl_qcd_daily_report' => $this->Backreport_model->Prod_report('QCD_REPORT_PD6'));
		$data['list_act_report'] = array('pcl_qcd_daily_report' => $this->Backreport_model->Prod_report('QCD_REPORT'), 
										 'pd6_qcd_daily_report' => $this->Backreport_model->Prod_report('QCD_REPORT_PD6'));
		$data['title'] = array("PCL QCD Daily Report", "PD6 QCD Daily Report") ;
		$data['filename'] = "pcl_qcd_daily_report";
		$data['colhead']  = "808080";
		$data['colhead_font']  = "0000000";
		//$data['ig'] = $this->Backreport_model->mysql_item_group();	
		//var_dump($data['ig']); exit;
		$this->load->view('Excel/from_report_qcd',$data);			
	}

}
