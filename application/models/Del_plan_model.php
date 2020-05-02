<?php
class Del_plan_model extends CI_Model 
{
    //var $objFa;
    public function __construct()
    { 
      parent::__construct();
      //echo  dirname(__FILE__) . '/query/query-production-history-fa.php'; exit;
      //require_once dirname(__FILE__) . '/query/query-production-history-fa.php';
      //$this->objFa = new FAHISTORY_ACTUAL();
    }

    public  function del_plan_report()
    {
      // $d = $this->getdate_req();
      $get_year = $this->input->get('year');
      $get_month = $this->input->get('month');
			$content = file_get_contents(sprintf("http://192.168.161.102/api_system/api_del_plan/sum_del_data?&year=%s&month=%s",$get_year,$get_month));
		 	$result  = json_decode($content);
      $recLoad =  json_decode(json_encode($result), true); 
      //$recLoad = $excEdt->result_array();      

      return $recLoad;
      //var_dump($recLoad); exit;
    }
  }  
?>
