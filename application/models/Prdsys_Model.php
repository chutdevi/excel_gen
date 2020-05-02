<?php
class Prdsys_Model extends CI_Model 
  {
    //var $objFa;
    public function __construct()
    { 
      parent::__construct();
      //echo  dirname(__FILE__) . '/query/query-production-history-fa.php'; exit;
      //require_once dirname(__FILE__) . '/query/query-production-history-fa.php';

      //$this->objFa = new FAHISTORY_ACTUAL();
 
    }
    public  function model_datareport( )
    {
			$d = $this->getdate_req();
			$content = file_get_contents("http://192.168.161.102/api_system/api_prdsys/prdsys_prodrept/?d=".$d);
		 	$result  = json_decode($content);
      $recLoad =  json_decode(json_encode($result), true); 
      
      //$recLoad = $excEdt->result_array();      

      return $recLoad;
      //var_dump($recLoad); exit;
    }
    public  function model_holiday( )
    {
			$d = $this->getdate_req();
			$content = file_get_contents("http://192.168.161.102/api_system/api_prdsys/prdsys_prodholiday?d=".$d);
		 	$result  = json_decode($content);
      $recLoad =  json_decode(json_encode($result), true); 
      
      //$recLoad = $excEdt->result_array();      

      return $recLoad;
      //var_dump($recLoad); exit;
    }
    public  function model_saturdaty( )
    {
			$d = $this->getdate_req();
			$content = @file_get_contents("http://192.168.161.102/api_system/api_prdsys/prdsys_prodsaturdy?d=".$d);
		 	$result  = json_decode($content);
      $recLoad =  json_decode(json_encode($result), true); 
      
      //$recLoad = $excEdt->result_array();      

      return $recLoad;
      //var_dump($recLoad); exit; prdsys_insert_prd_list
    }    
    public  function model_insert_file_list( $file_dir )
    { 
      //echo  "http://192.168.161.102/api_system/api_prdsys/prdsys_insert_prd_list?f=". $file_dir; exit;
      $content = @file_get_contents("http://192.168.161.102/api_system/api_prdsys/prdsys_insert_prd_list?f=". $file_dir);
      
		 	$result  = json_decode($content);
      $recLoad =  json_decode(json_encode($result), true); 
      
      //$recLoad = $excEdt->result_array();      

      return $recLoad;
      //var_dump($recLoad); exit; prdsys_insert_prd_list
    }
    public  function model_create_file_seq( $file_dir )
    { 

      $content = @file_get_contents("http://192.168.161.102/api_system/api_prdsys/prdsys_getdata_prd_seq?f=". $file_dir);
      //echo  $content; exit;
		 	$result  = json_decode($content);
      $recLoad =  json_decode(json_encode($result), true); 
      
      //$recLoad = $excEdt->result_array();      

      return $recLoad;
      //var_dump($recLoad); exit; prdsys_insert_prd_list
    }



    public function getdate_req()
      {
        $d = date('Y-m-d');
        if( $this->input->get('d') ){
          if(  strtoupper( $this->input->get('d') ) == "LAST" ){
              $d = date('Y-m-t', strtotime("- 1 month", strtotime(date('Y-m-01', strtotime($d) ) ) ) );
          }else{
              $d = $this->input->get('d');
          } 
        } 
        return $d;
      }
    public function getfilename_req()
      {
        $d = "daily-production-report-";
        if( $this->input->get('d') ){
          if(  strtoupper( $this->input->get('d') ) == "LAST" ){
             return "lastday-production-report-";
          }else{
             return "daily-production-report-";
          } 
        } 
        return $d;
      }  
        
  }  
?>
