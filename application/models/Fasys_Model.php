<?php
class Fasys_Model extends CI_Model 
  {
    var $objFa;
    public function __construct()
    { 
      parent::__construct();
      //echo  dirname(__FILE__) . '/query/query-production-history-fa.php'; exit;
      require_once dirname(__FILE__) . '/query/query-production-history-fa.php';

      $this->objFa = new FAHISTORY_ACTUAL();
 
    }
    public  function model_fahistory($f=10)
    {
      $this->fa = ($f == 10 ) ? $this->load->database('fa', true) : $this->load->database('f8', true) ;
   
      $excEdt  = $this->fa->query( $this->objFa->DB2GET_FAHISTORY() );
      $recLoad = $excEdt->result_array();      

      return $recLoad;
      //var_dump($recLoad); exit;
    }

  
        
  }  
?>
