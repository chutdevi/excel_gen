<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Bom extends CI_Controller {

    public function __construct()
    { 
        parent::__construct();



    }

    public function index()
    {


    }

    public function bm($part_nm)
    {

          $data_bom = array();
          $data_bom2 = array();
          $tmp_lvl = array();
          $lvl = array();


         // $part_nm = '898315-3732';                               //'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
          $lv = 1;
          $lvl_01 = $this->Backreport_model->list_bom(2, "WHERE PARENT_ITEM_CD = '".$part_nm."'");
          //var_dump($lvl_01);exit;
          array_push( $data_bom, $lvl_01 );

             // $lvl_01 = $this->Backreport_model->list_bom();

         echo $part_nm.' Master' . "<hr>"; 
        foreach ($lvl_01 as $ind => $value) 
          {

             $lvl = $this->Backreport_model->list_bom(3, "WHERE PARENT_ITEM_CD = '". $value['UNDERS'] ."'");
             echo "Level 2  ".$value['UNDERS'] ."<br>";
            //if( sizeof($lvl) > 0 ) array_push( $tmp_lvl, $lvl );

                 foreach ($lvl as $un => $v) 
                  {

                     $lo = $this->Backreport_model->list_bom(3, "WHERE PARENT_ITEM_CD = '". $v['UNDERS'] ."'");
                     //var_dump($lo); exit;

                     if( count($lo) > 0 ) 
                     {
                       echo "&emsp;&emsp;"."Level 3  ".$v['UNDERS']  ."<br>";
                         foreach ($lo as $u2 => $v2) 
                          {
                             $lo2 = $this->Backreport_model->list_bom(3, "WHERE PARENT_ITEM_CD = '". $v2['UNDERS'] ."'"); 
                         
                             if( count($lo2) > 0 ) 
                              {
                                  echo "&emsp;&emsp;&emsp;&emsp;"."Level 4  ".$v2['UNDERS']  ."<br>";                                  
                                   foreach ($lo2 as $u3 => $v3) 
                                    {
                                       $lo3 = $this->Backreport_model->list_bom(3, "WHERE PARENT_ITEM_CD = '". $v3['UNDERS'] ."'");                            
                                       if( count($lo3) > 0 ) 
                                        {
                                                var_dump($lo3); echo "<br>";

                                        }
                                       else echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 5  ".$v3['UNDERS']  ."<br>";
                                    }

                              }
                             else echo "&emsp;&emsp;&emsp;&emsp;"."Level 4  ".$v2['UNDERS']  ."<br>";
                          }
                      //var_dump($lo);
                     }
                     else echo "&emsp;&emsp;"."Level 3  ".$v['UNDERS']  ."<br>";
                                 

                  }             
              echo "<hr>";
          }
          exit; 
           var_dump($tmp_lvl); exit; 
          array_push( $data_bom, $tmp_lvl );
          
          $tmp_lvl = array();
        
        // foreach ($lvl_01 as $ind => $value) 
        //   {

        //      $lvl = $this->Backreport_model->list_bom(3, "WHERE PARENT_ITEM_CD = '". $value['UNDERS'] ."'");
        //      if( sizeof($lvl) > 0 ) array_push( $tmp_lvl, $lvl );
                         

        //   }

          var_dump($data_bom);
          exit;

    }
}

?>