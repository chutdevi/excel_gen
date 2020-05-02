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
          $lvl_01 = $this->Backreport_model->list_bom( "WHERE PARENT_ITEM_CD = '".$part_nm."'");
          //var_dump($lvl_01);exit;
          array_push( $data_bom, $lvl_01 );

             // $lvl_01 = $this->Backreport_model->list_bom();

         echo $part_nm.' Master' . "<hr>"; 
        foreach ($lvl_01 as $ind => $value) 
          {

             $lvl = $this->Backreport_model->list_bom( "WHERE PARENT_ITEM_CD = '". $value['UNDERS'] ."'");
             echo "Level 2  ".$value['UNDERS'] ."&emsp;&emsp;".number_format($value['UP'],2)."<br>";
            //if( sizeof($lvl) > 0 ) array_push( $tmp_lvl, $lvl );

                 foreach ($lvl as $un => $v) 
                  {

                     $lo = $this->Backreport_model->list_bom( "WHERE PARENT_ITEM_CD = '". $v['UNDERS'] ."'");
                     //var_dump($lo); exit;

                     if( count($lo) > 0 ) 
                     {
                       echo "&emsp;&emsp;"."Level 3  ".$v['UNDERS']."&emsp;&emsp;".number_format($v['UP'],2)  ."<br>";
                         foreach ($lo as $u2 => $v2) 
                          {
                             $lo2 = $this->Backreport_model->list_bom( "WHERE PARENT_ITEM_CD = '". $v2['UNDERS'] ."'"); 
                         
                             if( count($lo2) > 0 ) 
                              {
                                  echo "&emsp;&emsp;&emsp;&emsp;"."Level 4  ".$v2['UNDERS']."&emsp;&emsp;".number_format($v2['UP'],2)  ."<br>";                                  
                                   foreach ($lo2 as $u3 => $v3) 
                                    {
                                       $lo3 = $this->Backreport_model->list_bom( "WHERE PARENT_ITEM_CD = '". $v3['UNDERS'] ."'");                            
                                       if( count($lo3) > 0 ) 
                                        {
                                             echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 5  ".$v3['UNDERS']."&emsp;&emsp;".number_format($v3['UP'],2)  ."<br>";                                  
                                             foreach ($lo3 as $u4 => $v4) 
                                              {
                                                 $lo4 = $this->Backreport_model->list_bom( "WHERE PARENT_ITEM_CD = '". $v4['UNDERS'] ."'");                            
                                                  if( count($lo4) > 0 ) 
                                                    {
                                                        echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 6  ".$v4['UNDERS']."&emsp;&emsp;".number_format($v4['UP'],2)  ."<br>";                                  
                                                        foreach ($lo4 as $u5 => $v5) 
                                                         {
                                                            $lo5 = $this->Backreport_model->list_bom( "WHERE PARENT_ITEM_CD = '". $v5['UNDERS'] ."'");                            
                                                            if( count($lo5) > 0 ) 
                                                              {
                                                                var_dump($lo5); echo "<br>";

                                                              }
                                                            else echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 7  ".$v5['UNDERS']."&emsp;&emsp;".number_format($v5['UP'],2)  ."<br>";
                                                          }

                                                    }
                                                  else echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 6  ".$v4['UNDERS']."&emsp;&emsp;".number_format($v4['UP'],2)  ."<br>";
                                              }
                                        }
                                       else echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 5  ".$v3['UNDERS']."&emsp;&emsp;".number_format($v3['UP'],2)  ."<br>";
                                    }

                              }
                             else echo "&emsp;&emsp;&emsp;&emsp;"."Level 4  ".$v2['UNDERS']."&emsp;&emsp;".number_format($v2['UP'],2)  ."<br>";
                          }
                      //var_dump($lo);
                     }
                     else echo "&emsp;&emsp;"."Level 3  ".$v['UNDERS']."&emsp;&emsp;".number_format($v['UP'],2)  ."<br>";
                                 

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
    public function bm1($part_nm)
    {

          $data_bom = array();
          $data_bom2 = array();
          $tmp_lvl = array();
          $lvl = array();


         // $part_nm = '898315-3732';                               //'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
          $lv = 1;
          $lvl_01 = $this->Backreport_model->list_bom1( "PARENT_ITEM_CD = '".$part_nm."'");
          //var_dump($lvl_01);exit;
          array_push( $data_bom, $lvl_01 );

             // $lvl_01 = $this->Backreport_model->list_bom();

         echo $part_nm.' Master' . "<hr>"; 
        foreach ($lvl_01 as $ind => $value) 
          {

             $lvl = $this->Backreport_model->list_bom1( "PARENT_ITEM_CD = '". $value['UNDERS'] ."'");
             echo "Level 2  ".$value['UNDERS'] ."&emsp;&emsp;".number_format($value['UP'],2)."&emsp;&emsp;||||||||||".$value['ITEM_NAME']."&emsp;&emsp;||||||||||".$value['MODEL']."<br>";
            //if( sizeof($lvl) > 0 ) array_push( $tmp_lvl, $lvl );

                 foreach ($lvl as $un => $v) 
                  {

                     $lo = $this->Backreport_model->list_bom1( "PARENT_ITEM_CD = '". $v['UNDERS'] ."'");
                     //var_dump($lo); exit;

                     if( count($lo) > 0 ) 
                     {
                       echo "&emsp;&emsp;"."Level 3  ".$v['UNDERS']."&emsp;&emsp;".number_format($v['UP'],2)  ."&emsp;&emsp;||||||||||".$v['ITEM_NAME']."&emsp;&emsp;||||||||||".$v['MODEL']."<br>";
                       foreach ($lo as $u2 => $v2) 
                          {
                             $lo2 = $this->Backreport_model->list_bom1( " PARENT_ITEM_CD = '". $v2['UNDERS'] ."'"); 
                         
                             if( count($lo2) > 0 ) 
                              {
                                  echo "&emsp;&emsp;&emsp;&emsp;"."Level 4  ".$v2['UNDERS']."&emsp;&emsp;".number_format($v2['UP'],2). "&emsp;&emsp;||||||||||".$v2['ITEM_NAME']."&emsp;&emsp;||||||||||".$v2['MODEL']."<br>";                                 
                                   foreach ($lo2 as $u3 => $v3) 
                                    {
                                       $lo3 = $this->Backreport_model->list_bom1( "PARENT_ITEM_CD = '". $v3['UNDERS'] ."'");                            
                                       if( count($lo3) > 0 ) 
                                        {
                                             echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 5  ".$v3['UNDERS']."&emsp;&emsp;".number_format($v3['UP'],2) . "&emsp;&emsp;||||||||||".$v3['ITEM_NAME']."&emsp;&emsp;||||||||||".$v3['MODEL']."<br>";                                
                                             foreach ($lo3 as $u4 => $v4) 
                                              {
                                                 $lo4 = $this->Backreport_model->list_bom1( "PARENT_ITEM_CD = '". $v4['UNDERS'] ."'");                            
                                                  if( count($lo4) > 0 ) 
                                                    {
                                                        echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 6  ".$v4['UNDERS']."&emsp;&emsp;".number_format($v4['UP'],2).  "&emsp;&emsp;||||||||||".$v4['ITEM_NAME']."&emsp;&emsp;||||||||||".$v4['MODEL']."<br>";                                
                                                        foreach ($lo4 as $u5 => $v5) 
                                                         {
                                                            $lo5 = $this->Backreport_model->list_bom1( "PARENT_ITEM_CD = '". $v5['UNDERS'] ."'");                            
                                                            if( count($lo5) > 0 ) 
                                                              {
                                                                var_dump($lo5); echo "<br>";

                                                              }
                                                            else echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 7  ".$v5['UNDERS']."&emsp;&emsp;".number_format($v5['UP'],2).  "&emsp;&emsp;||||||||||".$va5['ITEM_NAME']."&emsp;&emsp;||||||||||".$v5['MODEL']."<br>";
                                                          }

                                                    }
                                                  else echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 6  ".$v4['UNDERS']."&emsp;&emsp;".number_format($v4['UP'],2).  "&emsp;&emsp;||||||||||".$v4['ITEM_NAME']."&emsp;&emsp;||||||||||".$v4['MODEL']."<br>";
                                              }
                                        }
                                       else echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 5  ".$v3['UNDERS']."&emsp;&emsp;".number_format($v3['UP'],2). "&emsp;&emsp;||||||||||".$v3['ITEM_NAME']."&emsp;&emsp;||||||||||".$v3['MODEL']."<br>";
                                    }

                              }
                             else echo "&emsp;&emsp;&emsp;&emsp;"."Level 4  ".$v2['UNDERS']."&emsp;&emsp;".number_format($v2['UP'],2) . "&emsp;&emsp;||||||||||".$v2['ITEM_NAME']."&emsp;&emsp;||||||||||".$v2['MODEL']."<br>";
                          }
                      //var_dump($lo);
                     }
                     else echo "&emsp;&emsp;"."Level 3  ".$v['UNDERS']."&emsp;&emsp;".number_format($v['UP'],2) . "&emsp;&emsp;||||||||||".$v['ITEM_NAME']."&emsp;&emsp;||||||||||".$v['MODEL']."<br>";
                                 

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

    public function re_bm($part_nm)
    {

          $data_bom = array();
          $data_bom2 = array();
          $tmp_lvl = array();
          $lvl = array();


         // $part_nm = '898315-3732';                               //'receive_history' => $this->Backreport_model->mysql_report_service('RECEIVE_MON_HIS'),
          $lv = 1;
          $lvl_01 = $this->Backreport_model->list_bom( "WHERE COMP_ITEM_CD = '".$part_nm."'");
          //var_dump($lvl_01);exit;
          array_push( $data_bom, $lvl_01 );

             // $lvl_01 = $this->Backreport_model->list_bom();

         echo $part_nm.' Master' . "<hr>"; 
        foreach ($lvl_01 as $ind => $value) 
          {

             $lvl = $this->Backreport_model->list_bom( "WHERE COMP_ITEM_CD = '". $value['HEAD'] ."'");
             echo "Level 2  ".$value['HEAD'] ."<br>";
            //if( sizeof($lvl) > 0 ) array_push( $tmp_lvl, $lvl );

                 foreach ($lvl as $un => $v) 
                  {

                     $lo = $this->Backreport_model->list_bom( "WHERE COMP_ITEM_CD = '". $v['HEAD'] ."'");
                     //var_dump($lo); exit;

                     if( count($lo) > 0 ) 
                     {
                       echo "&emsp;&emsp;"."Level 3  ".$v['HEAD']  ."<br>";
                         foreach ($lo as $u2 => $v2) 
                          {
                             $lo2 = $this->Backreport_model->list_bom( "WHERE COMP_ITEM_CD = '". $v2['HEAD'] ."'"); 
                         
                             if( count($lo2) > 0 ) 
                              {
                                  echo "&emsp;&emsp;&emsp;&emsp;"."Level 4  ".$v2['HEAD']  ."<br>";                                  
                                   foreach ($lo2 as $u3 => $v3) 
                                    {
                                       $lo3 = $this->Backreport_model->list_bom( "WHERE COMP_ITEM_CD = '". $v3['HEAD'] ."'");                            
                                       if( count($lo3) > 0 ) 
                                        {

                                           echo"&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 5  ".$v3['HEAD']  ."<br>";                                  
                                           foreach ($lo3 as $u4 => $v4) 
                                            {
                                               $lo4 = $this->Backreport_model->list_bom( "WHERE COMP_ITEM_CD = '". $v4['HEAD'] ."'");                            
                                               if( count($lo4) > 0 ) 
                                                {
                                                      echo"&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 6  ".$v4['HEAD']  ."<br>";                                  
                                                      foreach ($lo4 as $u5 => $v5) 
                                                        {
                                                         $lo5 = $this->Backreport_model->list_bom( "WHERE COMP_ITEM_CD = '". $v5['HEAD'] ."'");                            
                                                          if( count($lo5) > 0 ) 
                                                             {
                                                              var_dump($lo5); echo "<br>";

                                                             }
                                                          else echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 7  ".$v5['HEAD']  ."<br>";
                                                        }

                                                }
                                               else echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 6  ".$v4['HEAD']  ."<br>";
                                            }

                                        }
                                       else echo "&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;"."Level 5  ".$v3['HEAD']  ."<br>";
                                    }

                              }
                             else echo "&emsp;&emsp;&emsp;&emsp;"."Level 4  ".$v2['HEAD']  ."<br>";
                          }
                      //var_dump($lo);
                     }
                     else echo "&emsp;&emsp;"."Level 3  ".$v['HEAD']  ."<br>";
                                 

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