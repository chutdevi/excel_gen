<?php
class Backreport2_model extends CI_Model 
{
   // public   

     //public   $mn1= date('Y-m');


    public function __construct()
    {
        parent::__construct();
        date_default_timezone_set('Asia/Bangkok');
        ## asset config
        //session_destroy();
        ob_clean();
        flush();
       

       $this->EX = $this->load->database('dbj', true);
       $this->expk = $this->load->database('expk', true);

       // $month1 = date('t');
       // $month2 = date('t', strtotime("+ 1 month", strtotime(date('Y-m-d'))));
       // $month3 = date('t', strtotime("+ 2 month", strtotime(date('Y-m-d'))));
       
    }

    public function index() {     
         
    }

public function get_exec($str)
{
  //$sql = "SELECT CAL_DATE, HOLIDAY_FLG FROM M_CAL WHERE CAL_NO = 1 AND  CAL_DATE BETWEEN '$mst' AND '$men' ";

        $excEdt  = $this->expk->query($str);
        $recLoad = $excEdt->result_array();

        return $recLoad;
}
public function get_hol($mst, $men)
{
  $sql = "SELECT CAL_DATE, HOLIDAY_FLG FROM M_CAL WHERE CAL_NO = 1 AND  CAL_DATE BETWEEN '$mst' AND '$men' ";

        $excEdt  = $this->expk->query($sql);
        $recLoad = $excEdt->result_array();

        return $recLoad;
}
   public function demand_data( $where = '--' )
   {

   $qur = file_get_contents( dirname(__FILE__) . @"\query\demand.sql") or die("Unable to open file!");

   // echo $myfile;

        $where = "where FG.PARENT_SEC_CD = $where" ;

        $sqlEdt = "SELECT FG.* FROM ( $qur ) FG $where ";

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;

   }
   public function demand_data_1( $where = '--' )
   {

   //$this->Backreport2_model->$mn1 = date('Y-m');
   $qur = file_get_contents( dirname(__FILE__) . @"\query\demand_1.sql") or die("Unable to open file!");
       $month1 = date('t', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $month2 = date('t', strtotime("+ 1 month", strtotime(date('Y-m-01')))); //date('t', strtotime("+ 1 month", strtotime(date('Y-m-d'))));
       $month3 = date('t', strtotime("+ 2 month", strtotime(date('Y-m-01'))));


   $myfile =  $this->query_extend_dm( (int)$month1, (int)$month2, (int)$month3);
   $myfile .= $this->query_extend_pl( (int)$month1, (int)$month2, (int)$month3);
   $myfile .= $this->query_extend_ac( (int)$month1, (int)$month2, (int)$month3);
   $myfile .= $this->query_extend_df( (int)$month1, (int)$month2, (int)$month3);
   $myfile .= $this->query_extend_st( (int)$month1, (int)$month2, (int)$month3);
   $myfile .= $this->query_extend_sv( (int)$month1, (int)$month2, (int)$month3);
   $myfile .= $this->query_extend_db( (int)$month1, (int)$month2, (int)$month3);
   $myfile .= $this->query_extend_pb( (int)$month1, (int)$month2, (int)$month3);
   $myfile .= $this->query_extend_sb( (int)$month1, (int)$month2, (int)$month3);
   $myfile .= $this->query_extend_tb( (int)$month1, (int)$month2, (int)$month3);

        //  echo $month1 . " " . $month2 . " " . $month3; exit;
        //echo $myfile; exit;
        $where = "where FG.PARENT_SEC_CD = $where" ;

        $sqlEdt = "SELECT FG.* FROM ( $myfile ) FG $where ";

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;

   }

   public function query_extend_dm($m1, $m2, $m3  )
   {
      $m1_str = "\n\t\t";
      $m2_str = "\n\t\t";
      $m3_str = "\n\t\t";

       $mn1 = date('Y-m-', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $mn2 = date('Y-m-', strtotime("+ 1 month", strtotime(date('Y-m-01'))));
       $mn3 = date('Y-m-', strtotime("+ 2 month", strtotime(date('Y-m-01'))));

      $sql_se = 
      "SELECT            
        DC.PARENT_SEC_CD
        ,DC.SOURCE_CD
        ,DC.SOURCE_NAME
        ,DC.ITEM_CD
        ,DC.ITEM_NAME
        ,DC.MODEL
        ,DC.LCT
        ,DC.SNP
        ,1 DM_TYPE
        ,DL.DM_SM LM
      ";

      foreach (range(1, $m1) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.DM_$num"."_1 " . date('dS', strtotime( $mn1 . $num ) ) . "_d1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m2) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.DM_$num"."_2 " . date('dS', strtotime( $mn2 . $num ) ) . "_d2
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m3) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.DM_$num"."_3 " . date('dS', strtotime( $mn3 . $num ) ) . "_d3
        ";
      }
      $m1_str .= "\n\t\t";  

      $sql_fr = "
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        ";

      return $sql_se . $m1_str . $sql_fr;

   }
   public function query_extend_pl($m1, $m2, $m3  )
   {
      $m1_str = "\n\t\t";
      $m2_str = "\n\t\t";
      $m3_str = "\n\t\t";

       $mn1 = date('Y-m-', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $mn2 = date('Y-m-', strtotime("+ 1 month", strtotime(date('Y-m-01'))));
       $mn3 = date('Y-m-', strtotime("+ 2 month", strtotime(date('Y-m-01'))));      
      $sql_se = 
      "SELECT            
        DC.PARENT_SEC_CD
        ,DC.SOURCE_CD
        ,DC.SOURCE_NAME
        ,DC.ITEM_CD
        ,DC.ITEM_NAME
        ,DC.MODEL
        ,DC.LCT
        ,DC.SNP
        ,2 DM_TYPE
        ,DL.DM_SM LM
      ";

      foreach (range(1, $m1) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.PL_$num"."_1 " . date('dS', strtotime( $mn1 . $num ) ) . "_p1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m2) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.PL_$num"."_2 " . date('dS', strtotime( $mn3 . $num ) ) . "_p2
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m3) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.PL_$num"."_3 " . date('dS', strtotime( $mn2 . $num ) ) . "_p3
        ";
      }
      $m1_str .= "\n\t\t";  

      $sql_fr = "
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        ";
      $sql = $sql_se . $m1_str . $sql_fr;

      //echo $sql; exit;
      return $sql;

   }

   public function query_extend_ac($m1, $m2, $m3  )
   {
      $m1_str = "\n\t\t";
      $m2_str = "\n\t\t";
      $m3_str = "\n\t\t";

       $mn1 = date('Y-m-', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $mn2 = date('Y-m-', strtotime("+ 1 month", strtotime(date('Y-m-01'))));
       $mn3 = date('Y-m-', strtotime("+ 2 month", strtotime(date('Y-m-01'))));      
      $sql_se = 
      "SELECT            
        DC.PARENT_SEC_CD
        ,DC.SOURCE_CD
        ,DC.SOURCE_NAME
        ,DC.ITEM_CD
        ,DC.ITEM_NAME
        ,DC.MODEL
        ,DC.LCT
        ,DC.SNP
        ,3 DM_TYPE
        ,DL.DM_SM LM
      ";

      foreach (range(1, $m1) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.AC_$num"."_1 " . date('dS', strtotime( $mn1 . $num ) ) . "_a1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m2) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.AC_$num"."_2 " . date('dS', strtotime( $mn2 . $num ) ) . "_a1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m3) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.AC_$num"."_3 " . date('dS', strtotime( $mn3 . $num ) ) . "_a3
        ";
      }
      $m1_str .= "\n\t\t";  

      $sql_fr = "
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        ";
      $sql = $sql_se . $m1_str . $sql_fr;

      //echo $sql; exit;
      return $sql;

   }

   public function query_extend_df($m1, $m2, $m3  )
   {
      $m1_str = "\n\t\t";
      $m2_str = "\n\t\t";
      $m3_str = "\n\t\t";

       $mn1 = date('Y-m-', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $mn2 = date('Y-m-', strtotime("+ 1 month", strtotime(date('Y-m-01'))));
       $mn3 = date('Y-m-', strtotime("+ 2 month", strtotime(date('Y-m-01'))));      
      $sql_se = 
      "SELECT            
        DC.PARENT_SEC_CD
        ,DC.SOURCE_CD
        ,DC.SOURCE_NAME
        ,DC.ITEM_CD
        ,DC.ITEM_NAME
        ,DC.MODEL
        ,DC.LCT
        ,DC.SNP
        ,4 DM_TYPE
        ,DL.AC_SM - DL.PL_SM  LM
      ";

      foreach (range(1, $m1) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.AC_$num"."_1 - DC.PL_$num"."_1 " . date('dS', strtotime( $mn1 . $num ) ) . "_f1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m2) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.AC_$num"."_2 - DC.PL_$num"."_2 " . date('dS', strtotime( $mn2 . $num ) ) . "_f2
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m3) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",DC.AC_$num"."_3 - DC.PL_$num"."_3 " . date('dS', strtotime( $mn3 . $num ) ) . "_f3
        ";
      }
      $m1_str .= "\n\t\t";  

      $sql_fr = "
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        ";
      $sql = $sql_se . $m1_str . $sql_fr;

      //echo $sql; exit;
      return $sql;
   }

   public function query_extend_st($m1, $m2, $m3  )
   {
      $m1_str = "\n\t\t";
      $m2_str = "\n\t\t";
      $m3_str = "\n\t\t";

       $mn1 = date('Y-m-', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $mn2 = date('Y-m-', strtotime("+ 1 month", strtotime(date('Y-m-01'))));
       $mn3 = date('Y-m-', strtotime("+ 2 month", strtotime(date('Y-m-01'))));      
      $sql_se = 
      "SELECT            
        DC.PARENT_SEC_CD
        ,DC.SOURCE_CD
        ,DC.SOURCE_NAME
        ,DC.ITEM_CD
        ,DC.ITEM_NAME
        ,DC.MODEL
        ,DC.LCT
        ,DC.SNP
        ,5 DM_TYPE
        ,DS.LAST_MONTH LM
        ,0  01st
      ";

      foreach (range(2, $m1) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '$num' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END " . date('dS', strtotime( $mn1 . $num ) ) . "_s1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m2) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn2 . $num ) ) . "_s2
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m3) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn3 . $num ) ) . "_s3
        ";
      }
      $m1_str .= "\n\t\t";  

      $sql_fr = "
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN 
        ( SELECT * FROM DEMAND_STOCK  WHERE WH_CD IN ('K1MX', 'K2MX') ) DS
        ON DC.ITEM_CD = DS.ITEM_CD AND DC.PLANT_CD = DS.PLANT_CD

        UNION ALL

        ";
      $sql = $sql_se . $m1_str . $sql_fr;

      //echo $sql; exit;
      return $sql;
   }

   public function query_extend_sv($m1, $m2, $m3  )
   {
      $m1_str = "\n\t\t";
      $m2_str = "\n\t\t";
      $m3_str = "\n\t\t";

       $mn1 = date('Y-m-', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $mn2 = date('Y-m-', strtotime("+ 1 month", strtotime(date('Y-m-01'))));
       $mn3 = date('Y-m-', strtotime("+ 2 month", strtotime(date('Y-m-01'))));      
      $sql_se = 
      "SELECT            
        DC.PARENT_SEC_CD
        ,DC.SOURCE_CD
        ,DC.SOURCE_NAME
        ,DC.ITEM_CD
        ,DC.ITEM_NAME
        ,DC.MODEL
        ,DC.LCT
        ,DC.SNP
        ,6 DM_TYPE
        ,NULL  LM
      ";

      foreach (range(1, $m1) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn1 . $num ) ) . "_v1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m2) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn2 . $num ) ) . "_v2
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m3) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn3 . $num ) ) . "_v3
        ";
      }
      $m1_str .= "\n\t\t";  

      $sql_fr = "
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        ";
      $sql = $sql_se . $m1_str . $sql_fr;

      //echo $sql; exit;
      return $sql;
   }
   public function query_extend_db($m1, $m2, $m3  )
   {
      $m1_str = "\n\t\t";
      $m2_str = "\n\t\t";
      $m3_str = "\n\t\t";

       $mn1 = date('Y-m-', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $mn2 = date('Y-m-', strtotime("+ 1 month", strtotime(date('Y-m-01'))));
       $mn3 = date('Y-m-', strtotime("+ 2 month", strtotime(date('Y-m-01'))));      
      $sql_se = 
      "SELECT            
        DC.PARENT_SEC_CD
        ,DC.SOURCE_CD
        ,DC.SOURCE_NAME
        ,DC.ITEM_CD
        ,DC.ITEM_NAME
        ,DC.MODEL
        ,DC.LCT
        ,DC.SNP
        ,7 DM_TYPE
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DL.DM_SM / DC.SNP) END  LM
      ";

      foreach (range(1, $m1) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_$num"."_1 / DC.SNP) END " . date('dS', strtotime( $mn1 . $num ) ) . "_b1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m2) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_$num"."_2 / DC.SNP) END " . date('dS', strtotime( $mn2 . $num ) ) . "_b2
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m3) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_$num"."_3 / DC.SNP) END " . date('dS', strtotime( $mn3 . $num ) ) . "_b3
        ";
      }
      $m1_str .= "\n\t\t";  

      $sql_fr = "
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        ";
      $sql = $sql_se . $m1_str . $sql_fr;

      //echo $sql; exit;
      return $sql;
   }

   public function query_extend_pb($m1, $m2, $m3  )
   {
      $m1_str = "\n\t\t";
      $m2_str = "\n\t\t";
      $m3_str = "\n\t\t";

       $mn1 = date('Y-m-', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $mn2 = date('Y-m-', strtotime("+ 1 month", strtotime(date('Y-m-01'))));
       $mn3 = date('Y-m-', strtotime("+ 2 month", strtotime(date('Y-m-01'))));      
      $sql_se = 
      "SELECT            
        DC.PARENT_SEC_CD
        ,DC.SOURCE_CD
        ,DC.SOURCE_NAME
        ,DC.ITEM_CD
        ,DC.ITEM_NAME
        ,DC.MODEL
        ,DC.LCT
        ,DC.SNP
        ,8 DM_TYPE
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DL.PL_SM / DC.SNP) END  LM
      ";

      foreach (range(1, $m1) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_$num"."_1 / DC.SNP) END " . date('dS', strtotime( $mn1 . $num ) ) . "_c1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m2) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_$num"."_2 / DC.SNP) END " . date('dS', strtotime( $mn2 . $num ) ) . "_c2
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m3) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_$num"."_3 / DC.SNP) END " . date('dS', strtotime( $mn3 . $num ) ) . "_c3
        ";
      }
      $m1_str .= "\n\t\t";  

      $sql_fr = "
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        ";
      $sql = $sql_se . $m1_str . $sql_fr;

      //echo $sql; exit;
      return $sql;
   }
   public function query_extend_sb($m1, $m2, $m3  )
   {
      $m1_str = "\n\t\t";
      $m2_str = "\n\t\t";
      $m3_str = "\n\t\t";

       $mn1 = date('Y-m-', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $mn2 = date('Y-m-', strtotime("+ 1 month", strtotime(date('Y-m-01'))));
       $mn3 = date('Y-m-', strtotime("+ 2 month", strtotime(date('Y-m-01'))));      
      $sql_se = 
      "SELECT            
        DC.PARENT_SEC_CD
        ,DC.SOURCE_CD
        ,DC.SOURCE_NAME
        ,DC.ITEM_CD
        ,DC.ITEM_NAME
        ,DC.MODEL
        ,DC.LCT
        ,DC.SNP
        ,9 DM_TYPE
        ,NULL  LM
      ";

      foreach (range(1, $m1) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn1 . $num ) ) . "_e1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m2) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn2 . $num ) ) . "_e2
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m3) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn3 . $num ) ) . "_e3
        ";
      }
      $m1_str .= "\n\t\t";  

      $sql_fr = "
        FROM 

        TEMP_DEMAND_CONVERT DC

        UNION ALL

        ";
      $sql = $sql_se . $m1_str . $sql_fr;

      //echo $sql; exit;
      return $sql;
   } 
   public function query_extend_tb($m1, $m2, $m3  )
   {
      $m1_str = "\n\t\t";
      $m2_str = "\n\t\t";
      $m3_str = "\n\t\t";

       $mn1 = date('Y-m-', strtotime("+ 0 month", strtotime(date('Y-m-01'))));
       $mn2 = date('Y-m-', strtotime("+ 1 month", strtotime(date('Y-m-01'))));
       $mn3 = date('Y-m-', strtotime("+ 2 month", strtotime(date('Y-m-01'))));      
      $sql_se = 
      "SELECT            
        DC.PARENT_SEC_CD
        ,DC.SOURCE_CD
        ,DC.SOURCE_NAME
        ,DC.ITEM_CD
        ,DC.ITEM_NAME
        ,DC.MODEL
        ,DC.LCT
        ,DC.SNP
        ,10 DM_TYPE
        ,NULL  LM
      ";

      foreach (range(1, $m1) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn1 . $num ) ) . "_t1
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m2) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn2 . $num ) ) . "_t2
        ";
      }
      $m1_str .= "\n\t\t";
      foreach (range(1, $m3) as $dt) 
      {
        $num = str_pad($dt,2,"0",STR_PAD_LEFT);
        $m1_str .= ",0 " . date('dS', strtotime( $mn3 . $num ) ) . "_t3
        ";
      }
      $m1_str .= "\n\t\t";  

      $sql_fr = "
        FROM 

        TEMP_DEMAND_CONVERT DC

        ORDER BY 1,2,4,8

        ";
      $sql = $sql_se . $m1_str . $sql_fr;

      //echo $sql; exit;
      return $sql;
   }     
    }  
?>
