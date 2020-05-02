<?php
class Backreport_model extends CI_Model 
{

    public function sql_sv() 
    {
        //$this->EJ = $this->load->database('EJ', true);
        $sqlEdt = "SELECT * FROM qcd_stock;";
        $excEdt = $this->db->query($sqlEdt);
        $recLoad = $excEdt->result_array();


       // var_dump($recLoad); exit;
        return $recLoad;
    }  

    public function mysql_report_service($table='') 
    {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = "SELECT * FROM $table;";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        //var_dump($recLoad); exit;
        return $recLoad;
    } 

     public function mysql_report_service_sup($table='') 
    {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = "                 SELECT 
                              SU.PD
                              ,SU.LINE_CD   
                              ,VM.sec_nm AS 'LINE_NAME'                        
                              ,SU.ITEM_CD
                              ,SU.ITEM_NAME
                              ,SY.MODEL
                              ,SU.PLAN_QTY
                              ,SY.STOCK_ON_HAND_QTY  AS 'STOCK_TODAY'
                              ,CASE WHEN  ISNULL(RE.STC_LVL) THEN ROUND(SY.STOCK_ON_HAND_QTY/SU.PLAN_QTY,2) ELSE RE.STC_LVL  END  AS 'LEVEL_Str'
                              ,CASE WHEN NA.VEND_NAME IS NULL THEN SU.SUP_FROM  ELSE NA.VEND_NAME END AS 'SUP_FROM'
                              ,SU.LOCATION
                              ,SU.WI_NO
                              ,SY.PRODUCT_TYP AS 'PRODUCT_TYP'
                              ,ROUND((CY.CYCLE_TIME),2) AS 'CYCLE_TIME(Min)'
                              ,ROUND(((SU.PLAN_QTY * CY.CYCLE_TIME)/60),2) AS 'PRODUC_TIME(Hour)'
                              ,SN.PKG_UNIT_QTY AS 'PKG'
                              ,FLOOR(SU.PLAN_QTY / SN.PKG_UNIT_QTY) AS 'BOX'
                              ,SU.PLAN_QTY - ((FLOOR(SU.PLAN_QTY / SN.PKG_UNIT_QTY)* SN.PKG_UNIT_QTY)) as 'Remain'
                              ,NULL AS 'REMARK'

            FROM $table SU 
                        LEFT OUTER JOIN STOCK_FOR_SUPPLY SY 
                              ON SU.ITEM_CD = SY.ITEM_CD 
                        LEFT OUTER JOIN SUP_NAME NA 
                              ON NA.VEND_CD = SU.SUP_FROM 
                        LEFT OUTER JOIN CYCLE_TIME CY 
                              ON SU.ITEM_CD = CY.ITEM_CD AND SU.LINE_CD = CY.SOURCE_CODE 
                        LEFT OUTER JOIN RECEIVE_STOCK_LVL RE 
                              ON RE.ITEM_CD = SU.ITEM_CD 
                        LEFT OUTER JOIN SNP_CHECK SN 
                              ON SU.ITEM_CD = SN.ITEM_CD 
                        LEFT OUTER JOIN Production_report_TEST VM 
                              ON SU.LINE_CD = VM.line_cd
            ORDER BY 3, 12, 10 ASC
;";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();
        //var_dump($recLoad); exit;
        return $recLoad;
    } 
     public function mysql_product_cost($table='') 
    {

        $this->oee = $this->load->database('oee', true);
        $sqlEdt = "SELECT
                     OM.PD
                    ,OM.LINE_CD
                    ,OM.ITEM_CD
                    ,OM.ITEM_NAME
                    ,OM.MODEL
                    ,CY.CYCLE_TIME AS 'STD_CT'
                    ,OM.MAN AS 'MAN'
                    ,SUM(OM.JITU_SU) AS 'TOTAL_QTY'
                    ,((SUM(OM.TOTAL_TIME)) - SUM(OM.TOTAL_BREAK)) * OM.MAN AS 'TOTAL_TIME_ALLMAN'
                    ,SUM(OM.LOSS) AS 'LOSS_TIME'
                    ,(((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) * OM.MAN ) AS 'WORKING_TIME'
                    ,((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) AS 'PRODUCTION_TIME'
                    ,ROUND(((((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) * OM.MAN ) / SUM(OM.JITU_SU)),2) AS 'WORKING_TIME_PCS'
                    ,ROUND(((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) / SUM(OM.JITU_SU),2) AS 'PRODUCTION_TIME_PCS'
                    
                  FROM
                    OEE_WORK_MONTH OM
                  LEFT OUTER JOIN
                    OEE_CYCLE_TIME CY
                  ON CY.ITEM_CD = OM.ITEM_CD  AND  CY.SOURCE_CODE = OM.LINE_CD 

                  WHERE
                    OM.PD IN ('PD01','PD02','PD03','PD04','PD05','PD06')

                  GROUP BY
                     OM.PD
                    ,OM.LINE_CD
                    ,OM.ITEM_CD
                    ,OM.ITEM_NAME
                    ,OM.MODEL
                  ORDER BY 
                         OM.PD , OM.LINE_CD ASC
                    

;";
        $excEdt = $this->oee->query($sqlEdt);
        $recLoad = $excEdt->result_array();
     //   var_dump($recLoad); exit;
     return $recLoad;
    }

    public function mysql_fa_cost($table='') 
    {

        $this->oee = $this->load->database('oee', true);
        $sqlEdt = " SELECT
                             OM.PD
                            ,OM.LINE_CD
                            ,OM.ITEM_CD
                            ,OM.ITEM_NAME
                            ,OM.MODEL
                            ,CY.CYCLE_TIME AS 'STD_CT'
                            ,SI.SUM_MAN AS 'TOTAL_MAN'
                            ,MI.WORK_DAY 
                            ,SI.COUNT_SHIFT as 'COUNT_SHIFT'
                            ,ROUND(((SI.SUM_MAN )/(SI.COUNT_SHIFT)),2) AS 'AVER_MAN'
                            ,IFNULL(SH.COUNT_SHIFT,0) as 'SHIFT_B'
                            ,(SUM(OM.JITU_SU) - IFNULL((OA.ACT_OT),0))  AS 'NORMAL_QTY'
                            ,IFNULL((ACT_OT),0) AS 'OT_QTY'
                            ,SUM(OM.JITU_SU) AS 'TOTAL_QTY'
                          # ,(OW.TOTAL_TIME_ALL_MAN) AS 'TOTAL_TIME'
                            ,((OW.TOTAL_TIME_ALL_MAN) - (OW.BK_MAN)) AS 'TOTAL_TIME_ALL_MAN'
                            ,(OW.LOSS) AS 'TOTAL_LOSS_ALL_MAN'
                            ,(((OW.TOTAL_TIME_ALL_MAN) - (OW.BK_MAN)) - (OW.LOSS)) AS 'WORKING_TIME'
                            ,OW.PRODUC_TEST AS 'PRODUCTION_TIME'
                            ,ROUND((((OW.TOTAL_TIME_ALL_MAN) - (OW.BK_MAN)) - (OW.LOSS)) / SUM(OM.JITU_SU),2) AS 'WORKING_TIME_PCS'
                            ,ROUND((OW.PRODUC_TEST / SUM(OM.JITU_SU)),2) AS 'PRODUCTION_TIME_PCS'
                            ,TY.PRODUCT_TYP AS 'PRODUCT_TYP'
                            
                          FROM
                            OEE_WORK_MONTH OM
                          LEFT OUTER JOIN
                            OEE_CYCLE_TIME CY
                          ON CY.ITEM_CD = OM.ITEM_CD  AND  CY.SOURCE_CODE = OM.LINE_CD 
                           LEFT OUTER JOIN 
                                            (SELECT
                                                PD,
                                                LINE_CD,
                                                ITEM_CD,
                                                SUM(MAN) AS 'SUM_MAN',
                                                COUNT(ITEM_CD) AS 'COUNT_SHIFT'
                                            FROM
                                                OEE_SHIFT_DAY
                                            WHERE 
                                                SHIFT IN ('B','Q')
                                            GROUP BY
                                                ITEM_CD,
                                                LINE_CD
                                              ORDER BY
                                                LINE_CD,
                                                ITEM_CD ASC
                            
                                            ) SH 
                                            ON OM.LINE_CD = SH.LINE_CD AND OM.ITEM_CD = SH.ITEM_CD 
                           LEFT OUTER JOIN 
                                            (SELECT
                                                PD,
                                                LINE_CD,
                                                ITEM_CD,
                                                SUM(MAN) AS 'SUM_MAN',
                                                COUNT(ITEM_CD) AS 'COUNT_SHIFT'
                                            FROM
                                                OEE_SHIFT_DAY
                                            GROUP BY
                                                ITEM_CD,
                                                LINE_CD
                                              ORDER BY
                                                LINE_CD,
                                                ITEM_CD ASC
                            
                                            ) SI
                                            ON OM.LINE_CD = SI.LINE_CD AND OM.ITEM_CD = SI.ITEM_CD
                          LEFT OUTER JOIN
                                            (
                                            SELECT 
                                                    ITEM_CD, 
                                                    LINE_CD, 
                                                    SUM(ACT_NORMAL) ACT_NORMAL, 
                                                    SUM(ACT_OT) ACT_OT 
                                                    FROM OEE_ACT_OT 
                                                    GROUP BY 
                                                    ITEM_CD, LINE_CD ) OA
                                            ON OM.ITEM_CD = OA.ITEM_CD AND OM.LINE_CD = OA.LINE_CD
                                      LEFT OUTER JOIN
                                           (SELECT 
                                                                  PD
                                                                ,LINE_CD
                                                                ,ITEM_CD
                                                                ,WORK_DAY     
                                                          FROM OEE_WORK_DAY

                                                          GROUP BY
                                                                PD
                                                                ,LINE_CD
                                                                ,ITEM_CD
                                                                ,WORK_DAY 
                                            )MI
                                            ON OM.LINE_CD = MI.LINE_CD AND OM.ITEM_CD = MI.ITEM_CD 
                              LEFT OUTER JOIN 
                              (SELECT  
                                                    ITEM_CD, 
                                                    LINE_CD, 
                                                    SUM(MAN), 
                                                    SUM(PLAN_SU) AS 'PLAN',  
                                                    SUM(JITU_SU) , 
                                                    SUM(TOTAL_TIME) AS 'TOTAL_TIME', 
                                                    SUM(TOTAL_TIME * MAN) AS 'TOTAL_TIME_ALL_MAN',
                                                    SUM(LOSS) * SUM(MAN) AS 'LOSS_ALL_MAN',
                                                    SUM(TOTAL_BREAK) AS 'TOTAL_BREAK',
                                                    SUM(TOTAL_BREAK * MAN) AS 'BK_MAN',
                                                    SUM(OT_TIME) AS 'OT_TIME' ,
                                                    SUM(LOSS * MAN) AS 'LOSS',
                                                    SUM(TOTAL_TIME) - SUM(TOTAL_BREAK) - SUM(LOSS) AS 'PRODUC_TEST'
                                                    
                                                    FROM OEE_WORK_MONTH 

                              GROUP BY LINE_CD , ITEM_CD  ) OW
                  
                              ON OM.ITEM_CD = OW.ITEM_CD AND OM.LINE_CD = OW.LINE_CD
                              LEFT OUTER JOIN PRODUCT_TYPE TY
                              ON OM.ITEM_CD = TY.ITEM_CD 

                        WHERE
                          OM.PD IN ('PD01','PD02','PD03','PD04','PD05','PD06')
                        #AND  OM.ITEM_CD = '1300A033'
                        GROUP BY
                           OM.PD
                          ,OM.LINE_CD
                          ,OM.ITEM_CD
                          -- ,OM.ITEM_NAME
                          -- ,OM.MODEL
                        ORDER BY 
                               OM.PD , OM.LINE_CD ,OM.ITEM_CD ASC
  
;";
        $excEdt = $this->oee->query($sqlEdt);
        $recLoad = $excEdt->result_array();
     //   var_dump($recLoad); exit;
     return $recLoad;
    }

  public function mysql_fa_monthly($table='') 
    {

        $this->oee = $this->load->database('oee', true);
        $sqlEdt = " SELECT
                   OM.PD
                  ,OM.LINE_CD
                  ,OM.ITEM_CD
                  ,OM.ITEM_NAME
                  ,OM.MODEL
                  ,CY.CYCLE_TIME STD_CT
                  ,ROUND((((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) / SUM(OM.JITU_SU)),2) ACTUAL_CT
                  ,MI.WORK_DAY                  
                  ,IFNULL(IT.CAVITY,0) CAVITY
              #   ,SUM(OM.MAN)  TOTAL_MAN
              #   ,OM.MAN AVER_MAN

                  ,SUM(OM.JITU_SU) AS 'TOTAL_QTY'
                  ,SUM(OM.JITU_SU) - IFNULL((OA.ACT_OT),0) AS 'NORMAL_QTY'
                  ,IFNULL((OA.ACT_OT),0) AS 'OT_QTY'



                  ,(SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) TOTAL_TIME
                  ,((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) TOTAL_WORKING_TIME

                  ,ROUND((((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK))) * ROUND(((SH.SUM_MAN )/(SH.COUNT_SHIFT)),2)),2) TOTAL_TIME_ALLMAN
              #   ,(((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) * OM.MAN ) TOTAL_WORKING_TIME_ALLMAN
                  ,IFNULL(SUM(OM.LOSS),0) LOSS_TIME
                  ,ROUND(((((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK))) * ROUND(((SH.SUM_MAN )/(SH.COUNT_SHIFT)),2) ) / SUM(OM.JITU_SU)),2) MANHOUR_PCS
    
                  ,ROUND((((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS))/ SUM(OM.JITU_SU)),2) WORKING_TIME_PCS  
                  ,ROUND((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) / SUM(OM.JITU_SU),2) PRODUCTION_TIME_PCS
               #   ,ROUND((((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) / SUM(OM.JITU_SU)),2) ACTUAL_CT
                  ,CASE WHEN OM.PD = 'PD04' THEN IFNULL(ROUND(((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) / (SUM(OM.JITU_SU) / IFNULL(IT.CAVITY,0)),2),0) 
                   ELSE IFNULL(CY.CYCLE_TIME /(IFNULL(ROUND(((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) / (SUM(OM.JITU_SU)),2),0)),0) END EFF
                   
                  ,SH.SUM_MAN AS 'TOTAL_MAN'
                  ,SH.COUNT_SHIFT as 'SHIFT_BY_ITEM'
                  ,ROUND(((SH.SUM_MAN )/(SH.COUNT_SHIFT)),2) AS 'MAN'

                  ,IFNULL(CF.A ,0) AS 'A '
                  ,IFNULL(CF.B ,0) AS 'B '
                  ,IFNULL(CF.C ,0) AS 'C '
                  ,IFNULL(CF.C1,0) AS 'C1'
                  ,IFNULL(CF.D ,0) AS 'D '
                  ,IFNULL(CF.E ,0) AS 'E '
                  ,IFNULL(CF.F ,0) AS 'F '
                  ,IFNULL(CF.F1,0) AS 'F1'
                  ,IFNULL(CF.F2,0) AS 'F2'
                  ,IFNULL(CF.G ,0) AS 'G '
                  ,IFNULL(CF.G1,0) AS 'G1'
                  ,IFNULL(CF.H ,0) AS 'H '
                  ,IFNULL(CF.I ,0) AS 'I '
                  ,IFNULL(CF.J ,0) AS 'J '
                  ,IFNULL(CF.K ,0) AS 'K '
                  ,IFNULL(CF.K1,0) AS 'K1'
                  ,IFNULL(CF.K2,0) AS 'K2'
                  ,IFNULL(CF.K3,0) AS 'K3'
                  ,IFNULL(CF.L ,0) AS 'L '
                  ,IFNULL(CF.L1,0) AS 'L1'
                  ,IFNULL(CF.M ,0) AS 'M '
                  ,IFNULL(CF.N ,0) AS 'N '
                  ,IFNULL(CF.O ,0) AS 'O '
                  ,IFNULL(CF.P ,0) AS 'P '
                  ,IFNULL(CF.Q ,0) AS 'Q '
                  ,IFNULL(CF.Q1,0) AS 'Q1'
                  ,IFNULL(CF.R ,0) AS 'R '
                  ,IFNULL(CF.S ,0) AS 'S '
                  ,IFNULL(CF.T ,0) AS 'T '
                  ,IFNULL(CF.U ,0) AS 'U '
                  ,IFNULL(CF.V ,0) AS 'V '
                  ,IFNULL(CF.W ,0) AS 'W '
                                
            FROM
              OEE_WORK_MONTH OM
            LEFT OUTER JOIN
              OEE_CYCLE_TIME CY
            ON CY.ITEM_CD = OM.ITEM_CD  AND  CY.SOURCE_CODE = OM.LINE_CD 
            LEFT OUTER JOIN
              OEE_ITEM_CAVITY IT
            ON IT.ITEM_CD = OM.ITEM_CD
            LEFT OUTER JOIN
                 (SELECT 
                                        PD
                                      ,LINE_CD
                                      ,ITEM_CD
                                      ,WORK_DAY     
                                FROM OEE_WORK_DAY

                                GROUP BY
                                      PD
                                      ,LINE_CD
                                      ,ITEM_CD
                                      ,WORK_DAY 
                  )MI
                  ON OM.LINE_CD = MI.LINE_CD AND OM.ITEM_CD = MI.ITEM_CD 

                  LEFT OUTER JOIN OEE_CODE_ACCUM CF 

                  ON OM.PD = CF.PD AND OM.LINE_CD = CF.LINE_CD AND OM.ITEM_CD = CF.HINBAN 
                  LEFT OUTER JOIN 
                  (SELECT
                      PD,
                      LINE_CD,
                      ITEM_CD,
                      SUM(MAN) AS 'SUM_MAN',
                      COUNT(ITEM_CD) AS 'COUNT_SHIFT'
                  FROM
                      OEE_SHIFT_DAY
                  GROUP BY
                      ITEM_CD,
                      LINE_CD
                    ORDER BY
                      LINE_CD,
                      ITEM_CD ASC
  
                  ) SH 
                 ON OM.LINE_CD = SH.LINE_CD AND OM.ITEM_CD = SH.ITEM_CD  
 LEFT OUTER JOIN
                  (
                  SELECT 
                          ITEM_CD, 
                          LINE_CD, 
                          SUM(ACT_NORMAL) ACT_NORMAL, 
                          SUM(ACT_OT) ACT_OT 
                          FROM OEE_ACT_OT 
                          GROUP BY 
                          ITEM_CD, LINE_CD ) OA
                  ON OM.ITEM_CD = OA.ITEM_CD AND OM.LINE_CD = OA.LINE_CD



           
 WHERE
 #  OM.PD = 'PD01'
    OM.PD IN ('PCL1','PD01','PD02','PD03','PD04','PD05','PD06')
#AND OM.LINE_CD = 'K1M153'
GROUP BY
                    OM.PD
                  ,OM.LINE_CD
                  ,OM.ITEM_CD
                  -- ,OM.ITEM_NAME
                  -- ,OM.MODEL

ORDER BY OM.PD ,OM.LINE_CD ASC;
       
       

;";
        $excEdt = $this->oee->query($sqlEdt);
        $recLoad = $excEdt->result_array();
    //   var_dump($recLoad); exit;
     return $recLoad;
    }

  public function mysql_fa_weekly($table='') 
    {

        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = " SELECT
                   OM.PD
                  ,OM.LINE_CD
                  ,OM.ITEM_CD
                  ,OM.ITEM_NAME
                  ,OM.MODEL
                  ,CY.CYCLE_TIME STD_CT
                  ,ROUND((((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) / SUM(OM.JITU_SU)),2) ACTUAL_CT
                  ,MI.WORK_DAY
                  ,IFNULL(IT.CAVITY,0) CAVITY
        
                  ,SUM(OM.JITU_SU) TOTAL_QTY
                  ,SUM(OM.JITU_SU) - IFNULL((OA.ACT_OT),0) AS 'NORMAL_QTY'
                  ,IFNULL((OA.ACT_OT),0) AS 'OT_QTY'

                  ,(SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) TOTAL_TIME
                  ,((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) TOTAL_WORKING_TIME

                  ,ROUND((((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK))) * ROUND(((SH.SUM_MAN )/(SH.COUNT_SHIFT)),2)),2) TOTAL_TIME_ALLMAN
                  ,IFNULL(SUM(OM.LOSS),0) LOSS_TIME
                  ,ROUND(((((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK))) * ROUND(((SH.SUM_MAN )/(SH.COUNT_SHIFT)),2) ) / SUM(OM.JITU_SU)),2) MANHOUR_PCS
          
                  ,ROUND((((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS))/ SUM(OM.JITU_SU)),2) WORKING_TIME_PCS  
                  ,ROUND((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) / SUM(OM.JITU_SU),2) PRODUCTION_TIME_PCS
                  
                  ,CASE WHEN OM.PD = 'PD04' THEN IFNULL(ROUND(((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) / (SUM(OM.JITU_SU) / IFNULL(IT.CAVITY,0)),2),0) 
                   ELSE IFNULL(CY.CYCLE_TIME /(IFNULL(ROUND(((SUM(OM.TOTAL_TIME) - SUM(OM.TOTAL_BREAK)) - SUM(OM.LOSS)) / (SUM(OM.JITU_SU)),2),0)),0) END EFF

                  ,SH.SUM_MAN AS 'TOTAL_MAN'
                  ,SH.COUNT_SHIFT as 'SHIFT_BY_ITEM'
                  ,ROUND(((SH.SUM_MAN )/(SH.COUNT_SHIFT)),2) AS 'MAN'

                  ,IFNULL(CF.A ,0) AS 'A '
                  ,IFNULL(CF.B ,0) AS 'B '
                  ,IFNULL(CF.C ,0) AS 'C '
                  ,IFNULL(CF.C1,0) AS 'C1'
                  ,IFNULL(CF.D ,0) AS 'D '
                  ,IFNULL(CF.E ,0) AS 'E '
                  ,IFNULL(CF.F ,0) AS 'F '
                  ,IFNULL(CF.F1,0) AS 'F1'
                  ,IFNULL(CF.F2,0) AS 'F2'
                  ,IFNULL(CF.G ,0) AS 'G '
                  ,IFNULL(CF.G1,0) AS 'G1'
                  ,IFNULL(CF.H ,0) AS 'H '
                  ,IFNULL(CF.I ,0) AS 'I '
                  ,IFNULL(CF.J ,0) AS 'J '
                  ,IFNULL(CF.K ,0) AS 'K '
                  ,IFNULL(CF.K1,0) AS 'K1'
                  ,IFNULL(CF.K2,0) AS 'K2'
                  ,IFNULL(CF.K3,0) AS 'K3'
                  ,IFNULL(CF.L ,0) AS 'L '
                  ,IFNULL(CF.L1,0) AS 'L1'
                  ,IFNULL(CF.M ,0) AS 'M '
                  ,IFNULL(CF.N ,0) AS 'N '
                  ,IFNULL(CF.O ,0) AS 'O '
                  ,IFNULL(CF.P ,0) AS 'P '
                  ,IFNULL(CF.Q ,0) AS 'Q '
                  ,IFNULL(CF.Q1,0) AS 'Q1'
                  ,IFNULL(CF.R ,0) AS 'R '
                  ,IFNULL(CF.S ,0) AS 'S '
                  ,IFNULL(CF.T ,0) AS 'T '
                  ,IFNULL(CF.U ,0) AS 'U '
                  ,IFNULL(CF.V ,0) AS 'V '
                  ,IFNULL(CF.W ,0) AS 'W '
                                
            FROM
              WK_OEE_WORK_MONTH OM
            LEFT OUTER JOIN
              WK_OEE_CYCLE_TIME CY
            ON CY.ITEM_CD = OM.ITEM_CD  AND  CY.SOURCE_CODE = OM.LINE_CD 
            LEFT OUTER JOIN
              WK_OEE_ITEM_CAVITY IT
            ON IT.ITEM_CD = OM.ITEM_CD
            LEFT OUTER JOIN
                 (SELECT 
                                        PD
                                      ,LINE_CD
                                      ,ITEM_CD
                                      ,WORK_DAY     
                                FROM WK_OEE_WORK_DAY

                                GROUP BY
                                      PD
                                      ,LINE_CD
                                      ,ITEM_CD
                                      ,WORK_DAY 
                  )MI
                  ON OM.LINE_CD = MI.LINE_CD AND OM.ITEM_CD = MI.ITEM_CD 

                  LEFT OUTER JOIN WK_OEE_CODE_ACCUM CF 

                  ON OM.PD = CF.PD AND OM.LINE_CD = CF.LINE_CD AND OM.ITEM_CD = CF.HINBAN 
                  LEFT OUTER JOIN 
                  (SELECT 
                       PD,
                      LINE_CD,
                      ITEM_CD,
                      SUM(MAN) AS 'SUM_MAN',
                      COUNT(ITEM_CD) AS 'COUNT_SHIFT'
                      FROM WK_OEE_SHIFT_DAY 
                    GROUP BY
                      ITEM_CD,
                      LINE_CD
                    ORDER BY
                      LINE_CD,
                      ITEM_CD ASC
                  ) SH 
                   ON OM.LINE_CD = SH.LINE_CD AND OM.ITEM_CD = SH.ITEM_CD  
LEFT OUTER JOIN
                  (
                  SELECT 
                          ITEM_CD, 
                          LINE_CD, 
                          SUM(ACT_NORMAL) ACT_NORMAL, 
                          SUM(ACT_OT) ACT_OT 
                          FROM WK_OEE_ACT_OT 
                          GROUP BY 
                          ITEM_CD, LINE_CD ) OA
                  ON OM.ITEM_CD = OA.ITEM_CD AND OM.LINE_CD = OA.LINE_CD
 WHERE
#OM.PD IN ('PD04')
          OM.PD IN ('PCL1','PD01','PD02','PD03','PD04','PD05','PD06')
          AND OM.START_DATE_TIME BETWEEN DATE_FORMAT(CURDATE(), '%Y/%m/01') AND DATE_FORMAT(CURDATE(), '%Y/%m/%d 07:59:59')
         # AND OM.START_DATE_TIME < CURDATE()


GROUP BY
                    OM.PD
                  ,OM.LINE_CD
                  ,OM.ITEM_CD
                  ,OM.ITEM_NAME
                  ,OM.MODEL

ORDER BY OM.PD , OM.LINE_CD ASC;
       
       
  

;";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();
       // var_dump($recLoad); exit;
     return $recLoad;
}

    public function Prod_report($tbl, $whe=null) 
    {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = "SELECT * FROM $tbl $whe;";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();
        return $recLoad;
    }

    // public function Daily_ship1() 
    // {
    //     $this->EJ = $this->load->database('EJ', true);
    //     $sqlEdt = "SELECT * FROM DAILY_SHIP_1;";
    //     $excEdt = $this->EJ->query($sqlEdt);
    //     $recLoad = $excEdt->result_array();
    //     //var_dump($recLoad); exit;
    //     return $recLoad;
    // }

       public function mysql_fa_report() 
    {
        $this->EJ = $this->load->database('mindoee', true);
        $sqlEdt = "
                     SELECT     
                     
                     MS.NO      
                    ,MS.PD
                    ,MS.LINE_CD
                    ,MS.LOT_NO
                    ,MS.STAFF 
                    ,MS.SHIFT 
                    ,MS.PLAN_DATE 
                    ,MS.SEQ 
                    ,MS.ITEM_CD 
                    ,MS.ITEM_NAME 
                    ,MS.MODEL
                    ,IFNULL((CY.CYCLE_TIME),0) AS 'CYCLE_TIME'
                    ,MS.PLAN
                    ,MS.ACTUAL 
                    ,MS.DIFF 
                    ,MS.START_DATE_TIME 
                    ,MS.END_DATE_TIME 
                    ,MS.WI_NO 
                    ,MS.TOTAL_TIME AS 'TOTALTIME+BREAK'

                    ,(MS.TOTAL_TIME - MS.TOTAL_BREAK ) AS 'TOTAL_TIME'
                    ,MS.TOTAL_BREAK

                    ,CASE WHEN MS.SHIFT = 'M' OR MS.SHIFT = 'N' THEN 0
                      ELSE ((MS.TOTAL_TIME - MS.LOSS)- MS.TOTAL_BREAK) END  'WORK_TIME' 
                    ,MS.OT_TIME 
                      
                    ,ROUND((MS.LOSS),0)AS 'LOSS'
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'G' THEN CF.LOSS END),0) G
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'H' THEN CF.LOSS END),0) H
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'K' THEN CF.LOSS END),0) K
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'K1' THEN CF.LOSS END),0) K1
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'K3' THEN CF.LOSS END),0) K3
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'L' THEN CF.LOSS END),0) L
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'L1' THEN CF.LOSS END),0) L1
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'N' THEN CF.LOSS END),0) N
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'O' THEN CF.LOSS END),0) O
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'Q' THEN CF.LOSS END),0) Q
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'Q1' THEN CF.LOSS END),0) Q1
                    ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'S' THEN CF.LOSS END),0) S

                    FROM OEE_REPORT MS

                    LEFT OUTER JOIN IMPOR_DAILY_CODE CF
                    ON CONCAT(MS.ITEM_CD, MS.LINE_CD, MS.LOT_NO, MS.SEQ, MS.SHIFT, MS.PLAN_DATE) = CONCAT(CF.ITEM_CD, CF.LINE_CD, CF.LOT_NO, CF.SEQ, CF.SHIFT, CF.PLAN_DATE) 

                    LEFT OUTER JOIN CYCLE_TIME CY 
                    ON MS.ITEM_CD = CY.ITEM_CD AND MS.LINE_CD = CY.SOURCE_CODE 
                        
                    GROUP BY
                     MS.NO
                    ,MS.PD
                    ,MS.LINE_CD
                    ,MS.LOT_NO
                    ,MS.SHIFT
                    ,MS.PLAN_DATE
                    ,MS.SEQ
                    ,MS.ITEM_CD
                    ,MS.ITEM_NAME
                    ,MS.MODEL
                    ,MS.PLAN
                    ,MS.ACTUAL
                    ,MS.DIFF
                    ,MS.WI_NO
                    ,MS.STAFF
                    ,MS.TOTAL_TIME
                    ,MS.LOSS

                    ORDER BY MS.NO ASC

        ";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        //var_dump($recLoad); exit;
        return $recLoad;
    }
    public function mysql_fa_report_new() 
    {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = " SELECT 
                           MS.NO      
                          ,MS.PD
                          ,MS.LINE_CD
                          ,MS.LOT_NO
                          ,MS.STAFF 
                          ,MS.SHIFT 
                          ,MS.PLAN_DATE 
                          ,MS.SEQ 
                          ,MS.ITEM_CD 
                          ,MS.ITEM_NAME 
                          ,MS.MODEL
                          ,IFNULL((CY.CYCLE_TIME),0) AS 'STD_CT'
                          ,ROUND((((MS.TOTAL_TIME - MS.LOSS)- MS.TOTAL_BREAK) / MS.ACTUAL ),2) AS 'ACTUAL_CT'
                          ,MS.PLAN
                          ,MS.ACTUAL 
                          ,MS.DIFF 
                          ,MS.START_DATE_TIME 
                          ,MS.END_DATE_TIME 
                          ,MS.WI_NO 
                          ,MS.TOTAL_TIME AS 'TOTALTIME+BREAK'
                          ,(MS.TOTAL_TIME - MS.TOTAL_BREAK ) AS 'TOTAL_TIME'
                          ,MS.TOTAL_BREAK

                          ,CASE WHEN MS.SHIFT = 'M' OR MS.SHIFT = 'N' THEN 0
                            ELSE ((MS.TOTAL_TIME - MS.LOSS) - MS.TOTAL_BREAK) END  'WORK_TIME'
                          ,MS.OT_TIME 
                          ,CASE WHEN MS.PD = 'PD04' THEN IC.CAVITY ELSE 0 END 'CAVITY'
                          ,ROUND((MS.LOSS),0)AS 'LOSS'
                          
                          ,NULL AS 'BANK1'
                          ,NULL AS 'BANK2'
                          ,NULL AS 'BANK3'
                          ,NULL AS 'BANK4'
                      #    ,ROUND((MS.LOSS),0) AS 'LOSS'
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'G' THEN CF.LOSS END),0) G
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'H' THEN CF.LOSS END),0) H
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'I' THEN CF.LOSS END),0) I
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'K' THEN CF.LOSS END),0) K
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'K1' THEN CF.LOSS END),0) K1
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'K2' THEN CF.LOSS END),0) K2
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'K3' THEN CF.LOSS END),0) K3
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'L' THEN CF.LOSS END),0) L
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'N' THEN CF.LOSS END),0) N
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'O' THEN CF.LOSS END),0) O
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'Q' THEN CF.LOSS END),0) Q
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'Q1' THEN CF.LOSS END),0) Q1
                          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'S' THEN CF.LOSS END),0) S
                          ,TY.PRODUCT_TYP AS 'PRODUCT_TYPE'
                          

                          FROM OEE_REPORT MS
                          LEFT OUTER JOIN CYCLE_TIME CY 
                          ON MS.ITEM_CD = CY.ITEM_CD AND MS.LINE_CD = CY.SOURCE_CODE
                          LEFT OUTER JOIN IMPOR_DAILY_CODE CF
                          ON CONCAT(MS.ITEM_CD, MS.LINE_CD, MS.LOT_NO, MS.SEQ, MS.SHIFT, MS.PLAN_DATE ) = CONCAT(CF.ITEM_CD, CF.LINE_CD, CF.LOT_NO, CF.SEQ, CF.SHIFT, CF.PLAN_DATE) 
                          LEFT OUTER JOIN ITEM_CAVITY IC
                          ON IC.ITEM_CD = MS.ITEM_CD 
                          LEFT OUTER JOIN PRODUCT_TYPE TY
                          ON MS.ITEM_CD = TY.ITEM_CD  

                            GROUP BY
                             MS.NO
                            ,MS.PD
                            ,MS.LINE_CD
                            ,MS.LOT_NO
                            ,MS.SHIFT
                            ,MS.PLAN_DATE
                            ,MS.SEQ
                            ,MS.ITEM_CD
                            ,MS.ITEM_NAME
                            ,MS.MODEL
                            ,MS.PLAN
                            ,MS.ACTUAL
                            ,MS.DIFF
                            ,MS.START_DATE_TIME
                            ,MS.END_DATE_TIME
                            ,MS.WI_NO
                            ,MS.STAFF
                            ,MS.TOTAL_TIME
                            ,MS.LOSS
                            ,IC.CAVITY
                            ORDER BY 2 ,3 ,6  


        ";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        //var_dump($recLoad); exit;
        return $recLoad;
    } 

      public function mysql_fa_mon_ac() 
    {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = "     SELECT 
    
                          AC.PD
                        ,AC.LINE_CD
                        ,AC.ITEM_CD 
                        ,AC.ITEM_NAME
                        ,AC.MODEL 
                        ,AC.MAN 
                        ,DA.DAY AS 'WORK_DAY'
                        ,ROUND((AC.MAN / DA.DAY)) as 'AVER_MAN'       
                        ,AC.PLAN  
                        ,AC.ACTUAL
                        ,IFNULL((ST.PART_SET),0) AS 'PART_SET'
                        ,NULL AS 'BANK2'
                        ,NULL AS 'BANK3'
                        ,NULL AS 'BANK4' 

                        ,AC.TOTAL_TIME
                        ,IFNULL((LO.LOSS),0) AS 'LOSS'
                        ,AC.TOTAL_TIME AS 'PRODUC_TIME' 
                        ,ROUND((AC.MAN / DA.DAY)* AC.TOTAL_TIME )  AS 'WORKING_TIME'


                        FROM SUM_ACTUAL AC
                        LEFT OUTER JOIN SUM_LOSS LO 
                                        ON AC.PD = LO.PD  AND  AC.LINE_CD = LO.LINE_CD AND AC.ITEM_CD = LO.ITEM_CD 
                        LEFT OUTER JOIN SUM_DAYWORK DA 
                                        ON AC.ITEM_CD = DA.ITEM_CD AND AC.LINE_CD = DA.LINE_CD
                        LEFT OUTER JOIN PART_SET ST 
                                       ON AC.LINE_CD = ST.LINE_CD AND AC.ITEM_CD = ST.ITEM_CD
                        GROUP BY
                                  AC.PD
                                ,AC.LINE_CD
                                ,AC.ITEM_CD 
                                ,AC.ITEM_NAME
                                ,AC.MODEL 
                                ,AC.MAN 
                                ,AC.PLAN  
                                ,AC.ACTUAL 
                                ,AC.TOTAL_TIME

                        ORDER BY AC.PD , AC.LINE_CD ASC

        ";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        //var_dump($recLoad); exit;
        return $recLoad;
    }
    public function mysql_report_for_ce() 
        {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = "SELECT 

                            LO.PD
                          ,LO.LINE_CD
                          ,LO.PLAN_HI 
                          ,LO.SEQ 
                          ,LO.ITEM_CD 
                          ,LO.ITEM_NAME 
                          ,LO.MODEL 
                          ,LO.SHIFT 
                          ,LO.LOT_NO
                          ,LO.CODE  
                          ,CY.DETAIL
                          ,LO.START_DATE 
                          ,LO.START_TIME 
                          ,LO.END_DATE 
                          ,LO.END_TIME 
                          ,LO.SUM_LOSS 
                          ,LO.UPDATE_DATE 
                          ,LO.KEY_DUP

                          FROM DAILY_REPORT_FOR_CE LO
                          LEFT OUTER JOIN LOSS_CODE_COPY CY
                          ON LO.CODE = CY.CODE      

        ";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        // var_dump($recLoad); exit;
        return $recLoad;
    }
       public function mysql_report_manu_for_ce() 
        {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = "SELECT 
        
                         *

                   FROM LOSS_MANUAL
        ";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        // var_dump($recLoad); exit;
        return $recLoad;
    }


      public function mysql_loss_code() 
    {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = " SELECT 
                    *
                    FROM LOSS_CODE       

        ";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        //var_dump($recLoad); exit;
        return $recLoad;
    }
    public function mysql_shift_code() 
    {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = " SELECT 
                    *
                    FROM SHIFT_MASTER       

        ";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        //var_dump($recLoad); exit;
        return $recLoad;
    }
       public function mysql_loss_report() 
    {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = " SELECT 
                    *
                    FROM IMPOR_DAILY_CODE       

        ";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        //var_dump($recLoad); exit;
        return $recLoad;
    } 

    public function mysql_fa_remain($table='') 
    {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = "SELECT 

                       PD
                      ,LINE_CD
                      ,PLAN_DATE
                      ,SEQ
                      ,WI_NO
                      ,ITEM_NO
                      ,ITEM_NAME
                      ,MODEL
                      
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 1 DAY,'%Y%m%d')THEN PLAN END,0) 01st
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 2 DAY,'%Y%m%d')THEN PLAN END,0) 02nd 
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 3 DAY,'%Y%m%d')THEN PLAN END,0) 03rd 
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 4 DAY,'%Y%m%d')THEN PLAN END,0) 04th 
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 5 DAY,'%Y%m%d')THEN PLAN END,0) 05th 
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 6 DAY,'%Y%m%d')THEN PLAN END,0) 06th 
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 7 DAY,'%Y%m%d')THEN PLAN END,0) 07th 
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 8 DAY,'%Y%m%d')THEN PLAN END,0) 08th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 9 DAY,'%Y%m%d')THEN PLAN END,0) 09th 
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 10 DAY,'%Y%m%d')THEN PLAN END,0) 10th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 11 DAY,'%Y%m%d')THEN PLAN END,0) 11th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 12 DAY,'%Y%m%d')THEN PLAN END,0) 12th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 13 DAY,'%Y%m%d')THEN PLAN END,0) 13th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 14 DAY,'%Y%m%d')THEN PLAN END,0) 14th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 15 DAY,'%Y%m%d')THEN PLAN END,0) 15th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 16 DAY,'%Y%m%d')THEN PLAN END,0) 16th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 17 DAY,'%Y%m%d')THEN PLAN END,0) 17th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 18 DAY,'%Y%m%d')THEN PLAN END,0) 18th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 19 DAY,'%Y%m%d')THEN PLAN END,0) 19th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 20 DAY,'%Y%m%d')THEN PLAN END,0) 20th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 21 DAY,'%Y%m%d')THEN PLAN END,0) 21st
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 22 DAY,'%Y%m%d')THEN PLAN END,0) 22nd
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 23 DAY,'%Y%m%d')THEN PLAN END,0) 23rd
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 24 DAY,'%Y%m%d')THEN PLAN END,0) 24th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 25 DAY,'%Y%m%d')THEN PLAN END,0) 25th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 26 DAY,'%Y%m%d')THEN PLAN END,0) 26th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 27 DAY,'%Y%m%d')THEN PLAN END,0) 27th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 28 DAY,'%Y%m%d')THEN PLAN END,0) 28th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 29 DAY,'%Y%m%d')THEN PLAN END,0) 29th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 30 DAY,'%Y%m%d')THEN PLAN END,0) 30th
                      ,IFNULL(CASE WHEN PLAN_DATE = DATE_FORMAT(LAST_DAY(CURDATE()-INTERVAL 1 MONTH) + INTERVAL 31 DAY,'%Y%m%d')THEN PLAN END,0) 31st


                  FROM $table ;";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();
        //var_dump($recLoad); exit;
        return $recLoad;
    } 

   public function oracle_request_report() 
    {
        $this->ex = $this->load->database('expk', true);
        $sqlEdt = "
                    SELECT 
                       tp.puch_odr_cd ORDER_NO,
                       tp.item_cd,
                       mi.item_name,
                       tp.vend_cd,
                       mv.vend_name,
                       CASE WHEN to_char(tp.puch_odr_dlv_date,'YYYY/MM/DD') < to_char(SYSDATE +1,'YYYY/MM/DD') THEN tp.confirm_dlv_date else tp.puch_odr_dlv_date end puch_odr_dlv_date,
                       TP.WH_CD,
                       tp.puch_odr_qty QTY,
                       us.USER_NAME

                    FROM 
                       t_rlsd_puch_odr tp,
                       m_vend_ctrl mv,
                       m_item mi,
                       USER_MST us,
                       T_ACPT_RSLT ac

                    WHERE 
                        tp.vend_cd = mv.vend_cd(+)
                    AND tp.item_cd = mi.item_cd(+)
                    AND mv.OWN_PERSON_CD = US.USER_CD (+)
                    AND tp.PUCH_ODR_CD =  ac.PUCH_ODR_CD (+)

                    AND ac.PUCH_ODR_CD IS NULL
                    AND tp.puch_odr_sts_typ = 2 
                    AND tp.ODR_CANCEL_SLIP_ISS_FLG = 0 
                    AND substr(tp.item_cd, -3) in ( 'P30' ,'P20')
                    --AND to_char(tp.puch_odr_dlv_date,'YYYY/MM/DD') BETWEEN to_char(SYSDATE+1,'YYYY/MM/DD') AND to_char(SYSDATE+1,'YYYY/MM/DD')
                    AND to_char(tp.puch_odr_dlv_date,'YYYY/MM/DD') = to_char(SYSDATE +1,'YYYY/MM/DD')
                    order by 1,6,8
                      
                  ";

            $excEdt = $this->ex->query($sqlEdt);
            $recLoad = $excEdt->result_array(); 
        //var_dump($recLoad); exit;
        return $recLoad;
    }
   public function oracle_request_vend() 
    {
        $this->ex = $this->load->database('expk', true);
        $sqlEdt = "
                    SELECT 
                       -- tp.puch_odr_cd ORDER_NO,
                       -- tp.item_cd,
                       -- mi.item_name,
                       tp.vend_cd
                       -- mv.vend_name,
                       -- CASE WHEN to_char(tp.puch_odr_dlv_date,'YYYY/MM/DD') < to_char(SYSDATE,'YYYY/MM/DD') THEN tp.confirm_dlv_date else tp.puch_odr_dlv_date end puch_odr_dlv_date,
                       -- TP.WH_CD,
                       -- tp.puch_odr_qty QTY,
                       -- us.USER_NAME

                    FROM 
                       t_rlsd_puch_odr tp,
                       m_vend_ctrl mv,
                       m_item mi,
                       USER_MST us,
                       T_ACPT_RSLT ac

                    WHERE 
                        tp.vend_cd = mv.vend_cd(+)
                    AND tp.item_cd = mi.item_cd(+)
                    AND mv.OWN_PERSON_CD = US.USER_CD (+)
                    AND tp.PUCH_ODR_CD =  ac.PUCH_ODR_CD (+)

                    AND ac.PUCH_ODR_CD IS NULL
                    AND tp.puch_odr_sts_typ = 2 
                    AND tp.ODR_CANCEL_SLIP_ISS_FLG = 0 
                    AND substr(tp.item_cd, -3) in ( 'P30' ,'P20')
                    AND to_char(tp.puch_odr_dlv_date,'YYYY/MM/DD') BETWEEN to_char(SYSDATE +1,'YYYY/MM/DD') AND to_char(SYSDATE +1,'YYYY/MM/DD')
                    group by tp.vend_cd, CASE WHEN to_char(tp.puch_odr_dlv_date,'YYYY/MM/DD') < to_char(SYSDATE +1,'YYYY/MM/DD') THEN tp.confirm_dlv_date else tp.puch_odr_dlv_date end
                    order by 1
                      
                  ";

            $excEdt = $this->ex->query($sqlEdt);
            $recLoad = $excEdt->result_array(); 
        //var_dump($recLoad); exit;
        return $recLoad;
    }

    public function mysql_item_group() 
    {
        $this->EJ = $this->load->database('dbj', true);
        $sqlEdt = "SELECT NUM_ID, `GROUP` AS GP FROM ITEM_GROUP ;";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        //var_dump($recLoad); exit;
        return $recLoad;
    } 

    public function work_day() 
    {
            $this->ex = $this->load->database('expk', true);
            $sqlEdt = "SELECT 
                        COUNT(CASE WHEN SUBSTR(CAL_DATE,0,7) = TO_CHAR(ADD_MONTHS(SYSDATE,0),'YYYY/MM') THEN CAL_DATE END) MONTH1
                       ,COUNT(CASE WHEN SUBSTR(CAL_DATE,0,7) = TO_CHAR(ADD_MONTHS(SYSDATE,1),'YYYY/MM') THEN CAL_DATE END) MONTH2
                       ,COUNT(CASE WHEN SUBSTR(CAL_DATE,0,7) = TO_CHAR(ADD_MONTHS(SYSDATE,2),'YYYY/MM') THEN CAL_DATE END) MONTH3
                       ,COUNT(CASE WHEN SUBSTR(CAL_DATE,0,7) = TO_CHAR(ADD_MONTHS(SYSDATE,3),'YYYY/MM') THEN CAL_DATE END) MONTH4
                       ,COUNT(CASE WHEN SUBSTR(CAL_DATE,0,7) = TO_CHAR(ADD_MONTHS(SYSDATE,4),'YYYY/MM') THEN CAL_DATE END) MONTH5
                        FROM 
                        M_CAL
                        WHERE
                        CAL_NO = 1
                        AND HOLIDAY_FLG = 0";

            $excEdt = $this->ex->query($sqlEdt);
            $recLoad = $excEdt->result_array();           

       // var_dump($recLoad); exit;
        //    return $recLoad;
        return array("m1" => intval($recLoad[0]['MONTH1']), "m2" => intval($recLoad[0]['MONTH2']), "m3" => intval($recLoad[0]['MONTH3']), "m4" => intval($recLoad[0]['MONTH4']), "m5" => intval($recLoad[0]['MONTH5']) );
    }




    public function sale_report($dateSt='2018-01-01', $dateEn='2018-01-02') 
    {
        $this->EX = $this->load->database('expk', true);
        $sqlEdt = "
                    SELECT TO_CHAR (T.SHIP_DATE, 'YYYY-MM-DD') AS SHIP_DATE,
                      T.INVOICE_NO                             AS INV_NO,
                      T.CUST_ODR_NO                            AS ODR_NO,
                      T.CUST_CD                                AS CUST_CD,
                      M.CUST_NAME                              AS CUST_NAME,
                      T.CUST_ITEM_CD                           AS CUST_ITEM_CD,
                      T.CUST_ITEM_NAME                         AS CUST_ITEM_NAME,
                      T.ITEM_CD,
                      MI.ITEM_NAME,
                      NVL (SUM(T.SHIP_QTY), 0)                 AS QTY,
                      'PCS'                                    AS UNIT,
                      T.SHIP_UNIT_PRICE                        AS UNIT_PRICE,
                      NVL (SUM(T.SHIP_AMOUNT), 0) AMOUNT,
                      NULL AS VAT,
                      CASE WHEN SUBSTR (T.INVOICE_NO, 1, 2) IN ('F1', 'F2', 'F3', 'F4', 'E1', 'E2', 'E3')  --OR T.CUST_CD IN ('D20230', 'D20312')
                        THEN 0
                        ELSE NULL
                        END  AS TOTAL_VAT,
                      T.CUR_CD
                    FROM T_SHIP T,
                      M_CUST M,
                      M_ITEM MI
                    WHERE T.CUST_CD = M.CUST_CD (+)
                    AND T.ITEM_CD   = MI.ITEM_CD (+)
                    AND TO_CHAR (T.SHIP_DATE, 'YYYY/MM/DD') BETWEEN '$dateSt' AND '$dateEn'
                    GROUP BY 
                      T.SHIP_DATE,
                      T.INVOICE_NO,
                      T.CUST_ODR_NO,
                      T.CUST_CD,
                      M.CUST_NAME,
                      T.CUST_ITEM_CD,
                      T.CUST_ITEM_NAME,
                      T.ITEM_CD,
                      MI.ITEM_NAME,
                      T.SHIP_UNIT_PRICE,
                      T.CUR_CD
                     
                    UNION ALL
                     
                    SELECT 
                      NULL         AS SHIP_DATE,
                      T.INVOICE_NO AS INV_NO,
                      NULL         AS ODR_NO,
                      T.CUST_CD    AS CUST_CD,
                      M.CUST_NAME  AS CUST_NAME,
                      NULL         AS CUST_ITEM_CD,
                      NULL         AS CUST_ITEM_NAME,
                      NULL         AS ITEM_CD,
                      NULL         AS ITEM_NAME,
                      NULL         AS QTY,
                      NULL         AS UNIT,
                      NULL         AS UNIT_PRICE,
                      NVL (SUM(T.SHIP_AMOUNT), 0) AMOUNT,
                      NVL ( ROUND ( SUM ( CASE WHEN SUBSTR (T.INVOICE_NO, 1, 2) IN ('F1', 'F2', 'F3', 'F4', 'E1', 'E2', 'E3')  --OR T.CUST_CD IN ('D20230', 'D20312')
                        THEN 0
                        ELSE (T.SHIP_AMOUNT * 0.07)
                        END ), 2 ), 0 ) AS VAT,
                      NVL ( ROUND ( SUM ( CASE WHEN SUBSTR (T.INVOICE_NO, 1, 2) IN ('F1', 'F2', 'F3', 'F4', 'E1', 'E2', 'E3')  --OR T.CUST_CD IN ('D20230', 'D20312')
                        THEN (T.SHIP_AMOUNT + 0)
                        ELSE (T.SHIP_AMOUNT + (T.SHIP_AMOUNT * 0.07))
                        END ), 2 ), 0 ) AS TOTAL_VAT,
                      T.CUR_CD
                    FROM T_SHIP T,
                      M_CUST M,
                      M_ITEM MI
                    WHERE T.CUST_CD   = M .CUST_CD (+)
                    AND   T.ITEM_CD   = MI.ITEM_CD (+)
                    AND TO_CHAR (T.SHIP_DATE, 'YYYY/MM/DD') BETWEEN '$dateSt' AND '$dateEn'
                    GROUP BY 
                      M.CUST_NAME,
                      T.CUST_CD,
                      T.INVOICE_NO,
                      T.CUR_CD
                     
                    UNION ALL
                     
                    SELECT 'TOTAL' AS SHIP_DATE,
                      NULL         AS INV_NO,
                      NULL         AS ODR_NO,
                      NULL         AS CUST_CD,
                      TS.CUST_NAME AS CUST_NAME,
                      NULL         AS CUST_ITEM_CD,
                      NULL         AS CUST_ITEM_NAME,
                      NULL         AS ITEM_CD,
                      NULL         AS ITEM_NAME,
                      NULL         AS QTY,
                      NULL         AS UNIT,
                      NULL         AS UNIT_PRICE,
                      NVL (SUM(TS.AMOUNT), 0) AMOUNT,
                      NVL (ROUND(SUM(TS.VAT), 2), 0)       AS VAT,
                      NVL (ROUND(SUM(TS.TOTAL_VAT), 2), 0) AS TOTAL_VAT,
                      TS.CUR_CD
                    FROM
                      (SELECT MM.CUST_NAME,
                        TT.INVOICE_NO,
                        NVL ( SUM(TT.SHIP_AMOUNT), 0) AMOUNT,
                        NVL ( ROUND ( SUM (
                        CASE
                          WHEN SUBSTR (TT.INVOICE_NO, 1, 2) IN ('F1', 'F2', 'F3', 'F4', 'E1', 'E2', 'E3') --OR TT.CUST_CD IN ('D20230', 'D20312')
                          THEN 0
                          ELSE (TT.SHIP_AMOUNT * 0.07)
                        END ), 2 ), 0 ) AS VAT,
                        NVL ( ROUND ( SUM (
                        CASE
                          WHEN SUBSTR (TT.INVOICE_NO, 1, 2) IN ('F1', 'F2', 'F3', 'F4', 'E1', 'E2', 'E3') --OR TT.CUST_CD IN ('D20230', 'D20312')
                          THEN (TT.SHIP_AMOUNT + 0)
                          ELSE (TT.SHIP_AMOUNT +(TT.SHIP_AMOUNT * 0.07))
                        END ), 2 ), 0 ) AS TOTAL_VAT,
                        TT.CUR_CD
                      FROM T_SHIP TT,
                        M_CUST MM
                      WHERE TT.CUST_CD = MM.CUST_CD (+)
                      AND TO_CHAR (TT.SHIP_DATE, 'YYYY/MM/DD') BETWEEN '$dateSt' AND '$dateEn'
                      GROUP BY TT.INVOICE_NO,
                        MM.CUST_NAME,
                        TT.CUR_CD
                      ) TS
                    GROUP BY TS.CUST_NAME,
                      TS.CUR_CD
                    ORDER BY 5,2,3
                ";

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;
    }     



   public function boi() 
    {
        $this->FI = $this->load->database('fin_db', true);
        $sqlEdt = "
                    SELECT * FROM sys_boi
  
                  ";

        $excEdt = $this->FI->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;
    }

 public function boi_rel($dataSale, $dataBoi)
    {
      $data_rel = $dataSale;
        foreach ($dataSale as $saleIndex => $salevalue) 
        {
          //echo $saleIndex . "<hr>";
          foreach ($dataBoi as $boiIndex => $boivalue) 
          {
            if ( $salevalue['ITEM_CD'] == $boivalue['item_no'] )
            {
                $data_rel[$saleIndex]['BOI_CD']   = $boivalue['boi_cd'];
                $data_rel[$saleIndex]['BOI_NAME'] = $boivalue['boi_name'];   
                break;
            }
            else
            {
                $data_rel[$saleIndex]['BOI_CD']   = null;
                $data_rel[$saleIndex]['BOI_NAME'] = null;                          
            }
          }
        }
        return $data_rel;
    }

    public function inf_pu( $dateSt = 2019/01/01, $dateEn = 2019/01/01, $whe="--", $com='') 
    {
        $this->EX = $this->load->database('expk', true);
        $sqlEdt = "
                    SELECT
                     INTERNAL_CTRL_CD 
                    ,COMPANY_CD 
                    ,BUSINESS_PATTERN_CD 
                    ,SEQ_NO 
                    ,REFERENCE_ORG_TYP 
                    ,SALES_SEQ_NO 
                    ,ODR_CTL_NO 
                    ,SALES_SLIP_CD 
                    ,PUCH_ODR_CD 
                    ,ACPT_NO 
                    ,INSPC_ACPT_NO 
                    ,ONEROUS_CONS_NO 
                    ,TEMP_ODR_CD 
                    ,DEFECT_DISPOSAL_SLIP_CD 
                    ,DEFECT_DISPOSAL_CRCT_NO 
                    ,DEFECT_DISPOSAL_CRCT_TYP 
                    ,EXCE_COST_SLIP_CD 
                    ,EXCE_COST_CRCT_NO 
                    ,EXCE_COST_CRCT_TYP 
                    ,GNR_TYP 
                    ,PAYEE_CD 
                    ,VEND_CD 
                    ,PLANT_CD 
                    ,ITEM_CD 
                    ,ITEM_NAME 
                    ,UNIT_COST 
                    ,UNIT_COST_TYP 
                    ,INSPC_ACPT_QTY 
                    ,STOCK_UNIT 
                    ,INSPC_ACPT_AMOUNT 
                    ,SAVING_AMOUNT 
                    ,TO_CHAR(INSPC_ACPT_DATE,'YYYY/MM/DD')  INSPC_ACPT_DATE
                    ,TAX_CD 
                    ,AP_IF_FLG 
                    ,TO_CHAR(AP_IF_EXEC_DATE,'YYYY/MM/DD')  AP_IF_EXEC_DATE
                    ,AI_IF_FLG 
                    ,TO_CHAR(AI_IF_EXEC_DATE,'YYYY/MM/DD')  AI_IF_EXEC_DATE
                    ,AP_DENKBN_B 
                    ,AP_DENKBN_R 
                    ,AP_USSOSHIKI 
                    ,AP_SAIMUKANJYOKBN 
                    ,AP_TORIKBN_B 
                    ,AP_TORIKBN_R 
                    ,AP_HINCD 
                    ,AP_DENYOBIM02_B 
                    ,AP_DENYOBIM02_R 
                    ,AI_L_IF_COL_VAL 
                    ,AI_R_IF_COL_VAL 
                    ,AI_L_SLIPTYPE_CD 
                    ,AI_R_SLIPTYPE_CD 
                    ,AI_L_ACCRUAL_DEPT 
                    ,AI_R_ACCRUAL_DEPT 
                    ,AI_L_ACCOUNT_CD 
                    ,AI_R_ACCOUNT_CD 
                    ,CASE WHEN ( (NOT( VEND_CD  LIKE 'T%' OR VEND_CD LIKE 'M%')) OR ( VEND_CD = 'T10100' OR VEND_CD = 'T11200' OR VEND_CD = 'T11300') ) THEN 1 ELSE 2 END AI_L_SUBJECT_CD 
                    ,AI_R_SUBJECT_CD 
                    ,AI_L_DEPT_CD 
                    ,AI_R_DEPT_CD 
                    ,AI_L_DESCRIPTION1 
                    ,AI_R_DESCRIPTION1 
                    ,AI_DESCRIPTION_NAME_TYP 
                    ,AI_L_TAX_JUDGE_TYP 
                    ,AI_R_TAX_JUDGE_TYP 
                    ,AI_TAX_CD_NOT_CALC 
                    ,TAX_TYP 
                    ,TAX_RATE_1 
                    ,TAX_RATE_2 
                    ,TAX_RATE_3 
                    ,DEFECT_FACTOR_CLASS_TYP 
                    ,DEFECT_FACTOR_CD 
                    ,BEFORE_DEFECT_LINE_CD 
                    ,BEFORE_DEFECT_VEND_CD 
                    ,PROGRESS_PER 
                    ,CUR_CD 
                    ,CREATE_SLIP_TYP 
                    ,SLIP_CTRL_COMPANY_CD 
                    ,SHIP_RTN_FLG 
                    ,GRP_COMPANY_TYP 
                    ,TRN_TYP 
                    ,VEND_GRP_TYP 
                    ,DEFECT_DISPOSAL_TYP 
                    ,COSTPROC_TYP 
                    ,COUNTERPARTY_CD 
                    ,COST_CTRL_COMPANY_CD 
                    ,COST_OBTAIN_TYP 
                    ,COST_OBTAIN_TRADER_CD 
                    ,TO_CHAR(COST_STD_DATE,'YYYY/MM/DD')  COST_STD_DATE
                    ,COST_STD_QTY 
                    ,JNL_JUDGE_TYP 
                    ,INVOICE_NO 
                    ,NULL AS NA
                    ,TO_CHAR(ASIA_IF_EXEC_DATE,'YYYY/MM/DD') ASIA_IF_EXEC_DATE

                    FROM
                    UT_SLIP_INSPC
                    WHERE
                    PUCH_ODR_CD IS NOT NULL AND
                    ASIA_IF_EXEC_DATE IS NOT NULL AND
                    $whe
                    NOT( VEND_CD = 'D20230' OR VEND_CD = 'D20220'  OR VEND_CD = 'D20210' OR VEND_CD = 'L48070' OR  VEND_CD = 'L48190'  OR VEND_CD = 'L48050' OR VEND_CD = 'L48030' OR VEND_CD = 'L48170' OR VEND_CD = 'L48120' OR VEND_CD = 'L48020' OR VEND_CD = 'L48180' OR VEND_CD = 'L48010' OR VEND_CD = 'L48140' OR VEND_CD = 'L48150' OR VEND_CD = 'L48040' OR VEND_CD = 'L48160') AND 
                    VEND_CD <> 'L40810' $com AND
                    -- NOT(ITEM_CD LIKE 'P%')  AND 
                    $com TO_CHAR(CREATED_DATE,'YYYY/MM/DD') BETWEEN '$dateSt' AND '$dateEn'

                    ORDER BY 1 ASC
                ";

       //echo $sqlEdt; exit;

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();
//var_dump($recLoad); exit;
        return $recLoad;
    } 
    public function inf_sa($dateSt='2018-01-01', $dateEn='2018-01-02' , $whe="--", $com='') 
    {
        $this->EX = $this->load->database('expk', true);
        $sqlEdt = "
                    SELECT
                      INTERNAL_CTRL_CD
                     ,COMPANY_CD
                     ,BUSINESS_PATTERN_CD
                     ,SEQ_NO
                     ,REFERENCE_ORG_TYP
                     ,SALES_SEQ_NO
                     ,ODR_CTL_NO
                     ,SALES_SLIP_CD
                     ,PUCH_ODR_CD
                     ,ACPT_NO
                     ,INSPC_ACPT_NO
                     ,ONEROUS_CONS_NO
                     ,TEMP_ODR_CD
                     ,DEFECT_DISPOSAL_SLIP_CD
                     ,DEFECT_DISPOSAL_CRCT_NO
                     ,DEFECT_DISPOSAL_CRCT_TYP
                     ,EXCE_COST_SLIP_CD
                     ,EXCE_COST_CRCT_NO
                     ,EXCE_COST_CRCT_TYP
                     ,GNR_TYP
                     ,CUST_CD
                     ,CUST_ITEM_CD
                     ,CUST_ITEM_NAME
                     ,FINAL_DLV_LOC_CD
                     ,CUST_ODR_NO
                     ,PLANT_CD
                     ,ITEM_CD
                     ,ITEM_NAME
                     ,UNIT_COST
                     ,UNIT_COST_TYP
                     ,INSPC_ACPT_QTY
                     ,STOCK_UNIT
                     ,INSPC_ACPT_AMOUNT
                     ,TO_CHAR(INSPC_ACPT_DATE,'YYYY/MM/DD')  INSPC_ACPT_DATE
                     ,TAX_CD
                     ,AI_IF_FLG
                     ,TO_CHAR(AI_IF_EXEC_DATE,'YYYY/MM/DD')  AI_IF_EXEC_DATE
                     ,AI_L_IF_COL_VAL
                     ,AI_R_IF_COL_VAL
                     ,AI_L_SLIPTYPE_CD
                     ,AI_R_SLIPTYPE_CD
                     ,AI_L_ACCRUAL_DEPT
                     ,AI_R_ACCRUAL_DEPT
                     ,AI_L_ACCOUNT_CD
                     ,AI_R_ACCOUNT_CD
                     ,AI_L_SUBJECT_CD
                     ,CASE WHEN ( (NOT(CUST_CD  LIKE 'T%' OR CUST_CD LIKE 'F%')) OR ( CUST_CD = 'T10100' OR CUST_CD = 'T11200' OR CUST_CD = 'T11300') ) THEN 1 ELSE 2 END AI_R_SUBJECT_CD
                     ,AI_L_DEPT_CD
                     ,AI_R_DEPT_CD
                     ,AI_L_DESCRIPTION1
                     ,AI_R_DESCRIPTION1
                     ,AI_DESCRIPTION_NAME_TYP
                     ,AI_L_TAX_JUDGE_TYP
                     ,AI_R_TAX_JUDGE_TYP
                     ,AI_TAX_CD_NOT_CALC
                     ,TAX_TYP
                     ,TAX_RATE_1
                     ,TAX_RATE_2
                     ,TAX_RATE_3
                     ,CUR_CD
                     ,CREATE_SLIP_TYP
                     ,SLIP_CTRL_COMPANY_CD
                     ,SHIP_RTN_FLG
                     ,GRP_COMPANY_TYP
                     ,TRN_TYP
                     ,VEND_GRP_TYP
                     ,DEFECT_DISPOSAL_TYP
                     ,COSTPROC_TYP
                     ,COUNTERPARTY_CD
                     ,COST_CTRL_COMPANY_CD
                     ,COST_OBTAIN_TYP
                     ,COST_OBTAIN_TRADER_CD
                     ,TO_CHAR(COST_STD_DATE,'YYYY/MM/DD') COST_STD_DATE
                     ,COST_STD_QTY
                     ,JNL_JUDGE_TYP
                     ,INVOICE_NO
                     ,NULL ASIA_IF_FLG
                     ,TO_CHAR(ASIA_IF_EXEC_DATE,'YYYY/MM/DD') ASIA_IF_EXEC_DATE
                    FROM
                    UT_SLIP_SALES
                    WHERE
                    CUST_ODR_NO IS NOT NULL 
                    AND ASIA_IF_EXEC_DATE IS NOT NULL
                    $whe 
                    $com AND TO_CHAR(CREATED_DATE,'YYYY/MM/DD') BETWEEN '$dateSt' AND '$dateEn'


                ";

               //echo $sqlEdt; exit;;

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;
    } 


   public function picking_list( $dateUse = '2019/03/18', $cust = 'BANK', $cust_po = '--' )
   {

   		$cust = ( $cust == 'MEC-SUM'  )  ?  "AND (CS.CUST_ANAME = 'MEC-SUPPLEMENT Free Zone' OR CS.CUST_ANAME = 'MEC-Free zone')"  : "AND CS.CUST_ANAME = '$cust'" ;
        $this->EX = $this->load->database('expk', true);
        $sqlEdt = "
                          SELECT 
                          --CS.CUST_ANAME,
                          TR.ITEM_CD,
                          $cust_po TR.CUST_ITEM_CD,
                          ITT.ITEM_NAME,
                          IT.MODEL AS MODEL,
                          
                          $cust_po TR.CUST_ODR_NO,
                          IT.pkg_unit_qty AS SNP,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:01' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 11:59:59' THEN TR.ODR_QTY END),0) 
                          - MOD (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:01' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 11:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS PERIOD_1to2,


                          MOD (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:01' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 11:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_1,

                          CEIL(ROUND 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:01' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 11:59:59' THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_1,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          --NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = TO_CHAR(SYSDATE+1,'YYYY/MM/DD') THEN TR.ODR_QTY END),0) AS 'PLAN NEXT DAY',
                          --NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = TO_CHAR(SYSDATE+2,'YYYY/MM/DD') THEN TR.ODR_QTY END),0) AS 'PLAN NEXT 2 DAYS'
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------                          
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 12:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 16:59:59' THEN TR.ODR_QTY END),0) 
                          - MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 12:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 16:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          )AS PERIOD_3to4,


                          MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 12:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 16:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_2,

                          CEIL(ROUND (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 12:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 16:59:59' THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_2,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 17:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0) 
                          - MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 17:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          )AS PERIOD_5to6,


                          MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 17:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_3,

                          CEIL(ROUND 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 17:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_3,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse 00:00:00'  THEN TR.ODR_QTY END),0) 
                          - MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse 00:00:00'  THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS PERIOD_7tp8,
                          MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse 00:00:00'  THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_4,

                          CEIL(ROUND 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse 00:00:00'  THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_4,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          NULL AS REMARK
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                          FROM
                            (SELECT 
                            CUST_CD, 
                            CUST_ITEM_CD, 
                            CUST_ITEM_NAME, 
							$cust_po CUST_ODR_NO,
                            ITEM_CD, 
                            CASE WHEN CUST_CD IN ( 'T00100', 'T10200', 'T10600', 'T10300' ) THEN STNDRD_RCV_DESINATED_DLV_DATE ELSE DESINATED_DLV_DATE END DESINATED_DLV_DATE,
                            ODR_QTY, 
                            TOTAL_SHIP_QTY, 
                            DEL_FLG 
                            FROM 
                            T_ODR) TR, 

                          M_CUST CS,
                          M_PLANT_ITEM IT,
                          M_ITEM ITT

                          WHERE
                          TR.ITEM_CD = IT.ITEM_CD (+)
                          AND TR.ITEM_CD = ITT.ITEM_CD (+)
                          AND TR.CUST_CD = CS.CUST_CD (+)
                          AND TR.DEL_FLG = 0
                          AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = '$dateUse'-- AND '$dateUse 23:59:59'
                          --AND CS.CUST_ANAME LIKE '%$cust%'
                          $cust
                          GROUP BY 
                          --TR.CUST_CD,
                          --CS.CUST_ANAME,
                          TR.ITEM_CD,
                          ITT.ITEM_NAME,
                          IT.MODEL,
                          $cust_po TR.CUST_ITEM_CD,
                          $cust_po TR.CUST_ODR_NO,
                          IT.pkg_unit_qty
						  
                          ORDER BY
                          IT.MODEL desc
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                ";

               //echo $sqlEdt; exit;

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;

   } 
   public function picking_list_iemt( $dateUse = '2019/03/18', $cust = 'BANK', $cust_po = '--' )
   {

   		$cust_iemt = ( $cust == 'IEMT-SUM'  )  ?  " CUST_CD  IN ('D20110','D20111','D20112') "  : " CUST_CD = 'D20113'" ;


   		$dateUse_N = date('Y/m/d', strtotime("+1 day", strtotime($dateUse) ) ) ;
        $this->EX = $this->load->database('expk', true);
        $sqlEdt = "
                          SELECT 
                          --CS.CUST_ANAME,
                          TR.ITEM_CD,
                          
                          ITT.ITEM_NAME,
                          IT.MODEL AS MODEL,
                          $cust_po IT.CUST_ITEM_CD,
                          $cust_po TR.CUST_ODR_NO,
                          IT.pkg_unit_qty AS SNP,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 10:59:59' THEN TR.ODR_QTY END),0) 
                          - MOD (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 10:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS PERIOD_1to2,


                          MOD (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 10:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_1,

                          CEIL(ROUND 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 10:59:59' THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_1,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          --NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = TO_CHAR(SYSDATE+1,'YYYY/MM/DD') THEN TR.ODR_QTY END),0) AS 'PLAN NEXT DAY',
                          --NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = TO_CHAR(SYSDATE+2,'YYYY/MM/DD') THEN TR.ODR_QTY END),0) AS 'PLAN NEXT 2 DAYS'
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------                          
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 11:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 15:59:59' THEN TR.ODR_QTY END),0) 
                          - MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 11:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 15:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          )AS PERIOD_3to4,


                          MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 11:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 15:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_2,

                          CEIL(ROUND (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 11:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 15:59:59' THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_2,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 16:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0) 
                          - MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 16:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          )AS PERIOD_5to6,


                          MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 16:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_3,

                          CEIL(ROUND 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 16:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_3,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00'  THEN TR.ODR_QTY END),0) 
                          - MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00'  THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS PERIOD_7tp8,
                          MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00'  THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_4,

                          CEIL(ROUND 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00'  THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_4,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          NULL AS REMARK
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                          FROM
                            (SELECT 
                            CUST_CD, 
                            CUST_ITEM_CD, 
                            CUST_ITEM_NAME, 
							$cust_po CUST_ODR_NO,
                            ITEM_CD, 
                            CASE 
								                WHEN CUST_CD = 'T00100' THEN TO_DATE(TO_CHAR(STNDRD_RCV_DESINATED_DLV_DATE, 'YYYY/MM/DD') || ' ' || TO_CHAR(STNDRD_RCV_DESINATED_DLV_DATE, 'HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS')
								                WHEN CUST_CD = 'D20110' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '10:00:00','YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20111' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '19:00:00','YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20112' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '19:00:00','YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20113' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '23:00:00','YYYY/MM/DD HH24:MI:SS')
								ELSE TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD') || ' ' || TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS') END DESINATED_DLV_DATE,  
                            ODR_QTY, 
                            TOTAL_SHIP_QTY, 
                            DEL_FLG 
                            FROM 
                            T_ODR



                            WHERE
                              --CUST_CD  IN ('D20110','D20111','D20112')
                              $cust_iemt
                              AND( ( TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD') = '$dateUse' AND TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') != '$dateUse 04:30:00' ) OR TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00' )                           
                              AND DEL_FLG = 0
 
                            ) TR, 

                          M_CUST CS,
                          M_PLANT_ITEM IT,
                          M_ITEM ITT

                          WHERE
                          TR.ITEM_CD = IT.ITEM_CD (+)
                          AND TR.ITEM_CD = ITT.ITEM_CD (+)
                          AND TR.CUST_CD = CS.CUST_CD (+)

                          --$cust
                          --AND ( TO_CHAR(TR.DESINATED_DLV_DATE, 'YYYY/MM/DD') = '$dateUse' AND TO_CHAR(TR.DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') <> '$dateUse 04:30:00' ) OR TO_CHAR(TR.DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00'
                          --AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = '$dateUse'-- AND '$dateUse 23:59:59'
                          --AND CS.CUST_ANAME LIKE '%$cust%'
                          
                          GROUP BY 
                          --TR.CUST_CD,
                          --CS.CUST_ANAME,
                          TR.ITEM_CD,
                          ITT.ITEM_NAME,
                          IT.MODEL,
                          $cust_po TR.CUST_ODR_NO,
                          IT.pkg_unit_qty
						  
                          ORDER BY
                          IT.MODEL desc
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                ";

               // /echo $sqlEdt; exit;

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;

   } 
   public function picking_list_skc( $dateUse = '2019/03/18', $cust = 'BANK', $cust_po = '--' )
   {

      $cust_iemt = ( $cust == 'SKC-SUM'  )  ?  " CUST_CD  IN ('D20310','D20311') "  : " CUST_CD = 'D20113'" ;


      $dateUse_N = date('Y/m/d', strtotime("+1 day", strtotime($dateUse) ) ) ;
        $this->EX = $this->load->database('expk', true);
        $sqlEdt = "
                          SELECT 
                          --CS.CUST_ANAME,
                          TR.ITEM_CD,
                          
                          ITT.ITEM_NAME,
                          IT.MODEL AS MODEL,
                          
                          $cust_po TR.CUST_ODR_NO,
                          IT.pkg_unit_qty AS SNP,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 10:59:59' THEN TR.ODR_QTY END),0) 
                          - MOD (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 10:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS PERIOD_1to2,


                          MOD (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 10:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_1,

                          CEIL(ROUND 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 00:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 10:59:59' THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_1,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          --NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = TO_CHAR(SYSDATE+1,'YYYY/MM/DD') THEN TR.ODR_QTY END),0) AS 'PLAN NEXT DAY',
                          --NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = TO_CHAR(SYSDATE+2,'YYYY/MM/DD') THEN TR.ODR_QTY END),0) AS 'PLAN NEXT 2 DAYS'
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------                          
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 11:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 15:59:59' THEN TR.ODR_QTY END),0) 
                          - MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 11:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 15:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          )AS PERIOD_3to4,


                          MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 11:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 15:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_2,

                          CEIL(ROUND (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 11:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 15:59:59' THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_2,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 16:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0) 
                          - MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 16:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          )AS PERIOD_5to6,


                          MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 16:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_3,

                          CEIL(ROUND 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') >= '$dateUse 16:00:00' AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') <= '$dateUse 23:59:59' THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_3,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00'  THEN TR.ODR_QTY END),0) 
                          - MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00'  THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS PERIOD_7tp8,
                          MOD 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00'  THEN TR.ODR_QTY END),0),
                          IT.pkg_unit_qty
                          ) AS remain_4,

                          CEIL(ROUND 
                          (
                          NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00'  THEN TR.ODR_QTY END),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_4,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          NULL AS REMARK
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                          FROM
                            (SELECT 
                            CUST_CD, 
                            CUST_ITEM_CD, 
                            CUST_ITEM_NAME, 
              $cust_po CUST_ODR_NO,
                            ITEM_CD, 
                            CASE 
                                WHEN CUST_CD = 'T00100' THEN TO_DATE(TO_CHAR(STNDRD_RCV_DESINATED_DLV_DATE, 'YYYY/MM/DD') || ' ' || TO_CHAR(STNDRD_RCV_DESINATED_DLV_DATE, 'HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20110' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '10:00:00','YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20111' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '23:00:00','YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20112' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '23:00:00','YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20113' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '23:00:00','YYYY/MM/DD HH24:MI:SS')
                ELSE TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD') || ' ' || TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS') END DESINATED_DLV_DATE,  
                            ODR_QTY, 
                            TOTAL_SHIP_QTY, 
                            DEL_FLG 
                            FROM 
                            T_ODR



                            WHERE
                              --CUST_CD  IN ('D20110','D20111','D20112')
                              $cust_iemt
                              AND( ( TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD') = '$dateUse' AND TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') != '$dateUse 04:30:00' ) OR TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00' )                           
                              AND DEL_FLG = 0
 
                            ) TR, 

                          M_CUST CS,
                          M_PLANT_ITEM IT,
                          M_ITEM ITT

                          WHERE
                          TR.ITEM_CD = IT.ITEM_CD (+)
                          AND TR.ITEM_CD = ITT.ITEM_CD (+)
                          AND TR.CUST_CD = CS.CUST_CD (+)
                          --$cust
                          --AND ( TO_CHAR(TR.DESINATED_DLV_DATE, 'YYYY/MM/DD') = '$dateUse' AND TO_CHAR(TR.DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') <> '$dateUse 04:30:00' ) OR TO_CHAR(TR.DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') = '$dateUse_N 04:30:00'
                          --AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = '$dateUse'-- AND '$dateUse 23:59:59'
                          --AND CS.CUST_ANAME LIKE '%$cust%'
                          
                          GROUP BY 
                          --TR.CUST_CD,
                          --CS.CUST_ANAME,
                          TR.ITEM_CD,
                          ITT.ITEM_NAME,
                          IT.MODEL,
                          $cust_po TR.CUST_ODR_NO,
                          IT.pkg_unit_qty
              
                          ORDER BY
                          IT.MODEL desc
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                ";

               //echo $sqlEdt; exit;

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;

   }    
   public function picking_list_packing( $dateUse = '2019/03/18', $cust = 'BANK', $cust_po = '--' )
   {

      $cust_iemt = ( $cust == 'IEMT-SUM'  )  ?  " CUST_CD  IN ('D20110','D20111','D20112') "  : " CUST_CD = 'D20113'" ;

      //$id_cust = 


      $dateUse_N = date('Y/m/d', strtotime("+1 day", strtotime($dateUse) ) ) ;
        $this->EX = $this->load->database('expk', true);
        $sqlEdt = "

                          SELECT 
                          TR.CUST_CD,
                          TR.ITEM_CD,
                          
                          ITT.ITEM_NAME,
                          IT.MODEL AS MODEL,
                          
                          TR.CUST_ODR_NO,
                          IT.pkg_unit_qty AS SNP,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                          NVL(SUM(TR.ODR_QTY),0) 
                          - MOD (
                          NVL(SUM(TR.ODR_QTY),0),
                          IT.pkg_unit_qty
                          ) AS PERIOD_1to2,


                          MOD (
                          NVL(SUM(TR.ODR_QTY),0),
                          IT.pkg_unit_qty
                          ) AS remain_1,

                          CEIL(ROUND 
                          (
                          NVL(SUM(TR.ODR_QTY),0) /
                          IT.pkg_unit_qty,
                          2
                          )) box_1,
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          --NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = TO_CHAR(SYSDATE+1,'YYYY/MM/DD') THEN TR.ODR_QTY END),0) AS 'PLAN NEXT DAY',
                          --NVL(SUM(CASE WHEN TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = TO_CHAR(SYSDATE+2,'YYYY/MM/DD') THEN TR.ODR_QTY END),0) AS 'PLAN NEXT 2 DAYS'
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------                          

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          NULL AS REMARK
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                          FROM
                            (SELECT 
                            CUST_CD, 
                            CUST_ITEM_CD, 
                            CUST_ITEM_NAME,
                            --CUST_ITEM_ANAME,
                            CUST_ODR_NO,
                            ITEM_CD, 
                            CASE 
                                WHEN CUST_CD IN ( 'T00100', 'T10200' ) THEN TO_DATE(TO_CHAR(STNDRD_RCV_DESINATED_DLV_DATE, 'YYYY/MM/DD') || ' ' || TO_CHAR(STNDRD_RCV_DESINATED_DLV_DATE, 'HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20110' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '10:00:00','YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20111' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '23:00:00','YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20112' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '23:00:00','YYYY/MM/DD HH24:MI:SS')
                                WHEN CUST_CD = 'D20113' AND  TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS') = '00:00:00'   THEN TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY-MM-DD') || ' ' || '23:00:00','YYYY/MM/DD HH24:MI:SS')
                                ELSE TO_DATE(TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD') || ' ' || TO_CHAR(DESINATED_DLV_DATE, 'HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS') END DESINATED_DLV_DATE,  
                            ODR_QTY, 
                            TOTAL_SHIP_QTY, 
                            DEL_FLG 
                            FROM 
                            T_ODR
                            WHERE
                              --CUST_CD  IN ('D20110','D20111','D20112')
                               CUST_CD  IN ('D20510') 
                              AND( ( TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD') = '2019/04/26' AND TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') != '2019/04/26 04:30:00' ) OR TO_CHAR(DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') = '2019/04/27 04:30:00' )                           
                            
 
                            ) TR, 

                          M_CUST CS,
                          M_PLANT_ITEM IT,
                          M_ITEM ITT

                          WHERE
                          TR.ITEM_CD = IT.ITEM_CD (+)
                          AND TR.ITEM_CD = ITT.ITEM_CD (+)
                          AND TR.CUST_CD = CS.CUST_CD (+)
                          --IEMT-SUM
                          --AND ( TO_CHAR(TR.DESINATED_DLV_DATE, 'YYYY/MM/DD') = '2019/04/26' AND TO_CHAR(TR.DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') <> '2019/04/26 04:30:00' ) OR TO_CHAR(TR.DESINATED_DLV_DATE, 'YYYY/MM/DD HH24:MI:SS') = '2019/04/27 04:30:00'
                          --AND TO_CHAR(TR.DESINATED_DLV_DATE,'YYYY/MM/DD') = '2019/04/26'-- AND '2019/04/26 23:59:59'
                          --AND CS.CUST_ANAME LIKE '%IEMT-SUM%'
                          
                          GROUP BY 
                          TR.CUST_CD,
                          --CS.CUST_ANAME,
                          TR.ITEM_CD,
                          ITT.ITEM_NAME,
                          IT.MODEL,
                          TR.CUST_ODR_NO,
                          IT.pkg_unit_qty
              
                          ORDER BY
                          IT.MODEL desc
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                

                ";

              // echo $sqlEdt; exit;

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;

   } 
   public function list_cust(  $cust = 'BANK' )
   {

    

        $this->EX = $this->load->database('expk', true);
        $sqlEdt = " SELECT CUST_CD, CUST_ANAME, CUST_NAME FROM M_CUST WHERE CUST_ANAME = '$cust' AND ROWNUM = 1 ";

                  

                



        //echo $sqlEdt; exit;

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;

   }
   public function list_bom( $where = '--' )
   {

    

        $this->EX = $this->load->database('expk', true);
        $sqlEdt = " SELECT  PARENT_ITEM_CD HEAD, COMP_ITEM_CD UNDERS, PS_UNIT_NUMERATOR UP FROM M_PLANT_PS  $where ";

                  

                



        //echo $sqlEdt; exit;

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;

   }
   public function list_bom1( $where = '--' )
   {

    

        $this->EX = $this->load->database('expk', true);
        $sqlEdt = " SELECT
                    PS.PARENT_ITEM_CD HEAD,
                    PS.COMP_ITEM_CD UNDERS,
                    PS.PS_UNIT_NUMERATOR UP,
                    MI.ITEM_NAME,
                    MPI. MODEL
                    FROM
                      M_PLANT_PS PS,
                      M_ITEM MI,
                      M_PLANT_ITEM MPI
                    WHERE
                      PS.COMP_ITEM_CD = MI.ITEM_CD
                    AND PS.COMP_ITEM_CD = MPI.ITEM_CD
                    AND $where  ";

                    

                



        //echo $sqlEdt; exit;

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;

   }
    public function d_pur_report($picking_start, $picking_stop)
    {
        $this->EX = $this->load->database('expk', true);
        //$this->exp = $this->load->database('exp_db', TRUE);
       // $sqlExp = "SELECT MENU_NAME FROM menu_mst GROUP BY MENU_NAME ORDER BY MENU_CD "; 
        $sqlEx = " 
                SELECT 
                        TO_CHAR(TP.INSPC_ACPT_DATE,'YYYY-MM-DD') INSPC_ACPT_DATE,
                        TP.INVOICE_NO,
                        TP.PUCH_ODR_CD,
                        TP.VEND_CD,
                        VC.VEND_NAME,
                        TP.ITEM_CD,
                        TP.ITEM_NAME,
                        MP.MODEL,
                        SUM(TP.INSPC_ACPT_QTY) QTY,
                        TP.STOCK_UNIT UNIT,
                        TP.UNIT_COST PRICE,
                        SUM(TP.INSPC_ACPT_AMOUNT) PRICE_TOTAL,
                        NULL DEBIT,
                        NULL CAEDIT,
                        NULL NET_AMOUNT,
                        NULL VAT,
                        NULL TOTAL_VAT,
                        VC.CUR_CD,
                        PO.CREATED_BY,
                        UM.USER_NAME,
                       -- NULL SECTION
                        CASE 
                          WHEN UM.USER_CD    = 'SYSTEM-50' THEN 'PC'
                          WHEN UM.SECTION_CD = 'K1PL00'    THEN 'PC'
                          WHEN UM.SECTION_CD = 'K2PL00'    THEN 'PC'
                          WHEN UM.SECTION_CD = 'K1PU00'    THEN 'PU'
                          ELSE SUBSTR(UM.SECTION_CD,3,2) 
                        END SECTION
                          
                FROM  
                    T_PAST_INSPC_ACPT TP,
                    T_RLSD_PUCH_ODR PO,
                    M_VEND_CTRL VC,
                    M_PLANT_ITEM MP,
                    USER_MST UM

                WHERE 

                     TP.VEND_CD = VC.VEND_CD(+)
                     AND TP.PUCH_ODR_CD = PO.PUCH_ODR_CD(+)

                     AND TP.PUCH_ODR_CD = PO.PUCH_ODR_CD(+)
                     AND PO.CREATED_BY = UM.USER_CD(+)
                     AND TP.PLANT_CD = MP.PLANT_CD
                     AND TP.ITEM_CD = MP.ITEM_CD(+)                     
                     AND TP.INVOICE_NO <> '-'  
                     AND (NOT(TP.VEND_CD  LIKE 'T%' OR TP.VEND_CD LIKE 'M%') OR (TP.VEND_CD = 'T10100' OR TP.VEND_CD = 'T11200' OR TP.VEND_CD = 'T11300'))
                     AND TO_CHAR(TP.INSPC_ACPT_DATE , 'YYYY/MM/DD') BETWEEN '$picking_start' AND '$picking_stop'
                GROUP BY
                     TP.INSPC_ACPT_DATE,
                     TP.INVOICE_NO,
                     TP.PUCH_ODR_CD,
                     TP.VEND_CD,
                     VC.VEND_NAME,
                     TP.ITEM_CD,
                     TP.ITEM_NAME,
                     MP.MODEL,
                     TP.STOCK_UNIT,
                     TP.UNIT_COST,
                     VC.CUR_CD,
                     PO.CREATED_BY,
                     UM.USER_NAME,
                        CASE 
                          WHEN UM.USER_CD    = 'SYSTEM-50' THEN 'PC'
                          WHEN UM.SECTION_CD = 'K1PL00'    THEN 'PC'
                          WHEN UM.SECTION_CD = 'K2PL00'    THEN 'PC'
                          WHEN UM.SECTION_CD = 'K1PU00'    THEN 'PU'
                          ELSE SUBSTR(UM.SECTION_CD,3,2) 
                        END 
                  
                UNION ALL

                SELECT 
                        NULL INSPC_ACPT_DATE,
                        TP.INVOICE_NO,
                        NULL PUCH_ODR_CD,
                        TP.VEND_CD,
                        VC.VEND_NAME,
                        NULL ITEM_CD,
                        NULL ITEM_NAME,
                        NULL MODEL,
                        SUM(TP.INSPC_ACPT_QTY) QTY,
                        TP.STOCK_UNIT UNIT,
                        NULL PRICE,
                        SUM(TP.INSPC_ACPT_AMOUNT) PRICE_TOTAL ,
                        NULL DEBIT,
                        NULL CAEDIT,
                        NULL NET_AMOUNT,                        
                        ROUND( SUM(TP.INSPC_ACPT_AMOUNT) *0.07 ,2) VAT,
                        (ROUND( SUM(TP.INSPC_ACPT_AMOUNT) *0.07 ,2)) + SUM(TP.INSPC_ACPT_AMOUNT) TOTAL_VAT,
                        VC.CUR_CD,
                        NULL UPDATED_BY,
                        NULL USER_NAME,
                        NULL SECTION
                FROM  
                    T_PAST_INSPC_ACPT TP,
                    M_VEND_CTRL VC

                WHERE 

                     TP.VEND_CD = VC.VEND_CD(+)
                     AND TP.INVOICE_NO <> '-' 
                     AND (NOT(TP.VEND_CD  LIKE 'T%' OR TP.VEND_CD LIKE 'M%') OR (TP.VEND_CD = 'T10100' OR TP.VEND_CD = 'T11200' OR TP.VEND_CD = 'T11300'))
                     AND TO_CHAR(TP.INSPC_ACPT_DATE , 'YYYY/MM/DD')  BETWEEN '$picking_start' AND '$picking_stop'      
                 

                GROUP BY
                        TP.INVOICE_NO,
                        TP.VEND_CD,
                        VC.VEND_NAME,
                        TP.STOCK_UNIT,
                        VC.CUR_CD                
                UNION ALL

                SELECT 
                        NULL INSPC_ACPT_DATE,
                        NULL INVOICE_NO,
                        NULL PUCH_ODR_CD,
                        NULL VEND_CD,
                        TP.VEND_NAME,
                        NULL ITEM_CD,
                        NULL MODEL,
                        NULL ITEM_NAME,
                        SUM(TP.INSPC_ACPT_QTY) QTY,
                        TP.STOCK_UNIT UNIT,
                        NULL PRICE,
                        SUM(TP.INSPC_ACPT_AMOUNT) PRICE_TOTAL ,
                        NULL DEBIT,
                        NULL CAEDIT,
                        NULL NET_AMOUNT,                        
                        SUM(TP.VAT)  VAT,
                        SUM(TP.TOTAL_VAT)  TOTAL_VAT,
                        TP.CUR_CD,
                        NULL UPDATED_BY,
                        NULL USER_NAME,
                        NULL SECTION
                FROM  
                   (
                    SELECT 
                        NULL INSPC_ACPT_DATE,
                        TP.INVOICE_NO,
                        NULL PUCH_ODR_CD,
                        TP.VEND_CD,
                        VC.VEND_NAME,
                        NULL ITEM_CD,
                        NULL ITEM_NAME,
                        SUM(TP.INSPC_ACPT_QTY) INSPC_ACPT_QTY,
                        TP.STOCK_UNIT,
                        NULL PRICE,
                        SUM(TP.INSPC_ACPT_AMOUNT) INSPC_ACPT_AMOUNT ,
                        ROUND( SUM(TP.INSPC_ACPT_AMOUNT) *0.07 ,2) VAT,
                        (ROUND( SUM(TP.INSPC_ACPT_AMOUNT) *0.07 ,2)) + SUM(TP.INSPC_ACPT_AMOUNT) TOTAL_VAT,
                        VC.CUR_CD
                    FROM  
                      T_PAST_INSPC_ACPT TP,
                      M_VEND_CTRL VC


                    WHERE 

                        TP.VEND_CD = VC.VEND_CD(+)
                        AND TP.INVOICE_NO <> '-' 
                        AND (NOT(TP.VEND_CD  LIKE 'T%' OR TP.VEND_CD LIKE 'M%') OR (TP.VEND_CD = 'T10100' OR TP.VEND_CD = 'T11200' OR TP.VEND_CD = 'T11300'))
                        AND TO_CHAR(TP.INSPC_ACPT_DATE , 'YYYY/MM/DD')  BETWEEN '$picking_start' AND '$picking_stop'      
                 

                    GROUP BY
                        TP.INVOICE_NO,
                        TP.VEND_CD,
                        VC.VEND_NAME,
                        TP.STOCK_UNIT,
                        VC.CUR_CD
                    ) TP
                    

                GROUP BY
                        TP.VEND_NAME,
                        TP.STOCK_UNIT,
                        TP.CUR_CD      
                  
                ORDER BY 5,2,1
                 "; 
        $excEx = $this->EX->query($sqlEx);
        $recEx = $excEx->result_array();


        //var_dump($recEx); //exit;
        return $recEx;

    }

    public function o_pur_report($picking_start, $picking_stop)
    {
        $this->EX = $this->load->database('expk', true);
        //$this->exp = $this->load->database('exp_db', TRUE);
       // $sqlExp = "SELECT MENU_NAME FROM menu_mst GROUP BY MENU_NAME ORDER BY MENU_CD "; 
        $sqlEx = " 
                SELECT 
                        TO_CHAR(TP.INSPC_ACPT_DATE,'YYYY-MM-DD') INSPC_ACPT_DATE,
                        TP.INVOICE_NO,
                        TP.PUCH_ODR_CD,
                        TP.VEND_CD,
                        VC.VEND_NAME,
                        TP.ITEM_CD,
                        TP.ITEM_NAME,
                        SUM(TP.INSPC_ACPT_QTY) QTY,
                        TP.STOCK_UNIT UNIT,
                        TP.UNIT_COST PRICE,
                        SUM(TP.INSPC_ACPT_AMOUNT) PRICE_TOTAL,
                        NULL VAT,
                        NULL TOTAL_VAT,
                        VC.CUR_CD
                FROM  
                    T_PAST_INSPC_ACPT TP,
                    M_VEND_CTRL VC

                WHERE 

                     TP.VEND_CD = VC.VEND_CD(+)
                     AND TP.INVOICE_NO <> '-'  
                     AND ((TP.VEND_CD  LIKE 'T%' OR TP.VEND_CD LIKE 'M%') AND NOT(TP.VEND_CD = 'T10100' OR TP.VEND_CD = 'T11200' OR TP.VEND_CD = 'T11300'))
                     AND TO_CHAR(TP.INSPC_ACPT_DATE , 'YYYY/MM/DD') BETWEEN '$picking_start' AND '$picking_stop'
                GROUP BY
                     TP.INSPC_ACPT_DATE,
                     TP.INVOICE_NO,
                     TP.PUCH_ODR_CD,
                     TP.VEND_CD,
                     VC.VEND_NAME,
                     TP.ITEM_CD,
                     TP.ITEM_NAME,
                     TP.STOCK_UNIT,
                     TP.UNIT_COST,
                     VC.CUR_CD    
                  
                UNION ALL

                SELECT 
                        NULL INSPC_ACPT_DATE,
                        TP.INVOICE_NO,
                        NULL PUCH_ODR_CD,
                        TP.VEND_CD,
                        VC.VEND_NAME,
                        NULL ITEM_CD,
                        NULL ITEM_NAME,
                        SUM(TP.INSPC_ACPT_QTY) QTY,
                        TP.STOCK_UNIT UNIT,
                        NULL PRICE,
                        SUM(TP.INSPC_ACPT_AMOUNT) PRICE_TOTAL ,
                        0 VAT,
                        0 + SUM(TP.INSPC_ACPT_AMOUNT) TOTAL_VAT,
                        VC.CUR_CD
                FROM  
                        T_PAST_INSPC_ACPT TP,
                        M_VEND_CTRL VC

                WHERE 

                     TP.VEND_CD = VC.VEND_CD(+)
                     AND TP.INVOICE_NO <> '-' 
                     AND ((TP.VEND_CD  LIKE 'T%' OR TP.VEND_CD LIKE 'M%') AND NOT(TP.VEND_CD = 'T10100' OR TP.VEND_CD = 'T11200' OR TP.VEND_CD = 'T11300'))
                     AND TO_CHAR(TP.INSPC_ACPT_DATE , 'YYYY/MM/DD') BETWEEN '$picking_start' AND '$picking_stop'      
                 

                GROUP BY
                        TP.INVOICE_NO,
                        TP.VEND_CD,
                        VC.VEND_NAME,
                        TP.STOCK_UNIT,
                        VC.CUR_CD
                  
                UNION ALL

                SELECT 
                        NULL INSPC_ACPT_DATE,
                        NULL INVOICE_NO,
                        NULL PUCH_ODR_CD,
                        NULL VEND_CD,
                        TP.VEND_NAME,
                        NULL ITEM_CD,
                        NULL ITEM_NAME,
                        SUM(TP.INSPC_ACPT_QTY) QTY,
                        TP.STOCK_UNIT UNIT,
                        NULL PRICE,
                        SUM(TP.INSPC_ACPT_AMOUNT) PRICE_TOTAL ,
                        0  VAT,
                        0 + SUM(TP.INSPC_ACPT_AMOUNT) TOTAL_VAT,
                        TP.CUR_CD
                FROM  
                   (
                    SELECT 
                        NULL INSPC_ACPT_DATE,
                        TP.INVOICE_NO,
                        NULL PUCH_ODR_CD,
                        TP.VEND_CD,
                        VC.VEND_NAME,
                        NULL ITEM_CD,
                        NULL ITEM_NAME,
                        SUM(TP.INSPC_ACPT_QTY) INSPC_ACPT_QTY,
                        TP.STOCK_UNIT,
                        NULL PRICE,
                        SUM(TP.INSPC_ACPT_AMOUNT) INSPC_ACPT_AMOUNT ,
                        NULL VAT,
                        0 + SUM(TP.INSPC_ACPT_AMOUNT) TOTAL_VAT,
                        VC.CUR_CD
                    FROM  
                        T_PAST_INSPC_ACPT TP,
                        M_VEND_CTRL VC

                    WHERE 

                        TP.VEND_CD = VC.VEND_CD(+)
                        AND TP.INVOICE_NO <> '-' 
                        AND ((TP.VEND_CD  LIKE 'T%' OR TP.VEND_CD LIKE 'M%') AND NOT(TP.VEND_CD = 'T10100' OR TP.VEND_CD = 'T11200' OR TP.VEND_CD = 'T11300'))
                        AND TO_CHAR(TP.INSPC_ACPT_DATE , 'YYYY/MM/DD') BETWEEN '$picking_start' AND '$picking_stop'      
                 

                    GROUP BY
                        TP.INVOICE_NO,
                        TP.VEND_CD,
                        VC.VEND_NAME,
                        TP.STOCK_UNIT,
                        VC.CUR_CD
                    ) TP
                    

                GROUP BY
                        TP.VEND_NAME,
                        TP.STOCK_UNIT,
                        TP.CUR_CD      
                  
                ORDER BY 5,2,1
                 "; 
        $excEx = $this->EX->query($sqlEx);
        $recEx = $excEx->result_array();


        //var_dump($recEx); //exit;
        return $recEx;

    }
      
    }  
?>
