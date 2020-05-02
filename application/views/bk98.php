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
        $sqlEdt = "SELECT 
                              SU.PD
                              ,SU.LINE_CD                           
                              ,SU.ITEM_CD
                              ,SU.ITEM_NAME
                              ,SY.MODEL
                              ,SU.PLAN_QTY
                              ,CASE WHEN SU.SUP_FROM = '' THEN SU.SUP_FROM ELSE SY.STOCK_ON_HAND_QTY END AS 'STOCK'
                              ,CASE WHEN SU.SUP_FROM = '' THEN SU.SUP_FROM ELSE ROUND(SY.STOCK_ON_HAND_QTY/SU.PLAN_QTY,2) END  AS 'LEVEL Str'
                              ,SU.SUP_FROM
                              ,SU.LOCATION
                              ,SU.WI_NO
                              ,NULL AS 'REMARK'

                          FROM $table SU 

                          LEFT OUTER JOIN STOCK_FOR_SUPPLY SY 

                          ON SU.ITEM_CD = SY.ITEM_CD

                          ORDER BY 2,11,9 ASC
                           

;";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        //var_dump($recLoad); exit;
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
          ,IFNULL((CY.CYCLE_TIME),0) AS 'CYCLE_TIME'
          ,MS.PLAN
          ,MS.ACTUAL 
          ,MS.DIFF 
          ,MS.START_DATE_TIME 
          ,MS.END_DATE_TIME 
          ,MS.WI_NO 
          ,MS.TOTAL_TIME AS 'TOTAL_TIME'
          ,MS.OT_TIME 
          ,(MS.TOTAL_TIME - MS.TOTAL_BREAK ) AS 'TOTALTIME'
          ,MS.TOTAL_BREAK
          ,(((MS.TOTAL_TIME - MS.OT_TIME ) - MS.LOSS) - MS.TOTAL_BREAK) AS 'WORKING_TIME'
          ,ROUND((MS.LOSS),0)AS 'LOSS'
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'G' THEN CF.LOSS END), 0) G
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'DH' THEN CF.LOSS END), 0) DH
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'H' THEN CF.LOSS END), 0) H
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'L' THEN CF.LOSS END), 0) L
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'DL' THEN CF.LOSS END), 0) DL
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'DK' THEN CF.LOSS END), 0) DK
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'K' THEN CF.LOSS END), 0) K
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'DN' THEN CF.LOSS END), 0) DN
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'N' THEN CF.LOSS END), 0) N
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'O' THEN CF.LOSS END), 0) O
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'R' THEN CF.LOSS END), 0) R
          ,IFNULL(SUM(CASE WHEN CF.ERROR_CD = 'PE' THEN CF.LOSS END), 0) PE

          FROM OEE_REPORT MS
          LEFT OUTER JOIN CYCLE_TIME CY 
          ON MS.ITEM_CD = CY.ITEM_CD AND MS.LINE_CD = CY.SOURCE_CODE
          LEFT OUTER JOIN IMPOR_DAILY_CODE CF
          ON CONCAT(MS.ITEM_CD, MS.LINE_CD, MS.LOT_NO, MS.SEQ, MS.SHIFT, MS.PLAN_DATE ) = CONCAT(CF.ITEM_CD, CF.LINE_CD, CF.LOT_NO, CF.SEQ, CF.SHIFT, CF.PLAN_DATE) 

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

                ORDER BY MS.NO ASC

        ";
        $excEdt = $this->EJ->query($sqlEdt);
        $recLoad = $excEdt->result_array();


        //var_dump($recLoad); exit;
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
                       CASE WHEN to_char(tp.puch_odr_dlv_date,'YYYY/MM/DD') < to_char(SYSDATE,'YYYY/MM/DD') THEN tp.confirm_dlv_date else tp.puch_odr_dlv_date end puch_odr_dlv_date,
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
                    AND to_char(tp.puch_odr_dlv_date,'YYYY/MM/DD') = to_char(SYSDATE,'YYYY/MM/DD')
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
                    AND to_char(tp.puch_odr_dlv_date,'YYYY/MM/DD') BETWEEN to_char(SYSDATE,'YYYY/MM/DD') AND to_char(SYSDATE,'YYYY/MM/DD')
                    group by tp.vend_cd, CASE WHEN to_char(tp.puch_odr_dlv_date,'YYYY/MM/DD') < to_char(SYSDATE,'YYYY/MM/DD') THEN tp.confirm_dlv_date else tp.puch_odr_dlv_date end
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
                      CASE WHEN SUBSTR (T.INVOICE_NO, 1, 2) IN ('F1', 'F2', 'F3', 'F4', 'E1', 'E2', 'E3')  OR T.CUST_CD IN ('D20230', 'D20312')
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
                      NVL ( ROUND ( SUM ( CASE WHEN SUBSTR (T.INVOICE_NO, 1, 2) IN ('F1', 'F2', 'F3', 'F4', 'E1', 'E2', 'E3')  OR T.CUST_CD IN ('D20230', 'D20312')
                        THEN 0
                        ELSE (T.SHIP_AMOUNT * 0.07)
                        END ), 2 ), 0 ) AS VAT,
                      NVL ( ROUND ( SUM ( CASE WHEN SUBSTR (T.INVOICE_NO, 1, 2) IN ('F1', 'F2', 'F3', 'F4', 'E1', 'E2', 'E3')  OR T.CUST_CD IN ('D20230', 'D20312')
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
                          WHEN SUBSTR (TT.INVOICE_NO, 1, 2) IN ('F1', 'F2', 'F3', 'F4', 'E1', 'E2', 'E3') OR TT.CUST_CD IN ('D20230', 'D20312')
                          THEN 0
                          ELSE (TT.SHIP_AMOUNT * 0.07)
                        END ), 2 ), 0 ) AS VAT,
                        NVL ( ROUND ( SUM (
                        CASE
                          WHEN SUBSTR (TT.INVOICE_NO, 1, 2) IN ('F1', 'F2', 'F3', 'F4', 'E1', 'E2', 'E3') OR TT.CUST_CD IN ('D20230', 'D20312')
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
                    ,CASE WHEN ( (NOT(VEND_CD  LIKE 'T%' OR VEND_CD LIKE 'M%')) OR ( VEND_CD = 'T10100' OR VEND_CD = 'T11200' OR VEND_CD = 'T11300') ) THEN 1 ELSE 2 END AI_L_SUBJECT_CD 
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
                    NOT(VEND_CD = 'L40860' OR VEND_CD = 'D20230' OR VEND_CD = 'D20220'  OR VEND_CD = 'D20210' ) AND
                    VEND_CD <> 'L40810' AND
                    NOT(ITEM_CD LIKE 'P%') $com AND 
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
                            CASE WHEN CUST_CD = 'T00100' THEN STNDRD_RCV_DESINATED_DLV_DATE ELSE DESINATED_DLV_DATE END DESINATED_DLV_DATE,
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

//               echo $sqlEdt; exit;

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
   public function list_bom(  $lvl = 1, $where = '--' )
   {

    

        $this->EX = $this->load->database('expk', true);
        $sqlEdt = " SELECT $lvl AS LVL, PARENT_ITEM_CD HEAD, COMP_ITEM_CD UNDERS, PS_UNIT_NUMERATOR UP FROM M_PLANT_PS  $where ";

                  

                



        //echo $sqlEdt; exit;

        $excEdt = $this->EX->query($sqlEdt);
        $recLoad = $excEdt->result_array();

        return $recLoad;

   }

      
    }  
?>
