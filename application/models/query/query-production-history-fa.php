<?php
date_default_timezone_set('Asia/Bangkok');
class FAHISTORY_ACTUAL
{


	public function __construct()
	 { 
		//parent::__construct();
	 }


    public function DB2GET_FAHISTORY()
	    { 
            $qur = "";
            //$mc = date('m', strtotime("- 1 month", strtotime( date('Y-m-01') ) ));
            foreach(range(1, 12 ) as $m )
            {
                $mn = date('Ym', strtotime("+". ($m-1) ."month", strtotime( date("Y-01-01") ) ) );
                $am = date('F', strtotime("+". ($m-1) ."month", strtotime( date("Y-01-01") ) )  );
                $qur .= sprintf(",NVL(SUM(CASE WHEN AP.MN = '%s' THEN AP.ACTU END ),0 ) %s \n", $mn, $am);
            }

            $str_sql = sprintf( 
                "SELECT 
                    CASE WHEN AP.PD = 'PCL1' THEN 'PL00' ELSE AP.PD END PD
                    ,AP.LINE_CD
                    ,AP.ITEM_CD
                    ,AP.ITEM_NM
                    ,SUM(AP.ACTU) ACTU_SUM
                    $qur
                FROM(
                            SELECT
                                LM.SYOZK_CD PD
                            ,SH.LINE_CD
                            ,SH.HINBAN ITEM_CD
                            ,SH.HINMEI ITEM_NM
                            ,SUM(SH.JITU_SU) ACTU
                            ,SUBSTR(JITU_SD, 1, 6) MN
                            
                            FROM 
                                SEISAN_H SH
                            ,LINE_MST LM
                            
                            WHERE
                                        LM.LINE_CD = SH.LINE_CD
                                    AND SH.JITU_SD > '%s'	
                                    AND TO_CHAR(TO_DATE(CONCAT(SH.JITU_SD,SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') BETWEEN '%s' AND '%s'
                                    AND SH.JITU_SU > 0
                                    AND LM.SYOZK_CD NOT IN ('Z999')
                            --	 AND SH.HINBAN = '1320A047'
                            GROUP BY 
                                SH.LINE_CD
                                ,LM.SYOZK_CD
                                ,SH.KISYUMEI
                                ,SH.HINMEI
                                ,SH.HINBAN
                                ,SUBSTR(JITU_SD, 1, 6)
                ) AP
                
                GROUP BY 
                        CASE WHEN AP.PD = 'PCL1' THEN 'PL00' ELSE AP.PD END
                    ,AP.LINE_CD
                    ,AP.ITEM_CD
                    ,AP.ITEM_NM
                ORDER BY 1 , 2  ASC
                "
                ,date('Ymd', strtotime( date("Y-01-01") ))
                ,date('Y/m/d H:i:s', strtotime( date("Y-01-01 08:00:00") ))
                ,date('Y/m/d H:i:s', strtotime("+1 month" ,strtotime( date("Y-m-01 08:00:00") ) ) )
            );

            return $str_sql; //exit;

	    }
























}









?>