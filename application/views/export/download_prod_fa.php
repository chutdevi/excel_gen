<?php

date_default_timezone_set("Asia/Bangkok");

		$date_tm = date('Ymd', strtotime("+1 day", strtotime(date('Y/m/d') ) ));
		$date_to = date('Ymd', strtotime("+0 day", strtotime(date('Y/m/d') ) ));

		$date_to = date('Ymd', strtotime("+0 day", strtotime(date('Y/m/d') ) ));


		$date_da = date('d',   strtotime("+0 day", strtotime(date('Y/m/d') ) ));
		$date_cn = date('t',   strtotime("+0 day", strtotime(date('Y/m/d') ) ));

		//echo $date_cn; exit;
	foreach ( range(1, ( $date_da - 1 ) ) as $cdate) 
	{
		$date_tm = date('Ymd', strtotime("+1 day", strtotime(date('Y/m/').$cdate ) ));
		$date_to = date('Ymd', strtotime("+0 day", strtotime(date('Y/m/').$cdate ) ));
		$sql = "
				SELECT
				    LM.SYOZK_CD PD
				   ,SH.LINE_CD
				   ,SH.JITU_LOT LOT_NO
				   ,SH.JININ_SU MAN
				   ,SH.CYOKU_K SHIFT
				   ,TO_CHAR(TO_DATE(SH.PLAN_HI, 'YYYYMMDD'), 'YYYY/MM/DD') PLAN_DATE 
				   ,SH.PLAN_JUN SEQ
				   ,SH.HINBAN ITEM_CD
				   ,SH.HINMEI ITEM_NAME
				   ,SH.KISYUMEI MODEL
				   ,SH.PLAN_SU
				   ,SH.JITU_SU
				   ,SH.JITU_SU - SH.PLAN_SU DIFF
				   ,TO_CHAR(TO_DATE(CONCAT(SH.JITU_SD,SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') START_DATE_TIME
				   ,TO_CHAR(TO_DATE(CONCAT(SH.JITU_ED,SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') END_DATE_TIME
				   ,SH.SAGYO_SIJI_NO WI_NO
				   ,NVL(TIMESTAMPDIFF (4, CHAR(TIMESTAMP(CONCAT(SH.JITU_ED,SH.JITU_ET),'YYYY.MM.DD.HH.MI.SS')- TIMESTAMP(CONCAT(SH.JITU_SD,SH.JITU_ST),'YYYY.MM.DD.HH.MI.SS') )),0) AS TOTAL_TIME   
				   ,CASE   WHEN SH.CYOKU_K = 'P' AND SH.JITU_ET > 173000
				           THEN
								CASE WHEN SH.JITU_ST < 173000 
									 THEN NVL(TIMESTAMPDIFF (4, CHAR(TIMESTAMP(CONCAT(SH.JITU_ED,SH.JITU_ET),'YYYY.MM.DD.HH.MI.SS') - TIMESTAMP(CONCAT(SH.JITU_SD,173000),'YYYY.MM.DD.HH.MI.SS') )),0) 										
									 ELSE NVL(TIMESTAMPDIFF (4, CHAR(TIMESTAMP(CONCAT(SH.JITU_ED,SH.JITU_ET),'YYYY.MM.DD.HH.MI.SS') - TIMESTAMP(CONCAT(SH.JITU_SD,JITU_ST),'YYYY.MM.DD.HH.MI.SS') )),0) 	
								END
				            WHEN SH.CYOKU_K = 'M' AND SH.JITU_ET > 173000
							THEN
								 CASE WHEN SH.JITU_ST < 173000 
								 	  THEN NVL(TIMESTAMPDIFF (4, CHAR(TIMESTAMP(CONCAT(SH.JITU_ED,SH.JITU_ET),'YYYY.MM.DD.HH.MI.SS') - TIMESTAMP(CONCAT(SH.JITU_SD,173000),'YYYY.MM.DD.HH.MI.SS') )),0) 
								 	  ELSE NVL(TIMESTAMPDIFF (4, CHAR(TIMESTAMP(CONCAT(SH.JITU_ED,SH.JITU_ET),'YYYY.MM.DD.HH.MI.SS') - TIMESTAMP(CONCAT(SH.JITU_SD,JITU_ST),'YYYY.MM.DD.HH.MI.SS'))),0) 
								 END
				            WHEN SH.CYOKU_K = 'Q' AND TO_CHAR(TO_DATE(CONCAT(SH.JITU_ED,SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '{$date_tm}053000','YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
							THEN
								 CASE WHEN TO_CHAR(TO_DATE(CONCAT(SH.JITU_SD,SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '{$date_tm}053000','YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
								 	  THEN NVL(TIMESTAMPDIFF (4, CHAR(TIMESTAMP(CONCAT(SH.JITU_ED,SH.JITU_ET),'YYYY.MM.DD.HH.MI.SS') - TIMESTAMP( '{$date_tm}053000','YYYY.MM.DD.HH.MI.SS') )),0)
									  ELSE NVL(TIMESTAMPDIFF (4, CHAR(TIMESTAMP(CONCAT(SH.JITU_ED,SH.JITU_ET),'YYYY.MM.DD.HH.MI.SS') - TIMESTAMP(CONCAT(SH.JITU_SD,JITU_ST),'YYYY.MM.DD.HH.MI.SS')  )),0)
								 END
				            WHEN SH.CYOKU_K = 'N' AND TO_CHAR(TO_DATE(CONCAT(SH.JITU_ED,SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE('{$date_tm}053000','YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )    
							THEN
				                 CASE WHEN TO_CHAR(TO_DATE(CONCAT(SH.JITU_SD,SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE('{$date_tm}053000','YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                 	  THEN NVL(TIMESTAMPDIFF (4, CHAR(TIMESTAMP(CONCAT(SH.JITU_ED,SH.JITU_ET),'YYYY.MM.DD.HH.MI.SS') - TIMESTAMP('{$date_tm}053000','YYYY.MM.DD.HH.MI.SS') )),0)
				                      ELSE NVL(TIMESTAMPDIFF (4, CHAR(TIMESTAMP(CONCAT(SH.JITU_ED,SH.JITU_ET),'YYYY.MM.DD.HH.MI.SS') - TIMESTAMP(CONCAT(SH.JITU_SD,JITU_ST),'YYYY.MM.DD.HH.MI.SS')  )),0)
				                 END
							ELSE 0
				    END OT_TIME

				   ,SUM(TIMESTAMPDIFF (4, CHAR(TIMESTAMP(CONCAT(LT.KYUSI_ED,LT.KYUSI_ET),'YYYY.MM.DD.HH.MI.SS')  - TIMESTAMP(CONCAT(LT.KYUSI_SD,LT.KYUSI_ST ),'YYYY.MM.DD.HH.MI.SS') ))) AS LOSS  
				   ,CASE  WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 100000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				               					 AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 100000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
						  THEN 10 
				          ELSE 0 
				     END A1       
				    ,CASE  WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 120000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				               					  AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 124000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				           THEN 40
				           ELSE 0 
				     END A2       
				    ,CASE  WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 150000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )  
				               					  AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 151000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				           THEN 10
				           ELSE 0 
				     END A3 
				    ,CASE WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 170000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     			 AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 173000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
						  THEN 30 
				          ELSE 0 
				     END A4      
				    ,CASE WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 100000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     			 AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 221000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				          THEN 10 
				          ELSE 0  
				     END B1
				    ,CASE WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_tm 000000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     			 AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_tm 004000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				          THEN 40
				          ELSE 0
				     END B2
				    ,CASE WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_tm 030000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     			 AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_tm 031000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )                        
				          THEN 10
				          ELSE 0
				     END B3
				    ,CASE WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_tm 050000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     			 AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_tm 053000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				          THEN 30
				          ELSE 0
				     END B4
				    ,CASE WHEN SH.CYOKU_K = 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 190000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     			AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 191000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				          THEN 10
				          ELSE 0
				     END S1
				    ,CASE
				                     WHEN SH.CYOKU_K = 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 210000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     					   AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 214000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				           THEN 40
				           ELSE 0
				     END S2
				    ,CASE
				                     WHEN SH.CYOKU_K = 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_tm 000000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				                     					   AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_tm 001000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )   
				           THEN 10
				           ELSE 0
				     END S3
				    ,CASE
				                     WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 100000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     						AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 101000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     THEN 10
				                     ELSE 0
				                END +
				                CASE
				                     WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 120000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     						AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 124000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     THEN 40
				                     ELSE 0
				                END +
				                CASE
				                     WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 150000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     					   	AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 151000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     THEN 10
				                     ELSE 0
				                END +
				                CASE
				                     WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 170000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     					    AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 173000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     THEN 30
				                     ELSE 0
				                END +
				                CASE
				                     WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 220000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     						AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 221000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     THEN 10
				                     ELSE 0
				                END +
				                CASE
				                     WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_tm 000000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     						AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_tm 004000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )  
				                     THEN 40
				                     ELSE 0
				                END +
				                CASE
				                     WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_tm 030000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     						AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_tm 031000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     THEN 10
				                     ELSE 0
				                END +
				                CASE
				                     WHEN SH.CYOKU_K <> 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_tm 050000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     					    AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_tm 053000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     THEN 30
				                     ELSE 0
				                END +
				                CASE
				                     WHEN SH.CYOKU_K = 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 190000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     					   AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 191000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				                     THEN 10
				                     ELSE 0
				                END +
				                CASE
				                     WHEN SH.CYOKU_K = 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_to 210000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				                      					   AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_to 214000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				                     THEN 40
				                     ELSE 0
				                END +
				                CASE
				                     WHEN SH.CYOKU_K = 'S' AND TO_CHAR(TO_DATE((SH.JITU_SD||SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') < TO_CHAR(TO_DATE( '$date_tm 000000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) 
				                     					   AND TO_CHAR(TO_DATE((SH.JITU_ED||SH.JITU_ET), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') > TO_CHAR(TO_DATE( '$date_tm 001000', 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				                     THEN 10
				                     ELSE 0
				                END TOTAL_BREAK
				FROM
				               SEISAN_H SH
				LEFT OUTER JOIN LINE_MST LM
				               ON LM.LINE_CD = SH.LINE_CD
				LEFT OUTER JOIN
				(
				 SELECT
				 LINE_CD
				,PLAN_HI
				,PLAN_JUN
				,KYUSI_ED
				,KYUSI_ET
				,KYUSI_SD
				,KYUSI_ST
				,CYOKU_K
				FROM KYUSIJIT_F
				WHERE
				    KYUSI_SD >= '$date_to'
				    AND (TO_CHAR(TO_DATE(CONCAT(KYUSI_SD,KYUSI_ST),'YYYYMMDDHH24MISS'),'YYYY/MM/DD HH24:MI:SS') >= TO_CHAR(TO_DATE('$date_to 080000','YYYY.MM.DD.HH24.MI.SS'),'YYYY/MM/DD HH24:MI:SS') 
				    AND  TO_CHAR(TO_DATE(CONCAT(KYUSI_SD,KYUSI_ST),'YYYYMMDDHH24MISS'),'YYYY/MM/DD HH24:MI:SS') <= TO_CHAR(TO_DATE('$date_tm 075959','YYYY.MM.DD.HH24.MI.SS'),'YYYY/MM/DD HH24:MI:SS'))
				) LT
				ON LM.LINE_CD = LT.LINE_CD AND SH.PLAN_JUN = LT.PLAN_JUN AND SH.CYOKU_K = LT.CYOKU_K AND SH.PLAN_HI = LT.PLAN_HI
				WHERE
				    SH.JITU_SD >= '$date_to'
				    AND TO_CHAR(TO_DATE(CONCAT(SH.JITU_SD,SH.JITU_ST), 'YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS') BETWEEN TO_CHAR(TO_DATE('$date_to 080000','YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' ) AND TO_CHAR(TO_DATE('$date_tm 075959','YYYYMMDDHH24MISS'), 'YYYY/MM/DD HH24:MI:SS' )
				    AND SH.JITU_SU > 0
				GROUP BY
				   SH.LINE_CD
				  ,LM.SYOZK_CD
				  ,SH.SAGYO_SIJI_NO
				  ,SH.JITU_ED
				  ,SH.JITU_ET
				  ,SH.JITU_SD
				  ,SH.JITU_ST
				  ,SH.JITU_SU
				  ,SH.PLAN_SU
				  ,SH.KISYUMEI
				  ,SH.HINMEI
				  ,SH.HINBAN
				  ,SH.PLAN_SU
				  ,SH.PLAN_JUN
				  ,SH.PLAN_HI
				  ,SH.CYOKU_K
				  ,SH.JININ_SU
				  ,SH.JITU_LOT
				ORDER BY LM.SYOZK_CD , SH.LINE_CD ASC";

			$dir_date = date('Ym');
			if ( !is_dir("G:/vbs_prod_v1/work/Query/" . $dir_date) ) 
			{
			    mkdir( "G:/vbs_prod_v1/work/Query/" . $dir_date );
			   // echo "The directory $dir_date was successfully created.";
			} else
			{
			   // echo "The directory $dir_date exists.";
			}
				
			//exit;


	$myfile = fopen("G:/vbs_prod_v1/work/Query/" . $dir_date . "/prod_".$date_to.".sql", 'w') or die("Unable to open file!");		
    fwrite($myfile,$sql);
    fclose($myfile);				
	}



       // output_file($filename); 
    



function output_file($namefile){
        //$namefile = "Query_pods_data.sql";
        $file = $namefile; 
        //echo basename($file); exit;
        header("Content-Description: File Transfer"); 
        header("Content-Type: application/octet-stream"); 
        header("Content-Disposition: attachment; filename=".basename($file) ); 
        readfile ($file);
        exit;
}

?>

 
