
SELECT            
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
      
		,DC.DM_01_1 01st
        ,DC.DM_02_1 02nd
        ,DC.DM_03_1 03rd
        ,DC.DM_04_1 04th
        ,DC.DM_05_1 05th
        ,DC.DM_06_1 06th
        ,DC.DM_07_1 07th
        ,DC.DM_08_1 08th
        ,DC.DM_09_1 09th
        ,DC.DM_10_1 10th
        ,DC.DM_11_1 11th
        ,DC.DM_12_1 12th
        ,DC.DM_13_1 13th
        ,DC.DM_14_1 14th
        ,DC.DM_15_1 15th
        ,DC.DM_16_1 16th
        ,DC.DM_17_1 17th
        ,DC.DM_18_1 18th
        ,DC.DM_19_1 19th
        ,DC.DM_20_1 20th
        ,DC.DM_21_1 21st
        ,DC.DM_22_1 22nd
        ,DC.DM_23_1 23rd
        ,DC.DM_24_1 24th
        ,DC.DM_25_1 25th
        ,DC.DM_26_1 26th
        ,DC.DM_27_1 27th
        ,DC.DM_28_1 28th
        ,DC.DM_29_1 29th
        ,DC.DM_30_1 30th
        ,DC.DM_31_1 31st
        
		,DC.DM_01_2 01st
        ,DC.DM_02_2 02nd
        ,DC.DM_03_2 03rd
        ,DC.DM_04_2 04th
        ,DC.DM_05_2 05th
        ,DC.DM_06_2 06th
        ,DC.DM_07_2 07th
        ,DC.DM_08_2 08th
        ,DC.DM_09_2 09th
        ,DC.DM_10_2 10th
        ,DC.DM_11_2 11th
        ,DC.DM_12_2 12th
        ,DC.DM_13_2 13th
        ,DC.DM_14_2 14th
        ,DC.DM_15_2 15th
        ,DC.DM_16_2 16th
        ,DC.DM_17_2 17th
        ,DC.DM_18_2 18th
        ,DC.DM_19_2 19th
        ,DC.DM_20_2 20th
        ,DC.DM_21_2 21st
        ,DC.DM_22_2 22nd
        ,DC.DM_23_2 23rd
        ,DC.DM_24_2 24th
        ,DC.DM_25_2 25th
        ,DC.DM_26_2 26th
        ,DC.DM_27_2 27th
        ,DC.DM_28_2 28th
        ,DC.DM_29_2 29th
        ,DC.DM_30_2 30th
        
		,DC.DM_01_3 01st
        ,DC.DM_02_3 02nd
        ,DC.DM_03_3 03rd
        ,DC.DM_04_3 04th
        ,DC.DM_05_3 05th
        ,DC.DM_06_3 06th
        ,DC.DM_07_3 07th
        ,DC.DM_08_3 08th
        ,DC.DM_09_3 09th
        ,DC.DM_10_3 10th
        ,DC.DM_11_3 11th
        ,DC.DM_12_3 12th
        ,DC.DM_13_3 13th
        ,DC.DM_14_3 14th
        ,DC.DM_15_3 15th
        ,DC.DM_16_3 16th
        ,DC.DM_17_3 17th
        ,DC.DM_18_3 18th
        ,DC.DM_19_3 19th
        ,DC.DM_20_3 20th
        ,DC.DM_21_3 21st
        ,DC.DM_22_3 22nd
        ,DC.DM_23_3 23rd
        ,DC.DM_24_3 24th
        ,DC.DM_25_3 25th
        ,DC.DM_26_3 26th
        ,DC.DM_27_3 27th
        ,DC.DM_28_3 28th
        ,DC.DM_29_3 29th
        ,DC.DM_30_3 30th
        ,DC.DM_31_3 31st
        
		
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        SELECT            
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
      
		,DC.PL_01_1 01st
        ,DC.PL_02_1 02nd
        ,DC.PL_03_1 03rd
        ,DC.PL_04_1 04th
        ,DC.PL_05_1 05th
        ,DC.PL_06_1 06th
        ,DC.PL_07_1 07th
        ,DC.PL_08_1 08th
        ,DC.PL_09_1 09th
        ,DC.PL_10_1 10th
        ,DC.PL_11_1 11th
        ,DC.PL_12_1 12th
        ,DC.PL_13_1 13th
        ,DC.PL_14_1 14th
        ,DC.PL_15_1 15th
        ,DC.PL_16_1 16th
        ,DC.PL_17_1 17th
        ,DC.PL_18_1 18th
        ,DC.PL_19_1 19th
        ,DC.PL_20_1 20th
        ,DC.PL_21_1 21st
        ,DC.PL_22_1 22nd
        ,DC.PL_23_1 23rd
        ,DC.PL_24_1 24th
        ,DC.PL_25_1 25th
        ,DC.PL_26_1 26th
        ,DC.PL_27_1 27th
        ,DC.PL_28_1 28th
        ,DC.PL_29_1 29th
        ,DC.PL_30_1 30th
        ,DC.PL_31_1 31st
        
		,DC.PL_01_2 01st
        ,DC.PL_02_2 02nd
        ,DC.PL_03_2 03rd
        ,DC.PL_04_2 04th
        ,DC.PL_05_2 05th
        ,DC.PL_06_2 06th
        ,DC.PL_07_2 07th
        ,DC.PL_08_2 08th
        ,DC.PL_09_2 09th
        ,DC.PL_10_2 10th
        ,DC.PL_11_2 11th
        ,DC.PL_12_2 12th
        ,DC.PL_13_2 13th
        ,DC.PL_14_2 14th
        ,DC.PL_15_2 15th
        ,DC.PL_16_2 16th
        ,DC.PL_17_2 17th
        ,DC.PL_18_2 18th
        ,DC.PL_19_2 19th
        ,DC.PL_20_2 20th
        ,DC.PL_21_2 21st
        ,DC.PL_22_2 22nd
        ,DC.PL_23_2 23rd
        ,DC.PL_24_2 24th
        ,DC.PL_25_2 25th
        ,DC.PL_26_2 26th
        ,DC.PL_27_2 27th
        ,DC.PL_28_2 28th
        ,DC.PL_29_2 29th
        ,DC.PL_30_2 30th
        
		,DC.PL_01_3 01st
        ,DC.PL_02_3 02nd
        ,DC.PL_03_3 03rd
        ,DC.PL_04_3 04th
        ,DC.PL_05_3 05th
        ,DC.PL_06_3 06th
        ,DC.PL_07_3 07th
        ,DC.PL_08_3 08th
        ,DC.PL_09_3 09th
        ,DC.PL_10_3 10th
        ,DC.PL_11_3 11th
        ,DC.PL_12_3 12th
        ,DC.PL_13_3 13th
        ,DC.PL_14_3 14th
        ,DC.PL_15_3 15th
        ,DC.PL_16_3 16th
        ,DC.PL_17_3 17th
        ,DC.PL_18_3 18th
        ,DC.PL_19_3 19th
        ,DC.PL_20_3 20th
        ,DC.PL_21_3 21st
        ,DC.PL_22_3 22nd
        ,DC.PL_23_3 23rd
        ,DC.PL_24_3 24th
        ,DC.PL_25_3 25th
        ,DC.PL_26_3 26th
        ,DC.PL_27_3 27th
        ,DC.PL_28_3 28th
        ,DC.PL_29_3 29th
        ,DC.PL_30_3 30th
        ,DC.PL_31_3 31st
        
		
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        SELECT            
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
      
		,DC.AC_01_1 01st
        ,DC.AC_02_1 02nd
        ,DC.AC_03_1 03rd
        ,DC.AC_04_1 04th
        ,DC.AC_05_1 05th
        ,DC.AC_06_1 06th
        ,DC.AC_07_1 07th
        ,DC.AC_08_1 08th
        ,DC.AC_09_1 09th
        ,DC.AC_10_1 10th
        ,DC.AC_11_1 11th
        ,DC.AC_12_1 12th
        ,DC.AC_13_1 13th
        ,DC.AC_14_1 14th
        ,DC.AC_15_1 15th
        ,DC.AC_16_1 16th
        ,DC.AC_17_1 17th
        ,DC.AC_18_1 18th
        ,DC.AC_19_1 19th
        ,DC.AC_20_1 20th
        ,DC.AC_21_1 21st
        ,DC.AC_22_1 22nd
        ,DC.AC_23_1 23rd
        ,DC.AC_24_1 24th
        ,DC.AC_25_1 25th
        ,DC.AC_26_1 26th
        ,DC.AC_27_1 27th
        ,DC.AC_28_1 28th
        ,DC.AC_29_1 29th
        ,DC.AC_30_1 30th
        ,DC.AC_31_1 31st
        
		,DC.AC_01_2 01st
        ,DC.AC_02_2 02nd
        ,DC.AC_03_2 03rd
        ,DC.AC_04_2 04th
        ,DC.AC_05_2 05th
        ,DC.AC_06_2 06th
        ,DC.AC_07_2 07th
        ,DC.AC_08_2 08th
        ,DC.AC_09_2 09th
        ,DC.AC_10_2 10th
        ,DC.AC_11_2 11th
        ,DC.AC_12_2 12th
        ,DC.AC_13_2 13th
        ,DC.AC_14_2 14th
        ,DC.AC_15_2 15th
        ,DC.AC_16_2 16th
        ,DC.AC_17_2 17th
        ,DC.AC_18_2 18th
        ,DC.AC_19_2 19th
        ,DC.AC_20_2 20th
        ,DC.AC_21_2 21st
        ,DC.AC_22_2 22nd
        ,DC.AC_23_2 23rd
        ,DC.AC_24_2 24th
        ,DC.AC_25_2 25th
        ,DC.AC_26_2 26th
        ,DC.AC_27_2 27th
        ,DC.AC_28_2 28th
        ,DC.AC_29_2 29th
        ,DC.AC_30_2 30th
        
		,DC.AC_01_3 01st
        ,DC.AC_02_3 02nd
        ,DC.AC_03_3 03rd
        ,DC.AC_04_3 04th
        ,DC.AC_05_3 05th
        ,DC.AC_06_3 06th
        ,DC.AC_07_3 07th
        ,DC.AC_08_3 08th
        ,DC.AC_09_3 09th
        ,DC.AC_10_3 10th
        ,DC.AC_11_3 11th
        ,DC.AC_12_3 12th
        ,DC.AC_13_3 13th
        ,DC.AC_14_3 14th
        ,DC.AC_15_3 15th
        ,DC.AC_16_3 16th
        ,DC.AC_17_3 17th
        ,DC.AC_18_3 18th
        ,DC.AC_19_3 19th
        ,DC.AC_20_3 20th
        ,DC.AC_21_3 21st
        ,DC.AC_22_3 22nd
        ,DC.AC_23_3 23rd
        ,DC.AC_24_3 24th
        ,DC.AC_25_3 25th
        ,DC.AC_26_3 26th
        ,DC.AC_27_3 27th
        ,DC.AC_28_3 28th
        ,DC.AC_29_3 29th
        ,DC.AC_30_3 30th
        ,DC.AC_31_3 31st
        
		
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        SELECT            
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
      
		,DC.AC_01_1 - DC.PL_01_1 01st
        ,DC.AC_02_1 - DC.PL_02_1 02nd
        ,DC.AC_03_1 - DC.PL_03_1 03rd
        ,DC.AC_04_1 - DC.PL_04_1 04th
        ,DC.AC_05_1 - DC.PL_05_1 05th
        ,DC.AC_06_1 - DC.PL_06_1 06th
        ,DC.AC_07_1 - DC.PL_07_1 07th
        ,DC.AC_08_1 - DC.PL_08_1 08th
        ,DC.AC_09_1 - DC.PL_09_1 09th
        ,DC.AC_10_1 - DC.PL_10_1 10th
        ,DC.AC_11_1 - DC.PL_11_1 11th
        ,DC.AC_12_1 - DC.PL_12_1 12th
        ,DC.AC_13_1 - DC.PL_13_1 13th
        ,DC.AC_14_1 - DC.PL_14_1 14th
        ,DC.AC_15_1 - DC.PL_15_1 15th
        ,DC.AC_16_1 - DC.PL_16_1 16th
        ,DC.AC_17_1 - DC.PL_17_1 17th
        ,DC.AC_18_1 - DC.PL_18_1 18th
        ,DC.AC_19_1 - DC.PL_19_1 19th
        ,DC.AC_20_1 - DC.PL_20_1 20th
        ,DC.AC_21_1 - DC.PL_21_1 21st
        ,DC.AC_22_1 - DC.PL_22_1 22nd
        ,DC.AC_23_1 - DC.PL_23_1 23rd
        ,DC.AC_24_1 - DC.PL_24_1 24th
        ,DC.AC_25_1 - DC.PL_25_1 25th
        ,DC.AC_26_1 - DC.PL_26_1 26th
        ,DC.AC_27_1 - DC.PL_27_1 27th
        ,DC.AC_28_1 - DC.PL_28_1 28th
        ,DC.AC_29_1 - DC.PL_29_1 29th
        ,DC.AC_30_1 - DC.PL_30_1 30th
        ,DC.AC_31_1 - DC.PL_31_1 31st
        
		,DC.AC_01_2 - DC.PL_01_2 01st
        ,DC.AC_02_2 - DC.PL_02_2 02nd
        ,DC.AC_03_2 - DC.PL_03_2 03rd
        ,DC.AC_04_2 - DC.PL_04_2 04th
        ,DC.AC_05_2 - DC.PL_05_2 05th
        ,DC.AC_06_2 - DC.PL_06_2 06th
        ,DC.AC_07_2 - DC.PL_07_2 07th
        ,DC.AC_08_2 - DC.PL_08_2 08th
        ,DC.AC_09_2 - DC.PL_09_2 09th
        ,DC.AC_10_2 - DC.PL_10_2 10th
        ,DC.AC_11_2 - DC.PL_11_2 11th
        ,DC.AC_12_2 - DC.PL_12_2 12th
        ,DC.AC_13_2 - DC.PL_13_2 13th
        ,DC.AC_14_2 - DC.PL_14_2 14th
        ,DC.AC_15_2 - DC.PL_15_2 15th
        ,DC.AC_16_2 - DC.PL_16_2 16th
        ,DC.AC_17_2 - DC.PL_17_2 17th
        ,DC.AC_18_2 - DC.PL_18_2 18th
        ,DC.AC_19_2 - DC.PL_19_2 19th
        ,DC.AC_20_2 - DC.PL_20_2 20th
        ,DC.AC_21_2 - DC.PL_21_2 21st
        ,DC.AC_22_2 - DC.PL_22_2 22nd
        ,DC.AC_23_2 - DC.PL_23_2 23rd
        ,DC.AC_24_2 - DC.PL_24_2 24th
        ,DC.AC_25_2 - DC.PL_25_2 25th
        ,DC.AC_26_2 - DC.PL_26_2 26th
        ,DC.AC_27_2 - DC.PL_27_2 27th
        ,DC.AC_28_2 - DC.PL_28_2 28th
        ,DC.AC_29_2 - DC.PL_29_2 29th
        ,DC.AC_30_2 - DC.PL_30_2 30th
        
		,DC.AC_01_3 - DC.PL_01_3 01st
        ,DC.AC_02_3 - DC.PL_02_3 02nd
        ,DC.AC_03_3 - DC.PL_03_3 03rd
        ,DC.AC_04_3 - DC.PL_04_3 04th
        ,DC.AC_05_3 - DC.PL_05_3 05th
        ,DC.AC_06_3 - DC.PL_06_3 06th
        ,DC.AC_07_3 - DC.PL_07_3 07th
        ,DC.AC_08_3 - DC.PL_08_3 08th
        ,DC.AC_09_3 - DC.PL_09_3 09th
        ,DC.AC_10_3 - DC.PL_10_3 10th
        ,DC.AC_11_3 - DC.PL_11_3 11th
        ,DC.AC_12_3 - DC.PL_12_3 12th
        ,DC.AC_13_3 - DC.PL_13_3 13th
        ,DC.AC_14_3 - DC.PL_14_3 14th
        ,DC.AC_15_3 - DC.PL_15_3 15th
        ,DC.AC_16_3 - DC.PL_16_3 16th
        ,DC.AC_17_3 - DC.PL_17_3 17th
        ,DC.AC_18_3 - DC.PL_18_3 18th
        ,DC.AC_19_3 - DC.PL_19_3 19th
        ,DC.AC_20_3 - DC.PL_20_3 20th
        ,DC.AC_21_3 - DC.PL_21_3 21st
        ,DC.AC_22_3 - DC.PL_22_3 22nd
        ,DC.AC_23_3 - DC.PL_23_3 23rd
        ,DC.AC_24_3 - DC.PL_24_3 24th
        ,DC.AC_25_3 - DC.PL_25_3 25th
        ,DC.AC_26_3 - DC.PL_26_3 26th
        ,DC.AC_27_3 - DC.PL_27_3 27th
        ,DC.AC_28_3 - DC.PL_28_3 28th
        ,DC.AC_29_3 - DC.PL_29_3 29th
        ,DC.AC_30_3 - DC.PL_30_3 30th
        ,DC.AC_31_3 - DC.PL_31_3 31st
        
		
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        SELECT            
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
      
		,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '02' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 02nd
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '03' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 03rd
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '04' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 04th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '05' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 05th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '06' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 06th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '07' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 07th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '08' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 08th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '09' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 09th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '10' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 10th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '11' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 11th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '12' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 12th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '13' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 13th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '14' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 14th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '15' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 15th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '16' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 16th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '17' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 17th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '18' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 18th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '19' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 19th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '20' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 20th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '21' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 21st
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '22' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 22nd
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '23' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 23rd
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '24' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 24th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '25' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 25th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '26' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 26th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '27' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 27th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '28' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 28th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '29' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 29th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '30' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 30th
        ,CASE WHEN DATE_FORMAT( CURDATE() - INTERVAL 1 DAY , '%d') = '31' AND  DATE_FORMAT( CURDATE(), '%d') != '01' THEN  DS.ON_HAND ELSE 0 END 31st
        
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        ,0 31st
        
		
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN 
        ( SELECT * FROM DEMAND_STOCK  WHERE WH_CD IN ('K1MX', 'K2MX') ) DS
        ON DC.ITEM_CD = DS.ITEM_CD AND DC.PLANT_CD = DS.PLANT_CD

        UNION ALL

        SELECT            
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
      
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        ,0 31st
        
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        ,0 31st
        
		
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        SELECT            
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
      
		,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_01_1 / DC.SNP) END 01st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_02_1 / DC.SNP) END 02nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_03_1 / DC.SNP) END 03rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_04_1 / DC.SNP) END 04th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_05_1 / DC.SNP) END 05th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_06_1 / DC.SNP) END 06th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_07_1 / DC.SNP) END 07th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_08_1 / DC.SNP) END 08th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_09_1 / DC.SNP) END 09th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_10_1 / DC.SNP) END 10th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_11_1 / DC.SNP) END 11th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_12_1 / DC.SNP) END 12th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_13_1 / DC.SNP) END 13th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_14_1 / DC.SNP) END 14th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_15_1 / DC.SNP) END 15th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_16_1 / DC.SNP) END 16th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_17_1 / DC.SNP) END 17th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_18_1 / DC.SNP) END 18th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_19_1 / DC.SNP) END 19th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_20_1 / DC.SNP) END 20th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_21_1 / DC.SNP) END 21st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_22_1 / DC.SNP) END 22nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_23_1 / DC.SNP) END 23rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_24_1 / DC.SNP) END 24th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_25_1 / DC.SNP) END 25th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_26_1 / DC.SNP) END 26th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_27_1 / DC.SNP) END 27th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_28_1 / DC.SNP) END 28th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_29_1 / DC.SNP) END 29th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_30_1 / DC.SNP) END 30th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_31_1 / DC.SNP) END 31st
        
		,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_01_2 / DC.SNP) END 01st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_02_2 / DC.SNP) END 02nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_03_2 / DC.SNP) END 03rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_04_2 / DC.SNP) END 04th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_05_2 / DC.SNP) END 05th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_06_2 / DC.SNP) END 06th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_07_2 / DC.SNP) END 07th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_08_2 / DC.SNP) END 08th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_09_2 / DC.SNP) END 09th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_10_2 / DC.SNP) END 10th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_11_2 / DC.SNP) END 11th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_12_2 / DC.SNP) END 12th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_13_2 / DC.SNP) END 13th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_14_2 / DC.SNP) END 14th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_15_2 / DC.SNP) END 15th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_16_2 / DC.SNP) END 16th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_17_2 / DC.SNP) END 17th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_18_2 / DC.SNP) END 18th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_19_2 / DC.SNP) END 19th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_20_2 / DC.SNP) END 20th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_21_2 / DC.SNP) END 21st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_22_2 / DC.SNP) END 22nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_23_2 / DC.SNP) END 23rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_24_2 / DC.SNP) END 24th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_25_2 / DC.SNP) END 25th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_26_2 / DC.SNP) END 26th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_27_2 / DC.SNP) END 27th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_28_2 / DC.SNP) END 28th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_29_2 / DC.SNP) END 29th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_30_2 / DC.SNP) END 30th
        
		,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_01_3 / DC.SNP) END 01st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_02_3 / DC.SNP) END 02nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_03_3 / DC.SNP) END 03rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_04_3 / DC.SNP) END 04th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_05_3 / DC.SNP) END 05th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_06_3 / DC.SNP) END 06th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_07_3 / DC.SNP) END 07th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_08_3 / DC.SNP) END 08th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_09_3 / DC.SNP) END 09th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_10_3 / DC.SNP) END 10th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_11_3 / DC.SNP) END 11th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_12_3 / DC.SNP) END 12th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_13_3 / DC.SNP) END 13th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_14_3 / DC.SNP) END 14th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_15_3 / DC.SNP) END 15th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_16_3 / DC.SNP) END 16th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_17_3 / DC.SNP) END 17th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_18_3 / DC.SNP) END 18th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_19_3 / DC.SNP) END 19th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_20_3 / DC.SNP) END 20th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_21_3 / DC.SNP) END 21st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_22_3 / DC.SNP) END 22nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_23_3 / DC.SNP) END 23rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_24_3 / DC.SNP) END 24th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_25_3 / DC.SNP) END 25th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_26_3 / DC.SNP) END 26th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_27_3 / DC.SNP) END 27th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_28_3 / DC.SNP) END 28th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_29_3 / DC.SNP) END 29th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_30_3 / DC.SNP) END 30th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.DM_31_3 / DC.SNP) END 31st
        
		
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        SELECT            
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
      
		,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_01_1 / DC.SNP) END 01st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_02_1 / DC.SNP) END 02nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_03_1 / DC.SNP) END 03rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_04_1 / DC.SNP) END 04th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_05_1 / DC.SNP) END 05th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_06_1 / DC.SNP) END 06th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_07_1 / DC.SNP) END 07th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_08_1 / DC.SNP) END 08th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_09_1 / DC.SNP) END 09th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_10_1 / DC.SNP) END 10th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_11_1 / DC.SNP) END 11th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_12_1 / DC.SNP) END 12th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_13_1 / DC.SNP) END 13th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_14_1 / DC.SNP) END 14th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_15_1 / DC.SNP) END 15th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_16_1 / DC.SNP) END 16th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_17_1 / DC.SNP) END 17th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_18_1 / DC.SNP) END 18th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_19_1 / DC.SNP) END 19th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_20_1 / DC.SNP) END 20th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_21_1 / DC.SNP) END 21st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_22_1 / DC.SNP) END 22nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_23_1 / DC.SNP) END 23rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_24_1 / DC.SNP) END 24th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_25_1 / DC.SNP) END 25th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_26_1 / DC.SNP) END 26th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_27_1 / DC.SNP) END 27th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_28_1 / DC.SNP) END 28th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_29_1 / DC.SNP) END 29th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_30_1 / DC.SNP) END 30th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_31_1 / DC.SNP) END 31st
        
		,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_01_2 / DC.SNP) END 01st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_02_2 / DC.SNP) END 02nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_03_2 / DC.SNP) END 03rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_04_2 / DC.SNP) END 04th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_05_2 / DC.SNP) END 05th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_06_2 / DC.SNP) END 06th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_07_2 / DC.SNP) END 07th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_08_2 / DC.SNP) END 08th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_09_2 / DC.SNP) END 09th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_10_2 / DC.SNP) END 10th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_11_2 / DC.SNP) END 11th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_12_2 / DC.SNP) END 12th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_13_2 / DC.SNP) END 13th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_14_2 / DC.SNP) END 14th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_15_2 / DC.SNP) END 15th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_16_2 / DC.SNP) END 16th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_17_2 / DC.SNP) END 17th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_18_2 / DC.SNP) END 18th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_19_2 / DC.SNP) END 19th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_20_2 / DC.SNP) END 20th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_21_2 / DC.SNP) END 21st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_22_2 / DC.SNP) END 22nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_23_2 / DC.SNP) END 23rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_24_2 / DC.SNP) END 24th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_25_2 / DC.SNP) END 25th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_26_2 / DC.SNP) END 26th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_27_2 / DC.SNP) END 27th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_28_2 / DC.SNP) END 28th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_29_2 / DC.SNP) END 29th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_30_2 / DC.SNP) END 30th
        
		,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_01_3 / DC.SNP) END 01st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_02_3 / DC.SNP) END 02nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_03_3 / DC.SNP) END 03rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_04_3 / DC.SNP) END 04th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_05_3 / DC.SNP) END 05th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_06_3 / DC.SNP) END 06th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_07_3 / DC.SNP) END 07th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_08_3 / DC.SNP) END 08th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_09_3 / DC.SNP) END 09th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_10_3 / DC.SNP) END 10th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_11_3 / DC.SNP) END 11th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_12_3 / DC.SNP) END 12th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_13_3 / DC.SNP) END 13th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_14_3 / DC.SNP) END 14th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_15_3 / DC.SNP) END 15th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_16_3 / DC.SNP) END 16th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_17_3 / DC.SNP) END 17th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_18_3 / DC.SNP) END 18th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_19_3 / DC.SNP) END 19th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_20_3 / DC.SNP) END 20th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_21_3 / DC.SNP) END 21st
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_22_3 / DC.SNP) END 22nd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_23_3 / DC.SNP) END 23rd
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_24_3 / DC.SNP) END 24th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_25_3 / DC.SNP) END 25th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_26_3 / DC.SNP) END 26th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_27_3 / DC.SNP) END 27th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_28_3 / DC.SNP) END 28th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_29_3 / DC.SNP) END 29th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_30_3 / DC.SNP) END 30th
        ,CASE WHEN DC.SNP > 10000 THEN NULL ELSE CEIL(DC.PL_31_3 / DC.SNP) END 31st
        
		
        FROM 

        TEMP_DEMAND_CONVERT DC
        LEFT OUTER JOIN
        DEMAND_CONVERT_LM DL
        ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD

        UNION ALL

        SELECT            
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
      
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        ,0 31st
        
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        ,0 31st
        
		
        FROM 

        TEMP_DEMAND_CONVERT DC

        UNION ALL

        SELECT            
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
      
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        ,0 31st
        
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        
		,0 01st
        ,0 02nd
        ,0 03rd
        ,0 04th
        ,0 05th
        ,0 06th
        ,0 07th
        ,0 08th
        ,0 09th
        ,0 10th
        ,0 11th
        ,0 12th
        ,0 13th
        ,0 14th
        ,0 15th
        ,0 16th
        ,0 17th
        ,0 18th
        ,0 19th
        ,0 20th
        ,0 21st
        ,0 22nd
        ,0 23rd
        ,0 24th
        ,0 25th
        ,0 26th
        ,0 27th
        ,0 28th
        ,0 29th
        ,0 30th
        ,0 31st
        
		
        FROM 

        TEMP_DEMAND_CONVERT DC

        ORDER BY 1,2,4,8

        