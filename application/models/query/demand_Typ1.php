

	
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

		
			,DC.DM_01 01st
			,DC.DM_02 02nd
			,DC.DM_03 03rd
			,DC.DM_04 04th
			,DC.DM_05 05th
			,DC.DM_06 06th
			,DC.DM_07 07th
			,DC.DM_08 08th
			,DC.DM_09 09th
			,DC.DM_10 10th
			,DC.DM_11 11th
			,DC.DM_12 12th
			,DC.DM_13 13th
			,DC.DM_14 14th
			,DC.DM_15 15th
			,DC.DM_16 16th
			,DC.DM_17 17th
			,DC.DM_18 18th
			,DC.DM_19 19th
			,DC.DM_20 20th
			,DC.DM_21 21st
			,DC.DM_22 22nd
			,DC.DM_23 23rd
			,DC.DM_24 24th
			,DC.DM_25 25th
			,DC.DM_26 26th
			,DC.DM_27 27th
			,DC.DM_28 28th
			,DC.DM_29 29th
			,DC.DM_30 30th
			,DC.DM_31 31st


			FROM 

			DEMAND_CONVERT DC

			LEFT OUTER JOIN

			DEMAND_CONVERT_LM DL

			ON DC.ITEM_CD = DL.ITEM_CD AND DC.SOURCE_CD = DL.SOURCE_CD