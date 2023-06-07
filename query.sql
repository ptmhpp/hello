SELECT cctr.PNR_NBR ,cctr.PNR_CREATE_DT ,cctr.CREATE_DATE ,cctr.TCKT_NBR ,
concat(FRST_NM,concat(concat(MDL_NM,' '),concat(concat(MDL_INIT_NM,' '),concat(LST_NM,' ')))) as Name,
cce.EMAIL_ADDR ,ccp2.PHN_NBR ,
concat(cca.LN_1_ADDR,concat(concat(cca.LN_2_ADDR,' '),concat(concat(cca.LN_3_ADDR,' '),concat(cca.CITY_NM ,' ')))) as address,
--cca2.AGR_NBR ,
ccp2.LST_UPDT_CHNL_ID 
FROM ors_customer.C_CUST_TCKT_RLTN cctr
left JOIN ors_customer.C_CUST_PROFILE ccp ON cctr.CUST_ROW_ID =ccp.ROWID_OBJECT 
left JOIN ORS_CUSTOMER.C_CUST_EMAIL cce ON cce.CUST_ROW_ID =ccp.ROWID_OBJECT 
left JOIN ORS_CUSTOMER.C_CUST_PHONE ccp2 ON ccp2.CUST_ROW_ID =ccp.ROWID_OBJECT 
left JOIN ORS_CUSTOMER.C_CUST_ADDRESS cca ON cca.CUST_ROW_ID =ccp.ROWID_OBJECT
--LEFT JOIN ORS_CUSTOMER.C_CUST_AGR cca2 ON cca2.CUST_ROW_ID =ccp.ROWID_OBJECT 
WHERE cctr.PNR_NBR = '4E8227' 
AND cctr.PNR_CREATE_DT> TO_DATE('2023-04-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS')-- AND ccp2.LAST_ROWID_SYSTEM='ARROW'
ORDER BY cctr.LAST_UPDATE_DATE DESC;