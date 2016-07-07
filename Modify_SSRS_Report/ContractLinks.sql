
;WITH CTE AS 
( 
	SELECT 
		 * 
		,CAST(ZOCO_CO_UID AS varchar(36)) AS in_contract_uid 
		,CAST(ZOCO_PR_UID AS varchar(36)) AS in_premise_uid 
	FROM T_ZO_Contract_Object 
	WHERE ZOCO_PR_UID IS NOT NULL 
	AND ZOCO_Status = 1 
) 
SELECT 
	'https://www7.cor-asp.ch/SwissRe_Productiv_Dev/test.ashx?in_contract_uid=' + in_contract_uid + '&in_premise_uid=' + in_premise_uid 
FROM CTE 
