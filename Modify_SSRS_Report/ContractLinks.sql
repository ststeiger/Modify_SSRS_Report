
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
	CASE 
		WHEN SUSER_SNAME() LIKE 'COR[\\]%' 
		THEN 
			(
				SELECT TOP 1 
					REPLACE
					(
						REPLACE
						(
							REPLACE(FC_Value + '/', '//', '/') 
							,'_Saml'
							,''
						)
						,'_Portal'
						,'_Dev'
					) AS ExternalPortalLink 
				FROM T_FMS_Configuration 
				WHERE FC_Key = 'portalLink' 
			) 
		ELSE (
				SELECT TOP 1 
					REPLACE(FC_Value + '/', '//', '/') AS InternalPortalLink 
				FROM T_FMS_Configuration 
				WHERE FC_Key = 'portalLink' 
			 )
	END 
	+ 'test.ashx?in_contract_uid=' + in_contract_uid + '&in_premise_uid=' + in_premise_uid AS TestLink 
FROM CTE 

-- SELECT * FROM T_FMS_Configuration 


SELECT 
	REPLACE(FC_Value + '/', '//', '/') 
	+ 'COR/devtool.aspx' 
	AS BasicLink  
FROM T_FMS_Configuration 
WHERE FC_Key = 'basicLink' 
