WITH ActivePlatIDs AS (
    SELECT platform_id
    FROM "m-datawarehouse"."current_companies_expanded"
    WHERE isactive
      AND NOT isdemo
      AND platform_id IN (
          SELECT platform_clientid
          FROM "m-datawarehouse"."current_company_configuration"
          WHERE isactive
            AND LOWER(owner_ultimate_parent) LIKE '%united%health'
      )
),
RankedRecords AS (
    SELECT 
        v.*, 
        ROW_NUMBER() OVER (PARTITION BY v.customer_id ORDER BY v.process_date DESC) AS rn
    FROM (
        SELECT DISTINCT 
            product,
            platform,
            companyid,
            CONCAT(CAST(platform AS VARCHAR), '-', CAST(companyid AS VARCHAR)) AS PLAT_ID
        FROM "m-datawarehouse"."Transactions"
        WHERE product = 'Login'
          AND CONCAT(CAST(platform AS VARCHAR), '-', CAST(companyid AS VARCHAR)) IN (
              SELECT platform_id FROM ActivePlatIDs
          )
    ) e
    JOIN "unification"."verified_carrier_data" v 
        ON e.Plat_ID = v.carrier
    WHERE v.carrier = 'UHC' 
      AND v.funding_type = 'LF'
)
SELECT *
FROM RankedRecords
WHERE rn = 1
ORDER BY process_date DESC;

