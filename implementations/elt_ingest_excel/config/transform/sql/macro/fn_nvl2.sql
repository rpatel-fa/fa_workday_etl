CREATE OR REPLACE MACRO fn_nvl2(expr, value_if_null, value_if_not_null) AS (
    CASE
        WHEN NULLIF(TRIM(expr), '') IS NULL
        THEN value_if_null
        ELSE value_if_not_null
    END
)
;
