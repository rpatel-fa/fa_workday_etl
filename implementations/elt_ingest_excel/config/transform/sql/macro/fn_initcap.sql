CREATE OR REPLACE MACRO fn_initcap(s) AS (
    CASE
        WHEN s IS NULL OR NULLIF(TRIM(s), '') IS NULL
        THEN NULL
        ELSE list_aggregate(
           list_transform(
               regexp_split_to_array(lower(TRIM(s)), '\s+')
             , w -> upper(w[1]) || lower(substr(w, 2))
           )
         , 'string_agg'
         , ' '
        )
    END
)
;