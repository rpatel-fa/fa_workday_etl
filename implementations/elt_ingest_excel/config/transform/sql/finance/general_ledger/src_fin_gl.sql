DROP TABLE IF EXISTS src_fin_gl
;
CREATE TABLE src_fin_gl
    AS
  WITH cte_base
    AS (    
SELECT t.*
  FROM src_fin_gl_base                 t
       )
SELECT 
       ROW_NUMBER() OVER (
         ORDER BY 
               t.business_unit
             , t.account
                         )             row_id
     , t.*
  FROM cte_base                        t
;
