CREATE OR REPLACE MACRO fn_unaccent(s) AS (
    translate(
        replace(
            replace(
                replace(
                    replace(
                        replace(
                            replace(
                                replace(s, 'Æ', 'AE')
                              , 'æ', 'ae'
                            )
                          , 'Œ', 'OE'
                        )
                      , 'œ', 'oe'
                    )
                  , 'ß', 'ss'
                )
              , 'Þ', 'TH'
            )
          , 'þ', 'th'
        )
      , 'ÁÀÂÃÄÅĀĂĄáàâãäåāăąÇĆČçćčÐĎĐðďđÉÈÊËĒĔĖĘĚéèêëēĕėęěÍÌÎÏĪĬİĮíìîïīĭıįŁĽĹĻłľĺļÑŃŇñńňÓÒÔÕÖØŌŎŐóòôõöøōŏőŔŘŕřŚŠŞśšşÚÙÛÜŪŬŮŰŲúùûüūŭůűųÝŸýÿŽŹŻžźż'
      , 'AAAAAAAAAaaaaaaaaaCCCcccDDDdddEEEEEEEEEeeeeeeeeeIIIIIIIIiiiiiiiiLLLLllllNNNnnnOOOOOOOOOoooooooooRRrrSSSsssUUUUUUUUUuuuuuuuuuYYyyZZZzzz'
    )
)
;

CREATE OR REPLACE MACRO unaccent(s) AS (
    fn_unaccent(s)
)
;
