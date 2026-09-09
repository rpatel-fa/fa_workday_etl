DROP TABLE IF EXISTS ref_post_code_district
;
CREATE TABLE ref_post_code_district (
    postcode           VARCHAR,
    latitude           VARCHAR,
    longitude          VARCHAR,
    easting            VARCHAR,
    northing           VARCHAR,
    grid_reference     VARCHAR,
    town_area          VARCHAR,
    region             VARCHAR,
    postcodes          VARCHAR,
    active_postcodes   VARCHAR,
    population         VARCHAR,
    households         VARCHAR,
    nearby_districts   VARCHAR,
    uk_region          VARCHAR,
    post_town          VARCHAR
);
COPY ref_post_code_district FROM 'config/data/ref_post_code_district.csv' (HEADER TRUE, DELIMITER ',')
;
