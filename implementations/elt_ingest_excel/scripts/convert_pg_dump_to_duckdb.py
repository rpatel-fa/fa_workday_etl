
import os
import re

input_file = '.tmp/postcodesio-2025-12-18-1017.sql'
# Use absolute path for output dir to ensure DuckDB can find CSVs regardless of CWD
output_dir = os.path.abspath('elt_ingest_excel/config/transform/sql/ref/postcodes/data')
output_sql = 'elt_ingest_excel/config/transform/sql/ref/postcodes/load_postcodes_minimal.duckdb.sql'

# Relevant tables for Postcode -> County/City/Town/District mapping
# 'postcodes' contains the main data and foreign keys (admin_county_id, admin_district_id, etc.)
# 'counties', 'districts', 'wards' map the IDs to names.
# 'places' might contain town/city names for outcodes or specific places.
relevant_tables = {
    'postcodes',
    'counties',
    'districts',
    'wards',
    # 'places', # Usually places is huge and might not be strictly needed for basic postcode->admin mapping, but useful for town names. Keeping it for now.
    'outcodes' # Useful for broad area mapping
}

os.makedirs(output_dir, exist_ok=True)

with open(input_file, 'r', encoding='utf-8') as f, open(output_sql, 'w', encoding='utf-8') as sql_f:
    sql_f.write("-- DuckDB Load Script (Minimal - Postcode Mappings Only)\n")
    sql_f.write("INSTALL spatial; LOAD spatial;\n\n")

    # Add idempotent DROP TABLE logic
    def write_drop_if_exists(table_name):
        sql_f.write(f"DROP TABLE IF EXISTS {table_name};\n")

    current_table = None
    csv_f = None
    in_copy = False
    in_create_table = False
    skip_table = False

    for line in f:
        # --- Handle CREATE TABLE ---
        if line.startswith('CREATE TABLE'):
            # Check if table is relevant
            match = re.search(r'public\.(\w+)', line)
            if match:
                table_name = match.group(1)
                if table_name in relevant_tables:
                    write_drop_if_exists(table_name) # Ensure idempotency
                    in_create_table = True
                    skip_table = False
                    clean_line = line.replace('public.', '').replace('character varying', 'VARCHAR')
                    sql_f.write(clean_line)
                else:
                    skip_table = True
            continue

        if skip_table:
            if line.startswith(');'):
                skip_table = False # Reset but don't write
            continue

        if in_create_table:
            if line.startswith(');'):
                sql_f.write(line + "\n")
                in_create_table = False
            else:
                # Process column definitions
                # Replace character varying FIRST as it is the most common string type in PG
                clean_line = line.replace('character varying', 'VARCHAR')

                # Remove COLLATE
                clean_line = re.sub(r'COLLATE\s+pg_catalog\."\w+"', '', clean_line)
                if 'public.geography' in clean_line or 'public.geometry' in clean_line:
                    clean_line = re.sub(r'public\.ge(ography|ometry)\(.*\)', 'VARCHAR', clean_line)

                # Replace Array types (e.g. VARCHAR(255)[], integer[]) with VARCHAR
                # Regex looks for [] at the end of the type definition
                clean_line = re.sub(r'\[\]', '', clean_line)

                # Now replace standard numeric/other types with VARCHAR
                # We do this by replacing common PG types.
                # Note: We must be careful not to replace parts of column names, but usually types are distinct words.

                replacements = [
                    ('integer', 'VARCHAR'),
                    ('bigint', 'VARCHAR'),
                    ('double precision', 'VARCHAR'),
                    ('numeric', 'VARCHAR'),
                    ('boolean', 'VARCHAR'),
                    ('date', 'VARCHAR'),
                    ('timestamp without time zone', 'VARCHAR'),
                    ('timestamp with time zone', 'VARCHAR'),
                    ('text', 'VARCHAR'),
                    # character varying is already handled
                ]

                for pg_type, duck_type in replacements:
                    # Use word boundaries to avoid replacing substrings
                    clean_line = re.sub(r'\b' + pg_type + r'\b', duck_type, clean_line)

                # Ensure NOT NULL constraints remain, but type is VARCHAR.

                # Clean up double spaces
                clean_line = re.sub(r'\s+', ' ', clean_line).replace(' ,', ',')
                # Restore newline
                clean_line += "\n"

                sql_f.write(clean_line)
            continue

        # --- Handle COPY ---
        copy_match = re.match(r'COPY public\.(\w+) \((.*?)\) FROM stdin;', line)
        if copy_match:
            current_table = copy_match.group(1)

            if current_table in relevant_tables:
                columns = copy_match.group(2)
                csv_path = os.path.join(output_dir, f"{current_table}.csv")

                sql_f.write(f"COPY {current_table} FROM '{csv_path}' (DELIMITER '\t', NULL '\\N');\n")

                csv_f = open(csv_path, 'w', encoding='utf-8')
                in_copy = True
            else:
                in_copy = False # Skip this copy block
                current_table = None # Skip writing
            continue

        if in_copy and csv_f:
            if line.strip() == '\\.':
                csv_f.close()
                in_copy = False
                current_table = None
                continue
            csv_f.write(line)

print(f"Conversion complete. Artifacts in {output_dir}. SQL: {output_sql}")
