import pandas as pd
import re


# ── SQL keyword sets ──────────────────────────────────────────────────────────
_TABLE_KEYWORDS = r'\b(?:FROM|JOIN|INNER\s+JOIN|LEFT\s+JOIN|RIGHT\s+JOIN|FULL\s+JOIN|CROSS\s+JOIN|INTO|UPDATE)\b'

_SKIP_KEYWORDS = {
    'SELECT','FROM','WHERE','JOIN','INNER','LEFT','RIGHT','FULL','CROSS','OUTER',
    'ON','AND','OR','NOT','IN','IS','NULL','AS','BY','GROUP','ORDER','HAVING',
    'LIMIT','OFFSET','UNION','ALL','DISTINCT','CASE','WHEN','THEN','ELSE','END',
    'SET','INTO','VALUES','EXISTS','BETWEEN','LIKE','ILIKE','ASC','DESC','WITH',
    'OVER','PARTITION','ROWS','RANGE','UNBOUNDED','PRECEDING','FOLLOWING',
    'CURRENT','ROW','FILTER','LATERAL','RECURSIVE','TRUE','FALSE','TOP',
    'UPDATE','DELETE','INSERT','MERGE','TRUNCATE','CREATE','DROP','ALTER',
    'COUNT','SUM','AVG','MIN','MAX','COALESCE','NULLIF','CAST','CONVERT',
    'TRIM','UPPER','LOWER','LEN','LENGTH','SUBSTR','SUBSTRING','REPLACE',
    'DATEPART','DATEDIFF','DATEADD','GETDATE','NOW','NVL','DECODE','IFF',
    'ROW_NUMBER','RANK','DENSE_RANK','NTILE','LAG','LEAD','FIRST_VALUE','LAST_VALUE'
}


def _clean_sql(sql: str) -> str:
    """Strip comments from SQL."""
    sql = re.sub(r'--[^\n]*', ' ', sql)
    sql = re.sub(r'/\*.*?\*/', ' ', sql, flags=re.DOTALL)
    return sql


def extract_table_names(sql: str) -> list[str]:
    """Extract unique table names following FROM/JOIN/INTO/UPDATE keywords."""
    if not isinstance(sql, str) or not sql.strip():
        return ["NO_TABLE_FOUND"]

    sql = _clean_sql(sql)
    pattern = rf'{_TABLE_KEYWORDS}\s+([\w\.\`\"\[\]]+)'
    matches = re.findall(pattern, sql, flags=re.IGNORECASE)

    seen, tables = set(), []
    for m in matches:
        name = re.sub(r'[\`\"\[\]]', '', m).strip()
        if name.upper() not in _SKIP_KEYWORDS and name.lower() not in seen:
            seen.add(name.lower())
            tables.append(name)

    return tables if tables else ["NO_TABLE_FOUND"]


def extract_column_names(sql: str) -> str:
    """
    Extract column/field names from anywhere in the SQL text.
    Covers:
      - SELECT col, table.col, col AS alias
      - WHERE / ON / HAVING  col = ...
      - GROUP BY / ORDER BY  col
      - SET col = ...  (UPDATE)
      - Handles schema.table.col → keeps only the column part
    Returns a comma-separated string of unique column names.
    """
    if not isinstance(sql, str) or not sql.strip():
        return "NO_COLUMNS_FOUND"

    sql = _clean_sql(sql)
    columns = []
    seen = set()

    # ── 1. SELECT list ───────────────────────────────────────────────────────
    # Grab everything between SELECT and FROM (handles multiline)
    select_match = re.search(r'\bSELECT\b(.*?)\bFROM\b', sql, re.IGNORECASE | re.DOTALL)
    if select_match:
        select_block = select_match.group(1)
        # Each comma-separated token; strip alias (AS ...) and functions
        for token in select_block.split(','):
            token = re.sub(r'\bAS\b\s+\w+', '', token, flags=re.IGNORECASE)  # remove alias
            token = token.strip()
            # Resolve table.col or schema.table.col → last part is column
            parts = re.split(r'\.', token)
            col = parts[-1].strip()
            col = re.sub(r'[\`\"\[\]\(\)]', '', col).strip()
            # Skip wildcards, literals, numbers, functions with parens
            if col and col != '*' and not col.isdigit() and col.upper() not in _SKIP_KEYWORDS:
                key = col.lower()
                if key not in seen:
                    seen.add(key)
                    columns.append(col)

    # ── 2. WHERE / ON / HAVING / SET clauses ─────────────────────────────────
    # Find patterns like:  col = , col >, col <, col IN, col IS, col LIKE
    clause_pattern = r'([\w\.\`\"\[\]]+)\s*(?:=|!=|<>|>=|<=|>|<|\bIN\b|\bIS\b|\bLIKE\b|\bBETWEEN\b)'
    for raw in re.findall(clause_pattern, sql, flags=re.IGNORECASE):
        parts = re.split(r'\.', raw)
        col = re.sub(r'[\`\"\[\]]', '', parts[-1]).strip()
        if col and not col.isdigit() and col.upper() not in _SKIP_KEYWORDS:
            key = col.lower()
            if key not in seen:
                seen.add(key)
                columns.append(col)

    # ── 3. GROUP BY / ORDER BY ───────────────────────────────────────────────
    for clause in re.findall(r'\b(?:GROUP|ORDER)\s+BY\s+([\w\s\.\,\`\"\[\]]+?)(?:\bHAVING\b|\bORDER\b|\bLIMIT\b|\bUNION\b|$)',
                              sql, flags=re.IGNORECASE):
        for token in clause.split(','):
            token = re.sub(r'\b(ASC|DESC)\b', '', token, flags=re.IGNORECASE)
            parts = re.split(r'\.', token.strip())
            col = re.sub(r'[\`\"\[\]]', '', parts[-1]).strip()
            if col and not col.isdigit() and col.upper() not in _SKIP_KEYWORDS:
                key = col.lower()
                if key not in seen:
                    seen.add(key)
                    columns.append(col)

    return ', '.join(columns) if columns else "NO_COLUMNS_FOUND"


def process_excel(input_path: str, sql_column: str, output_path: str) -> pd.DataFrame:
    """
    Read Excel → extract table names + column names → explode rows → write to Excel.

    Args:
        input_path:  Path to the source .xlsx file
        sql_column:  Column name containing SQL text
        output_path: Path for the output .xlsx file

    Returns:
        Expanded DataFrame (one row per table name, with columns listed)
    """
    # 1. Load Excel
    df = pd.read_excel(input_path)

    if sql_column not in df.columns:
        raise ValueError(f"Column '{sql_column}' not found. Available: {list(df.columns)}")

    # 2. Extract table names (list) and column names (list)
    df['table_name']   = df[sql_column].apply(extract_table_names)
    df['column_names'] = df[sql_column].apply(
        lambda sql: extract_column_names(sql).split(', ') if extract_column_names(sql) != "NO_COLUMNS_FOUND" else ["NO_COLUMNS_FOUND"]
    )

    # 3. Explode table_name first (replicates rows per table)
    df_expanded = df.explode('table_name').reset_index(drop=True)

    # 4. Explode column_names — each column gets its own row
    df_expanded = df_expanded.explode('column_names').reset_index(drop=True)
    df_expanded.rename(columns={'column_names': 'column_name'}, inplace=True)

    # 4. Write to Excel
    df_expanded.to_excel(output_path, index=False)
    print(f"Done! {len(df_expanded)} rows written to '{output_path}'")
    print(f"  Original rows : {len(df)}")
    print(f"  Expanded rows : {len(df_expanded)}")

    return df_expanded


# ── Run ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    INPUT_FILE  = "your_file.xlsx"       # <-- your input Excel file
    SQL_COLUMN  = "sql_text"             # <-- your SQL column name
    OUTPUT_FILE = "output_expanded.xlsx"

    result_df = process_excel(INPUT_FILE, SQL_COLUMN, OUTPUT_FILE)
    print(result_df)
