"""
Convert a cleaned xlsx file to Semantic Excel-AST (markdown format).
Produces a structured, AI-readable representation of the spreadsheet.
"""
import openpyxl, os, sys, json
from collections import Counter
from openpyxl.utils import get_column_letter

MAX_SAMPLE_ROWS = 5  # sample rows per column
MAX_AST_COLS = 200   # don't dump columns beyond this in full table mode

def infer_col_type(column_cells):
    """Infer the semantic type of a column from its values."""
    types = Counter()
    non_empty = 0
    for val in column_cells:
        if val is None:
            continue
        non_empty += 1
        if isinstance(val, (int, float)):
            types['number'] += 1
        elif isinstance(val, str):
            s = val.strip()
            if s.startswith('='):
                types['formula'] += 1
            elif s.startswith('*') or s.startswith('【') or s.startswith('《'):
                types['category_header'] += 1
            elif '万' in s or '亿' in s or s.endswith('元'):
                types['amount_with_unit'] += 1
            elif len(s) > 50:
                types['long_text'] += 1
            else:
                types['text'] += 1
        elif isinstance(val, bool):
            types['boolean'] += 1

    if non_empty == 0:
        return 'empty'

    dominant = types.most_common(1)[0]
    if dominant[1] >= non_empty * 0.7:
        return dominant[0]
    if 'category_header' in types:
        return 'header_text'
    if 'formula' in types and types.get('number', 0) > non_empty * 0.3:
        return 'calculated_number'
    return 'mixed'

def sanitize(val):
    """Make a value safe for markdown."""
    if val is None:
        return ''
    s = str(val)
    s = s.replace('|', '\\|').replace('\n', ' ').replace('\r', '')
    if len(s) > 80:
        s = s[:77] + '...'
    return s

def build_sheet_ast(ws):
    """Build AST for a single sheet."""
    ast = {'name': ws.title, 'dimensions': f'{ws.min_column},{ws.max_column},{ws.min_row},{ws.max_row}'}

    # Determine actual used columns (skip trailing empty columns)
    rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 3000),
                              values_only=True))

    if not rows:
        ast['rows'] = 0
        ast['columns'] = 0
        ast['note'] = 'empty sheet'
        return ast

    # Find actual column count (non-empty)
    max_col = 0
    for row in rows:
        for ci, val in enumerate(row):
            if val is not None and str(val).strip():
                max_col = max(max_col, ci + 1)

    ast['rows'] = len(rows)
    ast['columns'] = max_col

    # Analyze header rows (first 3 rows)
    header_info = []
    for ri in range(min(3, len(rows))):
        row_vals = []
        for ci in range(min(max_col, MAX_AST_COLS)):
            row_vals.append(sanitize(rows[ri][ci] if ci < len(rows[ri]) else None))
        header_info.append({'row': ri + 1, 'values': row_vals})

    ast['header_rows'] = header_info

    # Column analysis
    columns = []
    for ci in range(min(max_col, MAX_AST_COLS)):
        vals = []
        for ri in range(len(rows)):
            if ci < len(rows[ri]):
                vals.append(rows[ri][ci])
            else:
                vals.append(None)
        col_type = infer_col_type(vals)
        non_empty = sum(1 for v in vals if v is not None and str(v).strip())
        samples = [sanitize(v) for v in vals[:MAX_SAMPLE_ROWS] if v is not None]
        columns.append({
            'col': get_column_letter(ci + 1),
            'type': col_type,
            'non_empty_cells': non_empty,
            'samples': samples[:5],
        })

    ast['columns_analysis'] = columns

    # Table body (for sheets <= 50 columns, dump first 30 rows as markdown table)
    if max_col <= 50 and len(rows) <= 100:
        table_data = []
        for ri in range(min(30, len(rows))):
            row_vals = []
            for ci in range(max_col):
                row_vals.append(sanitize(rows[ri][ci] if ci < len(rows[ri]) else None))
            table_data.append(row_vals)
        ast['preview_table'] = table_data

    # Summary
    type_summary = Counter(c['type'] for c in columns)
    ast['column_type_summary'] = dict(type_summary.most_common())

    return ast

def ast_to_markdown(ast, output_path):
    """Write AST as markdown document."""
    lines = []
    lines.append(f"# Excel-AST: {os.path.basename(output_path).replace('.md', '')}")
    lines.append("")
    lines.append(f"> Semantic representation of cleaned workbook structure")
    lines.append("")

    # Workbook overview
    total_rows = sum(s['rows'] for s in ast['sheets'])
    total_cols = sum(s['columns'] for s in ast['sheets'])
    lines.append("## Workbook Overview")
    lines.append("")
    lines.append(f"| Metric | Value |")
    lines.append(f"|--------|-------|")
    lines.append(f"| Sheets | {len(ast['sheets'])} |")
    lines.append(f"| Total rows (all sheets) | {total_rows} |")
    lines.append(f"| Total columns (all sheets) | {total_cols} |")
    lines.append("")

    # Per-sheet detail
    for sheet in ast['sheets']:
        lines.append(f"## Sheet: `{sheet['name']}`")
        lines.append("")
        lines.append(f"- **Rows:** {sheet['rows']}")
        lines.append(f"- **Columns:** {sheet['columns']}")
        lines.append(f"- **Range:** {sheet.get('dimensions', 'N/A')}")
        if 'note' in sheet:
            lines.append(f"- **Note:** {sheet['note']}")
            lines.append("")
            continue
        lines.append("")

        # Column type distribution
        if sheet.get('column_type_summary'):
            lines.append("### Column Type Distribution")
            lines.append("")
            lines.append("| Type | Count |")
            lines.append("|------|-------|")
            for col_type, count in sheet['column_type_summary'].items():
                lines.append(f"| {col_type} | {count} |")
            lines.append("")

        # Header preview
        if sheet.get('header_rows'):
            lines.append("### Header Preview (first 3 rows)")
            lines.append("")
            for hr in sheet['header_rows']:
                header_vals = hr['values'][:30]
                lines.append(f"**Row {hr['row']}:** {' | '.join(header_vals)}")
                lines.append("")
            lines.append("")

        # Column detail
        cols = sheet.get('columns_analysis', [])
        if cols:
            lines.append("### Column Analysis")
            lines.append("")
            lines.append("| Col | Letter | Type | Non-empty | Samples |")
            lines.append("|-----|--------|------|-----------|---------|")
            for ci, c in enumerate(cols):
                samples_str = ', '.join(c['samples'][:3])
                lines.append(f"| {ci+1} | {c['col']} | {c['type']} | {c['non_empty_cells']} | {samples_str} |")
            lines.append("")

        # Table preview (small sheets only)
        if sheet.get('preview_table'):
            lines.append("### Data Preview (first 30 rows)")
            lines.append("")
            pt = sheet['preview_table']
            for ri, row in enumerate(pt):
                lines.append('| ' + ' | '.join(str(v) for v in row) + ' |')
                if ri == 0:
                    lines.append('| ' + ' | '.join('---' for _ in row) + ' |')
            lines.append("")

    # Write
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))

    return output_path


if __name__ == '__main__':
    input_file = '4、拉迪亚码头-二航院_clean.xlsx'
    output_file = '4、拉迪亚码头-二航院_AST.md'

    print(f"Reading: {input_file}")

    wb = openpyxl.load_workbook(input_file, read_only=True, data_only=True)
    sheets_ast = []
    for name in wb.sheetnames:
        print(f"  Processing sheet: {name}")
        ws = wb[name]
        sheet_ast = build_sheet_ast(ws)
        sheets_ast.append(sheet_ast)

    wb.close()

    ast = {'sheets': sheets_ast}

    ast_to_markdown(ast, output_file)
    print(f"\nAST written to: {output_file}")

    # Summary
    for s in sheets_ast:
        print(f"  [{s['name']}] {s['rows']} rows x {s['columns']} cols")
        if s.get('column_type_summary'):
            print(f"    Types: {dict(s['column_type_summary'])}")
