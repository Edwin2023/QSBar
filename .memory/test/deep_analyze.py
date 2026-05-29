import zipfile, os, json, xml.etree.ElementTree as ET
from collections import defaultdict

filename = '4、拉迪亚码头-二航院.xlsx'
target_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'

report = {}

with zipfile.ZipFile(filename, 'r') as z:
    all_files = z.namelist()
    report['total_files'] = len(all_files)

    # 1. Analyze all .rels files
    rels_files = sorted([n for n in all_files if n.endswith('.rels')])
    report['rels_files'] = len(rels_files)

    ext_refs = []
    for rf in rels_files:
        tree = ET.parse(z.open(rf))
        root = tree.getroot()
        for rel in root:
            target = rel.get('Target', '')
            rel_type = rel.get('Type', '')
            if 'externalLink' in target or 'externalLink' in rel_type:
                ext_refs.append({'from': rf, 'target': target, 'type': rel_type})

    report['external_rels_references'] = len(ext_refs)
    report['ext_refs_sample'] = ext_refs[:10]

    # 2. Analyze workbook.xml - defined names and other nasties
    ns_main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    if 'xl/workbook.xml' in all_files:
        tree = ET.parse(z.open('xl/workbook.xml'))
        root = tree.getroot()

        # Defined names
        dns = root.findall('.//{'+ns_main+'}definedNames/{'+ns_main+'}definedName')
        names_clean = []
        names_bad = []
        for dn in dns:
            name = dn.get('name', '')
            text = dn.text or ''
            if ('[' in text or '#REF' in text or
                (']' in text and '!' in text) or
                (':' in text and '\\\\' in text)):
                names_bad.append({'name': name, 'refersTo': text[:150]})
            else:
                names_clean.append({'name': name, 'refersTo': text[:150]})

        report['definedNames_total'] = len(dns)
        report['definedNames_clean'] = len(names_clean)
        report['definedNames_bad'] = len(names_bad)
        report['bad_names_sample'] = names_bad[:15]

        # Check for CalculationProperties
        calc = root.find('.//{'+ns_main+'}calcPr')
        if calc is not None:
            report['calcPr'] = {k: v for k, v in calc.attrib.items()}

        # Check for externalReferences in workbook
        ext_refs_wb = root.findall('.//{'+ns_main+'}externalReferences/{'+ns_main+'}externalReference')
        report['workbook_externalReferences'] = len(ext_refs_wb)

    # 3. Check sheet xmls for external references in formulas
    sheet_files = sorted([n for n in all_files if n.startswith('xl/worksheets/') and n.endswith('.xml')])
    sheet_ext_summary = {}
    for sf in sheet_files:
        tree = ET.parse(z.open(sf))
        root = tree.getroot()
        cells = root.findall('.//{'+ns_main+'}c')
        ext_count = 0
        for c in cells:
            for child in c:
                text = child.text or ''
                if '[' in text or '#REF!' in text:
                    ext_count += 1
                    break
        if ext_count > 0:
            sheet_ext_summary[sf] = ext_count
    report['sheets_with_external_formulas'] = sheet_ext_summary

    # 4. Check for chartsheets, charts with external refs
    chart_files = [n for n in all_files if 'chart' in n.lower() and n.endswith('.xml')]
    report['chart_files'] = len(chart_files)

    # 5. Check pivot caches
    pivot_files = [n for n in all_files if 'pivot' in n.lower() and (n.endswith('.xml') or n.endswith('.rels'))]
    report['pivot_files'] = len(pivot_files)

    # 6. External links folder count
    ext_link_xml = [n for n in all_files if 'externalLinks/' in n and n.endswith('.xml')]
    ext_link_rels = [n for n in all_files if 'externalLinks/' in n and n.endswith('.rels')]
    report['externalLink_xml_count'] = len(ext_link_xml)
    report['externalLink_rels_count'] = len(ext_link_rels)

    # 7. Check for other binary external refs in rels
    ole_files = [n for n in all_files if 'oleObject' in n]
    report['oleObject_files'] = len(ole_files)

    # 8. Sheet rels with external references
    sheet_rels = [n for n in all_files if 'worksheets/_rels/' in n]
    sheet_rels_bad = {}
    for sr in sheet_rels:
        tree = ET.parse(z.open(sr))
        root = tree.getroot()
        for rel in root:
            target = rel.get('Target', '')
            if target.startswith('..') and ('externalLink' in target.lower()):
                sheet_rels_bad[sr] = sheet_rels_bad.get(sr, 0) + 1
    report['sheet_rels_external'] = sheet_rels_bad

print(json.dumps(report, indent=2, ensure_ascii=False))
