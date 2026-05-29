import zipfile, json, xml.etree.ElementTree as ET

filename = '4、拉迪亚码头-二航院_clean.xlsx'
ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

with zipfile.ZipFile(filename, 'r') as z:
    all_files = z.namelist()
    print(f"Total files: {len(all_files)}")

    # Check for any remaining external references
    bad = [n for n in all_files if 'externalLink' in n.lower()]
    print(f"External link files remaining: {len(bad)}")

    # Check defined names
    if 'xl/workbook.xml' in all_files:
        tree = ET.parse(z.open('xl/workbook.xml'))
        root = tree.getroot()
        dns = root.findall('.//{'+ns+'}definedNames/{'+ns+'}definedName')
        print(f"\nRemaining defined names: {len(dns)}")

        bad_found = 0
        for dn in dns:
            text = dn.text or ''
            if '[' in text or '#REF!' in text:
                bad_found += 1
                if bad_found <= 5:
                    print(f"  UNEXPECTED BAD: {dn.get('name')} -> {text[:100]}")
        print(f"Bad names remaining: {bad_found}")

        if len(dns) <= 30:
            for dn in dns:
                print(f"  [{dn.get('name')}] -> {dn.text}")

    # Check workbook.xml.rels
    rels_ns_pkg = 'http://schemas.openxmlformats.org/package/2006/relationships'
    if 'xl/_rels/workbook.xml.rels' in all_files:
        tree = ET.parse(z.open('xl/_rels/workbook.xml.rels'))
        root = tree.getroot()
        ext_rels = []
        for rel in root:
            if 'externalLink' in rel.get('Type', '') or 'externalLink' in rel.get('Target', '').lower():
                ext_rels.append(rel.get('Target'))
        print(f"\nExternal rels remaining in workbook.xml.rels: {len(ext_rels)}")

    # Check [Content_Types].xml
    if '[Content_Types].xml' in all_files:
        tree = ET.parse(z.open('[Content_Types].xml'))
        root = tree.getroot()
        ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
        ext_ct = []
        for ov in root.findall('.//{'+ct_ns+'}Override'):
            if 'externalLink' in ov.get('PartName', '').lower():
                ext_ct.append(ov.get('PartName'))
        print(f"\nExternal content types remaining: {len(ext_ct)}")

    print("\n=== Clean file structure ===")
    # Show directory tree
    dirs = sorted(set(
        d.rsplit('/', 1)[0] + '/'
        for d in all_files if '/' in d
    ))
    for d in sorted(dirs):
        files = sorted([f for f in all_files if f.startswith(d) and f != d])
        print(f"\n  {d}")
        for f in files:
            size = z.getinfo(f).file_size
            print(f"    {f} ({size:,} bytes)")
