import zipfile, xml.etree.ElementTree as ET

filename = '4、拉迪亚码头-二航院.xlsx'
with zipfile.ZipFile(filename, 'r') as z:
    # External references files
    ext_rels = [n for n in z.namelist() if 'externalLink' in n.lower() or 'external' in n.lower()]
    print('External-related files:', len(ext_rels))
    for n in ext_rels[:30]:
        print(' ', n)
    print()

    # Defined names
    print('=== Defined Names ===')
    if 'xl/workbook.xml' in z.namelist():
        tree = ET.parse(z.open('xl/workbook.xml'))
        root = tree.getroot()
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        dns = root.findall('.//{'+ns+'}definedNames/{'+ns+'}definedName')
        print('Total defined names:', len(dns))
        bad_count = 0
        for dn in dns:
            name = dn.get('name', '')
            text = dn.text or ''
            has_bracket = '[' in text
            has_ref = '#REF' in text
            has_path = ('!' in text) and (':' in text or '\\\\' in text)
            is_bad = has_bracket or has_ref or has_path
            if is_bad:
                bad_count += 1
                if bad_count <= 30:
                    print('  BAD [{}] => {}'.format(name, text[:200]))
        print('Names with external/error refs:', bad_count)
    print()

    # Ext links
    ext_link_files = [n for n in z.namelist() if 'externalLinks' in n and n.endswith('.xml')]
    print('External link XML files:', len(ext_link_files))
    for elf in ext_link_files[:5]:
        tree = ET.parse(z.open(elf))
        root = tree.getroot()
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        for eb in root.findall('.//{'+ns+'}externalBook'):
            rid = eb.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            print('  {} externalBook r:id={}'.format(elf, rid))
