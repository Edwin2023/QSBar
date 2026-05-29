"""
Clean xlsx file: strip external links and bad defined names at ZIP/XML level.
"""
import zipfile, os, shutil, sys, io, xml.etree.ElementTree as ET

SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
EXT_LINK_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
EXT_REF_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath'

def is_bad_name(text):
    """Check if a defined name formula refers to external or broken reference."""
    if not text:
        return True
    t = text
    if '#REF!' in t:
        return True
    if '[' in t:
        return True
    if ']' in t and '!' in t:
        return True
    if '!' in t and (':\\' in t or '\\\\' in t):
        return True
    return False

def clean_rels(rel_xml_bytes):
    """Remove external link relationships from a .rels file."""
    ET.register_namespace('', RELS_NS)
    tree = ET.fromstring(rel_xml_bytes)
    root = tree
    to_remove = []
    for rel in root:
        rel_type = rel.get('Type', '')
        target = rel.get('Target', '')
        if 'externalLink' in rel_type or 'externalLink' in target.lower():
            to_remove.append(rel)
    for rel in to_remove:
        root.remove(rel)
    return ET.tostring(root, encoding='UTF-8', xml_declaration=True)

def clean_workbook_xml(wb_xml_bytes, names_to_delete_count):
    """Remove bad defined names and externalReferences from workbook.xml."""
    ET.register_namespace('', SPREADSHEET_NS)
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    tree = ET.fromstring(wb_xml_bytes)
    root = tree

    # Remove bad definedNames
    dns_parents = root.findall('.//{'+SPREADSHEET_NS+'}definedNames')
    removed = 0
    for dns_parent in dns_parents:
        to_remove = []
        for dn in dns_parent:
            text = dn.text if dn.text else ''
            if is_bad_name(text):
                to_remove.append(dn)
        for dn in to_remove:
            dns_parent.remove(dn)
            removed += 1

    # If all definedNames are gone, remove the empty element
    for dns_parent in dns_parents:
        if len(dns_parent) == 0:
            root.remove(dns_parent)

    # Remove externalReferences
    for ext_refs in root.findall('.//{'+SPREADSHEET_NS+'}externalReferences'):
        root.remove(ext_refs)

    # Remove externalReference children inside externalReferences
    for ext_ref in root.findall('.//{'+SPREADSHEET_NS+'}externalReference'):
        parent = root.find('.//{'+SPREADSHEET_NS+'}externalReferences/..')
        # just remove all externalReferences elements
        pass

    result = ET.tostring(root, encoding='UTF-8', xml_declaration=True)
    return result, removed

def clean_file(input_path, output_path):
    """Main cleanup: strip external links and bad names from xlsx."""
    stats = {'ext_link_files_removed': 0, 'names_removed': 0, 'rels_cleaned': 0}

    with zipfile.ZipFile(input_path, 'r') as zin:
        namelist = zin.namelist()

        # Build list of files to keep
        keep_files = []
        skip_files = []

        for name in namelist:
            if 'externalLinks/' in name:
                skip_files.append(name)
                stats['ext_link_files_removed'] += 1
            else:
                keep_files.append(name)

        print(f"Files to remove: {stats['ext_link_files_removed']}")
        print(f"Files to keep: {len(keep_files)}")

        # Create output zip
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name in keep_files:
                data = zin.read(name)

                # Clean .rels files that reference external links
                if name.endswith('.rels') and name != '_rels/.rels':
                    try:
                        new_data = clean_rels(data)
                        stats['rels_cleaned'] += 1
                        zout.writestr(name, new_data)
                        continue
                    except Exception as e:
                        print(f"  WARNING: failed to clean {name}: {e}")

                # Clean workbook.xml
                if name == 'xl/workbook.xml':
                    try:
                        new_data, removed = clean_workbook_xml(data, 0)
                        stats['names_removed'] = removed
                        zout.writestr(name, new_data)
                        continue
                    except Exception as e:
                        print(f"  WARNING: failed to clean workbook.xml: {e}")

                # Default: copy as-is
                zout.writestr(name, data)

    return stats

if __name__ == '__main__':
    input_file = '4、拉迪亚码头-二航院.xlsx'
    output_file = '4、拉迪亚码头-二航院_clean.xlsx'

    print(f"Cleaning: {input_file}")
    print(f"Output: {output_file}")
    print()

    stats = clean_file(input_file, output_file)

    print()
    print("=== Cleanup Summary ===")
    print(f"  External link files removed: {stats['ext_link_files_removed']}")
    print(f"  Bad defined names removed:   {stats['names_removed']}")
    print(f"  .rels files cleaned:         {stats['rels_cleaned']}")

    # Verify
    with zipfile.ZipFile(output_file, 'r') as z:
        remaining = [n for n in z.namelist() if 'externalLink' in n.lower()]
        print(f"  Remaining external refs:     {len(remaining)}")

    orig_size = os.path.getsize(input_file)
    new_size = os.path.getsize(output_file)
    print(f"\nFile size: {orig_size/1024:.0f} KB -> {new_size/1024:.0f} KB ({(1-new_size/orig_size)*100:.1f}% reduced)")
