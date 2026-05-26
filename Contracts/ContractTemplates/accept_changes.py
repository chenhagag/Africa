"""
Accept all tracked changes and disable Track Changes in Word .docx files.

For each file in Old Templates/:
1. Accept all insertions (keep inserted content, remove w:ins wrapper)
2. Accept all deletions (remove w:del and its content entirely)
3. Remove property change records (rPrChange, pPrChange, sectPrChange, etc.)
4. Remove self-closing revision markers (w:ins inside pPr, etc.)
5. Convert w:delText to w:t (for any edge cases)
6. Turn off trackChanges in settings.xml
7. Remove people.xml (revision author metadata)
"""

import zipfile
import os
import re
from lxml import etree
from copy import deepcopy

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = f'{{{W_NS}}}'

INPUT_DIR = "Old Templates"

# Revision element local names
PROPERTY_CHANGES = {'rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange',
                    'tcPrChange', 'trPrChange', 'tblGridChange', 'numberingChange',
                    'cellIns', 'cellDel', 'cellMerge'}


def accept_revisions_in_tree(root):
    """Accept all tracked changes in an lxml element tree."""
    changed = True
    passes = 0

    # Multiple passes since accepting changes can expose nested revisions
    while changed and passes < 10:
        changed = False
        passes += 1

        # 1. Accept insertions: unwrap w:ins (keep children)
        for ins in root.findall(f'.//{W}ins'):
            parent = ins.getparent()
            if parent is None:
                continue
            idx = list(parent).index(ins)
            children = list(ins)

            # Preserve text before first child (ins.text)
            if ins.text:
                if idx > 0:
                    prev = parent[idx - 1]
                    prev.tail = (prev.tail or '') + ins.text
                else:
                    parent.text = (parent.text or '') + ins.text

            # Move children to parent
            for i, child in enumerate(children):
                parent.insert(idx + i, child)

            # Preserve tail text
            if ins.tail:
                if children:
                    last_child = children[-1]
                    last_child.tail = (last_child.tail or '') + ins.tail
                elif idx > 0:
                    prev = parent[idx + len(children) - 1] if children else parent[idx - 1]
                    prev.tail = (prev.tail or '') + ins.tail
                else:
                    parent.text = (parent.text or '') + ins.tail

            parent.remove(ins)
            changed = True

        # 2. Accept deletions: remove w:del entirely
        for del_elem in root.findall(f'.//{W}del'):
            parent = del_elem.getparent()
            if parent is None:
                continue
            # Preserve tail text of the deleted element
            if del_elem.tail:
                idx = list(parent).index(del_elem)
                if idx > 0:
                    prev = parent[idx - 1]
                    prev.tail = (prev.tail or '') + del_elem.tail
                else:
                    parent.text = (parent.text or '') + del_elem.tail
            parent.remove(del_elem)
            changed = True

        # 3. Accept moveFrom (remove) and moveTo (unwrap, like ins)
        for mf in root.findall(f'.//{W}moveFrom'):
            parent = mf.getparent()
            if parent is None:
                continue
            if mf.tail:
                idx = list(parent).index(mf)
                if idx > 0:
                    prev = parent[idx - 1]
                    prev.tail = (prev.tail or '') + mf.tail
                else:
                    parent.text = (parent.text or '') + mf.tail
            parent.remove(mf)
            changed = True

        for mt in root.findall(f'.//{W}moveTo'):
            parent = mt.getparent()
            if parent is None:
                continue
            idx = list(parent).index(mt)
            children = list(mt)
            if mt.text:
                if idx > 0:
                    prev = parent[idx - 1]
                    prev.tail = (prev.tail or '') + mt.text
                else:
                    parent.text = (parent.text or '') + mt.text
            for i, child in enumerate(children):
                parent.insert(idx + i, child)
            if mt.tail:
                if children:
                    children[-1].tail = (children[-1].tail or '') + mt.tail
                elif idx > 0:
                    prev = parent[idx - 1]
                    prev.tail = (prev.tail or '') + mt.tail
                else:
                    parent.text = (parent.text or '') + mt.tail
            parent.remove(mt)
            changed = True

        # 4. Remove property change records
        for tag_name in PROPERTY_CHANGES:
            for elem in root.findall(f'.//{W}{tag_name}'):
                parent = elem.getparent()
                if parent is None:
                    continue
                if elem.tail:
                    idx = list(parent).index(elem)
                    if idx > 0:
                        prev = parent[idx - 1]
                        prev.tail = (prev.tail or '') + elem.tail
                    else:
                        parent.text = (parent.text or '') + elem.tail
                parent.remove(elem)
                changed = True

        # 5. Remove moveFromRangeStart/End and moveToRangeStart/End
        for tag_name in ['moveFromRangeStart', 'moveFromRangeEnd',
                         'moveToRangeStart', 'moveToRangeEnd']:
            for elem in root.findall(f'.//{W}{tag_name}'):
                parent = elem.getparent()
                if parent is None:
                    continue
                if elem.tail:
                    idx = list(parent).index(elem)
                    if idx > 0:
                        prev = parent[idx - 1]
                        prev.tail = (prev.tail or '') + elem.tail
                    else:
                        parent.text = (parent.text or '') + elem.tail
                parent.remove(elem)
                changed = True

    # 6. Convert any remaining w:delText to w:t (shouldn't happen but safety net)
    for dt in root.findall(f'.//{W}delText'):
        dt.tag = f'{W}t'

    # 7. Remove rsidDel attributes from w:r elements
    for r in root.findall(f'.//{W}r'):
        if f'{W}rsidDel' in r.attrib:
            del r.attrib[f'{W}rsidDel']

    return root


def disable_track_changes(settings_xml_bytes):
    """Remove trackChanges element from settings.xml."""
    root = etree.fromstring(settings_xml_bytes)

    # Remove w:trackChanges
    for tc in root.findall(f'.//{W}trackChanges'):
        parent = tc.getparent()
        if parent is not None:
            parent.remove(tc)

    # Remove w:revisionView if present (forces markup view)
    for rv in root.findall(f'.//{W}revisionView'):
        parent = rv.getparent()
        if parent is not None:
            parent.remove(rv)

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def process_file(filepath):
    """Process a single .docx file: accept all changes and disable tracking."""

    with zipfile.ZipFile(filepath, 'r') as zin:
        file_list = zin.namelist()
        file_contents = {}
        for name in file_list:
            file_contents[name] = zin.read(name)

    # XML parts that can contain revision marks
    xml_parts = [n for n in file_list if n.endswith('.xml') and
                 (n.startswith('word/') and 'webextension' not in n and
                  'glossary' not in n and 'theme' not in n)]

    revisions_found = 0

    for part in xml_parts:
        if part not in file_contents:
            continue

        raw = file_contents[part]

        if part == 'word/settings.xml':
            file_contents[part] = disable_track_changes(raw)
            continue

        # Skip parts that don't have any revision marks
        raw_str = raw.decode('utf-8', errors='replace')
        has_revisions = any(marker in raw_str for marker in
                           ['<w:ins ', '<w:ins/', '<w:del ', '<w:del/',
                            ':rPrChange', ':pPrChange', ':sectPrChange',
                            ':tblPrChange', ':tcPrChange', ':trPrChange',
                            ':moveFrom', ':moveTo'])
        if not has_revisions:
            continue

        root = etree.fromstring(raw)
        root = accept_revisions_in_tree(root)
        file_contents[part] = etree.tostring(root, xml_declaration=True,
                                              encoding='UTF-8', standalone=True)
        revisions_found += 1

    # Remove people.xml (revision authors)
    files_to_remove = set()
    if 'word/people.xml' in file_list:
        files_to_remove.add('word/people.xml')

    # Also remove people.xml from relationships
    rels_file = 'word/_rels/document.xml.rels'
    if rels_file in file_contents:
        rels_str = file_contents[rels_file].decode('utf-8')
        # Remove relationship entry for people.xml
        rels_str = re.sub(r'<Relationship[^>]*Target="people\.xml"[^>]*/>', '', rels_str)
        file_contents[rels_file] = rels_str.encode('utf-8')

    # Remove people.xml from [Content_Types].xml
    ct_file = '[Content_Types].xml'
    if ct_file in file_contents:
        ct_str = file_contents[ct_file].decode('utf-8')
        ct_str = re.sub(r'<Override[^>]*PartName="/word/people\.xml"[^>]*/>', '', ct_str)
        file_contents[ct_file] = ct_str.encode('utf-8')

    # Write back
    tmp_path = filepath + '.tmp'
    with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in file_list:
            if name in files_to_remove:
                continue
            zout.writestr(name, file_contents[name])

    os.replace(tmp_path, filepath)
    return revisions_found


def main():
    files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.docx')]
    print(f"Processing {len(files)} templates\n")

    for filename in sorted(files):
        filepath = os.path.join(INPUT_DIR, filename)
        try:
            revs = process_file(filepath)
            status = "CLEANED" if revs > 0 else "OK (no revisions)"
            print(f"{status}: {filename}")
        except Exception as e:
            print(f"ERROR: {filename} - {e}")

    print("\nDone!")


if __name__ == '__main__':
    main()
