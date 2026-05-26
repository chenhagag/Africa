"""
Fix Old Templates — replicate what the add-in's insertContentControl() API creates.

For each SDT:
1. Extract the rPr (formatting) from the original sdtPr
2. Extract alias, tag (rename if needed)
3. Determine if block-level (has <w:p>) or inline
4. For block-level: preserve the original <w:pPr> from the content
5. Build a new clean SDT matching the API-created structure
6. Replace the old SDT with the new one

Also: remove SharePoint customXml, normalize placeholders.
"""
import zipfile
import re
import os
import random

TAG_MAPPINGS = {
    "cntTzadB_x002C__x0020_cntTzadC_x002C__x0020_cntTzadD": "cntTzadB",
    "cmtTzadAName": "cntPartyAName",
    "cntLocalAuth": "cntMunicipality",
    "cntContractType": "cntTemplateName",
    "_dlc_DocId": "cntContractNumber",
    "cntJobDesc": "cntWorkDescription",
}

# label lookup (from FIELD_CATALOG in taskpane.ts)
TAG_LABELS = {
    "cntContractNumber": "מספר חוזה",
    "cntContractVersion": "גרסת חוזה",
    "cntTemplateName": "תבנית",
    "cntProjectName": "שם הפרויקט",
    "cntSite": "אתר",
    "cntSupllierName": "ספק",
    "cntMunicipality": "רשות מקומית",
    "cntWorkDescription": "תיאור העבודה",
    "cntSignDate": "תאריך חתימה החוזה",
    "cntStartDate": "תאריך התחלה",
    "cntDurationMonths": "משך בחודשים",
    "cntExpectedEndDate": "תאריך סיום משוער",
    "cntStatus": "סטטוס",
    "cntTzadA": "צד א מסכם",
    "cntTzadB": "צדדים נוספים",
    "cntLocalAuth": "רשות מקומית",
    "cmtTzadAPercent": "צד א שם ואחוז",
    "cntMadadTypeTitle": "סוג מדד",
    "cntIsKnownTitle": "מדד בגין/ידוע",
    "cntMadadBase": "מדד בסיס",
    "cntMadadPoints": "נקודות מדד",
    "cntJobDesc": "תיאור העבודה",
    "cntPartyAName": "שם צד א",
    "cntCustomField1": "שדה מותאם 1",
    "cntCustomField2": "שדה מותאם 2",
    "cntCustomField3": "שדה מותאם 3",
    "cntCustomField4": "שדה מותאם 4",
    "cntCustomField5": "שדה מותאם 5",
    "cntCustomField6": "שדה מותאם 6",
    "cntCustomField7": "שדה מותאם 7",
    "cntCustomField8": "שדה מותאם 8",
}


def gen_hex8():
    return format(random.randint(0, 0xFFFFFFFF), '08X')


def gen_sdt_id():
    return str(random.randint(-2000000000, 2000000000))


def replace_sdts_in_xml(xml_bytes):
    """Replace each SDT with a clean API-style SDT."""
    text = xml_bytes.decode('utf-8')
    original = text

    # Collect existing paraIds to avoid duplicates
    existing_ids = set(re.findall(r'w14:paraId="([^"]+)"', text))

    def unique_hex(existing):
        while True:
            h = gen_hex8()
            if h not in existing:
                existing.add(h)
                return h

    def replace_one_sdt(match):
        full_sdt = match.group(0)

        # Extract sdtPr
        pr_match = re.search(r'<w:sdtPr>(.*?)</w:sdtPr>', full_sdt, re.DOTALL)
        if not pr_match:
            return full_sdt
        sdt_pr = pr_match.group(1)

        # Extract tag
        tag_match = re.search(r'<w:tag w:val="([^"]*)"', sdt_pr)
        if not tag_match:
            return full_sdt
        old_tag = tag_match.group(1)
        tag = TAG_MAPPINGS.get(old_tag, old_tag)

        # Get label from alias or lookup
        alias_match = re.search(r'<w:alias w:val="([^"]*)"', sdt_pr)
        alias = alias_match.group(1) if alias_match else TAG_LABELS.get(tag, tag)

        # Get label for placeholder text
        label = TAG_LABELS.get(tag, alias)

        # Extract rPr from sdtPr (the formatting)
        rpr_match = re.search(r'(<w:rPr>.*?</w:rPr>)', sdt_pr, re.DOTALL)
        rpr = rpr_match.group(1) if rpr_match else '<w:rPr><w:rFonts w:cs="David"/><w:color w:val="00B0F0"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>'

        # Ensure rPr has color 00B0F0 (the API always sets this)
        if '<w:color' not in rpr:
            rpr = rpr.replace('</w:rPr>', '<w:color w:val="00B0F0"/></w:rPr>')

        # Extract sdtContent to determine block vs inline and get pPr
        content_match = re.search(r'<w:sdtContent>(.*?)</w:sdtContent>', full_sdt, re.DOTALL)
        if not content_match:
            return full_sdt
        content = content_match.group(1)

        is_block = '<w:p ' in content or '<w:p>' in content

        sdt_id = gen_sdt_id()

        if is_block:
            # Extract pPr from original paragraph (preserve formatting)
            ppr_match = re.search(r'(<w:pPr>.*?</w:pPr>)', content, re.DOTALL)
            ppr = ppr_match.group(1) if ppr_match else ''

            para_id = unique_hex(existing_ids)
            text_id = unique_hex(existing_ids)

            new_sdt = (
                f'<w:sdt><w:sdtPr>{rpr}'
                f'<w:alias w:val="{alias}"/>'
                f'<w:tag w:val="{tag}"/>'
                f'<w:id w:val="{sdt_id}"/>'
                f'<w:placeholder><w:docPart w:val="DefaultPlaceholder_-1854013440"/></w:placeholder>'
                f'</w:sdtPr><w:sdtEndPr/><w:sdtContent>'
                f'<w:p w14:paraId="{para_id}" w14:textId="{text_id}">'
                f'{ppr}'
                f'<w:r>{rpr}<w:t>[{label}]</w:t></w:r>'
                f'</w:p></w:sdtContent></w:sdt>'
            )
        else:
            new_sdt = (
                f'<w:sdt><w:sdtPr>{rpr}'
                f'<w:alias w:val="{alias}"/>'
                f'<w:tag w:val="{tag}"/>'
                f'<w:id w:val="{sdt_id}"/>'
                f'<w:placeholder><w:docPart w:val="DefaultPlaceholder_-1854013440"/></w:placeholder>'
                f'</w:sdtPr><w:sdtEndPr/><w:sdtContent>'
                f'<w:r>{rpr}<w:t>[{label}]</w:t></w:r>'
                f'</w:sdtContent></w:sdt>'
            )

        return new_sdt

    text = re.sub(r'<w:sdt>.*?</w:sdt>', replace_one_sdt, text, flags=re.DOTALL)

    if text != original:
        return text.encode('utf-8')
    return None


def is_sp_metadata_xml(content):
    if 'contentTypeSchema' in content and 'contentTypeVersion="' in content:
        version_match = re.search(r'contentTypeVersion="(\d+)"', content)
        if version_match and int(version_match.group(1)) > 10:
            return True
    if 'customXsn' in content and 'xsnLocation' in content:
        return True
    return False


def process_docx(input_path, output_path):
    """Process a docx: input -> output (does not modify input)."""

    with zipfile.ZipFile(input_path, 'r') as zin:
        file_list = zin.namelist()
        file_contents = {n: zin.read(n) for n in file_list}

    # Find SharePoint customXml to remove
    items_to_remove = set()
    for name in file_list:
        if name.startswith('customXml/item') and name.endswith('.xml') and 'Props' not in name:
            content = file_contents[name].decode('utf-8', errors='replace')
            if is_sp_metadata_xml(content):
                items_to_remove.add(name)
                item_num = re.search(r'item(\d+)\.xml', name)
                if item_num:
                    items_to_remove.add(f'customXml/itemProps{item_num.group(1)}.xml')
                    items_to_remove.add(f'customXml/_rels/item{item_num.group(1)}.xml.rels')

    items_to_remove.update(n for n in file_list if n.startswith('[trash]'))

    # Replace SDTs in relevant XML parts
    xml_parts = [
        'word/document.xml',
        'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
        'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml',
        'word/endnotes.xml', 'word/footnotes.xml',
    ]

    for part in xml_parts:
        if part in file_contents:
            fixed = replace_sdts_in_xml(file_contents[part])
            if fixed:
                file_contents[part] = fixed

    # Fix Content_Types - remove refs to deleted customXml items
    if '[Content_Types].xml' in file_contents:
        ct = file_contents['[Content_Types].xml'].decode('utf-8')
        for name in items_to_remove:
            m = re.search(r'itemProps(\d+)\.xml', name)
            if m:
                ct = re.sub(
                    rf'<Override[^>]*PartName="/customXml/itemProps{m.group(1)}\.xml"[^>]*/>\s*',
                    '', ct)
        file_contents['[Content_Types].xml'] = ct.encode('utf-8')

    # Fix document.xml.rels
    if 'word/_rels/document.xml.rels' in file_contents:
        rels = file_contents['word/_rels/document.xml.rels'].decode('utf-8')
        for name in items_to_remove:
            m = re.search(r'item(\d+)\.xml', name)
            if m and 'Props' not in name and '_rels' not in name:
                rels = re.sub(
                    rf'<Relationship[^>]*Target="[^"]*customXml/item{m.group(1)}\.xml"[^>]*/>\s*',
                    '', rels)
        file_contents['word/_rels/document.xml.rels'] = rels.encode('utf-8')

    # Write output
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in file_list:
            if name in items_to_remove:
                continue
            zout.writestr(name, file_contents[name])


if __name__ == '__main__':
    base = os.path.dirname(os.path.abspath(__file__))
    old_dir = os.path.join(base, 'Old Templates')
    out_dir = os.path.join(base, 'NewFixedTemplates')

    target = 'אדריכל ראשי.docx'
    input_path = os.path.join(old_dir, target)
    output_path = os.path.join(out_dir, target)

    print(f'Processing: {target}')
    process_docx(input_path, output_path)
    print(f'Output: {output_path}')

    # Verify
    z = zipfile.ZipFile(output_path, 'r')
    doc = z.read('word/document.xml').decode('utf-8')
    print()
    print(f'SDTs: {doc.count("<w:sdt>")}')
    print(f'dataBinding: {doc.count("dataBinding")}')
    print(f'sdtEndPr: {doc.count("sdtEndPr")}')
    print(f'DefaultPlaceholder: {doc.count("DefaultPlaceholder")}')
    print(f'<w:date: {doc.count("<w:date")}')
    print(f'<w:text/>: {doc.count("<w:text/>")}')
    print(f'showingPlcHdr: {doc.count("showingPlcHdr")}')
    print(f'color 00B0F0: {doc.count("00B0F0")}')

    tags = re.findall(r'<w:tag w:val="([^"]+)"', doc)
    print(f'Tags: {sorted(set(tags))}')
