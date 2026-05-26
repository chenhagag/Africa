"""
Word Template Migration Script
Converts old SharePoint-bound templates to clean add-in-compatible templates.
See: # Word Template Migration Guide (Sh.txt
"""

import zipfile
import re
import os
import shutil
import uuid

# Tag mappings (old -> new)
TAG_MAPPINGS = {
    "cntTzadB_x002C__x0020_cntTzadC_x002C__x0020_cntTzadD": "cntTzadB",
    "cmtTzadAName": "cntPartyAName",
    "cntLocalAuth": "cntMunicipality",
    "cntContractType": "cntTemplateName",
    "_dlc_DocId": "cntContractNumber",
    "cntJobDesc": "cntWorkDescription",
}

# SharePoint customXml URIs to remove
SP_URIS = [
    "http://schemas.microsoft.com/office/2006/metadata/customXsn",
    "http://schemas.microsoft.com/office/2006/metadata/contentType",
]

INPUT_DIR = "Old Templates"
OUTPUT_DIR = "New Templates"


def clean_sdt_xml(xml_content):
    """Clean content controls in XML: remove bindings, fix tags, remove showingPlcHdr and text."""

    # 1. Remove <w:dataBinding .../> (self-closing, attributes may contain slashes in URLs)
    xml_content = re.sub(r'<w:dataBinding\b[^>]*?/>', '', xml_content)

    # 2. Remove <w:showingPlcHdr/>
    xml_content = re.sub(r'<w:showingPlcHdr/>', '', xml_content)

    # 3. Remove <w:text/> inside sdtPr blocks
    # We need to be careful to only remove <w:text/> that appears inside <w:sdtPr>
    def remove_text_in_sdtPr(match):
        sdtpr = match.group(0)
        sdtpr = re.sub(r'<w:text/>', '', sdtpr)
        return sdtpr

    xml_content = re.sub(r'<w:sdtPr>.*?</w:sdtPr>', remove_text_in_sdtPr, xml_content, flags=re.DOTALL)

    # 4. Apply tag mappings
    for old_tag, new_tag in TAG_MAPPINGS.items():
        xml_content = xml_content.replace(
            f'<w:tag w:val="{old_tag}"/>',
            f'<w:tag w:val="{new_tag}"/>'
        )

    # 5. Normalize placeholder docPart references to default
    xml_content = re.sub(
        r'<w:docPart w:val="[A-F0-9]{32}"/>',
        '<w:docPart w:val="DefaultPlaceholder_-1854013440"/>',
        xml_content
    )

    return xml_content


def is_sp_metadata_xml(content):
    """Check if a customXml item is SharePoint metadata that should be removed."""
    # Remove the large contentTypeSchema (item5.xml type - has contentTypeVersion="175" etc)
    if 'contentTypeSchema' in content and 'contentTypeVersion="' in content:
        # Check if version is high (SharePoint-generated, not the basic one)
        version_match = re.search(r'contentTypeVersion="(\d+)"', content)
        if version_match and int(version_match.group(1)) > 10:
            return True
    # Remove customXsn (SharePoint form config)
    if 'customXsn' in content and 'xsnLocation' in content:
        return True
    return False


def migrate_template(input_path, output_path):
    """Migrate a single .docx template."""

    # Read input
    with zipfile.ZipFile(input_path, 'r') as zin:
        file_list = zin.namelist()
        file_contents = {}
        for name in file_list:
            file_contents[name] = zin.read(name)

    # Track which customXml items to remove
    items_to_remove = set()

    # Check customXml items for SharePoint metadata
    for name in file_list:
        if name.startswith('customXml/item') and name.endswith('.xml') and 'Props' not in name:
            content = file_contents[name].decode('utf-8', errors='replace')
            if is_sp_metadata_xml(content):
                items_to_remove.add(name)
                # Also remove corresponding itemProps
                item_num = re.search(r'item(\d+)\.xml', name)
                if item_num:
                    props_name = f'customXml/itemProps{item_num.group(1)}.xml'
                    items_to_remove.add(props_name)
                    rels_name = f'customXml/_rels/item{item_num.group(1)}.xml.rels'
                    items_to_remove.add(rels_name)

    # Also remove [trash] folder if present
    trash_items = [n for n in file_list if n.startswith('[trash]')]
    items_to_remove.update(trash_items)

    # Process XML files that may contain content controls
    xml_parts = ['word/document.xml', 'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
                 'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml',
                 'word/endnotes.xml', 'word/footnotes.xml']

    for part in xml_parts:
        if part in file_contents:
            xml = file_contents[part].decode('utf-8')
            cleaned = clean_sdt_xml(xml)
            file_contents[part] = cleaned.encode('utf-8')

    # Write output
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in file_list:
            if name in items_to_remove:
                continue
            zout.writestr(name, file_contents[name])

    return True


def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    old_files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.docx')]

    print(f"Found {len(old_files)} templates to migrate\n")

    for filename in sorted(old_files):
        input_path = os.path.join(INPUT_DIR, filename)
        output_path = os.path.join(OUTPUT_DIR, filename)

        # Skip if already exists in output
        if os.path.exists(output_path):
            print(f"SKIP (exists): {filename}")
            continue

        try:
            migrate_template(input_path, output_path)
            print(f"OK: {filename}")
        except Exception as e:
            print(f"ERROR: {filename} - {e}")

    print("\nMigration complete!")


if __name__ == '__main__':
    main()
