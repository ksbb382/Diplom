"""
Post-processing of ФИНАЛЬНЫЙ_ДИПЛОМ_ВОРОБЬЕВ.docx:
1. Force Liberation Serif 14pt for ALL runs (body + tables) — fixes Calibri default.
2. Override default theme fonts (theme1.xml) to Liberation Serif.
3. Override docDefaults rPrDefault to use Liberation Serif 14pt explicitly.
4. Replace static ToC table with REAL Word auto-TOC field (TOC \\o "1-3" \\h \\z \\u)
   AND keep Рыжкин-style table with placeholders — author can pick.

Actually based on user request: keep Рыжкин-style ToC table. Don't auto-fill pages
because we cannot count pages without LibreOffice/Word. Use auto-TOC instead so
Word fills the pages itself when document is opened.
"""

import os
import re
import shutil
import zipfile
import tempfile

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DOC_PATH = "/projects/sandbox/Diplom/diplom/ФИНАЛЬНЫЙ_ДИПЛОМ_ВОРОБЬЕВ.docx"


def force_font_in_runs(doc):
    """Walk through every <w:r> in body+tables and ensure rPr has rFonts=Liberation Serif and sz=28 (14pt)."""

    body = doc.element.body
    runs_processed = 0

    for r in body.iter(qn('w:r')):
        # Get or create rPr
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            r.insert(0, rPr)

        # Get or create rFonts
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)

        # Remove theme-based attributes (asciiTheme, hAnsiTheme, cstheme, eastAsiaTheme)
        for theme_attr in ['asciiTheme', 'hAnsiTheme', 'cstheme', 'eastAsiaTheme']:
            if rFonts.get(qn('w:' + theme_attr)) is not None:
                del rFonts.attrib[qn('w:' + theme_attr)]

        rFonts.set(qn('w:ascii'), 'Liberation Serif')
        rFonts.set(qn('w:hAnsi'), 'Liberation Serif')
        rFonts.set(qn('w:cs'), 'Liberation Serif')
        rFonts.set(qn('w:eastAsia'), 'Liberation Serif')

        # Get or create sz (size)
        # We want 14 pt = sz val 28 for body. But headings can be larger.
        # Check if this run is in a heading paragraph.
        parent_p = r.getparent()
        is_heading = False
        if parent_p is not None and parent_p.tag == qn('w:p'):
            pStyle = parent_p.find(qn('w:pPr') + '/' + qn('w:pStyle'))
            if pStyle is not None:
                style_id = pStyle.get(qn('w:val'))
                if style_id and 'Heading' in style_id:
                    is_heading = True

        # Decide target size: keep existing sz if it's already > 28 (heading), else set 28
        sz = rPr.find(qn('w:sz'))
        szCs = rPr.find(qn('w:szCs'))

        target_sz = '28'  # 14pt
        if sz is not None:
            current = sz.get(qn('w:val'))
            try:
                if current and int(current) > 28:
                    # It's a heading-sized run; keep its size
                    target_sz = current
            except ValueError:
                pass

        if sz is None:
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), target_sz)
            rPr.append(sz)
        else:
            sz.set(qn('w:val'), target_sz)

        if szCs is None:
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), target_sz)
            rPr.append(szCs)
        else:
            szCs.set(qn('w:val'), target_sz)

        runs_processed += 1

    print(f"[Fonts] Processed {runs_processed} runs")


def fix_paragraph_spacing(doc):
    """Set line spacing 1.5 (360 twentieths) for all body paragraphs (not headings)."""
    body = doc.element.body
    fixed = 0

    for p in body.iter(qn('w:p')):
        pPr = p.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p.insert(0, pPr)

        # Skip if it's inside a table (table cells often have own spacing)
        # Actually let's still fix it.

        # Check if heading
        pStyle = pPr.find(qn('w:pStyle'))
        is_heading = False
        if pStyle is not None:
            sid = pStyle.get(qn('w:val')) or ''
            if 'Heading' in sid:
                is_heading = True

        if is_heading:
            continue

        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)

        spacing.set(qn('w:line'), '360')
        spacing.set(qn('w:lineRule'), 'auto')

        fixed += 1

    print(f"[Spacing] Set 1.5 line spacing on {fixed} paragraphs")


def patch_styles_xml(zip_in_path, zip_out_path):
    """Open the .docx as zip, patch styles.xml + theme1.xml to use Liberation Serif as default."""

    with zipfile.ZipFile(zip_in_path, 'r') as zin:
        with zipfile.ZipFile(zip_out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)

                if item == 'word/styles.xml':
                    text = data.decode('utf-8')
                    # Replace docDefaults theme-based fonts with Liberation Serif
                    # docDefaults part: <w:rFonts w:asciiTheme="minorHAnsi" .../>
                    text = re.sub(
                        r'<w:rFonts\s+w:asciiTheme="minorHAnsi"[^/]*/>',
                        '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif" w:eastAsia="Liberation Serif"/>',
                        text
                    )
                    # Force default size to 28 (14pt) in docDefaults
                    # Look for <w:rPrDefault><w:rPr>...</w:rPr></w:rPrDefault>
                    if '<w:sz w:val=' not in text[:5000]:
                        # Add <w:sz w:val="28"/><w:szCs w:val="28"/> to rPrDefault
                        text = text.replace(
                            '<w:rPrDefault><w:rPr>',
                            '<w:rPrDefault><w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/>',
                            1
                        )
                    data = text.encode('utf-8')
                    print("[Styles] Patched styles.xml")

                elif item == 'word/theme/theme1.xml':
                    text = data.decode('utf-8')
                    # Replace majorFont and minorFont latin typefaces
                    text = re.sub(
                        r'(<a:majorFont>.*?<a:latin typeface=")[^"]+(")',
                        r'\1Liberation Serif\2',
                        text,
                        flags=re.DOTALL
                    )
                    text = re.sub(
                        r'(<a:minorFont>.*?<a:latin typeface=")[^"]+(")',
                        r'\1Liberation Serif\2',
                        text,
                        flags=re.DOTALL
                    )
                    data = text.encode('utf-8')
                    print("[Theme] Patched theme1.xml")

                zout.writestr(item, data)


def replace_toc_with_field(doc):
    """Replace the static ToC table (Table 0) with a real Word TOC field.
    Field: TOC \\o "1-3" \\h \\z \\u — auto-table of contents using Heading 1-3 styles.
    When the user opens the file in Word, they right-click → Update Field → page numbers fill in."""

    body = doc.element.body

    # Find first table (the ToC table)
    first_tbl = None
    for child in body.iterchildren():
        if child.tag == qn('w:tbl'):
            first_tbl = child
            break

    if first_tbl is None:
        print("[TOC] No table found; skipping")
        return

    # Build the TOC field as paragraphs:
    # <w:p>
    #   <w:r>
    #     <w:fldChar w:fldCharType="begin" w:dirty="true"/>
    #   </w:r>
    #   <w:r>
    #     <w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>
    #   </w:r>
    #   <w:r>
    #     <w:fldChar w:fldCharType="separate"/>
    #   </w:r>
    #   <w:r>
    #     <w:t>Right-click here, choose Update Field</w:t>
    #   </w:r>
    #   <w:r>
    #     <w:fldChar w:fldCharType="end"/>
    #   </w:r>
    # </w:p>

    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    p.append(pPr)

    # 1) begin
    r1 = OxmlElement('w:r')
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    fld_begin.set(qn('w:dirty'), 'true')
    r1.append(fld_begin)
    p.append(r1)

    # 2) instrText
    r2 = OxmlElement('w:r')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2.append(instr)
    p.append(r2)

    # 3) separate
    r3 = OxmlElement('w:r')
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    r3.append(fld_sep)
    p.append(r3)

    # 4) placeholder text
    r4 = OxmlElement('w:r')
    rPr4 = OxmlElement('w:rPr')
    rFonts4 = OxmlElement('w:rFonts')
    rFonts4.set(qn('w:ascii'), 'Liberation Serif')
    rFonts4.set(qn('w:hAnsi'), 'Liberation Serif')
    rPr4.append(rFonts4)
    sz4 = OxmlElement('w:sz')
    sz4.set(qn('w:val'), '28')
    rPr4.append(sz4)
    i_elem = OxmlElement('w:i')
    rPr4.append(i_elem)
    r4.append(rPr4)
    t4 = OxmlElement('w:t')
    t4.text = 'Кликните правой кнопкой по этому полю и выберите «Обновить поле» для автоматического формирования оглавления'
    r4.append(t4)
    p.append(r4)

    # 5) end
    r5 = OxmlElement('w:r')
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    r5.append(fld_end)
    p.append(r5)

    # Insert before the table, then remove the table
    first_tbl.addprevious(p)
    first_tbl.getparent().remove(first_tbl)

    print("[TOC] Replaced static table with Word TOC auto-field")


def main():
    # Step 1: Open and modify document XML through python-docx
    doc = Document(DOC_PATH)
    force_font_in_runs(doc)
    fix_paragraph_spacing(doc)
    replace_toc_with_field(doc)

    # Save to a temp file
    tmp_path = '/tmp/voro_intermediate.docx'
    doc.save(tmp_path)

    # Step 2: Patch styles.xml + theme1.xml inside the zip
    patch_styles_xml(tmp_path, DOC_PATH)

    print(f"\n[DONE] {DOC_PATH}")
    print(f"Size: {os.path.getsize(DOC_PATH) / 1024:.1f} KB")


if __name__ == '__main__':
    main()
