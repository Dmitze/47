import zipfile
import io
import os
import xml.sax.saxutils as saxutils

def escape(s):
    if s is None:
        return ""
    return saxutils.escape(str(s))

def col_letter(col_idx):
    """1 -> A, 2 -> B, ..., 26 -> Z, 27 -> AA"""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

class SimpleXlsx:
    def __init__(self):
        self.sheets = [] # list of dicts: {'name': str, 'rows': list of dicts, 'cols': list of dicts}

    def add_sheet(self, name, col_widths=None):
        sheet = {
            'name': name,
            'col_widths': col_widths or {},
            'rows': [] # list of list of cell dicts: {'val': ..., 'type': 'str'|'num'|'formula', 'style': int}
        }
        self.sheets.append(sheet)
        return sheet

    def save(self, filepath):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
            # 1. [Content_Types].xml
            ct_parts = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                '  <Default Extension="xml" ContentType="application/xml"/>',
                '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
                '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            ]
            for idx in range(1, len(self.sheets) + 1):
                ct_parts.append(f'  <Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            ct_parts.append('</Types>')
            z.writestr('[Content_Types].xml', '\n'.join(ct_parts))

            # 2. _rels/.rels
            z.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''')

            # 3. xl/_rels/workbook.xml.rels
            wb_rels = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            ]
            for idx in range(1, len(self.sheets) + 1):
                wb_rels.append(f'  <Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>')
            wb_rels.append(f'  <Relationship Id="rId{len(self.sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
            wb_rels.append('</Relationships>')
            z.writestr('xl/_rels/workbook.xml.rels', '\n'.join(wb_rels))

            # 4. xl/workbook.xml
            wb_parts = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
                '  <sheets>'
            ]
            for idx, s in enumerate(self.sheets, 1):
                wb_parts.append(f'    <sheet name="{escape(s["name"])}" sheetId="{idx}" r:id="rId{idx}"/>')
            wb_parts.append('  </sheets>')
            wb_parts.append('</workbook>')
            z.writestr('xl/workbook.xml', '\n'.join(wb_parts))

            # 5. xl/styles.xml
            # Styles:
            # 0: Normal (Calibri 11)
            # 1: Title Header (Calibri 14 Bold, Navy fill #1F4E79, White text)
            # 2: Section Header (Calibri 11 Bold, Slate fill #2F5597, White text, center)
            # 3: Table Header (Calibri 10 Bold, Muted Blue #41719C, White text, center, border)
            # 4: Data Text (Calibri 10, border, left)
            # 5: Data Text Center (Calibri 10, border, center)
            # 6: Data Number (Calibri 10, border, right)
            # 7: Data Currency (Calibri 10, border, right, format #,##0.00)
            # 8: Total Bold (Calibri 10 Bold, Yellow fill #FFF2CC, border, center)
            # 9: Total Bold Number (Calibri 10 Bold, Yellow fill #FFF2CC, border, right)
            # 10: Alert Green (Calibri 10 Bold, Green fill #E2EFDA, text #385723, center, border)
            # 11: Alert Red (Calibri 10 Bold, Red fill #FCE4D6, text #C00000, center, border)
            # 12: Note Info (Calibri 10 Italic, Gray #595959)
            styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
  </numFmts>
  <fonts count="6">
    <font><sz val="10"/><name val="Calibri"/></font> <!-- 0: normal -->
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font> <!-- 1: title -->
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font> <!-- 2: header white -->
    <font><b/><sz val="10"/><name val="Calibri"/></font> <!-- 3: bold 10 -->
    <font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/></font> <!-- 4: italic gray -->
    <font><b/><sz val="10"/><color rgb="FFC00000"/><name val="Calibri"/></font> <!-- 5: alert red text -->
  </fonts>
  <fills count="8">
    <fill><patternFill patternType="none"/></fill> <!-- 0 -->
    <fill><patternFill patternType="gray125"/></fill> <!-- 1 -->
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill> <!-- 2: dark navy -->
    <fill><patternFill patternType="solid"><fgColor rgb="FF2F5597"/></patternFill></fill> <!-- 3: slate blue -->
    <fill><patternFill patternType="solid"><fgColor rgb="FF41719C"/></patternFill></fill> <!-- 4: medium blue header -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill> <!-- 5: light yellow total -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill> <!-- 6: light green -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/></patternFill></fill> <!-- 7: light red/peach -->
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border> <!-- 0: none -->
    <border> <!-- 1: thin all around -->
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="13">
    <!-- 0: default -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <!-- 1: Title Header -->
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 2: Section Subheader -->
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 3: Table Header Blue -->
    <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <!-- 4: Data Text Left -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <!-- 5: Data Text Center -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 6: Data Number Right -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 7: Data Currency -->
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 8: Total Center -->
    <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 9: Total Number Right -->
    <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <!-- 10: Alert Green -->
    <xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 11: Alert Red -->
    <xf numFmtId="0" fontId="5" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <!-- 12: Info italic -->
    <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>'''
            z.writestr('xl/styles.xml', styles_xml)

            # 6. Worksheets
            for idx, sheet in enumerate(self.sheets, 1):
                ws_parts = [
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                ]
                
                # Column widths
                if sheet['col_widths']:
                    ws_parts.append('  <cols>')
                    for col_num, width in sorted(sheet['col_widths'].items()):
                        ws_parts.append(f'    <col min="{col_num}" max="{col_num}" width="{width}" customWidth="1"/>')
                    ws_parts.append('  </cols>')
                
                ws_parts.append('  <sheetData>')
                for r_idx, row in enumerate(sheet['rows'], 1):
                    if not row:
                        ws_parts.append(f'    <row r="{r_idx}"/>')
                        continue
                    ws_parts.append(f'    <row r="{r_idx}">')
                    for c_idx, cell in enumerate(row, 1):
                        if cell is None:
                            continue
                        cell_ref = f"{col_letter(c_idx)}{r_idx}"
                        val = cell.get('val', '')
                        cell_type = cell.get('type', 'str')
                        style = cell.get('style', 0)

                        if cell_type == 'formula':
                            f_text = escape(val)
                            ws_parts.append(f'      <c r="{cell_ref}" s="{style}"><f>{f_text}</f></c>')
                        elif cell_type == 'num':
                            ws_parts.append(f'      <c r="{cell_ref}" s="{style}"><v>{val}</v></c>')
                        else: # inlineStr
                            txt = escape(val)
                            ws_parts.append(f'      <c r="{cell_ref}" t="inlineStr" s="{style}"><is><t>{txt}</t></is></c>')
                    ws_parts.append('    </row>')
                ws_parts.append('  </sheetData>')
                ws_parts.append('</worksheet>')
                z.writestr(f'xl/worksheets/sheet{idx}.xml', '\n'.join(ws_parts))

        with open(filepath, 'wb') as f:
            f.write(buf.getvalue())
        print(f"Workbook successfully saved to {filepath}")

print("SimpleXlsx engine defined")
