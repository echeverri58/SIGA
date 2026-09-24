# -*- coding: utf-8 -*-
"""Crea Siga_planilla.xlsx: copia de Siga.xlsx con SOLO la hoja 'asistencia_mod'.
Sirve para que la conversion a PDF (Excel o LibreOffice) produzca solo la planilla."""
import zipfile, re, os

SRC = r'C:\Users\ASUS vivobook\Documents\AplicativoSiga\SIGA-Web\Siga.xlsx'
DST = r'C:\Users\ASUS vivobook\Documents\AplicativoSiga\SIGA-Web\Siga_planilla.xlsx'

z = zipfile.ZipFile(SRC)
orig = {n: z.read(n) for n in z.namelist()}

# ---------- 1) workbook.xml ----------
wb = orig['xl/workbook.xml'].decode('utf-8')
wb = re.sub(r'<sheets>.*?</sheets>',
            '<sheets><sheet name="asistencia_mod" sheetId="47" r:id="rId2"/></sheets>',
            wb, flags=re.S)

# definedNames: conservar solo los que apuntan a asistencia_mod (y renumerar la hoja a 0)
m = re.search(r'<definedNames>.*?</definedNames>', wb, re.S)
if m:
    bloque = m.group(0)
    keep = []
    for n in re.findall(r'<definedName[^>]*>.*?</definedName>', bloque, re.S):
        if 'asistencia_mod' in re.sub(r'<[^>]+>', '', n):
            keep.append(re.sub(r'localSheetId="\d+"', 'localSheetId="0"', n))
    wb = wb.replace(bloque, '<definedNames>' + ''.join(keep) + '</definedNames>' if keep else '')
orig['xl/workbook.xml'] = wb.encode('utf-8')

# ---------- 2) workbook.xml.rels ----------
rels = orig['xl/_rels/workbook.xml.rels'].decode('utf-8')
for rid in ('rId1', 'rId3', 'rId4'):        # hojas que se eliminan
    rels = re.sub(r'<Relationship Id="%s"[^>]*/>' % rid, '', rels)
orig['xl/_rels/workbook.xml.rels'] = rels.encode('utf-8')

# ---------- 3) sheet2.xml: quitar validaciones que apuntan a 'lista elegibles' ----------
s2 = orig['xl/worksheets/sheet2.xml'].decode('utf-8')
i = s2.rfind('<extLst>')
if i != -1 and s2.rstrip().endswith('</worksheet>'):
    s2 = s2[:i] + s2[s2.rfind('</worksheet>'):]
orig['xl/worksheets/sheet2.xml'] = s2.encode('utf-8')

# ---------- 4) [Content_Types].xml ----------
ct = orig['[Content_Types].xml'].decode('utf-8')
for parte in ('/xl/worksheets/sheet1.xml', '/xl/worksheets/sheet3.xml', '/xl/worksheets/sheet4.xml',
              '/xl/drawings/drawing1.xml', '/xl/drawings/drawing3.xml'):
    ct = re.sub(r'<Override PartName="%s"[^>]*/>' % re.escape(parte), '', ct)
orig['[Content_Types].xml'] = ct.encode('utf-8')

# ---------- 5) docProps/app.xml (metadatos de hojas) ----------
app = orig['docProps/app.xml'].decode('utf-8')
app = app.replace('<vt:vector size="4" baseType="variant">', '<vt:vector size="4" baseType="variant">')
app = re.sub(r'<HeadingPairs>.*?</HeadingPairs>',
             '<HeadingPairs><vt:vector size="4" baseType="variant">'
             '<vt:variant><vt:lpstr>Hojas de cálculo</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant>'
             '<vt:variant><vt:lpstr>Rangos con nombre</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant>'
             '</vt:vector></HeadingPairs>', app, flags=re.S)
app = re.sub(r'<TitlesOfParts>.*?</TitlesOfParts>',
             '<TitlesOfParts><vt:vector size="2" baseType="lpstr">'
             '<vt:lpstr>asistencia_mod</vt:lpstr>'
             '<vt:lpstr>asistencia_mod!Área_de_impresión</vt:lpstr>'
             '</vt:vector></TitlesOfParts>', app, flags=re.S)
orig['docProps/app.xml'] = app.encode('utf-8')

# ---------- 6) Escribir el nuevo archivo ----------
ELIMINAR = {
    'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet3.xml', 'xl/worksheets/sheet4.xml',
    'xl/worksheets/_rels/sheet1.xml.rels', 'xl/worksheets/_rels/sheet3.xml.rels',
    'xl/worksheets/_rels/sheet4.xml.rels',
    'xl/drawings/drawing1.xml', 'xl/drawings/drawing3.xml',
    'xl/drawings/_rels/drawing1.xml.rels', 'xl/drawings/_rels/drawing3.xml.rels',
    'xl/media/image1.png', 'xl/media/image3.png',
    'xl/printerSettings/printerSettings1.bin', 'xl/printerSettings/printerSettings3.bin',
}

if os.path.exists(DST):
    os.remove(DST)
with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as out:
    for nombre, datos in orig.items():
        if nombre in ELIMINAR:
            continue
        out.writestr(nombre, datos)

print('Creado:', DST)
print('Entradas:', len([n for n in orig if n not in ELIMINAR]))
