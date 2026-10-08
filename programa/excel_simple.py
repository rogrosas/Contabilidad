# -*- coding: utf-8 -*-
"""Escritor mínimo de archivos .xlsx usando solo la biblioteca estándar de Python.

Soporta varias hojas, estilos básicos (encabezados, moneda, fechas, porcentajes,
totales), fórmulas con valor precalculado, ancho de columnas, fila congelada y
autofiltro. Suficiente para reportes contables sin instalar openpyxl.
"""
import datetime as dt
import re
import zipfile
from xml.sax.saxutils import escape

# Índices de estilo (coinciden con el orden de <cellXfs> en _STYLES)
NORMAL, HEADER, MONEY, DATE, MONEY_TOTAL, QTY, TOTAL, TITLE, PCT, PCT_TOTAL, QTY_TOTAL, SUBTITLE = range(12)

_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_EPOCH = dt.date(1899, 12, 30)


class F:
    """Fórmula (sin el '=') con su valor ya calculado, para que se vea aunque no se recalcule."""

    def __init__(self, formula, valor=None):
        self.formula = formula
        self.valor = valor


def col(i):
    """Índice de columna 0-based -> letra (0 -> A, 26 -> AA)."""
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _num(v):
    if isinstance(v, bool):
        return str(int(v))
    if isinstance(v, int):
        return str(v)
    return repr(round(float(v), 10))


def _celda(fila, c, v, s):
    ref = f"{col(c)}{fila}"
    a = f' r="{ref}"' + (f' s="{s}"' if s else "")
    if v is None or v == "":
        return f"<c{a}/>" if s else ""
    if isinstance(v, F):
        val = "" if v.valor is None else f"<v>{_num(v.valor)}</v>"
        return f"<c{a}><f>{escape(v.formula)}</f>{val}</c>"
    if isinstance(v, dt.datetime):
        d = v - dt.datetime(1899, 12, 30)
        return f"<c{a}><v>{d.days + d.seconds / 86400:.6f}</v></c>"
    if isinstance(v, dt.date):
        return f"<c{a}><v>{(v - _EPOCH).days}</v></c>"
    if isinstance(v, (int, float)):
        return f"<c{a}><v>{_num(v)}</v></c>"
    txt = escape(_CTRL.sub("", str(v)))
    return f'<c{a} t="inlineStr"><is><t xml:space="preserve">{txt}</t></is></c>'


class Hoja:
    def __init__(self, nombre):
        self.nombre = nombre
        self.filas = []
        self.anchos = {}      # {indice_columna: ancho}
        self.congelar = 0     # cantidad de filas superiores congeladas
        self.filtro = None    # rango de autofiltro, ej. "A1:H50"

    def fila(self, *celdas):
        """Agrega una fila. Cada celda es un valor o una tupla (valor, estilo). Devuelve el nº de fila."""
        self.filas.append(celdas)
        return len(self.filas)

    def _xml(self, seleccionada):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
        out.append('<sheetViews><sheetView workbookViewId="0"' + (' tabSelected="1"' if seleccionada else "") + ">")
        if self.congelar:
            n = self.congelar
            out.append(f'<pane ySplit="{n}" topLeftCell="A{n + 1}" activePane="bottomLeft" state="frozen"/>'
                       '<selection pane="bottomLeft"/>')
        out.append("</sheetView></sheetViews>")
        if self.anchos:
            out.append("<cols>")
            for i in sorted(self.anchos):
                out.append(f'<col min="{i + 1}" max="{i + 1}" width="{self.anchos[i]}" customWidth="1"/>')
            out.append("</cols>")
        out.append("<sheetData>")
        for r, celdas in enumerate(self.filas, start=1):
            partes = []
            for c, cel in enumerate(celdas):
                v, s = cel if isinstance(cel, tuple) else (cel, NORMAL)
                partes.append(_celda(r, c, v, s))
            out.append(f'<row r="{r}">{"".join(partes)}</row>')
        out.append("</sheetData>")
        if self.filtro:
            out.append(f'<autoFilter ref="{self.filtro}"/>')
        out.append("</worksheet>")
        return "".join(out)


class Libro:
    def __init__(self):
        self.hojas = []

    def hoja(self, nombre):
        limpio = "".join("_" if ch in '[]:*?/\\' else ch for ch in nombre)[:31] or f"Hoja{len(self.hojas) + 1}"
        h = Hoja(limpio)
        self.hojas.append(h)
        return h

    def guardar(self, ruta):
        n = len(self.hojas)
        ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
              '<Default Extension="xml" ContentType="application/xml"/>'
              '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
              '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
        for i in range(1, n + 1):
            ct.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                      'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        ct.append("</Types>")

        rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                '</Relationships>')

        wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>']
        for i, h in enumerate(self.hojas, start=1):
            wb.append(f'<sheet name="{escape(h.nombre, {chr(34): "&quot;"})}" sheetId="{i}" r:id="rId{i}"/>')
        wb.append("</sheets>")
        filtros = [(i, h) for i, h in enumerate(self.hojas) if h.filtro]
        if filtros:
            wb.append("<definedNames>")
            for i, h in filtros:
                a, b = h.filtro.split(":")
                abs_ref = lambda r: re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", r)
                nombre = h.nombre.replace("'", "''")
                wb.append(f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i}" hidden="1">'
                          f"{escape(chr(39) + nombre + chr(39))}!{abs_ref(a)}:{abs_ref(b)}</definedName>")
            wb.append("</definedNames>")
        wb.append('<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')

        wbrels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for i in range(1, n + 1):
            wbrels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
        wbrels.append(f'<Relationship Id="rId{n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
        wbrels.append("</Relationships>")

        with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", "".join(ct))
            z.writestr("_rels/.rels", rels)
            z.writestr("xl/workbook.xml", "".join(wb))
            z.writestr("xl/_rels/workbook.xml.rels", "".join(wbrels))
            z.writestr("xl/styles.xml", _STYLES)
            for i, h in enumerate(self.hojas, start=1):
                z.writestr(f"xl/worksheets/sheet{i}.xml", h._xml(i == 1))


_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy"/></numFmts>
<fonts count="5">
<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="16"/><color rgb="FF1F3864"/><name val="Calibri"/><family val="2"/></font>
<font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/><family val="2"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF305496"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left/><right/><top style="thin"><color rgb="FF305496"/></top><bottom style="double"><color rgb="FF305496"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="12">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="10" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""
