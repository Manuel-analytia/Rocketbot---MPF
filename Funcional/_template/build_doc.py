"""
build_doc.py - Generador de documentos corporativos MPF.

Lee `style.json` (paleta + fuente + margenes + footer) y produce un .docx con:
  - Pagina de portada con logo MPF + titulo + subtitulo + fecha
  - Tabla de contenido (campo TOC; Word lo regenera al abrir o con F9)
  - Pagina de cuerpo con headings (Heading 1/2/3 nativos para que entren al TOC)
  - Pie de pagina con numero de pagina + logo Analytia (desarrollador)
  - Tipografia Garamond y interlineado 1.5 globales

Uso programatico:

    from build_doc import DocumentBuilder
    db = DocumentBuilder()
    db.add_cover(title='...', subtitle='...', date='...')
    db.add_toc()                # TOC despues del cover
    db.h1('1. Titulo')
    db.p('Parrafo...')
    db.h2('1.1 Sub')
    db.bullet('item')
    db.code('comando')
    db.save('out.docx')

CLI:
    python build_doc.py demo                # genera demo.docx con todos los estilos
"""
from __future__ import annotations
import json
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


HERE = Path(__file__).resolve().parent
STYLE_PATH = HERE / 'style.json'
LOGO_PATH = HERE / 'logo.png'
LOGO_DEV_PATH = HERE / 'logo_analytia.png'


def _hex_to_rgb(h: str) -> RGBColor:
    h = h.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_run_font(run, name: str) -> None:
    """Forzar la fuente para ascii / hAnsi / cs / eastAsia (Word a veces ignora 'name')."""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rFonts.set(qn(attr), name)


class DocumentBuilder:
    """Wrapper minimal sobre python-docx con la paleta MPF aplicada."""

    def __init__(self, style_path: Path | str = STYLE_PATH):
        self.style = json.loads(Path(style_path).read_text(encoding='utf-8'))
        self.doc = Document()
        self._apply_page()
        self._apply_default_font_and_spacing()
        self._apply_heading_styles()
        self._setup_footer()

    @classmethod
    def from_existing(cls, docx_path: Path | str,
                      style_path: Path | str = STYLE_PATH):
        """Carga un docx existente y borra todo desde la primera H1 hacia
        abajo (preservando lo de arriba: portada, tabla de contenido manual,
        encabezados, etc.). Despues se puede appender body con h1/h2/p/etc.

        Si NO encuentra ninguna H1, no borra nada (apendea al final)."""
        inst = cls.__new__(cls)
        inst.style = json.loads(Path(style_path).read_text(encoding='utf-8'))
        inst.doc = Document(str(docx_path))
        # Reaplicar estilos Normal y Heading 1/2/3 desde style.json (source of truth).
        # Esto NO toca la TOC del usuario (la TOC usa TOC 1/2/3, no estos estilos).
        inst._apply_default_font_and_spacing()
        inst._apply_heading_styles()

        body = inst.doc.element.body
        # Buscar primera H1
        first_h1_el = None
        for p in inst.doc.paragraphs:
            if p.style.name == 'Heading 1':
                first_h1_el = p._element
                break
        if first_h1_el is not None:
            # Borrar todos los hijos del body desde la H1 hasta el final,
            # excepto la <w:sectPr> de cierre.
            children = list(body)
            try:
                idx = children.index(first_h1_el)
            except ValueError:
                return inst
            for child in children[idx:]:
                tag = child.tag.split('}')[-1]
                if tag == 'sectPr':
                    continue
                body.remove(child)
        return inst

    # ---------- setup ----------
    def _apply_page(self) -> None:
        page = self.style['page']
        for section in self.doc.sections:
            section.page_width = Inches(page['width_in'])
            section.page_height = Inches(page['height_in'])
            section.top_margin = Inches(page['margin_top_in'])
            section.bottom_margin = Inches(page['margin_bottom_in'])
            section.left_margin = Inches(page['margin_left_in'])
            section.right_margin = Inches(page['margin_right_in'])

    def _apply_default_font_and_spacing(self) -> None:
        """Garamond + interlineado 1.5 + sin espacios entre parrafos + justificado."""
        ts = self.style['typography']
        normal = self.doc.styles['Normal']
        normal.font.name = ts['font_family']
        normal.font.size = Pt(ts['size_body_pt'])
        pf = normal.paragraph_format
        pf.line_spacing = ts.get('line_spacing', 1.0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # Forzar fuente para todas las variantes
        rPr = normal.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
            rFonts.set(qn(attr), ts['font_family'])

        # Ajustar List Bullet / List Number: indentacion minima, sin espacios extra
        for list_style_name in ('List Bullet', 'List Number'):
            try:
                ls = self.doc.styles[list_style_name]
                lpf = ls.paragraph_format
                lpf.left_indent = Inches(0.25)
                lpf.first_line_indent = Inches(0)
                lpf.space_before = Pt(0)
                lpf.space_after = Pt(0)
                lpf.line_spacing = ts.get('line_spacing', 1.0)
            except KeyError:
                pass

    def _apply_heading_styles(self) -> None:
        """Sobreescribe Heading 1/2/3 (estilos nativos) con paleta y fuente MPF.
        Mantener los estilos nativos es importante para que el campo TOC los detecte."""
        ts = self.style['typography']
        hd = self.style['headings']
        for level, color_key, size_key, sp_before, sp_after in [
            (1, hd['h1_color'], hd['h1_size_pt'], 12, 4),
            (2, hd['h2_color'], hd['h2_size_pt'], 8,  2),
            (3, hd['h3_color'], hd['h3_size_pt'], 6,  2),
        ]:
            style = self.doc.styles[f'Heading {level}']
            style.font.name = ts['font_family']
            style.font.size = Pt(size_key)
            style.font.bold = True
            style.font.color.rgb = self._color(color_key)
            pf = style.paragraph_format
            pf.line_spacing = ts.get('line_spacing', 1.0)
            pf.space_before = Pt(sp_before)
            pf.space_after = Pt(sp_after)
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT  # headings nunca justificados
            # Forzar fuente
            rPr = style.element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
                rFonts.set(qn(attr), ts['font_family'])

    def _setup_footer(self) -> None:
        """Pie de pagina: tabla 1x3 con [vacio | numero pagina | logo Analytia]."""
        cfg = self.style.get('footer', {})
        if not cfg:
            return
        section = self.doc.sections[0]
        footer = section.footer

        # Vaciar el parrafo default del footer
        for p in list(footer.paragraphs):
            p._element.getparent().remove(p._element)

        page_w = self.style['page']['width_in']
        m_l = self.style['page']['margin_left_in']
        m_r = self.style['page']['margin_right_in']
        usable = page_w - m_l - m_r

        table = footer.add_table(rows=1, cols=3, width=Inches(usable))
        table.autofit = False
        # Distribuir: 1/3 vacio, 1/3 page#, 1/3 logo
        col_w = Inches(usable / 3)
        for col in table.columns:
            col.width = col_w
        for row in table.rows:
            for cell in row.cells:
                cell.width = col_w

        # Celda izquierda: vacia
        left_cell = table.cell(0, 0)
        left_cell.paragraphs[0].text = ''

        # Celda central: numero de pagina (campo PAGE)
        if cfg.get('show_page_number', True):
            mid_p = table.cell(0, 1).paragraphs[0]
            mid_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = mid_p.add_run()
            _set_run_font(run, self.style['typography']['font_family'])
            run.font.size = Pt(self.style['typography']['size_body_pt'])
            # Inyectar field PAGE
            for tag, txt in [('begin', None), ('separate', '1'), ('end', None)]:
                if tag == 'begin':
                    fc = OxmlElement('w:fldChar')
                    fc.set(qn('w:fldCharType'), 'begin')
                    run._r.append(fc)
                    instr = OxmlElement('w:instrText')
                    instr.set(qn('xml:space'), 'preserve')
                    instr.text = 'PAGE   \\* MERGEFORMAT'
                    run._r.append(instr)
                elif tag == 'separate':
                    fc = OxmlElement('w:fldChar')
                    fc.set(qn('w:fldCharType'), 'separate')
                    run._r.append(fc)
                    t = OxmlElement('w:t')
                    t.text = txt
                    run._r.append(t)
                else:
                    fc = OxmlElement('w:fldChar')
                    fc.set(qn('w:fldCharType'), 'end')
                    run._r.append(fc)

        # Celda derecha: logo Analytia
        right_p = table.cell(0, 2).paragraphs[0]
        right_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if LOGO_DEV_PATH.exists():
            r = right_p.add_run()
            r.add_picture(str(LOGO_DEV_PATH),
                          width=Inches(cfg.get('developer_logo_width_in', 0.7)))

        # Quitar bordes de la tabla del footer
        tbl = table._element
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        tblBorders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b = OxmlElement(f'w:{edge}')
            b.set(qn('w:val'), 'nil')
            tblBorders.append(b)
        tblPr.append(tblBorders)

    def _color(self, key: str) -> RGBColor:
        if key.startswith('#'):
            return _hex_to_rgb(key)
        return _hex_to_rgb(self.style['colors'][key])

    # ---------- portada ----------
    def add_cover(self, title: str, subtitle: str = '', date: str = '',
                  logo_path: Path | str | None = None) -> None:
        cov = self.style['cover']
        ts = self.style['typography']
        logo = Path(logo_path) if logo_path else LOGO_PATH

        # Logo centrado
        if logo.exists():
            p_logo = self.doc.add_paragraph()
            p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p_logo.add_run()
            run.add_picture(str(logo), width=Inches(cov['logo_max_width_in']))

        for _ in range(4):
            self.doc.add_paragraph()

        # Titulo
        p_title = self.doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_title.add_run(title)
        _set_run_font(run, ts['font_family'])
        run.font.size = Pt(cov['title_size_pt'])
        run.font.bold = True
        run.font.color.rgb = self._color(cov['title_color'])

        # Subtitulo
        if subtitle:
            p_sub = self.doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p_sub.add_run(subtitle)
            _set_run_font(r, ts['font_family'])
            r.font.size = Pt(cov['subtitle_size_pt'])
            r.font.color.rgb = self._color(cov['subtitle_color'])

        # Fecha
        if date:
            for _ in range(2):
                self.doc.add_paragraph()
            p_date = self.doc.add_paragraph()
            p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p_date.add_run(date)
            _set_run_font(r, ts['font_family'])
            r.font.size = Pt(ts['size_body_pt'])
            r.font.color.rgb = self._color('black_text')

        self.doc.add_page_break()

    # ---------- TOC ----------
    def add_toc(self, title: str = 'Tabla de contenido') -> None:
        """Inserta titulo + campo TOC. Word lo poblara al abrir el doc o con F9."""
        # Titulo de la TOC (sin que aparezca en la propia TOC)
        ts = self.style['typography']
        p_t = self.doc.add_paragraph()
        run = p_t.add_run(title)
        _set_run_font(run, ts['font_family'])
        run.font.size = Pt(self.style['headings']['h1_size_pt'])
        run.font.bold = True
        run.font.color.rgb = self._color(self.style['headings']['h1_color'])

        # Campo TOC
        p_toc = self.doc.add_paragraph()
        run = p_toc.add_run()

        fld_begin = OxmlElement('w:fldChar')
        fld_begin.set(qn('w:fldCharType'), 'begin')
        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = r'TOC \o "1-3" \h \z \u'
        fld_sep = OxmlElement('w:fldChar')
        fld_sep.set(qn('w:fldCharType'), 'separate')
        placeholder = OxmlElement('w:t')
        placeholder.text = ('Tabla de contenido vacia. En Word: clic derecho '
                            'sobre este texto > Actualizar campos > Actualizar toda la tabla.')
        fld_end = OxmlElement('w:fldChar')
        fld_end.set(qn('w:fldCharType'), 'end')

        for el in (fld_begin, instr, fld_sep, placeholder, fld_end):
            run._r.append(el)

        self.doc.add_page_break()

    # ---------- contenido ----------
    def h1(self, text: str) -> None:
        self.doc.add_paragraph(text, style='Heading 1')

    def h2(self, text: str) -> None:
        self.doc.add_paragraph(text, style='Heading 2')

    def h3(self, text: str) -> None:
        self.doc.add_paragraph(text, style='Heading 3')

    def p(self, text: str, bold: bool = False) -> None:
        ts = self.style['typography']
        p = self.doc.add_paragraph()
        run = p.add_run(text)
        _set_run_font(run, ts['font_family'])
        run.font.size = Pt(ts['size_body_pt'])
        run.font.bold = bold

    def bullet(self, text: str) -> None:
        ts = self.style['typography']
        p = self.doc.add_paragraph(style='List Bullet')
        run = p.add_run(text)
        _set_run_font(run, ts['font_family'])
        run.font.size = Pt(ts['size_body_pt'])

    def numbered(self, text: str) -> None:
        ts = self.style['typography']
        p = self.doc.add_paragraph(style='List Number')
        run = p.add_run(text)
        _set_run_font(run, ts['font_family'])
        run.font.size = Pt(ts['size_body_pt'])

    def code(self, text: str) -> None:
        """Bloque monoespaciado con fondo cream. Util para comandos de instalacion."""
        ts = self.style['typography']
        p = self.doc.add_paragraph()
        # Sombreado claro
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), self.style['colors']['cream_bg'].lstrip('#'))
        pPr.append(shd)
        # En bloques de codigo: interlineado 1.0 y alineacion izquierda (no justificado)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        _set_run_font(run, 'Consolas')
        run.font.size = Pt(ts['size_body_pt'])

    def table(self, headers: list[str], rows: list[list[str]],
              col_widths_in: list[float] | None = None) -> None:
        """Tabla simple con header marron + cuerpo Garamond.
        headers: lista de strings | rows: lista de listas de strings
        col_widths_in: anchos en pulgadas (opcional)."""
        ts = self.style['typography']
        n_cols = len(headers)
        table = self.doc.add_table(rows=1 + len(rows), cols=n_cols)
        table.autofit = False

        # Anchos de columna
        if col_widths_in:
            for col, w in zip(table.columns, col_widths_in):
                col.width = Inches(w)
            for row in table.rows:
                for i, cell in enumerate(row.cells):
                    if i < len(col_widths_in):
                        cell.width = Inches(col_widths_in[i])

        # Header
        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # tablas siempre LEFT, no justify
            run = p.add_run(h)
            _set_run_font(run, ts['font_family'])
            run.font.size = Pt(ts['size_body_pt'])
            run.font.bold = True
            run.font.color.rgb = self._color('white')
            # Sombreado marron
            tcPr = cell._tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), self.style['colors']['brown_primary'].lstrip('#'))
            tcPr.append(shd)

        # Filas de datos
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                cell = table.rows[1 + ri].cells[ci]
                cell.text = ''
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(str(val))
                _set_run_font(run, ts['font_family'])
                run.font.size = Pt(ts['size_body_pt'])

        # Bordes finos en todas las celdas
        tbl = table._element
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        tblBorders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b = OxmlElement(f'w:{edge}')
            b.set(qn('w:val'), 'single')
            b.set(qn('w:sz'), '4')
            b.set(qn('w:color'), '999999')
            tblBorders.append(b)
        tblPr.append(tblBorders)

    def page_break(self) -> None:
        self.doc.add_page_break()

    def save(self, out_path: str | Path) -> Path:
        out = Path(out_path)
        self.doc.save(str(out))
        return out


# ---------------- demo CLI ----------------
def _demo(out: str = 'demo.docx'):
    db = DocumentBuilder()
    db.add_cover(title='Documento de Prueba',
                 subtitle='Estilo corporativo MPF',
                 date='Mayo 2026')
    db.add_toc()
    db.h1('1. Encabezado nivel 1')
    db.p('Este es un parrafo normal con el cuerpo en Garamond 11pt e interlineado 1.5.')
    db.h2('1.1 Encabezado nivel 2')
    db.bullet('Item con bullet')
    db.bullet('Otro item con bullet')
    db.numbered('Paso numerado 1')
    db.numbered('Paso numerado 2')
    db.h3('1.1.1 Encabezado nivel 3')
    db.p('Comando de ejemplo:')
    db.code('python --version')
    db.code('pip install python-docx openpyxl')
    db.p('Texto en negrita:', bold=True)
    db.p('Fin del demo. El logo Analytia debe verse en el pie a la derecha.')
    out_path = db.save(out)
    print(f'Demo generado: {out_path.resolve()}')


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'demo':
        _demo(sys.argv[2] if len(sys.argv) > 2 else 'demo.docx')
    else:
        print(__doc__)
