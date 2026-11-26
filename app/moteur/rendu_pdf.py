"""Imprime en PDF les pages du plan financier et des bilans, à l'identique de l'Excel.

Le classeur généré est relu tel quel (styles, fusions, largeurs, sauts de page) ; les formules sont
calculées en Python (pycel), puis chaque page est dessinée avec reportlab : mêmes tableaux, mêmes
couleurs, mêmes libellés que l'onglet « Plan financier à imprimer » et que les bilans.
Aucune dépendance à Excel ou LibreOffice : fonctionne sur n'importe quel hébergeur Python.
"""
import datetime
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string as CI, get_column_letter as GL, range_boundaries
from pycel import ExcelCompiler
from pycel import excellib as _excellib

_round_origine = _excellib.round_


def _round(number, num_digits=0):
    # pycel passe parfois des scalaires numpy, que son ROUND ne sait pas convertir
    try:
        number = float(number)
    except (TypeError, ValueError):
        pass
    return _round_origine(number, num_digits)


_excellib.round_ = _round
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

POLICES_DIR = Path(__file__).resolve().parent.parent / 'data' / 'polices'
FAMILLES = {  # police Excel -> fichiers (regular, bold, italic, bolditalic)
    'Carlito': ('Carlito-Regular', 'Carlito-Bold', 'Carlito-Italic', 'Carlito-BoldItalic'),
    'Caladea': ('Caladea-Regular', 'Caladea-Bold', 'Caladea-Italic', 'Caladea-BoldItalic'),
    'LibSans': ('LiberationSans-Regular', 'LiberationSans-Bold', 'LiberationSans-Italic', 'LiberationSans-BoldItalic'),
    'LibSerif': ('LiberationSerif-Regular', 'LiberationSerif-Bold', 'LiberationSerif-Italic', 'LiberationSerif-BoldItalic'),
}
CORRESPONDANCE = {'calibri': 'Carlito', 'cambria': 'Caladea', 'arial': 'LibSans', 'verdana': 'LibSans', 'trebuchet ms': 'LibSans',
                  'century gothic': 'LibSans', 'helvetica': 'LibSans', 'georgia': 'LibSerif', 'times new roman': 'LibSerif'}
_enregistrees = False
THEME_COULEURS = {0: 'FFFFFF', 1: '000000', 2: 'E7E6E6', 3: '44546A', 4: '4472C4', 5: 'ED7D31'}
TRAITS = {'hair': (0.25, None), 'thin': (0.5, None), 'medium': (1.0, None), 'thick': (1.6, None), 'double': (1.2, None),
          'dashed': (0.5, (2, 1.5)), 'dotted': (0.5, (0.6, 1.2)), 'mediumDashed': (1.0, (3, 2)), 'dashDot': (0.5, (3, 1, 1, 1)),
          'mediumDashDot': (1.0, (3, 1, 1, 1)), 'dashDotDot': (0.5, (3, 1, 1, 1, 1, 1)), 'slantDashDot': (1.0, (3, 1, 1, 1))}


def _polices():
    global _enregistrees
    if _enregistrees:
        return
    for fam, fichiers in FAMILLES.items():
        for suffixe, f in zip(('', '-B', '-I', '-BI'), fichiers):
            pdfmetrics.registerFont(TTFont(fam + suffixe, str(POLICES_DIR / f'{f}.ttf')))
    _enregistrees = True


def _police(font):
    fam = CORRESPONDANCE.get((font.name or 'Calibri').lower(), 'Carlito')
    suffixe = '-BI' if font.b and font.i else '-B' if font.b else '-I' if font.i else ''
    return fam + suffixe


def _couleur(c, defaut=None):
    if c is None:
        return defaut
    try:
        if c.type == 'rgb' and isinstance(c.rgb, str) and len(c.rgb) >= 6:
            if c.rgb[:2] == '00' and len(c.rgb) == 8 and c.rgb != '00000000':
                return c.rgb[2:]
            return c.rgb[-6:] if c.rgb != '00000000' else defaut
        if c.type == 'theme':
            base = THEME_COULEURS.get(c.theme, '000000')
            t = c.tint or 0
            if t:
                r, g, b = (int(base[i:i + 2], 16) for i in (0, 2, 4))
                f = (lambda x: x + (255 - x) * t) if t > 0 else (lambda x: x * (1 + t))
                return ''.join(f'{int(round(f(x))):02X}' for x in (r, g, b))
            return base
        if c.type == 'indexed':
            return {8: '000000', 9: 'FFFFFF', 64: '000000', 65: 'FFFFFF'}.get(c.indexed, defaut)
    except Exception:
        return defaut
    return defaut


def _rgb(h):
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _separer_sections(nf):
    sections, cur, q = [], '', False
    for ch in nf:
        if ch == '"':
            q = not q
        if ch == ';' and not q:
            sections.append(cur); cur = ''
        else:
            cur += ch
    sections.append(cur)
    return sections


def _milliers(n):
    s = f'{n:,}'.replace(',', ' ')
    return s


def formater(v, nf):
    if v is None or v == '':
        return ''
    if isinstance(v, bool):
        return 'VRAI' if v else 'FAUX'
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%d/%m/%Y')
    if isinstance(v, str):
        return v
    try:
        x = float(v)
    except (TypeError, ValueError):
        return str(v)
    nf = nf or 'General'
    if re.search(r'(?<![A-Za-z])(d{1,2}|yy|yyyy)(?![A-Za-z])', nf.split(';')[0].replace('"', '')) and 'mm' in nf or 'yyyy' in nf:
        try:
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(x))).strftime('%d/%m/%Y')
        except (OverflowError, ValueError):
            return ''
    if nf == 'General':
        if abs(x - round(x)) < 1e-9:
            return _milliers(int(round(x))) if abs(x) >= 10000 else str(int(round(x)))
        return f'{x:.2f}'.replace('.', ',')
    secs = _separer_sections(nf)
    sec = secs[0]
    if x < 0 and len(secs) >= 2:
        sec = secs[1]
    if abs(x) < 0.005 and len(secs) >= 3:
        z = secs[2]
        if '-' in z.replace('_-', '').replace('\\-', '-'):
            devise = re.search(r'\[\$([^\]-]+)', nf)
            return '-' + ((' ' + devise.group(1)) if devise and '[$' in z else '')
        if '0' not in z:
            return ''
        sec = z
    propre = re.sub(r'"[^"]*"|\[\$[^\]]*\]|_.|\\.|\*.', '', sec)
    pct = '%' in propre
    dec = 0
    m = re.search(r'\.(0+)', propre)
    if m:
        dec = len(m.group(1))
    val = abs(x) * (100 if pct else 1)
    if ',' in propre:
        txt = f'{val:,.{dec}f}'.replace(',', ' ').replace('.', ',')
    else:
        txt = f'{val:.{dec}f}'.replace('.', ',')
    if pct:
        txt += ' %'
    devise = re.search(r'\[\$([^\]-]+)', sec)
    if devise:
        txt += ' ' + devise.group(1)
    for lit in re.findall(r'"([^"]+)"', sec):
        if lit.strip() and lit.strip() not in ('-',):
            txt += lit if lit.startswith(' ') else ' ' + lit
    if x < 0:
        txt = ('(' + txt + ')') if '(' in sec and len(secs) >= 2 else '-' + txt
    return txt


class Rendu:
    def __init__(self, chemin_xlsx, theme=None):
        _polices()
        self.wb = load_workbook(chemin_xlsx)
        self.xl = ExcelCompiler(filename=str(chemin_xlsx))
        self.theme = theme or {}

    def valeur(self, ws, c):
        v = c.value
        if isinstance(v, str) and v.startswith('='):
            try:
                v = self.xl.evaluate(f"'{ws.title}'!{c.coordinate}")
            except Exception:
                v = ''
            if hasattr(v, 'tolist'):
                v = v.tolist()
            if isinstance(v, str) and v.startswith('#'):
                v = ''
        return v

    def page(self, cv, ws, c0, c1, r0, r1, paysage=False):
        colonnes = [c for c in range(c0, c1 + 1) if not ws.column_dimensions[GL(c)].hidden]
        lignes = [r for r in range(r0, r1 + 1) if not ws.row_dimensions[r].hidden]
        larg = {c: ((ws.column_dimensions[GL(c)].width or 8.43) * 7 + 5) * 0.75 for c in colonnes}
        haut = {r: (ws.row_dimensions[r].height or 15) for r in lignes}
        # retire les colonnes / lignes vides en fin de page
        W, Hh = (A4[1], A4[0]) if paysage else A4
        cv.setPageSize((W, Hh))
        marge = 22
        tw, th = sum(larg.values()), sum(haut.values())
        s = min((W - 2 * marge) / tw, (Hh - 2 * marge - 12) / th, 1.0)
        x0 = (W - tw * s) / 2; y0 = Hh - marge
        xs, x = {}, x0
        for c in colonnes:
            xs[c] = x; x += larg[c] * s
        xs_fin = x
        ys, y = {}, y0
        for r in lignes:
            ys[r] = y; y -= haut[r] * s
        ys_fin = y
        fusion = {}
        couvert = set()
        for m in ws.merged_cells.ranges:
            if m.max_col < c0 or m.min_col > c1 or m.max_row < r0 or m.min_row > r1:
                continue
            fusion[(m.min_row, m.min_col)] = (m.max_row, m.max_col)
            for rr in range(m.min_row, m.max_row + 1):
                for cc in range(m.min_col, m.max_col + 1):
                    if (rr, cc) != (m.min_row, m.min_col):
                        couvert.add((rr, cc))

        def boite(r, c):
            r2, c2 = fusion.get((r, c), (r, c))
            cs = [k for k in colonnes if c <= k <= c2]; rs = [k for k in lignes if r <= k <= r2]
            if not cs or not rs:
                return None
            xa = xs[cs[0]]; xb = xs[cs[-1]] + larg[cs[-1]] * s
            ya = ys[rs[0]]; yb = ys[rs[-1]] - haut[rs[-1]] * s
            return xa, yb, xb, ya

        # 1. fonds
        for r in lignes:
            for c in colonnes:
                if (r, c) in couvert:
                    continue
                cell = ws.cell(r, c)
                f = cell.fill
                if f is not None and f.fill_type == 'solid':
                    col = _couleur(f.fgColor)
                    if col and col != 'FFFFFF':
                        b = boite(r, c)
                        if b:
                            cv.setFillColorRGB(*_rgb(col)); cv.rect(b[0], b[1], b[2] - b[0], b[3] - b[1], stroke=0, fill=1)
        # 2. bordures (dans une cellule fusionnée, seul le contour est tracé)
        zone_fusion = {}
        for (fr, fc), (tr, tc) in fusion.items():
            for rr in range(fr, tr + 1):
                for cc in range(fc, tc + 1):
                    zone_fusion[(rr, cc)] = (fr, fc, tr, tc)
        for r in lignes:
            for c in colonnes:
                bd = ws.cell(r, c).border
                z = zone_fusion.get((r, c))
                if bd is None:
                    continue
                xa, xb = xs[c], xs[c] + larg[c] * s
                ya, yb = ys[r], ys[r] - haut[r] * s
                for cote, seg in (('top', (xa, ya, xb, ya)), ('bottom', (xa, yb, xb, yb)), ('left', (xa, ya, xa, yb)), ('right', (xb, ya, xb, yb))):
                    if z and ((cote == 'top' and r > z[0]) or (cote == 'bottom' and r < z[2]) or (cote == 'left' and c > z[1]) or (cote == 'right' and c < z[3])):
                        continue
                    side = getattr(bd, cote)
                    if side is None or not side.style or side.style not in TRAITS:
                        continue
                    ep, tirets = TRAITS[side.style]
                    cv.setStrokeColorRGB(*_rgb(_couleur(side.color, '000000') or '000000'))
                    cv.setLineWidth(ep * max(s, 0.5)); cv.setDash(tirets or [])
                    cv.line(*seg)
        cv.setDash([])
        # 3. textes
        for r in lignes:
            for c in colonnes:
                if (r, c) in couvert:
                    continue
                cell = ws.cell(r, c)
                v = self.valeur(ws, cell)
                txt = formater(v, cell.number_format)
                if not txt:
                    continue
                b = boite(r, c)
                if not b:
                    continue
                fo = cell.font
                nom = _police(fo); taille = (fo.sz or 11) * s
                couleur = _couleur(fo.color, '000000') or '000000'
                cv.setFillColorRGB(*_rgb(couleur)); cv.setFont(nom, taille)
                al = cell.alignment
                nombre = isinstance(v, (int, float)) and not isinstance(v, bool)
                h = al.horizontal or ('right' if nombre else 'left')
                if nombre and '*' in (cell.number_format or ''):
                    h = 'right'  # format comptable : le remplissage aligne le nombre à droite
                pad = 2 * s + (al.indent or 0) * 6 * s
                largeur = b[2] - b[0] - 2 * pad
                lignes_txt = []
                for morceau in str(txt).split('\n'):
                    if al.wrap_text or fusion.get((r, c), (r, c))[0] > r:
                        mots, cur = morceau.split(' '), ''
                        for mot in mots:
                            essai = (cur + ' ' + mot).strip()
                            if cur and pdfmetrics.stringWidth(essai, nom, taille) > largeur:
                                lignes_txt.append(cur); cur = mot
                            else:
                                cur = essai
                        lignes_txt.append(cur)
                    else:
                        lignes_txt.append(morceau)
                lh = taille * 1.18
                bloc = lh * len(lignes_txt)
                va = al.vertical or 'bottom'
                if va == 'center':
                    yt = (b[1] + b[3]) / 2 + bloc / 2 - taille * 0.95
                elif va == 'top':
                    yt = b[3] - taille * 1.0
                else:
                    yt = b[1] + (len(lignes_txt) - 1) * lh + 2.2 * s
                for i, ligne in enumerate(lignes_txt):
                    yy = yt - i * lh
                    if h in ('center', 'centerContinuous', 'distributed'):
                        cv.drawCentredString((b[0] + b[2]) / 2, yy, ligne)
                    elif h == 'right':
                        cv.drawRightString(b[2] - pad, yy, ligne)
                    else:
                        cv.drawString(b[0] + pad, yy, ligne)
        # 4. images (logo)
        for img in getattr(ws, '_images', []):
            try:
                a = img.anchor._from
                c, r = a.col + 1, a.row + 1
                if c in xs and r in ys:
                    from reportlab.lib.utils import ImageReader
                    data = img._data()
                    import io
                    ir = ImageReader(io.BytesIO(data))
                    wpx, hpx = img.width, img.height
                    ext = getattr(img.anchor, 'ext', None)
                    if ext is not None and getattr(ext, 'width', 0):  # taille affichée dans le classeur (EMU), pas la taille d'origine du fichier
                        wpx, hpx = ext.width / 9525, ext.height / 9525
                    w_, h_ = wpx * 0.75 * s, hpx * 0.75 * s
                    cv.drawImage(ir, xs[c], ys[r] - h_, w_, h_, mask='auto')
            except Exception:
                pass
        marque = self.theme.get('marque')
        if marque:
            cv.setFillColorRGB(0.5, 0.5, 0.5); cv.setFont('Carlito-I', 7); cv.drawString(marge, 12, marque)
        cv.showPage()


def imprimer(chemin_xlsx, chemin_pdf, theme=None):
    """Écrit le PDF : pages du plan financier (une par tableau, comme le modèle), puis les bilans."""
    R = Rendu(chemin_xlsx, theme)
    wb = R.wb
    cv = canvas.Canvas(str(chemin_pdf), pagesize=A4)
    plan = next(ws for ws in wb.worksheets if ws.title.startswith('01_Plan'))
    cv.setTitle(f"Plan financier — {plan['M6'].value or ''}")
    zone = plan.print_area
    c_min, r_min, c_max, r_max = range_boundaries(zone[0] if isinstance(zone, (list, tuple)) else str(zone).split('!')[-1].replace('$', ''))
    bornes = [c_min - 1] + sorted(b.id for b in plan.col_breaks.brk) + [c_max]
    for d, f in zip(bornes, bornes[1:]):
        R.page(cv, plan, d + 1, f, r_min, min(r_max, 60))
    for ws in wb.worksheets:
        if 'Bilan' in ws.title and ws.title[:1] == '6' and ws.sheet_state == 'visible':
            R.page(cv, ws, 1, CI('K'), 1, 53, paysage=True)
    cv.save()
