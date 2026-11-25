"""Remplit les tableaux du modèle de plan financier du cabinet (gabarit_cabinet.xlsx).

Les tableaux reprennent la présentation du modèle (mêmes lignes, mêmes libellés, même mise en page) ;
seules les couleurs (couleur principale et secondaire du client) et les données changent. Toutes les
cellules sont des formules liées au dossier de travail (onglets 10 à 16), qui reste la seule zone de saisie.
Les onglets sans données (tableaux d'amortissement sans immobilisation) ne sont pas générés.
"""
import datetime
import re
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import get_column_letter as GL, column_index_from_string as CI

from .excel import _argb, clair
from .formulaire import CATEGORIES, COMPTES, FRAIS_DEMARRAGE

GABARIT = Path(__file__).resolve().parent.parent / 'data' / 'gabarit_cabinet.xlsx'
MARRON, BEIGE_MODELE = 'FF8B4513', 'FFDDD4B9'
MOIS = ['janv', 'févr', 'mars', 'avr', 'mai', 'juin', 'juil', 'août', 'sept', 'oct', 'nov', 'déc']
CAT = {k: v[0] for k, v in CATEGORIES.items()}
INCORP = [CAT['brevets'], CAT['logiciels']]
FIN = [CAT['cautions']]


def charger_gabarit():
    return load_workbook(GABARIT)


def _q(n):
    return "'" + n + "'"


def _ecrire(ws, ref, v):
    c = ws[ref]
    if type(c).__name__ == 'MergedCell':
        for rng in ws.merged_cells.ranges:
            if c.coordinate in rng:
                ws.cell(rng.min_row, rng.min_col).value = v
                return
        return
    c.value = v


def _style_comme(ws, ref_src, ref_dst):
    s, t = ws[ref_src], ws[ref_dst]
    t.font = copy(s.font); t.fill = copy(s.fill); t.border = copy(s.border)
    t.alignment = copy(s.alignment); t.number_format = s.number_format


def _style_comme_entre(ws, ref_src, ref_dst):
    s, t = ws[ref_src], ws[ref_dst]
    if type(t).__name__ == 'MergedCell':
        return
    t.font = copy(s.font); t.fill = copy(s.fill); t.border = copy(s.border)
    t.alignment = copy(s.alignment); t.number_format = s.number_format


_REF = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')


def _decaler(formule, off):
    """Décale de `off` colonnes les références internes au bloc (sans nom d'onglet)."""
    if not (isinstance(formule, str) and formule.startswith('=')):
        return formule
    tok = Tokenizer(formule)
    out = []
    for t in tok.items:
        v = t.value
        if t.type == 'OPERAND' and t.subtype == 'RANGE' and '!' not in v:
            v = _REF.sub(lambda m: m.group(1) + GL(CI(m.group(2)) + off) + m.group(3) + m.group(4), v)
        out.append(v)
    return '=' + ''.join(out)


def _transplanter(src, dst, off):
    """Recopie un bloc rempli (onglet de travail du gabarit) dans le plan financier, décalé de `off` colonnes."""
    for row in src.iter_rows(min_col=2):
        for c in row:
            if type(c).__name__ == 'MergedCell':
                continue
            t = dst.cell(c.row, c.column + off)
            if type(t).__name__ == 'MergedCell':
                continue
            t.value = _decaler(c.value, off)
            if c.has_style:
                t.font = copy(c.font); t.fill = copy(c.fill); t.border = copy(c.border)
                t.alignment = copy(c.alignment); t.number_format = c.number_format
    for k, d in src.column_dimensions.items():
        if CI(k) >= 2 and d.width:
            dst.column_dimensions[GL(CI(k) + off)].width = d.width


def recolorer(ws, primaire, secondaire):
    p, s = _argb(primaire), clair(secondaire)
    for row in ws.iter_rows():
        for c in row:
            f = c.fill
            if f is not None and f.fill_type == 'solid' and f.fgColor.type == 'rgb':
                if f.fgColor.rgb == MARRON:
                    c.fill = PatternFill('solid', fgColor=p)
                elif f.fgColor.rgb == BEIGE_MODELE:
                    c.fill = PatternFill('solid', fgColor=s)


def remplir(cfg, ctx):
    wb = ctx['wb']; theme = ctx['theme'] or {}
    H, M, C, I, R, B = map(_q, (ctx['SH'], ctx['SM'], ctx['SC'], ctx['SI'], ctx['SR'], ctx['SB']))
    years, cal, N = ctx['years'], ctx['cal'], ctx['N']
    ny = len(years)
    nbm = {y: sum(1 for a, _ in cal if a == y) for y in years}
    lab = {y: (f'{y}' if nbm[y] == 12 else f'{y} ({nbm[y]} mois)') for y in years}
    dern = {y: max(i for i, (a, _) in enumerate(cal) if a == y) for y in years}
    prem = {y: min(i for i, (a, _) in enumerate(cal) if a == y) for y in years}
    irow = ctx['irow']; drows = ctx['drows']
    i0, i1 = (irow[0], irow[-1]) if irow else (None, None)
    pp = sorted(ctx['prow_p'].values())
    p0, p1 = (pp[0], pp[-1]) if pp else (None, None)
    soc = cfg['societe']; porteur = cfg.get('porteur', '')
    dev = cfg.get('pays', {}).get('devise', 'XAF')
    a0, m0 = cal[0]
    NUM = '#,##0;-#,##0;"-"'
    invs = cfg.get('investissements', [])

    def inv_cat(cats, cond):
        """Somme des investissements des catégories `cats` dont le mois d'acquisition vérifie `cond` (critère SUMIFS)."""
        if not irow:
            return '0'
        return '+'.join(f'SUMIFS({H}!$B${i0}:$B${i1},{H}!$F${i0}:$F${i1},"{c}",{H}!$D${i0}:$D${i1},"{cond}")' for c in cats)

    def sumif_cr(code, k):
        return f'SUMIF({R}!$A:$A,"{code} —*",{R}!{GL(2 + k)}:{GL(2 + k)})'

    def sumif_m(code, col):
        return f'SUMIF({M}!$A:$A,"{code} —*",{M}!{col}:{col})'

    # ---------------------------------------------------------------- 01 Investissements et financements initiaux
    inv = wb['INV']
    y1 = years[0]; c1 = '<=2'  # financements initiaux : les 3 premiers mois (le reste figure au plan de financement)
    _ecrire(inv, 'D6', soc); _ecrire(inv, 'D7', porteur)
    _ecrire(inv, 'H9', f'Montant {dev} hors taxes'); _ecrire(inv, 'H41', f'Montant {dev} hors taxes')
    lignes = {13: [CAT['brevets']], 14: [CAT['logiciels']], 16: [CAT['terrains']], 17: [CAT['constructions']],
              18: [CAT['amenagements']], 19: [CAT['materiel_industriel']], 20: [CAT['agencements']], 21: [CAT['transport']],
              22: [CAT['bureau']], 23: [CAT['mobilier']]}
    for r, cats in lignes.items():
        _ecrire(inv, f'B{r}', '   ' + cats[0]); _ecrire(inv, f'H{r}', '=' + inv_cat(cats, c1))
    _ecrire(inv, 'H12', '=SUM(H13:H14)'); _ecrire(inv, 'H15', '=SUM(H16:H23)'); _ecrire(inv, 'H24', '=SUM(H25:H34)')
    frais = [FRAIS_DEMARRAGE[k][1] for k in ('etudes', 'consultants', 'permis', 'formation', 'communication', 'notaire', 'experts')]
    frais += ['Caution ou dépôt de garantie', FRAIS_DEMARRAGE['compteurs'][1], FRAIS_DEMARRAGE['creation'][1]]
    for k, lib_ in enumerate(frais):
        r = 25 + k; _ecrire(inv, f'B{r}', '   ' + lib_)
        if lib_.startswith('Caution'):
            f = inv_cat(FIN, c1)
        elif pp:
            f = f'SUMIFS({H}!$B${p0}:$B${p1},{H}!$D${p0}:$D${p1},"{lib_}",{H}!$C${p0}:$C${p1},"{c1}")'
        else:
            f = '0'
        _ecrire(inv, f'H{r}', '=' + f)
    _ecrire(inv, 'B35', 'Stock de matières au lancement'); _ecrire(inv, 'H35', 0)
    _ecrire(inv, 'B36', 'Besoin en fonds de roulement (BFR)'); _ecrire(inv, 'H36', 0)
    _ecrire(inv, 'B37', 'Trésorerie de démarrage'); _ecrire(inv, 'H37', '=H55-(H12+H15+H24+H35+H36)')
    _ecrire(inv, 'H38', None); _ecrire(inv, 'F39', 'TOTAL BESOINS'); _ecrire(inv, 'H39', '=H12+H15+H24+SUM(H35:H37)')
    _ecrire(inv, 'B44', 'Apport en fonds propres'); _ecrire(inv, 'H44', '=SUM(H45:H46)')
    _ecrire(inv, 'B45', '   Apport en numéraire'); _ecrire(inv, 'H45', f"={H}!$B${ctx['r_cap']}")
    _ecrire(inv, 'B46', '   Apport en nature'); _ecrire(inv, 'H46', 0)
    _ecrire(inv, 'B47', 'Emprunt'); _ecrire(inv, 'H47', '=SUM(H48:H50)')
    emps = cfg['financement'].get('emprunts', [])
    erow = ctx['erow']
    for k in range(3):
        r = 48 + k
        if k < len(emps):
            idx = list(range(k, len(emps))) if k == 2 else [k]
            nom = emps[k]['nom'] if len(idx) == 1 else 'Autres prêts'
            _ecrire(inv, f'B{r}', f'   Prêt n°{k + 1} ({nom})' if len(idx) == 1 else '   Autres prêts')
            _ecrire(inv, f'H{r}', '=' + '+'.join(f'IF({H}!$E${erow[j]}<=2,{H}!$B${erow[j]},0)' for j in idx))
        else:
            _ecrire(inv, f'B{r}', f'   Prêt n°{k + 1}'); _ecrire(inv, f'H{r}', 0)
    _ecrire(inv, 'B51', 'Subvention n°1'); _ecrire(inv, 'H51', 0)
    _ecrire(inv, 'B52', 'Subvention n°2'); _ecrire(inv, 'H52', 0)
    _ecrire(inv, 'B53', "Autre financement (comptes courants d'associés)"); _ecrire(inv, 'H53', f"={H}!$B${ctx['r_cca']}")
    _ecrire(inv, 'F55', 'TOTAL RESSOURCES'); _ecrire(inv, 'H55', '=SUM(H44,H47,H51:H53)')
    _ecrire(inv, 'B57', "Programme des trois premiers mois ; les investissements suivants figurent dans le plan de financement.")
    _ecrire(inv, 'B58', f"cf dossier de travail, onglet {ctx['SH']}")
    for r_ in (57, 58):
        inv[f'B{r_}'].font = Font(name='Calibri', size=9, italic=True)

    # ---------------------------------------------------------------- 02 Compte de résultat
    cr = wb['CR']
    _ecrire(cr, 'B2', f'Compte de résultats prévisionnel sur {ny} ans' if ny <= 5 else f'Compte de résultats prévisionnel ({ny} exercices)')
    _ecrire(cr, 'D6', soc); _ecrire(cr, 'D7', porteur)
    ycol = [GL(6 + k) for k in range(ny)]  # F, G, H, I, J (et K si 6 exercices)
    for k, col in enumerate(ycol):
        cr.column_dimensions[col].width = 18
        for r in range(8, 58):
            if k >= 3:
                _style_comme(cr, f'H{r}', f'{col}{r}')
        _ecrire(cr, f'{col}8', lab[years[k]])
    prods = cfg['produits']; car = ctx['car']; yc = ctx['yc']
    if len(prods) <= 2:
        plabs = [(f"Chiffre d'affaires HT vente {p['nom']}", [p['nom']]) for p in prods]
    else:
        plabs = [(f"Chiffre d'affaires HT vente {prods[0]['nom']}", [prods[0]['nom']]),
                 (f"Chiffre d'affaires HT autres ventes ({len(prods) - 1} produits et services)", [p['nom'] for p in prods[1:]])]
    for j in range(2):
        r = 11 + j
        _ecrire(cr, f'B{r}', ('   ' + plabs[j][0]) if j < len(plabs) else None)
        for k, col in enumerate(ycol):
            v = ('=' + '+'.join(f'{M}!{yc[years[k]]}{car[n]}' for n in plabs[j][1])) if j < len(plabs) else None
            _ecrire(cr, f'{col}{r}', v)
    codes = [c for c, _ in COMPTES if c not in ('65', '64')]
    for k, col in enumerate(ycol):
        f = lambda code: '=' + sumif_cr(code, k)
        _ecrire(cr, f'{col}10', f'=SUM({col}11:{col}12)')
        _ecrire(cr, f'{col}13', f'={col}14'); _ecrire(cr, f'{col}14', f('60'))
        _ecrire(cr, f'{col}16', f'={col}10-{col}13')
        _ecrire(cr, f'{col}17', f'=SUM({col}18:{col}41)')
        for j, code in enumerate(codes):
            _ecrire(cr, f'{col}{18 + j}', f(code))
        _ecrire(cr, f'{col}42', f'={col}43'); _ecrire(cr, f'{col}43', f('65'))
        _ecrire(cr, f'{col}45', f'={col}16-{col}17-{col}42')
        _ecrire(cr, f'{col}46', f('64')); _ecrire(cr, f'{col}47', f('6611')); _ecrire(cr, f'{col}48', f('6641'))
        _ecrire(cr, f'{col}49', f('6612')); _ecrire(cr, f'{col}50', f('6642'))
        _ecrire(cr, f'{col}51', f'={col}45-SUM({col}46:{col}50)')
        _ecrire(cr, f'{col}52', f"={R}!{GL(2 + k)}{ctx['cro']['int']}")
        _ecrire(cr, f'{col}53', f"={R}!{GL(2 + k)}{ctx['cro']['dot']}")
        _ecrire(cr, f'{col}54', f'={col}51-{col}52-{col}53')
        _ecrire(cr, f'{col}55', None)
        _ecrire(cr, f'{col}56', f"={R}!{GL(2 + k)}{ctx['cro']['imp']}")
        _ecrire(cr, f'{col}57', f'={col}54-SUM({col}55:{col}56)')
    for j, code in enumerate(codes):
        _ecrire(cr, f'B{18 + j}', f'{code} {dict(COMPTES)[code]}')
    _ecrire(cr, 'B43', "65 Autres charges (charges provisionnées d'exploitation, jetons de présence, dons…)")
    _ecrire(cr, 'B46', 'Impôts et taxes')
    _ecrire(cr, 'E50', None)
    _ecrire(cr, 'B56', "Impôt sur les sociétés (ou minimum de perception)")
    _ecrire(cr, 'B60', f"cf dossier de travail, onglet {ctx['SR']}")
    for r in (59, 62):
        for col in ('I', 'J'):
            _ecrire(cr, f'{col}{r}', None)

    # ---------------------------------------------------------------- 03 Plan de financement
    pf = wb['PF']
    _ecrire(pf, 'B2', f'Plan de financement à {ny} ans' if ny <= 5 else f'Plan de financement ({ny} exercices)')
    _ecrire(pf, 'D6', soc); _ecrire(pf, 'D7', porteur)
    pcol = [GL(6 + k) for k in range(ny)]
    nC = lambda k: GL(2 + N + k)
    for k, col in enumerate(pcol):
        if k >= 3:
            pf.column_dimensions[col].width = 17
            for r in range(12, 28):
                _style_comme(pf, f'H{r}', f'{col}{r}')
        y = years[k]
        _ecrire(pf, f'{col}12', lab[y])
        _ecrire(pf, f'{col}14', f"={I}!{GL(2 + k)}{ctx['r_cpx']}")
        _ecrire(pf, f'{col}15', 0); _ecrire(pf, f'{col}16', 0)
        _ecrire(pf, f'{col}17', f"={C}!{nC(k)}{ctx['r_tK']}")
        _ecrire(pf, f'{col}18', 0)
        _ecrire(pf, f'{col}19', f'=SUM({col}14:{col}18)')
        _ecrire(pf, f'{col}20', f"={H}!$B${ctx['r_cap']}" if k == 0 else 0)
        _ecrire(pf, f'{col}21', f"={C}!{nC(k)}{ctx['r_tD']}")
        _ecrire(pf, f'{col}22', 0)
        _ecrire(pf, f'{col}23', f"={H}!$B${ctx['r_cca']}" if k == 0 else 0)
        _ecrire(pf, f'{col}24', f"={R}!{GL(2 + k)}{ctx['cro']['caf']}")
        _ecrire(pf, f'{col}25', f'=SUM({col}20:{col}24)')
        _ecrire(pf, f'{col}26', f'={col}25-{col}19')
        _ecrire(pf, f'{col}27', f'={col}26' if k == 0 else f'={pcol[k - 1]}27+{col}26')
    _ecrire(pf, 'B16', 'Variation du besoin en fonds de roulement (non modélisée)')
    _ecrire(pf, 'B23', "Autres financements (comptes courants d'associés)")
    _ecrire(pf, 'B29', 'Rappel trésorerie début année 1 :'); _ecrire(pf, 'E29', 0)
    _ecrire(pf, 'B31', "L'excédent cumulé diffère de la trésorerie de fin d'exercice du décalage de paiement du solde d'impôt (payé l'exercice suivant).")
    for r in (50,):
        _ecrire(pf, f'I{r}', None)

    # ---------------------------------------------------------------- 04 Budget de trésorerie de la première année (12 mois)
    t = wb['TRE']
    mcols = ['E', 'F', 'G', 'H', 'I', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
    r_tres = ctx['r_tres']
    codes_ext = [c for c, _ in COMPTES if c not in ('64',)]
    _ecrire(t, 'B2', 'Budget prévisionnel de trésorerie'); _ecrire(t, 'L2', 'Budget prévisionnel de trésorerie (suite)')
    for c_ in ('D6', 'N6'):
        _ecrire(t, c_, soc)
    for c_ in ('D7', 'N7'):
        _ecrire(t, c_, porteur)
    _ecrire(t, 'B11', 'Première année')
    libs = {15: 'Apport en capital', 16: 'Emprunts', 17: 'Subventions', 18: "Autres financements (comptes courants d'associés)",
            19: 'Ventes de produits et services', 20: 'Autres encaissements', 21: "Chiffre d'affaires (total)",
            22: 'Immobilisations incorporelles', 23: 'Immobilisations corporelles et financières', 24: 'Immobilisations (total)',
            25: 'Acquisition stocks', 26: "Impôt sur les bénéfices (acomptes et solde)", 27: 'Achats consommés',
            28: 'Charges externes et autres charges', 29: 'Impôts et taxes', 30: 'Salaires employés', 31: 'Charges sociales employés',
            32: 'Salaire dirigeant', 33: 'Charges sociales dirigeant(s)', 34: 'Total charges de personnel',
            35: "Échéances d'emprunts (intérêts + capital)", 36: 'Total des décaissements', 37: 'Total des encaissements',
            38: 'Solde précédent', 39: 'Solde du mois', 40: 'Solde de trésorerie (cumul)'}
    for r, l_ in libs.items():
        _ecrire(t, f'B{r}', l_)
    for j, col in enumerate(mcols):
        idx = j
        mc = GL(2 + idx)
        a_, m_ = cal[idx]
        _ecrire(t, f'{col}13', f'{MOIS[m_ - 1]}-{str(a_)[2:]}')
        _ecrire(t, f'{col}15', f"={H}!$B${ctx['r_cap']}" if idx == 0 else 0)
        _ecrire(t, f'{col}16', f"={M}!{mc}{ctx['r_deb']}")
        _ecrire(t, f'{col}17', 0)
        _ecrire(t, f'{col}18', f"={H}!$B${ctx['r_cca']}" if idx == 0 else 0)
        _ecrire(t, f'{col}19', f"={M}!{mc}{ctx['r_ca']}")
        _ecrire(t, f'{col}20', 0)
        _ecrire(t, f'{col}21', f'=SUM({col}19:{col}20)')
        _ecrire(t, f'{col}22', '=' + inv_cat(INCORP, str(idx)))
        _ecrire(t, f'{col}23', f"={M}!{mc}{ctx['r_capex']}-{col}22")
        _ecrire(t, f'{col}24', f'=SUM({col}22:{col}23)')
        _ecrire(t, f'{col}25', 0)
        _ecrire(t, f'{col}26', f"={M}!{mc}{ctx['r_imp']}")
        _ecrire(t, f'{col}27', '=' + sumif_m('60', mc))
        _ecrire(t, f'{col}28', '=' + '+'.join(sumif_m(c_, mc) for c_ in codes_ext))
        _ecrire(t, f'{col}29', '=' + sumif_m('64', mc))
        for r, code in ((30, '6611'), (31, '6641'), (32, '6612'), (33, '6642')):
            _ecrire(t, f'{col}{r}', '=' + sumif_m(code, mc))
        _ecrire(t, f'{col}34', f'=SUM({col}30:{col}33)')
        _ecrire(t, f'{col}35', f"={M}!{mc}{ctx['r_serv']}")
        _ecrire(t, f'{col}36', f'=SUM({col}24:{col}29,{col}34:{col}35)')
        _ecrire(t, f'{col}37', f'=SUM({col}15:{col}20)')
        _ecrire(t, f'{col}38', 0 if j == 0 else f'={mcols[j - 1]}40')
        _ecrire(t, f'{col}39', f'={col}37-{col}36')
        _ecrire(t, f'{col}40', f'={col}38+{col}39')
        _ecrire(t, f'{col}41', f'=IF({col}40<0,{col}40,"")')
    for r in list(range(15, 38)) + [41]:
        _ecrire(t, f'S{r}', f'=SUM(E{r}:I{r},L{r}:R{r})')
    _ecrire(t, 'S13', 'TOTAL')
    _ecrire(t, 'S38', None); _ecrire(t, 'S39', None)
    _ecrire(t, 'S40', '=R40'); _ecrire(t, 'J41', None); _ecrire(t, 'K41', None)
    _ecrire(t, 'B44', "Ce tableau est établi hors TVA ; il intègre les acomptes et le solde de l'impôt sur les bénéfices et les échéances des emprunts.")
    _ecrire(t, 'B45', f"Les mois suivants figurent dans l'onglet {ctx['SM']} (trésorerie de fin de mois identique).")
    for c_ in ('I49', 'S49'):
        _ecrire(t, c_, None)

    # ---------------------------------------------------------------- #68 Tableaux des immobilisations et amortissements
    modele_am = wb['AMORT']
    am_sheets = []
    ordre = [k for k in CATEGORIES]
    if invs:
        for k, y in enumerate(years):
            presents = [(j, iv) for j, iv in enumerate(invs) if iv['mise_en_service'] <= dern[y]]
            if not presents:
                continue
            t = wb.copy_worksheet(modele_am); t.title = f'#68 Tabl amort {y}'; am_sheets.append(t)
            for r in range(10, 86):
                for c_ in range(1, 13):
                    cell = t.cell(r, c_)
                    if type(cell).__name__ != 'MergedCell':
                        cell.value = None
                        cell.font = copy(modele_am.cell(16, c_).font); cell.fill = PatternFill(fill_type=None)
                        cell.border = copy(modele_am.cell(16, c_).border)
            _ecrire(t, 'A3', f'TABLEAU DES IMMOBILISATIONS ET AMORTISSEMENTS AU {"31 DECEMBRE " + str(y) if cal[dern[y]][1] == 12 else "31/" + str(cal[dern[y]][1]).zfill(2) + "/" + str(y)}')
            d_deb = datetime.datetime(y, cal[prem[y]][1], 1)
            _ecrire(t, 'D8', d_deb); _ecrire(t, 'G8', datetime.datetime(y, 12, 31)); _ecrire(t, 'I8', d_deb - datetime.timedelta(days=1))
            for c_ in ('J8', 'K8', 'L8'):
                _ecrire(t, c_, datetime.datetime(y, 12, 31))
            r = 10; cat_rows = []
            for cle in ordre:
                lib_ = CATEGORIES[cle][0]
                items = [(j, iv) for j, iv in presents if iv['categorie'] == lib_]
                if not items:
                    continue
                rc = r; cat_rows.append(rc)
                for c_ in range(1, 13):
                    src = modele_am.cell(17, c_); dst = t.cell(rc, c_)
                    dst.font = copy(src.font); dst.fill = copy(src.fill); dst.border = copy(src.border); dst.alignment = copy(src.alignment); dst.number_format = src.number_format
                t.cell(rc, 2, lib_.upper()); t.cell(rc, 2).alignment = Alignment(horizontal='left', vertical='center')
                for j, iv in items:
                    r += 1
                    for c_ in range(1, 13):
                        src = modele_am.cell(18, c_); dst = t.cell(r, c_)
                        dst.font = copy(src.font); dst.fill = copy(src.fill); dst.border = copy(src.border); dst.alignment = copy(src.alignment); dst.number_format = src.number_format
                    hr = irow[j]; dr = drows[j]
                    t.cell(r, 2, f'={H}!$A${hr}')
                    t.cell(r, 3, f'=DATE({a0},{m0}+{H}!$D${hr},1)'); t.cell(r, 3).number_format = 'dd/mm/yyyy'
                    t.cell(r, 4, f'=IF({H}!$D${hr}<{prem[y]},{H}!$B${hr},0)')
                    t.cell(r, 5, f'=IF(AND({H}!$D${hr}>={prem[y]},{H}!$D${hr}<={dern[y]}),{H}!$B${hr},0)')
                    t.cell(r, 7, f'=D{r}+E{r}-F{r}')
                    t.cell(r, 8, f'=IF({H}!$C${hr}>0,1/{H}!$C${hr},"")'); t.cell(r, 8).number_format = '0.00%'
                    t.cell(r, 9, ('=' + '+'.join(f'{I}!{GL(2 + kk)}{dr}' for kk in range(k))) if k > 0 else 0)
                    t.cell(r, 10, f'={I}!{GL(2 + k)}{dr}')
                    t.cell(r, 11, f'=I{r}+J{r}'); t.cell(r, 12, f'=G{r}-K{r}')
                for c_ in (4, 5, 6, 7, 9, 10, 11, 12):
                    t.cell(rc, c_, f'=SUM({GL(c_)}{rc + 1}:{GL(c_)}{r})')
                r += 2
            rt = r
            for c_ in range(1, 13):
                src = modele_am.cell(82, c_); dst = t.cell(rt, c_)
                dst.font = copy(src.font); dst.fill = copy(src.fill); dst.border = copy(src.border); dst.alignment = copy(src.alignment); dst.number_format = src.number_format
            t.cell(rt, 2, 'TOTAL GENERAL')
            for c_ in (4, 5, 6, 7, 9, 10, 11, 12):
                t.cell(rt, c_, '=' + '+'.join(f'{GL(c_)}{x}' for x in cat_rows))
            t.cell(rt + 2, 2, f"cf dossier de travail, onglet {ctx['SI']}").font = Font(name='Calibri', size=9, italic=True)
            t.print_area = f'A1:L{rt + 2}'
            for c_ in ('M', 'N', 'O', 'P'):
                t.column_dimensions[c_].hidden = True
            t._images = []
    wb.remove(modele_am)

    # ---------------------------------------------------------------- Bilans
    modele_b = wb['BILAN']
    bil_sheets = []
    suff = ['bis', 'ter', 'quater', 'quinquies', 'sexies', 'septies']
    corp_autres = [CAT[k] for k in ('amenagements', 'materiel_industriel', 'agencements', 'transport', 'bureau', 'mobilier')]

    def remplir_bilan(t, k):
        """k = None : bilan d'ouverture ; sinon index de l'exercice."""
        for r in range(8, 51):
            for col in ('C', 'D', 'E', 'J'):
                c = t[f'{col}{r}']
                if type(c).__name__ != 'MergedCell' and not (isinstance(c.value, str) and not c.value.startswith('=')):
                    c.value = None
        for row in t.iter_rows(min_col=12, max_col=24):
            for c in row:
                if type(c).__name__ != 'MergedCell':
                    c.value = None
        brut = lambda cats: '=' + (inv_cat(cats, f'<={dern[years[k]]}') if k is not None else '0')

        def amort(cats):
            if k is None or not irow:
                return 0
            ys = f"{I}!$B${drows[0]}:${GL(2 + k)}${drows[-1]}"
            return '=' + '+'.join(f'SUMPRODUCT(({H}!$F${i0}:$F${i1}="{c}")*{ys})' for c in cats)
        _ecrire(t, 'B9', 'Logiciels, site internet et ERP'); _ecrire(t, 'B10', 'Brevets, concessions, …'); _ecrire(t, 'B11', 'Autres immobilisations incorporelles')
        for r, cats in ((9, [CAT['logiciels']]), (10, [CAT['brevets']]), (13, [CAT['terrains']]), (14, [CAT['constructions']]),
                        (15, corp_autres), (18, FIN)):
            _ecrire(t, f'C{r}', brut(cats)); _ecrire(t, f'D{r}', amort(cats) if r != 18 and r != 13 else 0)
        _ecrire(t, 'B18', 'Dépôts et cautionnements versés')
        for r in (11, 17, 19, 25, 26, 27, 28, 30, 31, 32, 38):
            _ecrire(t, f'C{r}', 0); _ecrire(t, f'D{r}', 0)
        for r in (9, 10, 11, 13, 14, 15, 17, 18, 19, 25, 26, 27, 28, 30, 31, 32, 33, 38):
            _ecrire(t, f'E{r}', f'=C{r}-D{r}')
        tres = f"{H}!$B${ctx['r_t0']}" if k is None else f"{M}!{ctx['yc'][years[k]]}{r_tres}"
        _ecrire(t, 'C33', f'=MAX({tres},0)'); _ecrire(t, 'D33', 0)
        for col in ('C', 'D', 'E'):
            _ecrire(t, f'{col}21', f'=SUM({col}8:{col}20)'); _ecrire(t, f'{col}35', f'=SUM({col}24:{col}34)')
            _ecrire(t, f'{col}49', f'=SUM({col}38:{col}45)')
        _ecrire(t, 'C50', '=C21+C35+C49'); _ecrire(t, 'D50', '=D21+D35+D49'); _ecrire(t, 'E50', '=E21+E35+E49')
        # passif
        cp_ = R
        _ecrire(t, 'J8', f"={H}!$B${ctx['r_cap']}")
        for r in (9, 11, 12, 13, 16, 24, 40, 41, 43, 44, 48):
            _ecrire(t, f'J{r}', 0)
        if k is None:
            _ecrire(t, 'J14', 0); _ecrire(t, 'J15', 0); _ecrire(t, 'J39', 0); _ecrire(t, 'J45', 0)
        else:
            _ecrire(t, 'J14', ('=' + '+'.join(f"{cp_}!{GL(2 + kk)}{ctx['cro']['rn']}" for kk in range(k))) if k > 0 else 0)
            _ecrire(t, 'J15', f"={cp_}!{GL(2 + k)}{ctx['cro']['rn']}")
            _ecrire(t, 'J39', f"={C}!{GL(2 + N + k)}{ctx['r_tCF']}")
            _ecrire(t, 'J45', f"={B}!{GL(3 + k)}{ctx['br']['dfi']}")
        _ecrire(t, 'J46', f"={H}!$B${ctx['r_cca']}")
        _ecrire(t, 'I47', ' Instruments de trésorerie (découvert bancaire)'); _ecrire(t, 'J47', f'=MAX(-{tres},0)')
        _ecrire(t, 'J21', '=SUM(J8:J20)'); _ecrire(t, 'J35', '=SUM(J24:J34)'); _ecrire(t, 'J49', '=SUM(J39:J48)')
        _ecrire(t, 'J50', '=J21+J35+J49')
        _ecrire(t, 'B52', "Contrôle (total actif - total passif, doit être 0) :"); _ecrire(t, 'E52', '=ROUND(E50-J50,0)')
        t['E52'].number_format = NUM

    def bilan(nom, titre, date_txt, k):
        t = wb.copy_worksheet(modele_b); t.title = nom; bil_sheets.append(t)
        t.print_area = 'A1:K53'; t.column_dimensions['F'].hidden = True; t.column_dimensions['D'].width = 26
        _ecrire(t, 'B2', titre); _ecrire(t, 'C5', date_txt); _ecrire(t, 'J5', date_txt)
        remplir_bilan(t, k)
    d0 = f'01/{str(m0).zfill(2)}/{a0}'
    bilan(f'6_Bilan Prévisionnel {d0[:2]}{d0[3:5]}{str(a0)[2:]}', f"Bilan prévisionnel (ouverture {d0})", d0, None)
    for k, y in enumerate(years):
        mm = cal[dern[y]][1]
        jj = (datetime.date(y + (mm == 12), mm % 12 + 1, 1) - datetime.timedelta(days=1)).day
        dt = f'{jj}/{str(mm).zfill(2)}/{y}'
        nom = f"6{suff[k]}_Bilan Prévisionnel {jj}{str(mm).zfill(2)}{str(y)[2:]}"
        if len(nom) > 31:
            nom = f"6{suff[k]}_Bilan {jj}{str(mm).zfill(2)}{str(y)[2:]}"
        bilan(nom, f'Bilan prévisionnel (clôture au {dt} avant affectation)', str(y), k)
    wb.remove(modele_b)

    # ---------------------------------------------------------------- Plan financier à imprimer (un seul onglet, comme le modèle)
    plan = wb['PLAN']; plan.title = '01_Plan financier à imprimer'
    PL = plan
    for bloc, off in (('INV', 9), ('CR', 29), ('PF', 66), ('TRE', 77)):
        _transplanter(wb[bloc], PL, off)
        wb.remove(wb[bloc])
    contact = cfg.get('contact', {})
    ny = len(years)
    crc = [GL(35 + k) for k in range(ny)]          # colonnes du compte de résultat (AI…)

    for c0 in ('V', 'AG', 'AR', 'BG', 'BR', 'CC', 'CM', 'M'):
        _ecrire(PL, f'{c0}6', soc); _ecrire(PL, f'{c0}7', porteur)

    # --- page de garde
    _ecrire(PL, 'B14', f'Etude financière \nprévisionnelle sur {ny} ans')
    _ecrire(PL, 'C23', porteur); _ecrire(PL, 'C28', soc)
    _ecrire(PL, 'C33', contact.get('forme_juridique') or cfg['pays']['nom'])
    _ecrire(PL, 'C42', contact.get('telephone', '')); _ecrire(PL, 'C43', contact.get('email', ''))
    _ecrire(PL, 'C44', ', '.join(x for x in (contact.get('ville', ''), cfg['pays']['nom']) if x))
    _ecrire(PL, 'C47', datetime.date.today()); PL['C47'].number_format = 'dd/mm/yyyy'
    _ecrire(PL, 'H8', None)
    if theme.get('logo'):
        try:
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(theme['logo']); ratio = 60 / img.height; img.height = 60; img.width = int(img.width * ratio)
            PL.add_image(img, 'C37')
        except Exception:
            pass

    def etendre(cols_src, cols_dst, rows):
        for cs, cd in zip(cols_src, cols_dst):
            for r in rows:
                _style_comme_entre(PL, f'{cs}{r}', f'{cd}{r}')

    # --- salaires et charges sociales / détail des amortissements (colonnes X à AB)
    sc = [GL(24 + k) for k in range(ny)]
    etendre(['Z'] * 2, ['AA', 'AB'], list(range(13, 22)) + list(range(29, 52)))
    _ecrire(PL, 'X9', contact.get('forme_juridique') or '')
    _ecrire(PL, 'X10', 'Rémunéré' if any(k_.startswith('6612') for k_ in cfg['charges_fixes_mensuelles']) else 'Non rémunéré')
    for k, col in enumerate(sc):
        cr_ = crc[k]
        _ecrire(PL, f'{col}13', lab[years[k]]); _ecrire(PL, f'{col}29', lab[years[k]])
        _ecrire(PL, f'{col}15', f'={cr_}49'); _ecrire(PL, f'{col}17', f'={cr_}50')
        _ecrire(PL, f'{col}19', f'={cr_}47'); _ecrire(PL, f'{col}21', f'={cr_}48')
        for r in (16, 20):
            _ecrire(PL, f'{col}{r}', '' if k == 0 else f'=IFERROR(({col}{r - 1}-{sc[k - 1]}{r - 1})/{sc[k - 1]}{r - 1},"")')
            PL[f'{col}{r}'].number_format = '0%'

        def dot(cats):
            if not irow:
                return 0
            return '=' + '+'.join(f'SUMPRODUCT(({H}!$F${i0}:$F${i1}="{c}")*{I}!{GL(2 + k)}${drows[0]}:{GL(2 + k)}${drows[-1]})' for c in cats)
        _ecrire(PL, f'{col}33', dot([CAT['brevets']])); _ecrire(PL, f'{col}34', dot([CAT['logiciels']]))
        for j, cle in enumerate(('terrains', 'constructions', 'amenagements', 'materiel_industriel', 'agencements', 'transport', 'bureau', 'mobilier')):
            _ecrire(PL, f'{col}{42 + j}', dot([CAT[cle]]))
        _ecrire(PL, f'{col}31', f'=SUM({col}33:{col}39)'); _ecrire(PL, f'{col}40', f'=SUM({col}42:{col}49)')
        _ecrire(PL, f'{col}51', f'=SUM({col}31,{col}40)')
    for r, ref in ((33, 'K13'), (34, 'K14')):
        _ecrire(PL, f'T{r}', f'={ref}')
    for j in range(8):
        _ecrire(PL, f'T{42 + j}', f'=K{16 + j}')
    _ecrire(PL, 'T54', f"cf dossier de travail, onglet {ctx['SI']}")
    _ecrire(PL, 'T55', f"cf dossier de travail, onglet {ctx['SR']}")

    # --- soldes intermédiaires de gestion et capacité d'autofinancement (AS/AT … BA/BB)
    vcol = [GL(45 + 2 * k) for k in range(ny)]; pcol_ = [GL(46 + 2 * k) for k in range(ny)]
    etendre(['AW', 'AX', 'AW', 'AX'], ['AY', 'AZ', 'BA', 'BB'], range(11, 30))
    for k, (v, pc_) in enumerate(zip(vcol, pcol_)):
        c = crc[k]
        _ecrire(PL, f'{v}11', lab[years[k]]); _ecrire(PL, f'{pc_}11', '%'); _ecrire(PL, f'{v}41', lab[years[k]])
        f = {13: f'={c}10', 14: f'={c}10', 15: f'={c}14', 16: f'={v}14-{v}15', 17: f'={c}17', 18: f'={c}42',
             19: f'={v}16-{v}17-{v}18', 20: f'={c}46', 21: f'=SUM({c}47:{c}50)', 22: f'={v}19-{v}20-{v}21',
             23: f'={c}53', 24: f'={v}22-{v}23', 25: f'={c}52', 26: f'=-{v}25', 27: f'={v}24+{v}26',
             28: f'={c}57', 29: f'={v}28+{v}23'}
        for r, fo in f.items():
            _ecrire(PL, f'{v}{r}', fo)
            _ecrire(PL, f'{pc_}{r}', f'=IFERROR({v}{r}/{v}$14,0)'); PL[f'{pc_}{r}'].number_format = '0%'
        _ecrire(PL, f'{v}43', f'={v}28'); _ecrire(PL, f'{v}44', f'={v}23'); _ecrire(PL, f'{v}45', f'={v}43+{v}44')
        _ecrire(PL, f'{v}46', f'={GL(72 + k)}17'); _ecrire(PL, f'{v}47', f'={v}45-{v}46')
    _ecrire(PL, 'AP18', 'Autres charges'); _ecrire(PL, 'AP46', ' - Remboursement emprunts')
    _ecrire(PL, 'AP49', "Le flux de trésorerie prévisionnel que l'entreprise génère résulte de ses opérations courantes.")
    _ecrire(PL, 'AP50', f"cf dossier de travail, onglet {ctx['SR']}")

    # --- seuil de rentabilité et besoin en fonds de roulement (BI … BM)
    scol = [GL(61 + k) for k in range(ny)]
    etendre(['BK', 'BK'], ['BL', 'BM'], list(range(9, 22)) + list(range(34, 40)))
    cvr_ = ctx.get('cvr', {})
    for k, col in enumerate(scol):
        c = crc[k]; y = years[k]
        cv = '+'.join(f"{M}!{yc[y]}{rw}" for rw in cvr_.values()) or '0'
        _ecrire(PL, f'{col}9', lab[y]); _ecrire(PL, f'{col}34', lab[y])
        _ecrire(PL, f'{col}11', f'={c}10'); _ecrire(PL, f'{col}12', f'={cv}'); _ecrire(PL, f'{col}13', f'={col}12')
        _ecrire(PL, f'{col}14', f'={col}11-{col}13'); _ecrire(PL, f'{col}15', f'=IFERROR({col}14/{col}11,0)')
        _ecrire(PL, f'{col}16', f'={col}11-{col}13-{col}18'); _ecrire(PL, f'{col}17', f'={col}13+{col}16')
        _ecrire(PL, f'{col}18', f'={c}54'); _ecrire(PL, f'{col}19', f'=IFERROR({col}16/{col}15,0)')
        _ecrire(PL, f'{col}20', f'={col}11-{col}19'); _ecrire(PL, f'{col}21', f'={col}19/{nbm[y] * 25}')
        _ecrire(PL, f'{col}36', f'={col}11/{nbm[y] * 30}*$BH36'); _ecrire(PL, f'{col}38', f'={col}12/{nbm[y] * 30}*$BH38')
        _ecrire(PL, f'{col}39', f'={col}36-{col}38')
        PL[f'{col}15'].number_format = '0%'
    _ecrire(PL, 'BE16', 'Coûts fixes'); _ecrire(PL, 'BE21', "Point mort en chiffre d'affaires par jour ouvré")
    _ecrire(PL, 'BH36', 0); _ecrire(PL, 'BH38', 0)
    _ecrire(PL, 'BE41', "Délais de 0 jour : le prévisionnel retient des ventes encaissées et des achats payés au comptant.")

    # --- plan de financement : les colonnes BT…BX viennent du bloc PF ; trésorerie : bloc TRE
    # nettoyage des repères numériques du modèle (numéros de page)
    for ref in ('AM59', 'AM62', 'BM49', 'CH49', 'CR49', 'BB50', 'BX50'):
        _ecrire(PL, ref, None)
    for k in range(ny, 5):  # prévisionnel sur 3 ans : colonnes d'années inutilisées masquées
        for col in (GL(35 + k), GL(24 + k), GL(45 + 2 * k), GL(46 + 2 * k), GL(61 + k), GL(72 + k)):
            PL.column_dimensions[col].hidden = True
    if ny < 5:  # la bordure droite du tableau passe sur la dernière colonne visible
        from copy import copy as _cp
        for base, pas in ((35, 1), (24, 1), (45, 2), (61, 1), (72, 1)):
            der, fin = GL(base + pas * ny - 1 + (pas - 1)), GL(base + pas * 5 - 1 + (pas - 1))
            for r in range(1, 63):
                src, dst = PL[f'{fin}{r}'], PL[f'{der}{r}']
                if type(dst).__name__ == 'MergedCell' or not src.border.right or not src.border.right.style:
                    continue
                b_ = _cp(dst.border); b_.right = _cp(src.border.right); dst.border = b_
    for col in crc + sc + vcol + scol + [GL(72 + k) for k in range(ny)]:
        PL.column_dimensions[col].width = 15
    for col in pcol_:
        PL.column_dimensions[col].width = 6
    for col in ('CD', 'CE', 'CF', 'CG', 'CH', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR'):
        PL.column_dimensions[col].width = 15
    PL.print_area = 'A1:CS62'

    # ---------------------------------------------------------------- Table des matières
    tdm = wb['TDM']; tdm.title = '0_Table des matières'
    _ecrire(tdm, 'B1', soc); _ecrire(tdm, 'G2', porteur); _ecrire(tdm, 'M1', ' - '.join(str(y) for y in years))
    _ecrire(tdm, 'M2', datetime.date.today()); tdm['M2'].number_format = 'dd/mm/yyyy'
    _ecrire(tdm, 'B2', 'PRÉVISIONNEL FINANCIER')
    ordre_onglets = (['0_Table des matières', PL.title]
                     + [s.title for s in am_sheets] + [s.title for s in bil_sheets]
                     + [ctx['SS'], ctx['SH'], ctx['SM'], ctx['SC'], ctx['SI'], ctx['SR'], ctx['SB']])
    st = copy(tdm['B8']._style)
    for r in range(8, 40):
        c = tdm.cell(r, 2)
        if type(c).__name__ != 'MergedCell':
            c.value = None; c.hyperlink = None
    # onglets de travail sans contenu (aucun emprunt / aucun investissement) : masqués, car les formules y renvoient
    masques = set()
    if not cfg.get('financement', {}).get('emprunts'):
        masques.add(ctx['SC'])
    if not invs:
        masques.add(ctx['SI'])
    for i, n in enumerate([n for n in ordre_onglets[1:] if n not in masques]):
        c = tdm.cell(8 + i, 2); c.value = n; c._style = copy(st); c.hyperlink = f"#'{n}'!A1"
        f_ = copy(c.font); f_.sz = 12; c.font = f_; tdm.row_dimensions[8 + i].height = 18
    fin_ = 8 + len(ordre_onglets) - len(masques)
    tdm.cell(fin_, 2, f"Onglets {ctx['SH'][:2]} à {ctx['SS'][:2]} : dossier de travail (seules cellules à saisir : cases bleues de {ctx['SH']}). "
                      "Le plan financier, les tableaux d'amortissement et les bilans en sont issus par formules.").font = Font(name='Calibri', size=10, italic=True)
    if theme.get('marque'):
        tdm.cell(fin_ + 1, 2, theme['marque']).font = Font(name='Calibri', size=9, italic=True, color='FF808080')
    tdm.print_area = f'A1:M{fin_ + 2}'
    if theme.get('logo'):
        try:
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(theme['logo']); ratio = 70 / img.height; img.height = 70; img.width = int(img.width * ratio)
            tdm.add_image(img, 'K4')
        except Exception:
            pass

    # mise en page : un tableau = une page (plan financier : une page par tableau, comme le modèle)
    for ws in (*bil_sheets, *am_sheets):
        ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    # comme le modèle : sauts de page manuels entre les tableaux, échelle unique choisie pour que la page la plus large tienne en A4
    PL.sheet_properties.pageSetUpPr.fitToPage = False
    bornes = [0] + sorted(b_.id for b_ in PL.col_breaks.brk) + [CI('CS')]
    larg = lambda c: (PL.column_dimensions[GL(c)].width or 8.43) if not PL.column_dimensions[GL(c)].hidden else 0
    pages_pouces = [sum(larg(c) * 7 + 5 for c in range(d + 1, f + 1)) / 96 for d, f in zip(bornes, bornes[1:])]
    haut = sum((PL.row_dimensions[r].height or 15) for r in range(1, 63)) / 72
    PL.page_margins.left = PL.page_margins.right = 0.2; PL.page_margins.top = 0.3; PL.page_margins.bottom = 0.3
    PL.page_margins.header = PL.page_margins.footer = 0.15
    PL.page_setup.scale = int(min(100, 750 / max(pages_pouces), 1050 / haut))
    PL.page_setup.paperSize = 9; PL.page_setup.orientation = 'portrait'

    # ordre des onglets, couleurs, pieds de page
    wb._sheets = [wb[n] for n in ordre_onglets]
    for ws in wb.worksheets:
        if ws.title in ordre_onglets[:-7]:
            recolorer(ws, theme.get('primaire', '#8B4513'), theme.get('secondaire', '#DDD4B9'))
        if dev != 'XAF':
            for row in ws.iter_rows():
                for c in row:
                    if 'XAF' in (c.number_format or ''):
                        c.number_format = c.number_format.replace('XAF', dev)
        if theme.get('marque'):
            ws.oddFooter.left.text = theme['marque']; ws.oddFooter.left.size = 8
    for n in masques:
        wb[n].sheet_state = 'hidden'
    wb.active = 0
