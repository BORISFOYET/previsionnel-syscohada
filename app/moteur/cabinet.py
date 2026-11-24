"""Remplit les tableaux du modèle de plan financier du cabinet (gabarit_cabinet.xlsx).

Les tableaux reprennent la présentation du modèle (mêmes lignes, mêmes libellés, même mise en page) ;
seules les couleurs (couleur principale et secondaire du client) et les données changent. Toutes les
cellules sont des formules liées au dossier de travail (onglets 10 à 16), qui reste la seule zone de saisie.
Les onglets sans données (tableaux d'amortissement sans immobilisation) ne sont pas générés.
"""
import datetime
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
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
    inv = wb['INV']; inv.title = '01_Invest & Financements'
    y1 = years[0]; c1 = f'<={dern[y1]}'
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
            _ecrire(inv, f'H{r}', '=' + '+'.join(f'IF({H}!$E${erow[j]}<={dern[y1]},{H}!$B${erow[j]},0)' for j in idx))
        else:
            _ecrire(inv, f'B{r}', f'   Prêt n°{k + 1}'); _ecrire(inv, f'H{r}', 0)
    _ecrire(inv, 'B51', 'Subvention n°1'); _ecrire(inv, 'H51', 0)
    _ecrire(inv, 'B52', 'Subvention n°2'); _ecrire(inv, 'H52', 0)
    _ecrire(inv, 'B53', "Autre financement (comptes courants d'associés)"); _ecrire(inv, 'H53', f"={H}!$B${ctx['r_cca']}")
    _ecrire(inv, 'F55', 'TOTAL RESSOURCES'); _ecrire(inv, 'H55', '=SUM(H44,H47,H51:H53)')
    _ecrire(inv, 'B57', f"Programme de l'exercice {lab[y1]}. Les investissements des exercices suivants figurent dans le plan de financement. "
                        f"Détail : onglet {ctx['SH']}.")

    # ---------------------------------------------------------------- 02 Compte de résultat
    cr = wb['CR']; cr.title = '02_Compte de résultat'
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
    pf = wb['PF']; pf.title = '03_Plan de financement'
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

    # ---------------------------------------------------------------- 04 Budgets de trésorerie (un onglet par exercice)
    modele_tre = wb['TRE']
    mcols = ['E', 'F', 'G', 'H', 'I', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
    r_tres = ctx['r_tres']
    codes_ext = [c for c, _ in COMPTES if c not in ('64',)]
    tres_sheets = []
    for k, y in enumerate(years):
        t = wb.copy_worksheet(modele_tre); t.title = f'04_Trésorerie {y}'; tres_sheets.append(t)
        t.print_area = modele_tre.print_area
        t.page_setup.orientation = 'landscape'; t.page_setup.fitToWidth = 2; t.page_setup.fitToHeight = 1
        t.sheet_properties.pageSetUpPr.fitToPage = True
        for b_ in modele_tre.col_breaks.brk:
            t.col_breaks.append(copy(b_))
        _ecrire(t, 'B2', f'Budget prévisionnel de trésorerie {lab[y]}'); _ecrire(t, 'L2', f'Budget prévisionnel de trésorerie {y} (suite)')
        for c_ in ('D6', 'N6'):
            _ecrire(t, c_, soc)
        for c_ in ('D7', 'N7'):
            _ecrire(t, c_, porteur)
        _ecrire(t, 'B11', f'Exercice {lab[y]}')
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
            idx = next((i for i, (a, m) in enumerate(cal) if a == y and m == j + 1), None)
            if idx is None:
                for r in range(13, 42):
                    _ecrire(t, f'{col}{r}', None)
                continue
            mc = GL(2 + idx)
            _ecrire(t, f'{col}13', f'{MOIS[j]}-{str(y)[2:]}')
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
            prev = mcols[j - 1] if j > 0 and any(a == y and m == j for a, m in cal) else None
            if prev:
                _ecrire(t, f'{col}38', f'={prev}40')
            else:
                _ecrire(t, f'{col}38', 0 if idx == 0 else f'={M}!{GL(1 + idx)}{r_tres}')
            _ecrire(t, f'{col}39', f'={col}37-{col}36')
            _ecrire(t, f'{col}40', f'={col}38+{col}39')
            _ecrire(t, f'{col}41', f'=IF({col}40<0,{col}40,"")')
        for r in list(range(15, 38)) + [41]:
            _ecrire(t, f'S{r}', f'=SUM(E{r}:I{r},L{r}:R{r})')
        _ecrire(t, 'S38', None); _ecrire(t, 'S39', None)
        lastm = mcols[cal[dern[y]][1] - 1]
        _ecrire(t, 'S40', f'={lastm}40'); _ecrire(t, 'J41', None); _ecrire(t, 'K41', None)
        _ecrire(t, 'B44', "Ce tableau est établi hors TVA ; il intègre les acomptes et le solde de l'impôt sur les bénéfices et les échéances des emprunts.")
        _ecrire(t, 'B45', f"La trésorerie de fin de mois est identique à celle de l'onglet {ctx['SM']}.")
        for c_ in ('I49', 'S49'):
            _ecrire(t, c_, None)
    wb.remove(modele_tre)

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
                t.cell(rc, 2, lib_.upper())
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

    # ---------------------------------------------------------------- Table des matières
    tdm = wb['TDM']; tdm.title = '0_Table des matières'
    _ecrire(tdm, 'B1', soc); _ecrire(tdm, 'G2', porteur); _ecrire(tdm, 'M1', ' - '.join(str(y) for y in years))
    _ecrire(tdm, 'M2', datetime.date.today()); tdm['M2'].number_format = 'dd/mm/yyyy'
    _ecrire(tdm, 'B2', 'PRÉVISIONNEL FINANCIER')
    ordre_onglets = (['0_Table des matières', '01_Invest & Financements', '02_Compte de résultat', '03_Plan de financement']
                     + [s.title for s in tres_sheets] + [s.title for s in am_sheets] + [s.title for s in bil_sheets]
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
                      "Les tableaux 01 à 6 en sont issus par formules.").font = Font(name='Calibri', size=10, italic=True)
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

    # mise en page : un tableau = une page (le budget de trésorerie : deux pages en largeur)
    for ws in (wb['01_Invest & Financements'], wb['02_Compte de résultat'], wb['03_Plan de financement'], *bil_sheets, *am_sheets):
        ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    for col in 'FGHIJK':
        wb['03_Plan de financement'].column_dimensions[col].width = 17
    for t in tres_sheets:
        for col in mcols + ['S']:
            t.column_dimensions[col].width = 15
        t.column_dimensions['B'].width = 34

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
