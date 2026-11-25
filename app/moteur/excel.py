#!/usr/bin/env python3
"""Génère un prévisionnel financier Excel entièrement en formules à partir d'une configuration JSON
(même format que modele_financier.py, voir assets/exemple_config.json).

Usage : python3 generer_previsionnel_excel.py config.json sortie.xlsx [--prefixe 07]

Onglets produits (le « dossier de travail ») :
  <p>_Hypothèses du projet   cellules d'entrée en bleu : produits, coûts, investissements, financement, fiscalité
  <p+1>_Détail mensuel        volumes mensuels (entrées) et calculs mois par mois, trésorerie
  <p+2>_Crédit bancaire       tableaux d'amortissement des emprunts (intérêts intercalaires payés)
  <p+3>_Détail investissements dotations par année (prorata de mise en service), immobilisations nettes
  <p+4>_Compte de résultat    par exercice civil, impôt forfaitaire ou réel avec minimum de perception, CAF, couverture
  <p+5>_Bilan simplifié       bilan d'ouverture et de clôture avec contrôle actif - passif
  <p+6>_Synthèse et contrôles indicateurs clés et contrôles
Les noms contiennent un espace pour rester entre apostrophes dans les formules (indispensable quand ils
commencent par un chiffre). Vérifier ensuite avec recalc.py et comparer à modele_financier.py.
"""
import json, sys
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as GL
from openpyxl.worksheet.datavalidation import DataValidation

MOIS_FR = ['janv', 'févr', 'mars', 'avr', 'mai', 'juin', 'juil', 'août', 'sept', 'oct', 'nov', 'déc']
BROWN, BEIGE, INFILL = 'FF331C77', 'FFEDE7F6', 'FFDDEBF7'
POLICE = 'Calibri'


def _argb(hexa):
    h = hexa.lstrip('#').upper()
    return 'FF' + h if len(h) == 6 else h


def _teinte(hexa, f=0.85):
    h = hexa.lstrip('#')
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return 'FF' + ''.join(f'{int(c + (255 - c) * f):02X}' for c in (r, g, b))


def clair(hexa):
    """Couleur d'en-tête : la couleur secondaire si elle est claire, sinon une teinte éclaircie."""
    h = hexa.lstrip('#')
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return _argb(hexa) if (0.299 * r + 0.587 * g + 0.114 * b) > 175 else _teinte(hexa, 0.7)


def _contraste(hexa):
    h = hexa.lstrip('#')
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return 'FF000000' if (0.299 * r + 0.587 * g + 0.114 * b) > 160 else 'FFFFFFFF'

NUM = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'
F = lambda **k: Font(**{'name': POLICE, 'size': 11, **k})
TITRE_TXT = 'FFFFFFFF'
thin = Side(style='thin', color='FFBFBFBF')


def generer(cfg, out, theme=None, p=1, apres=None, classeur=None):
    """theme : {'primaire': '#331C77', 'secondaire': '#FF2FED', 'police': 'Calibri', 'logo': chemin, 'marque': texte pied de page}"""
    global BROWN, BEIGE, POLICE, TITRE_TXT
    theme = theme or {}
    BROWN = _argb(theme.get('primaire', '#8B4513')); TITRE_TXT = _contraste(theme.get('primaire', '#8B4513'))
    BEIGE = clair(theme.get('secondaire', '#DDD4B9')); POLICE = theme.get('police') or 'Calibri'
    pays = cfg.get('pays', {})
    nm = lambda k, t: f"{p + k:02d}_{t}"
    SH, SM, SC, SI, SR, SB, SS = (nm(0, 'Hypothèses du projet'), nm(1, 'Détail mensuel'), nm(2, 'Crédit bancaire'), nm(3, 'Détail investissements'),
                                  nm(4, 'Compte de résultat'), nm(5, 'Bilan simplifié'), nm(6, 'Synthèse et contrôles'))
    q = lambda s: "'" + s + "'"
    H, M, C, I, R, B = map(q, (SH, SM, SC, SI, SR, SB))

    a0, m0 = map(int, cfg['debut'].split('-')); N = cfg['nb_mois']
    cal = [(a0 + (m0 - 1 + i) // 12, (m0 - 1 + i) % 12 + 1) for i in range(N)]
    years = sorted({y for y, _ in cal})
    prods = cfg['produits']; fixes = cfg.get('charges_fixes_mensuelles', {})
    invs = cfg.get('investissements', []); fin = cfg['financement']; emps = fin.get('emprunts', []); fisc = cfg['fiscalite']
    postes_var = []
    for pr in prods:
        for k in pr.get('couts_variables', {}):
            if k not in postes_var: postes_var.append(k)

    if classeur is not None:
        wb = classeur
    else:
        wb = Workbook(); wb.remove(wb.active)
    def title(ws, t, sub=None):
        ws['A1'] = t; ws['A1'].font = Font(name=POLICE, size=14, bold=True, color=TITRE_TXT)
        for cc in range(2, 9): ws.cell(1, cc).fill = PatternFill('solid', fgColor=BROWN)
        ws.row_dimensions[1].height = 24
        ws.oddFooter.left.text = theme.get('marque', '')
        ws.oddFooter.left.size = 8
        ws.oddFooter.right.text = 'Page &P / &N'
        ws.oddFooter.right.size = 8; ws['A1'].fill = PatternFill('solid', fgColor=BROWN)
        if sub: ws['A2'] = sub; ws['A2'].font = F(italic=True)
        ws.sheet_view.showGridLines = False
    def sec(ws, r, t, ncol):
        for c in range(1, ncol + 1): ws.cell(r, c).fill = PatternFill('solid', fgColor=BEIGE)
        ws.cell(r, 1, t).font = F(bold=True)
    def inp(c, v, fmt=NUM):
        c.value = v; c.font = F(color='FF0000FF'); c.fill = PatternFill('solid', fgColor=INFILL); c.number_format = fmt
    def val(c, v, fmt=NUM, bold=False):
        c.value = v; c.font = F(bold=bold); c.number_format = fmt
    def tot(ws, r, ncol):
        for c in range(1, ncol + 1):
            ws.cell(r, c).font = F(bold=True); ws.cell(r, c).fill = PatternFill('solid', fgColor=BEIGE)

    # ======================= Hypothèses =======================
    h = wb.create_sheet(SH); title(h, f"{cfg.get('societe', '')} — Hypothèses du prévisionnel", 'Cellules bleues : données à saisir. Toutes les autres cellules sont calculées.')
    h.column_dimensions['A'].width = 52
    for c in 'BCDEFGH': h.column_dimensions[c].width = 18
    r = 4; sec(h, r, 'A. Produits / services', 4); r += 1
    for j, t in enumerate(['Libellé', 'Unité de vente', 'Prix unitaire HT']): h.cell(r, 1 + j, t).font = F(bold=True)
    prow = {}
    for pr in prods:
        r += 1; inp(h.cell(r, 1), pr['nom'], 'General'); inp(h.cell(r, 2), pr.get('unite', ''), 'General'); inp(h.cell(r, 3), pr['prix']); prow[pr['nom']] = r
    r += 2; sec(h, r, 'B. Coûts variables unitaires (par unité vendue)', 2 + len(prods)); r += 1
    h.cell(r, 1, 'Poste').font = F(bold=True)
    for j, pr in enumerate(prods): h.cell(r, 2 + j, pr['nom']).font = F(bold=True)
    vrow = {}
    for k in postes_var:
        r += 1; h.cell(r, 1, k).font = F(); vrow[k] = r
        for j, pr in enumerate(prods): inp(h.cell(r, 2 + j), pr.get('couts_variables', {}).get(k, 0))
    r += 1; h.cell(r, 1, 'Total coûts variables unitaires').font = F(bold=True)
    for j in range(len(prods)):
        col = GL(2 + j); val(h.cell(r, 2 + j), f"=SUM({col}{vrow[postes_var[0]]}:{col}{r - 1})" if postes_var else 0, bold=True)
    r += 1; h.cell(r, 1, 'Marge unitaire sur coûts variables').font = F(bold=True)
    for j, pr in enumerate(prods):
        val(h.cell(r, 2 + j), f"=C{prow[pr['nom']]}-{GL(2 + j)}{r - 1}", bold=True)
    r += 2; sec(h, r, 'C. Charges fixes mensuelles', 2); frow = {}
    soc_cells = []
    for k, v in fixes.items():
        r += 1; h.cell(r, 1, k).font = F(); frow[k] = r
        if isinstance(v, dict) and 'taux_sur' in v:
            soc_cells.append((h.cell(r, 2), v['taux_sur']))
        else:
            inp(h.cell(r, 2), v)
    r += 1; h.cell(r, 1, 'Total charges fixes mensuelles').font = F(bold=True)
    val(h.cell(r, 2), f"=SUM(B{min(frow.values())}:B{r - 1})" if frow else 0, bold=True); r_fix_tot = r
    ponct = cfg.get('charges_ponctuelles', []); prow_p = {}
    if ponct:
        r += 2; sec(h, r, 'C bis. Charges ponctuelles (frais de démarrage, non immobilisés)', 3); r += 1
        for j, t in enumerate(['Charge', 'Montant', 'Mois (index, 0 = 1er mois)', 'Rubrique']): h.cell(r, 1 + j, t).font = F(bold=True)
        for pc_ in ponct:
            r += 1; inp(h.cell(r, 1), pc_['nom'], 'General'); inp(h.cell(r, 2), pc_['montant']); inp(h.cell(r, 3), pc_.get('mois', 0), '0'); prow_p[pc_['nom']] = r
            inp(h.cell(r, 4), pc_.get('rubrique', ''), 'General')
    r += 2; sec(h, r, 'D. Investissements', 6); r += 1
    for j, t in enumerate(['Bien', 'Montant', 'Durée (ans, 0 = non amorti)', 'Mois de mise en service (index, 0 = 1er mois)', "Au bilan d'ouverture (Oui/Non)", 'Catégorie']):
        h.cell(r, 1 + j, t).font = F(bold=True); h.cell(r, 1 + j).alignment = Alignment(wrap_text=True)
    irow = []
    dv = DataValidation(type='list', formula1='"Oui,Non"'); h.add_data_validation(dv)
    for iv in invs:
        r += 1; inp(h.cell(r, 1), iv['nom'], 'General'); inp(h.cell(r, 2), iv['montant']); inp(h.cell(r, 3), iv.get('duree_ans', 0), '0')
        inp(h.cell(r, 4), iv.get('mise_en_service', 0), '0'); inp(h.cell(r, 5), 'Oui' if iv.get('ouverture', True) else 'Non', 'General'); dv.add(h.cell(r, 5)); irow.append(r)
        inp(h.cell(r, 6), iv.get('categorie', ''), 'General')
    r += 1; h.cell(r, 1, 'Total des investissements').font = F(bold=True)
    val(h.cell(r, 2), f"=SUM(B{irow[0]}:B{irow[-1]})" if irow else 0, bold=True); r_inv_tot = r
    r += 1; h.cell(r, 1, "dont investissements au bilan d'ouverture").font = F()
    val(h.cell(r, 2), f'=SUMIF(E{irow[0]}:E{irow[-1]},"Oui",B{irow[0]}:B{irow[-1]})' if irow else 0); r_inv_ouv = r
    r += 2; sec(h, r, 'E. Financement', 8)
    r += 1; h.cell(r, 1, 'Capital social').font = F(); inp(h.cell(r, 2), fin['capital']); r_cap = r
    r += 1; h.cell(r, 1, 'Trésorerie de départ').font = F(); inp(h.cell(r, 2), fin['tresorerie_depart']); r_t0 = r
    r += 1; h.cell(r, 1, "Comptes courants d'associés").font = F(); inp(h.cell(r, 2), fin.get('comptes_courants', 0)); r_cca = r
    r += 1; h.cell(r, 1, "Contrôle ouverture (investissements d'ouverture + trésorerie - capital - comptes courants = 0)").font = F(italic=True)
    val(h.cell(r, 2), f"=B{r_inv_ouv}+B{r_t0}-B{r_cap}-B{r_cca}")
    r += 2
    for j, t in enumerate(['Emprunt', 'Montant', 'Taux annuel', 'Durée (mois)', 'Mois de déblocage (index)', '1re échéance (index)', 'Annuité calculée']):
        h.cell(r, 1 + j, t).font = F(bold=True); h.cell(r, 1 + j).alignment = Alignment(wrap_text=True)
    erow = []
    for e in emps:
        r += 1; inp(h.cell(r, 1), e['nom'], 'General'); inp(h.cell(r, 2), e['montant']); inp(h.cell(r, 3), e['taux_annuel'], '0.00%')
        inp(h.cell(r, 4), e['duree_mois'], '0'); inp(h.cell(r, 5), e['mois_deblocage'], '0'); inp(h.cell(r, 6), e['mois_premiere_echeance'], '0')
        val(h.cell(r, 7), f"=IF(D{r}>0,IF(C{r}=0,B{r}/D{r},B{r}*(C{r}/12)/(1-(1+C{r}/12)^-D{r})),0)", bold=True); erow.append(r)
    r += 1; h.cell(r, 1, 'Intérêts entre déblocage et 1re échéance : payés à leur date (intérêts intercalaires).').font = F(italic=True, size=9)
    r += 2; sec(h, r, 'F. Fiscalité par exercice', 7); r += 1
    for j, t in enumerate(['Exercice', 'Régime (forfait / reel)', 'Montant forfaitaire', 'Taux IS', 'Minimum de perception (% CA)', 'Seuil de CA du taux majoré (0 = aucun)', 'Taux IS au-delà du seuil']):
        h.cell(r, 1 + j, t).font = F(bold=True)
    yrow = {}
    for y in years:
        f = fisc[str(y)]; r += 1; h.cell(r, 1, y).font = F(bold=True); yrow[y] = r
        inp(h.cell(r, 2), f['regime'], 'General'); inp(h.cell(r, 3), f.get('montant', 0)); inp(h.cell(r, 4), f.get('taux_is', 0), '0.00%'); inp(h.cell(r, 5), f.get('minimum_ca', 0), '0.00%')
        inp(h.cell(r, 6), f.get('seuil_ca', 0)); inp(h.cell(r, 7), f.get('taux_majore', f.get('taux_is', 0)), '0.00%')
    r += 1; h.cell(r, 1, "Mois de l'année où le solde d'IS N-1 est payé (1 à 12)").font = F(); inp(h.cell(r, 2), cfg.get('mois_paiement_solde_is', 3), '0'); r_msolde = r
    r += 2; sec(h, r, 'G. Paramètres du pays', 3)
    r += 1; h.cell(r, 1, 'Pays').font = F(); val(h.cell(r, 2), pays.get('nom', ''), 'General', bold=True)
    r += 1; h.cell(r, 1, 'Devise').font = F(); val(h.cell(r, 2), cfg.get('devise', ''), 'General')
    r += 1; h.cell(r, 1, 'Taux de TVA (pour information : le prévisionnel est établi hors taxes)').font = F(); inp(h.cell(r, 2), pays.get('tva', 0), '0.00%')
    r += 1; h.cell(r, 1, 'Taux des charges sociales patronales (sur salaires bruts)').font = F(); inp(h.cell(r, 2), pays.get('charges_sociales_employeur', 0), '0.00%'); r_soc = r
    if pays.get('source'):
        r += 1; h.cell(r, 1, 'Source des taux : ' + pays['source']).font = F(italic=True, size=9)
    for cell, base in soc_cells:
        val(cell, f"=B{frow[base]}*B{r_soc}")
    h.freeze_panes = 'B4'

    # ======================= Détail mensuel =======================
    m = wb.create_sheet(SM); title(m, 'Activité, exploitation et trésorerie mensuelles', 'Volumes : cellules bleues à saisir.')
    m.column_dimensions['A'].width = 48
    LC = GL(1 + N)  # dernière colonne mois
    yc = {y: GL(2 + N + k) for k, y in enumerate(years)}
    for i in range(N + len(years)): m.column_dimensions[GL(2 + i)].width = 14
    r = 3
    m.cell(r, 1, 'Index du mois').font = F(italic=True)
    for i in range(N): val(m.cell(r, 2 + i), i, '0')
    r = 4; m.cell(r, 1, 'Exercice').font = F(italic=True)
    for i, (y, mo) in enumerate(cal): val(m.cell(r, 2 + i), y, '0')
    for y in years: val(m.cell(r, 1 + N + 1 + years.index(y)), y, '0', bold=True)
    r = 5; m.cell(r, 1, 'Mois de l\'année').font = F(italic=True)
    for i, (y, mo) in enumerate(cal): val(m.cell(r, 2 + i), mo, '0')
    r = 6; m.cell(r, 1, 'Mois').font = F(bold=True)
    for i, (y, mo) in enumerate(cal): c = m.cell(r, 2 + i, f"{MOIS_FR[mo - 1]}-{str(y)[2:]}"); c.font = F(bold=True); c.alignment = Alignment(horizontal='center')
    for y in years: c = m.cell(r, 2 + N + years.index(y), f'Total {y}'); c.font = F(bold=True)
    for cc in range(1, 2 + N + len(years)): m.cell(r, cc).fill = PatternFill('solid', fgColor=BEIGE)
    r = 7; m.cell(r, 1, 'Volumes vendus').font = F(bold=True)
    volr = {}
    for pr in prods:
        r += 1; m.cell(r, 1, f"Volume — {pr['nom']}").font = F(); volr[pr['nom']] = r
        for i in range(N): inp(m.cell(r, 2 + i), pr['volumes'][i])
    def annual(rw, kind='sum'):
        for y in years:
            col = yc[y]
            if kind == 'sum': f = f"=SUMIF($B$4:${LC}$4,{y},$B{rw}:${LC}{rw})"
            else:  # dernière valeur de l'exercice
                f = f"=INDEX($B{rw}:${LC}{rw},MATCH({y},$B$4:${LC}$4,1))"
            val(m[f"{col}{rw}"], f, bold=True)
    for k in volr.values(): annual(k)
    def line(label, fn, bold=False, kind='sum'):
        nonlocal r
        r += 1; m.cell(r, 1, label).font = F(bold=bold)
        for i in range(N): val(m.cell(r, 2 + i), fn(GL(2 + i), i), bold=bold)
        if kind: annual(r, kind)
        if bold: tot(m, r, 1 + N + len(years))
        return r
    r += 1
    car = {}
    for pr in prods:
        car[pr['nom']] = line(f"Chiffre d'affaires — {pr['nom']}", lambda c, i, pr=pr: f"={c}{volr[pr['nom']]}*{H}!$C${prow[pr['nom']]}")
    r_ca = line("Chiffre d'affaires total", lambda c, i: "=" + "+".join(f"{c}{x}" for x in car.values()), True)
    cvr = {}
    for k in postes_var:
        cvr[k] = line(k, lambda c, i, k=k: "=" + "+".join(f"{c}{volr[pr['nom']]}*{H}!${GL(2 + j)}${vrow[k]}" for j, pr in enumerate(prods)))
    cfr = {}
    for k in fixes:
        cfr[k] = line(k, lambda c, i, k=k: f"={H}!$B${frow[k]}")
    cpr = {}
    for k in prow_p:
        cpr[k] = line(k, lambda c, i, k=k: f"=IF({c}$3={H}!$C${prow_p[k]},{H}!$B${prow_p[k]},0)")
    tous = list(cvr.values()) + list(cfr.values()) + list(cpr.values())
    first = min(tous); last = max(tous)
    r_ch = line("Total charges d'exploitation", lambda c, i: f"=SUM({c}{first}:{c}{last})", True)
    r_ebe = line("Excédent brut d'exploitation", lambda c, i: f"={c}{r_ca}-{c}{r_ch}", True)
    r += 1
    r_capex = line('Investissements du mois (hors bilan d\'ouverture)', lambda c, i: (f'=SUMIFS({H}!$B${irow[0]}:$B${irow[-1]},{H}!$D${irow[0]}:$D${irow[-1]},{c}$3,{H}!$E${irow[0]}:$E${irow[-1]},"Non")' if irow else "=0"))
    r_deb = line('Déblocages d\'emprunts', lambda c, i: f"={C}!{c}{{DEB}}")
    r_serv = line('Échéances d\'emprunts (intérêts + capital)', lambda c, i: f"={C}!{c}{{SERV}}")
    r_imp = line('Impôts décaissés', lambda c, i: (f'=IF(INDEX({H}!$B${yrow[years[0]]}:$B${yrow[years[-1]]},MATCH({c}$4,{H}!$A${yrow[years[0]]}:$A${yrow[years[-1]]},0))="forfait",'
                                                 f'INDEX({H}!$C${yrow[years[0]]}:$C${yrow[years[-1]]},MATCH({c}$4,{H}!$A${yrow[years[0]]}:$A${yrow[years[-1]]},0))/COUNTIF($B$4:${LC}$4,{c}$4),'
                                                 f'INDEX({H}!$E${yrow[years[0]]}:$E${yrow[years[-1]]},MATCH({c}$4,{H}!$A${yrow[years[0]]}:$A${yrow[years[-1]]},0))*{c}{r_ca})'
                                                 f'+IF({c}$5={H}!$B${r_msolde},IFERROR(INDEX({R}!$B${{SOLDE}}:${GL(1 + len(years))}${{SOLDE}},MATCH({c}$4-1,{R}!$B$4:${GL(1 + len(years))}$4,0)),0),0)'))
    r_enc = line('Total encaissements', lambda c, i: f"={c}{r_ca}+{c}{r_deb}", True)
    r_dec = line('Total décaissements', lambda c, i: f"={c}{r_ch}+{c}{r_capex}+{c}{r_serv}+{c}{r_imp}", True)
    r_flux = line('Flux net de trésorerie', lambda c, i: f"={c}{r_enc}-{c}{r_dec}")
    # trésorerie (référence à la ligne elle-même)
    r += 1; r_tres = r; m.cell(r, 1, 'Trésorerie en fin de mois').font = F(bold=True)
    for i in range(N):
        c = GL(2 + i); val(m.cell(r, 2 + i), f"={H}!$B${r_t0}+{c}{r_flux}" if i == 0 else f"={GL(1 + i)}{r}+{c}{r_flux}", bold=True)
    annual(r, 'last'); tot(m, r, 1 + N + len(years))
    m.freeze_panes = 'B7'

    # ======================= Crédit =======================
    cr = wb.create_sheet(SC); title(cr, 'Tableaux d\'amortissement des emprunts')
    cr.column_dimensions['A'].width = 40
    for i in range(N + len(years)): cr.column_dimensions[GL(2 + i)].width = 14
    for i in range(N): val(cr.cell(3, 2 + i), i, '0'); val(cr.cell(4, 2 + i), cal[i][0], '0'); cr.cell(6, 2 + i, f"{MOIS_FR[cal[i][1] - 1]}-{str(cal[i][0])[2:]}").font = F(bold=True)
    cr.cell(3, 1, 'Index du mois').font = F(italic=True); cr.cell(4, 1, 'Exercice').font = F(italic=True); cr.cell(6, 1, 'Mois').font = F(bold=True)
    for y in years: cr.cell(6, 2 + N + years.index(y), f'Total {y}').font = F(bold=True)
    for cc in range(1, 2 + N + len(years)): cr.cell(6, cc).fill = PatternFill('solid', fgColor=BEIGE)
    def cann(rw, kind='sum'):
        for y in years:
            col = GL(2 + N + years.index(y))
            f = f"=SUMIF($B$4:${LC}$4,{y},$B{rw}:${LC}{rw})" if kind == 'sum' else f"=INDEX($B{rw}:${LC}{rw},MATCH({y},$B$4:${LC}$4,1))"
            val(cr[f"{col}{rw}"], f, bold=True)
    r = 7; blocks = []
    for k, e in enumerate(emps):
        er = erow[k]; r += 1; cr.cell(r, 1, e['nom']).font = F(bold=True); tot(cr, r, 1 + N + len(years))
        rD, rCD, rI, rP, rK, rCF = r + 1, r + 2, r + 3, r + 4, r + 5, r + 6
        for j, t in enumerate(['Déblocage', 'Capital restant dû en début de mois', 'Intérêts', 'Échéance payée', 'Capital remboursé', 'Capital restant dû en fin de mois']):
            cr.cell(r + 1 + j, 1, t).font = F()
        for i in range(N):
            c = GL(2 + i); pc = GL(1 + i)
            val(cr[f"{c}{rD}"], f"=IF({c}$3={H}!$E${er},{H}!$B${er},0)")
            val(cr[f"{c}{rCD}"], f"={c}{rD}" if i == 0 else f"={pc}{rCF}+{c}{rD}")
            val(cr[f"{c}{rI}"], f"={c}{rCD}*{H}!$C${er}/12")
            val(cr[f"{c}{rP}"], f"=IF(AND({c}$3>={H}!$F${er},{c}$3<{H}!$F${er}+{H}!$D${er}),{H}!$G${er},IF({c}$3>={H}!$E${er},{c}{rI},0))")
            val(cr[f"{c}{rK}"], f"=IF(AND({c}$3>={H}!$F${er},{c}$3<{H}!$F${er}+{H}!$D${er}),{c}{rP}-{c}{rI},0)")
            val(cr[f"{c}{rCF}"], f"=ROUND({c}{rCD}-{c}{rK},2)")
        for rw in (rD, rI, rP, rK): cann(rw)
        cann(rCD, 'last'); cann(rCF, 'last')
        blocks.append((rD, rI, rP, rK, rCF)); r = rCF + 1
    r += 1; r_tD, r_tI, r_tP, r_tK, r_tCF = r + 1, r + 2, r + 3, r + 4, r + 5
    cr.cell(r, 1, 'TOTAL DES EMPRUNTS').font = F(bold=True); tot(cr, r, 1 + N + len(years))
    for j, (t, idx) in enumerate([('Déblocages', 0), ('Intérêts', 1), ('Échéances payées', 2), ('Capital remboursé', 3), ('Capital restant dû en fin de mois', 4)]):
        rw = r + 1 + j; cr.cell(rw, 1, t).font = F(bold=True)
        for i in range(N + len(years)):
            c = GL(2 + i); val(cr[f"{c}{rw}"], ("=" + "+".join(f"{c}{b[idx]}" for b in blocks)) if blocks else 0, bold=True)
    # renseigner les références dans le détail mensuel
    for i in range(N):
        c = GL(2 + i)
        m[f"{c}{r_deb}"].value = m[f"{c}{r_deb}"].value.replace('{DEB}', str(r_tD))
        m[f"{c}{r_serv}"].value = m[f"{c}{r_serv}"].value.replace('{SERV}', str(r_tP))

    # ======================= Investissements =======================
    iv = wb.create_sheet(SI); title(iv, 'Investissements et amortissements', 'Amortissement linéaire au prorata des mois de service.')
    iv.column_dimensions['A'].width = 40
    for k in range(len(years) + 3): iv.column_dimensions[GL(2 + k)].width = 16
    iv.cell(4, 1, 'Dotations aux amortissements').font = F(bold=True)
    for k, y in enumerate(years): iv.cell(4, 2 + k, y).font = F(bold=True)
    for cc in range(1, 2 + len(years)): iv.cell(4, cc).fill = PatternFill('solid', fgColor=BEIGE)
    r = 4; drows = []
    for j, rowh in enumerate(irow):
        r += 1; drows.append(r); val(iv.cell(r, 1), f"={H}!A{rowh}", 'General')
        for k, y in enumerate(years):
            f = (f"=IF({H}!$C${rowh}>0,{H}!$B${rowh}/({H}!$C${rowh}*12)*SUMPRODUCT(({M}!$B$4:${LC}$4={y})*({M}!$B$3:${LC}$3>={H}!$D${rowh})"
                 f"*({M}!$B$3:${LC}$3<{H}!$D${rowh}+{H}!$C${rowh}*12)),0)")
            val(iv.cell(r, 2 + k), f)
    r += 1; r_dot = r; iv.cell(r, 1, 'Total dotations').font = F(bold=True)
    for k in range(len(years)):
        col = GL(2 + k); val(iv.cell(r, 2 + k), f"=SUM({col}{drows[0]}:{col}{drows[-1]})" if drows else 0, bold=True)
    tot(iv, r, 1 + len(years))
    r += 2; r_cpx = r; iv.cell(r, 1, "Investissements de l'exercice (hors ouverture)").font = F()
    for k, y in enumerate(years): val(iv.cell(r, 2 + k), f"=SUMIF({M}!$B$4:${LC}$4,{y},{M}!$B${r_capex}:${LC}${r_capex})")
    r += 1; r_brut = r; iv.cell(r, 1, 'Immobilisations brutes en fin d\'exercice').font = F()
    for k in range(len(years)):
        col = GL(2 + k); val(iv.cell(r, 2 + k), f"={H}!$B${r_inv_ouv}+SUM($B{r_cpx}:{col}{r_cpx})")
    r += 1; r_amc = r; iv.cell(r, 1, 'Amortissements cumulés').font = F()
    for k in range(len(years)):
        col = GL(2 + k); val(iv.cell(r, 2 + k), f"=SUM($B{r_dot}:{col}{r_dot})")
    r += 1; r_net = r; iv.cell(r, 1, 'Immobilisations nettes').font = F(bold=True)
    for k in range(len(years)):
        col = GL(2 + k); val(iv.cell(r, 2 + k), f"={col}{r_brut}-{col}{r_amc}", bold=True)
    tot(iv, r, 1 + len(years))

    # ======================= Compte de résultat =======================
    R_ = wb.create_sheet(SR); title(R_, 'Compte de résultat prévisionnel', 'Par exercice civil.')
    R_.column_dimensions['A'].width = 58
    for k in range(len(years)): R_.column_dimensions[GL(2 + k)].width = 18
    R_.cell(4, 1, 'En devise du projet').font = F(bold=True)
    for k, y in enumerate(years): val(R_.cell(4, 2 + k), y, '0', bold=True)
    for cc in range(1, 2 + len(years)): R_.cell(4, cc).fill = PatternFill('solid', fgColor=BEIGE)
    r = 4; cro = {}
    def crl(label, fn, bold=False, fmt=NUM, key=None):
        nonlocal r
        r += 1; R_.cell(r, 1, label).font = F(bold=bold)
        for k, y in enumerate(years): val(R_.cell(r, 2 + k), fn(GL(2 + k), yc[y], y), fmt, bold)
        if bold: tot(R_, r, 1 + len(years))
        if key: cro[key] = r
        return r
    for pr in prods: crl(f"Chiffre d'affaires — {pr['nom']}", lambda c, mc, y, pr=pr: f"={M}!{mc}{car[pr['nom']]}")
    crl("Chiffre d'affaires total", lambda c, mc, y: f"={M}!{mc}{r_ca}", True, key='ca')
    rr0 = r + 1
    for k in postes_var: crl(k, lambda c, mc, y, k=k: f"={M}!{mc}{cvr[k]}")
    for k in fixes: crl(k, lambda c, mc, y, k=k: f"={M}!{mc}{cfr[k]}")
    for k in cpr: crl(k, lambda c, mc, y, k=k: f"={M}!{mc}{cpr[k]}")
    rlast = r
    crl("Total charges d'exploitation", lambda c, mc, y: f"=SUM({c}{rr0}:{c}{rlast})", True, key='ch')
    crl("Excédent brut d'exploitation (EBE)", lambda c, mc, y: f"={c}{cro['ca']}-{c}{cro['ch']}", True, key='ebe')
    crl('Dotations aux amortissements', lambda c, mc, y: f"={I}!{c}{r_dot}", key='dot')
    crl("Résultat d'exploitation", lambda c, mc, y: f"={c}{cro['ebe']}-{c}{cro['dot']}", True, key='rex')
    crl('Charges financières (intérêts)', lambda c, mc, y: f"={C}!{GL(2 + N + years.index(y))}{r_tI}", key='int')
    crl('Résultat avant impôt', lambda c, mc, y: f"={c}{cro['rex']}-{c}{cro['int']}", True, key='rai')
    hy = lambda col, y: f"{H}!${col}${yrow[y]}"
    crl('Impôt calculé au taux normal (réel)', lambda c, mc, y: f'=IF({hy("B", y)}="reel",IF(AND({hy("F", y)}>0,{c}{cro["ca"]}>{hy("F", y)}),{hy("G", y)},{hy("D", y)})*MAX({c}{cro["rai"]},0),0)', key='iscalc')
    crl('Impôt de l\'exercice (forfait, ou max(IS, minimum de perception))', lambda c, mc, y: f'=IF({hy("B", y)}="forfait",{hy("C", y)},MAX({c}{cro["iscalc"]},{hy("E", y)}*{c}{cro["ca"]}))', key='imp')
    crl('Résultat net', lambda c, mc, y: f"={c}{cro['rai']}-{c}{cro['imp']}", True, key='rn')
    crl('Impôt versé dans l\'exercice (hors solde N-1)', lambda c, mc, y: f'=IF({hy("B", y)}="forfait",{hy("C", y)},{hy("E", y)}*{c}{cro["ca"]})', key='verse')
    crl("Solde d'impôt à payer en N+1", lambda c, mc, y: f"=MAX({c}{cro['imp']}-{c}{cro['verse']},0)", key='solde')
    r += 1
    crl("Capacité d'autofinancement (CAF)", lambda c, mc, y: f"={c}{cro['rn']}+{c}{cro['dot']}", True, key='caf')
    crl('Service de la dette (échéances payées)', lambda c, mc, y: f"={C}!{GL(2 + N + years.index(y))}{r_tP}", key='serv')
    crl('Couverture du service de la dette (CAF / service)', lambda c, mc, y: f'=IF(AND({c}{cro["serv"]}>0,{c}{cro["caf"]}>0),{c}{cro["caf"]}/{c}{cro["serv"]},"n.s.")', fmt='0.00', key='cov')
    for i in range(N):  # référence du solde d'IS dans le détail mensuel
        c = GL(2 + i); m[f"{c}{r_imp}"].value = m[f"{c}{r_imp}"].value.replace('{SOLDE}', str(cro['solde']))

    # ======================= Bilan =======================
    b = wb.create_sheet(SB); title(b, 'Bilan prévisionnel simplifié')
    b.column_dimensions['A'].width = 50
    for k in range(len(years) + 1): b.column_dimensions[GL(2 + k)].width = 18
    b.cell(4, 1, 'En devise du projet').font = F(bold=True); b.cell(4, 2, 'Ouverture').font = F(bold=True)
    for k, y in enumerate(years): val(b.cell(4, 3 + k), y, '0', bold=True)
    for cc in range(1, 3 + len(years)): b.cell(4, cc).fill = PatternFill('solid', fgColor=BEIGE)
    lines = [('ACTIF', None), ('Immobilisations nettes', 'imm'), ('Trésorerie', 'tre'), ('TOTAL ACTIF', 'ta'), ('PASSIF', None), ('Capital social', 'cap'),
             ('Report à nouveau', 'ran'), ("Résultat de l'exercice", 'res'), ('Capitaux propres', 'cp'), ("Comptes courants d'associés", 'cca'),
             ('Dettes financières', 'df'), ("Dette fiscale (solde d'impôt)", 'dfi'), ('TOTAL PASSIF', 'tp'), ('Contrôle actif - passif (doit être 0)', 'ctl')]
    br = {}
    for j, (t, k) in enumerate(lines):
        rw = 5 + j; b.cell(rw, 1, t).font = F(bold=k in (None, 'ta', 'tp', 'cp', 'ctl')); br[k] = rw
        if k in ('ta', 'tp', 'ctl'): tot(b, rw, 2 + len(years))
    def bset(k, col, f): val(b[f"{col}{br[k]}"], f, bold=k in ('ta', 'tp', 'cp', 'ctl'))
    bset('imm', 'B', f"={H}!B{r_inv_ouv}"); bset('tre', 'B', f"={H}!B{r_t0}"); bset('ta', 'B', f"=B{br['imm']}+B{br['tre']}")
    bset('cap', 'B', f"={H}!B{r_cap}"); bset('ran', 'B', 0); bset('res', 'B', 0); bset('cp', 'B', f"=SUM(B{br['cap']}:B{br['res']})")
    bset('cca', 'B', f"={H}!B{r_cca}"); bset('df', 'B', 0); bset('dfi', 'B', 0)
    bset('tp', 'B', f"=B{br['cp']}+B{br['cca']}+B{br['df']}+B{br['dfi']}"); bset('ctl', 'B', f"=ROUND(B{br['ta']}-B{br['tp']},0)")
    for k, y in enumerate(years):
        col = GL(3 + k); rc = GL(2 + k); pcol = GL(2 + k)
        bset('imm', col, f"={I}!{rc}{r_net}"); bset('tre', col, f"={M}!{yc[y]}{r_tres}"); bset('ta', col, f"={col}{br['imm']}+{col}{br['tre']}")
        bset('cap', col, f"={H}!$B${r_cap}"); bset('ran', col, f"={pcol}{br['ran']}+{pcol}{br['res']}" if k > 0 else 0)
        bset('res', col, f"={R}!{rc}{cro['rn']}"); bset('cp', col, f"=SUM({col}{br['cap']}:{col}{br['res']})")
        bset('cca', col, f"={H}!$B${r_cca}"); bset('df', col, f"={C}!{GL(2 + N + k)}{r_tCF}")
        dern_mois = cal[max(i for i, (yy, _) in enumerate(cal) if yy == y)][1]
        bset('dfi', col, f"={R}!{rc}{cro['solde']}" + (f"+IF({dern_mois}<{H}!$B${r_msolde},{R}!{GL(1 + k)}{cro['solde']},0)" if k > 0 else ''))
        bset('tp', col, f"={col}{br['cp']}+{col}{br['cca']}+{col}{br['df']}+{col}{br['dfi']}"); bset('ctl', col, f"=ROUND({col}{br['ta']}-{col}{br['tp']},0)")

    # ======================= Synthèse =======================
    s = wb.create_sheet(SS, 0); title(s, f"{cfg.get('societe', '')} — Synthèse du prévisionnel")
    s.column_dimensions['A'].width = 50
    for k in range(len(years)): s.column_dimensions[GL(2 + k)].width = 18
    s.cell(4, 1, 'Indicateur').font = F(bold=True)
    for k, y in enumerate(years): val(s.cell(4, 2 + k), y, '0', bold=True)
    for cc in range(1, 2 + len(years)): s.cell(4, cc).fill = PatternFill('solid', fgColor=BEIGE)
    items = [("Chiffre d'affaires", 'ca'), ('EBE', 'ebe'), ('Résultat net', 'rn'), ('CAF', 'caf'), ('Service de la dette', 'serv'), ('Couverture du service de la dette', 'cov')]
    r = 4
    for t, k in items:
        r += 1; s.cell(r, 1, t).font = F()
        for j in range(len(years)): val(s.cell(r, 2 + j), f"={R}!{GL(2 + j)}{cro[k]}", '0.00' if k == 'cov' else NUM)
    r += 1; s.cell(r, 1, "Trésorerie de fin d'exercice").font = F()
    for j, y in enumerate(years): val(s.cell(r, 2 + j), f"={M}!{yc[y]}{r_tres}")
    r += 1; s.cell(r, 1, 'Contrôle bilan (actif - passif)').font = F(bold=True)
    for j in range(len(years)): val(s.cell(r, 2 + j), f"={B}!{GL(3 + j)}{br['ctl']}", bold=True)
    r += 1; s.cell(r, 1, 'Trésorerie minimale sur la période').font = F(bold=True); val(s.cell(r, 2), f"=MIN({M}!B{r_tres}:{LC}{r_tres})", bold=True)
    r += 2; s.cell(r, 1, f"Onglets : {SH} (saisie), {SM}, {SC}, {SI}, {SR}, {SB}.").font = F(italic=True, size=9)
    if theme.get('marque'):
        r += 1; s.cell(r, 1, theme['marque']).font = F(italic=True, size=8, color='FF808080')
    if theme.get('logo'):
        try:
            img = XLImage(theme['logo']); ratio = 60 / img.height; img.height = 60; img.width = int(img.width * ratio)
            s.add_image(img, f"{GL(3 + len(years))}2")
        except Exception:
            pass
    for ws in wb.worksheets:
        ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.calculation.fullCalcOnLoad = True
    ctx = dict(wb=wb, SH=SH, SM=SM, SC=SC, SI=SI, SR=SR, SB=SB, SS=SS, years=years, cal=cal, N=N, LC=LC, yc=yc,
               prow=prow, vrow=vrow, frow=frow, prow_p=prow_p, irow=irow, erow=erow, yrow=yrow, drows=drows, r_dot=r_dot,
               r_cap=r_cap, r_cca=r_cca, r_t0=r_t0, r_soc=r_soc, cro=cro, br=br, car=car, cvr=cvr, cfr=cfr, cpr=cpr,
               r_ca=r_ca, r_ch=r_ch, r_ebe=r_ebe, r_capex=r_capex, r_deb=r_deb, r_serv=r_serv, r_imp=r_imp, r_tres=r_tres,
               r_flux=r_flux, r_enc=r_enc, r_dec=r_dec, r_tD=r_tD, r_tI=r_tI, r_tP=r_tP, r_tK=r_tK, r_tCF=r_tCF,
               blocks=blocks, r_brut=r_brut, r_amc=r_amc, r_net=r_net, r_cpx=r_cpx, theme=theme, F=F, BROWN=BROWN, BEIGE=BEIGE)
    if apres:
        apres(cfg, ctx)
    wb.save(out)
    return wb.sheetnames


def main():
    cfg = json.load(open(sys.argv[1], encoding='utf-8'))
    print(generer(cfg, sys.argv[2]))


if __name__ == '__main__':
    main()
