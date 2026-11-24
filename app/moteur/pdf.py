"""Synthèse PDF du prévisionnel, aux couleurs choisies par le client."""
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.platypus import (Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

POLICES_PDF = {'Arial': 'Helvetica', 'Calibri': 'Helvetica', 'Verdana': 'Helvetica', 'Trebuchet MS': 'Helvetica',
               'Century Gothic': 'Helvetica', 'Georgia': 'Times-Roman', 'Times New Roman': 'Times-Roman',
               'Cambria': 'Times-Roman'}


def fm(x):
    if x is None:
        return 'n.s.'
    s = f"{abs(x):,.0f}".replace(',', '\u00a0')
    return f"({s})" if x < -0.5 else s


def pc(x):
    return f"{x * 100:.2f}".replace('.', ',').replace(',00', '') + '\u00a0%'


def _hex(c, defaut):
    try:
        return colors.HexColor(c if c.startswith('#') else '#' + c)
    except Exception:
        return colors.HexColor(defaut)


def _clair(c, f=0.88):
    return colors.Color(c.red + (1 - c.red) * f, c.green + (1 - c.green) * f, c.blue + (1 - c.blue) * f)


def generer_pdf(cfg, res, out, theme=None):
    theme = theme or {}
    P = _hex(theme.get('primaire', '#331C77'), '#331C77')
    S = _hex(theme.get('secondaire', '#FF2FED'), '#FF2FED')
    base = POLICES_PDF.get(theme.get('police', ''), 'Helvetica')
    gras = {'Helvetica': 'Helvetica-Bold', 'Times-Roman': 'Times-Bold'}[base]
    dev = cfg.get('devise', 'FCFA')
    ex = res['exercices']; ans = list(ex.keys())

    st = lambda n, **k: ParagraphStyle(n, **{'fontName': base, 'fontSize': 9.5, 'leading': 13, 'alignment': TA_LEFT, **k})
    T1 = st('t1', fontName=gras, fontSize=20, leading=24, textColor=P)
    T2 = st('t2', fontName=gras, fontSize=13, leading=17, textColor=P, spaceBefore=10, spaceAfter=6)
    N = st('n'); PET = st('p', fontSize=7.5, leading=10, textColor=colors.grey)

    def tableau(lignes, larg_1=62 * mm, gras_lignes=()):
        n = len(lignes[0]) - 1
        w = [larg_1] + [(180 * mm - larg_1) / n] * n
        t = Table(lignes, colWidths=w, repeatRows=1)
        sty = [('FONT', (0, 0), (-1, -1), base, 8.5), ('FONT', (0, 0), (-1, 0), gras, 8.5),
               ('BACKGROUND', (0, 0), (-1, 0), P), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
               ('ALIGN', (1, 0), (-1, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
               ('LINEBELOW', (0, 0), (-1, -1), 0.25, colors.HexColor('#DDDDDD')),
               ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3)]
        for i in range(1, len(lignes)):
            if i % 2 == 0:
                sty.append(('BACKGROUND', (0, i), (-1, i), _clair(S, 0.92)))
        for i in gras_lignes:
            sty += [('FONT', (0, i), (-1, i), gras, 8.5), ('BACKGROUND', (0, i), (-1, i), _clair(P, 0.85))]
        t.setStyle(TableStyle(sty))
        return t

    def pied(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(P); canvas.rect(0, A4[1] - 6 * mm, A4[0], 6 * mm, stroke=0, fill=1)
        canvas.setFont(base, 7.5); canvas.setFillColor(colors.grey)
        canvas.drawString(15 * mm, 9 * mm, f"{cfg['societe']} — Prévisionnel financier {ans[0]}-{ans[-1]}")
        canvas.drawRightString(A4[0] - 15 * mm, 9 * mm, f"{theme.get('marque', '')}   page {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=16 * mm, bottomMargin=16 * mm,
                            title=f"Prévisionnel {cfg['societe']}", author=theme.get('marque', ''))
    el = []
    if theme.get('logo'):
        try:
            img = Image(theme['logo']); r = 22 * mm / img.imageHeight
            img.drawHeight = 22 * mm; img.drawWidth = img.imageWidth * r
            if img.drawWidth > 80 * mm:
                k = 80 * mm / img.drawWidth; img.drawWidth *= k; img.drawHeight *= k
            img.hAlign = 'LEFT'; el += [img, Spacer(1, 6 * mm)]
        except Exception:
            pass
    pays = cfg.get('pays', {})
    el += [Paragraph(cfg['societe'], T1),
           Paragraph(f"Prévisionnel financier sur 5 ans ({ans[0]} à {ans[-1]}) — {pays.get('nom', '')}", st('s', fontSize=12, leading=16, textColor=S)),
           Spacer(1, 4 * mm),
           Paragraph(f"Montants en {dev}, hors taxes. Démarrage de l'activité : {res['mois'][0]['label']}. "
                     f"Impôt sur les sociétés retenu : {pc(pays.get('is', 0))}, minimum de perception {pc(pays.get('minimum_ca', 0))} du chiffre d'affaires, "
                     f"charges sociales patronales {pc(pays.get('charges_sociales_employeur', 0))} des salaires bruts.", N)]

    # Indicateurs clés
    el.append(Paragraph('1. Indicateurs clés', T2))
    lig = [['', *ans]]
    def L(lib, cle, fn=fm):
        lig.append([lib, *[fn(ex[a][cle]) for a in ans]])
    L("Chiffre d'affaires", 'ca'); L("Excédent brut d'exploitation (EBE)", 'ebe'); L('Résultat net', 'resultat_net')
    L("Capacité d'autofinancement (CAF)", 'caf'); L('Service de la dette', 'service_dette')
    L('Couverture du service de la dette', 'couverture_dette', lambda v: 'n.s.' if v is None else f"{v:.2f}")
    L("Trésorerie de fin d'exercice", 'tresorerie_fin')
    el.append(tableau(lig, gras_lignes=(3,)))
    el.append(Spacer(1, 2 * mm))
    el.append(Paragraph(f"Trésorerie la plus basse : <b>{fm(res['tresorerie_min'])} {dev}</b> ({res['mois_tresorerie_min']}). "
                        + ("Elle devient négative : il faut augmenter les apports ou le financement, ou étaler les investissements."
                           if res['tresorerie_min'] < 0 else "Elle reste positive sur toute la période."), N))

    # Graphique CA / résultat net
    d = Drawing(180 * mm, 62 * mm)
    bc = VerticalBarChart(); bc.x, bc.y, bc.width, bc.height = 14 * mm, 10 * mm, 160 * mm, 44 * mm
    bc.data = [[ex[a]['ca'] for a in ans], [ex[a]['resultat_net'] for a in ans]]
    bc.categoryAxis.categoryNames = ans; bc.bars[0].fillColor = P; bc.bars[1].fillColor = S
    bc.valueAxis.labels.fontName = base; bc.valueAxis.labels.fontSize = 7
    bc.valueAxis.labelTextFormat = lambda v: f"{v / 1e6:,.0f} M".replace(',', ' ')
    bc.categoryAxis.labels.fontName = base; bc.categoryAxis.labels.fontSize = 8
    bc.barSpacing = 2; bc.groupSpacing = 10
    bc.valueAxis.valueMin = min(0, min(min(x) for x in bc.data))
    d.add(bc)
    d.add(String(14 * mm, 58 * mm, "Chiffre d'affaires", fontName=base, fontSize=8, fillColor=P))
    d.add(String(50 * mm, 58 * mm, 'Résultat net', fontName=base, fontSize=8, fillColor=S))
    el += [Spacer(1, 4 * mm), d]

    # Compte de résultat
    el.append(PageBreak()); el.append(Paragraph('2. Compte de résultat prévisionnel', T2))
    lig = [['', *ans]]
    L("Chiffre d'affaires", 'ca')
    postes_cv = list(ex[ans[0]]['couts_variables'].keys()); postes_cf = list(ex[ans[0]]['charges_fixes'].keys())
    for k in postes_cv:
        lig.append([k, *[fm(ex[a]['couts_variables'][k]) for a in ans]])
    for k in postes_cf:
        lig.append([k.strip(), *[fm(ex[a]['charges_fixes'][k]) for a in ans]])
    L("Total charges d'exploitation", 'charges'); i_ch = len(lig) - 1
    L('EBE', 'ebe'); i_ebe = len(lig) - 1
    L('Dotations aux amortissements', 'dotations'); L("Résultat d'exploitation", 'resultat_exploitation')
    L('Charges financières', 'interets'); L('Résultat avant impôt', 'resultat_avant_impot'); L('Impôt', 'impot')
    L('Résultat net', 'resultat_net'); i_rn = len(lig) - 1
    el.append(tableau(lig, gras_lignes=(1, i_ch, i_ebe, i_rn)))

    # Plan de financement
    el.append(Paragraph('3. Plan de financement', T2))
    fin = cfg['financement']
    lig = [['', *ans]]
    besoins = {a: ex[a]['capex'] + ex[a]['capital_rembourse'] for a in ans}
    ressources = {a: ex[a]['caf'] + ex[a]['deblocages'] + (fin['capital'] + fin.get('comptes_courants', 0) if k == 0 else 0) for k, a in enumerate(ans)}
    lig.append(['Investissements', *[fm(ex[a]['capex']) for a in ans]])
    lig.append(["Remboursement du capital des emprunts", *[fm(ex[a]['capital_rembourse']) for a in ans]])
    lig.append(['Total besoins', *[fm(besoins[a]) for a in ans]])
    lig.append(["Apports (capital et comptes courants)", *[fm(fin['capital'] + fin.get('comptes_courants', 0)) if k == 0 else fm(0) for k, a in enumerate(ans)]])
    lig.append(['Emprunts', *[fm(ex[a]['deblocages']) for a in ans]])
    lig.append(["Capacité d'autofinancement", *[fm(ex[a]['caf']) for a in ans]])
    lig.append(['Total ressources', *[fm(ressources[a]) for a in ans]])
    lig.append(['Excédent (déficit) de l\'exercice', *[fm(ressources[a] - besoins[a]) for a in ans]])
    el.append(tableau(lig, gras_lignes=(3, 7, 8)))
    el.append(Paragraph("La variation de trésorerie diffère légèrement de l'excédent à cause du décalage de paiement du solde d'impôt.", PET))

    # Bilans
    el.append(Paragraph('4. Bilans prévisionnels simplifiés', T2))
    B = res['bilans']; cols = ['ouverture', *ans]
    lig = [['', 'Ouverture', *ans]]
    def LB(lib, cle):
        lig.append([lib, *[fm(B[c].get(cle, 0) if c != 'ouverture' or cle in B[c] else 0) for c in cols]])
    lig.append(['Immobilisations nettes', fm(B['ouverture']['immobilisations']), *[fm(B[a]['immobilisations_nettes']) for a in ans]])
    LB('Trésorerie', 'tresorerie'); LB('Total actif', 'actif'); i_a = len(lig) - 1
    LB('Capital', 'capital'); LB('Report à nouveau', 'report_a_nouveau'); LB("Résultat de l'exercice", 'resultat')
    LB("Comptes courants d'associés", 'comptes_courants'); LB('Dettes financières', 'dettes_financieres')
    LB("Dette d'impôt", 'dette_fiscale'); LB('Total passif', 'passif'); i_p = len(lig) - 1
    el.append(tableau(lig, larg_1=52 * mm, gras_lignes=(i_a, i_p)))

    # Seuil et sensibilités
    el.append(PageBreak()); el.append(Paragraph('5. Seuil de rentabilité et sensibilités', T2))
    s = res['seuil']
    if s.get('ca_mensuel_seuil'):
        el.append(Paragraph(f"Avec un taux de marge sur coûts variables de {pc(s['taux_marge_cv'])} et {fm(s['charges_fixes_mensuelles'])} {dev} "
                            f"de charges fixes par mois, l'entreprise couvre ses charges fixes à partir de <b>{fm(s['ca_mensuel_seuil'])} {dev}</b> "
                            f"de chiffre d'affaires mensuel (hors amortissements et frais financiers).", N))
    else:
        el.append(Paragraph("La marge sur coûts variables est nulle ou négative : le prix de vente ne couvre pas les coûts variables.", N))
    sens = res.get('sensibilites', [])
    if sens:
        lig = [['Scénario', *[f'RN {a}' for a in ans], 'Tréso. min.']]
        for sc in sens:
            lig.append([sc['nom'], *[fm(sc[a]['resultat_net']) for a in ans], fm(sc['tresorerie_min'])])
        el += [Spacer(1, 3 * mm), tableau(lig, larg_1=40 * mm, gras_lignes=(1,))]

    el.append(Paragraph('6. Hypothèses et limites', T2))
    for t in [f"Taux fiscaux et sociaux du pays ({pays.get('nom', '')}) retenus d'après : {pays.get('source', 'à vérifier')}. Ils évoluent avec chaque loi de finances : faites-les confirmer par votre comptable.",
              "Le prévisionnel est établi hors taxes. Les délais de paiement des clients et des fournisseurs (BFR) ne sont pas modélisés : prévoyez une réserve de trésorerie.",
              "Les volumes de vente sont des hypothèses de l'entreprise. Le fichier Excel joint permet de les modifier mois par mois (cellules bleues) ; tous les tableaux se recalculent."]:
        el.append(Paragraph('• ' + t, N))
    doc.build(el, onFirstPage=pied, onLaterPages=pied)
