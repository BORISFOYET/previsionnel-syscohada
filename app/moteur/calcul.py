#!/usr/bin/env python3
"""Modèle financier mensuel générique pour business plan.

Usage : python3 modele_financier.py config.json resultats.json

Calcule, mois par mois puis par exercice civil : CA, charges, EBE, amortissements
(prorata de mise en service), intérêts (y compris intercalaires), impôt (forfait ou
réel avec minimum de perception), résultat, CAF, couverture du service de la dette,
trésorerie, bilans équilibrés, seuil de rentabilité, sensibilités.
Voir assets/exemple_config.json pour le format de configuration.
"""
import json, sys, copy

MOIS_FR = ['janv', 'févr', 'mars', 'avr', 'mai', 'juin', 'juil', 'août', 'sept', 'oct', 'nov', 'déc']


def calendrier(debut, n):
    a, m = map(int, debut.split('-'))
    out = []
    for i in range(n):
        mm = (m - 1 + i) % 12
        aa = a + (m - 1 + i) // 12
        out.append({'label': f"{MOIS_FR[mm]}-{str(aa)[2:]}", 'annee': str(aa), 'mois': mm + 1})
    return out


def pmt(r, n, pv):
    return pv / n if r == 0 else pv * r / (1 - (1 + r) ** -n)


def tableau_emprunt(e, n_mois):
    r = e['taux_annuel'] / 12
    d, p, n = e['mois_deblocage'], e['mois_premiere_echeance'], e['duree_mois']
    ann = pmt(r, n, e['montant'])
    crd = 0.0
    rows = []
    for i in range(n_mois):
        deb = e['montant'] if i == d else 0.0
        crd_debut = crd + deb
        it = crd_debut * r if crd_debut > 0 else 0.0
        if i == p:
            ann = pmt(r, n, crd_debut)  # annuité calculée sur le capital restant dû (intérêts capitalisés inclus)
        if p <= i < p + n:
            pay = ann
            pr = ann - it
        else:
            pr = 0.0
            pay = it if (i >= d and e.get('interets_intercalaires', 'payes') == 'payes') else 0.0
            if pay == 0.0 and i >= d:  # intérêts capitalisés
                pr = -it
        crd = crd_debut - pr
        rows.append({'mois': i, 'deblocage': deb, 'crd_debut': crd_debut, 'interets': it,
                     'capital': max(pr, 0.0), 'echeance': pay, 'crd_fin': max(crd, 0.0) if abs(crd) > 1e-6 else 0.0})
    return rows, ann


def calcule(cfg):
    cal = calendrier(cfg['debut'], cfg['nb_mois'])
    N = len(cal)
    annees = sorted({c['annee'] for c in cal})
    prods = cfg['produits']
    fixes = cfg.get('charges_fixes_mensuelles', {})
    fin = cfg['financement']
    fisc = cfg['fiscalite']

    # Amortissements
    invs = cfg.get('investissements', [])
    dot_m = [0.0] * N
    capex = [0.0] * N
    inv_ouv = 0.0
    for inv in invs:
        ms = inv.get('mise_en_service', 0)
        if inv.get('ouverture', True):
            inv_ouv += inv['montant']
        else:
            capex[ms] += inv['montant']
        if inv.get('duree_ans', 0) > 0 and inv['montant'] > 0:
            d = inv['montant'] / (inv['duree_ans'] * 12)
            for i in range(ms, min(N, ms + inv['duree_ans'] * 12)):
                dot_m[i] += d

    # Emprunts
    emp_rows = [tableau_emprunt(e, N) for e in fin.get('emprunts', [])]
    annuites = [a for _, a in emp_rows]

    mois = []
    treso = fin['tresorerie_depart']
    dette_fiscale = 0.0
    ys = {a: {} for a in annees}
    solde_is_a_payer = {}  # mois index -> montant
    for i, c in enumerate(cal):
        ca = sum(p['volumes'][i] * p['prix'] for p in prods)
        vol = sum(p['volumes'][i] for p in prods)
        cv = {}
        for p in prods:
            for k, v in p.get('couts_variables', {}).items():
                cv[k] = cv.get(k, 0.0) + p['volumes'][i] * v
        cf = dict(fixes)
        charges = sum(cv.values()) + sum(cf.values())
        it = sum(r[i]['interets'] for r, _ in emp_rows)
        serv = sum(r[i]['echeance'] for r, _ in emp_rows)
        prin = sum(r[i]['capital'] for r, _ in emp_rows)
        deb = sum(r[i]['deblocage'] for r, _ in emp_rows)
        f = fisc[c['annee']]
        if f['regime'] == 'forfait':
            vm = f.get('versements_mois_index')
            nb = len([j for j in range(N) if cal[j]['annee'] == c['annee']])
            tax_cash = f['montant'] / nb
        else:
            tax_cash = f['minimum_ca'] * ca
        tax_cash += solde_is_a_payer.get(i, 0.0)
        flux = ca + deb - charges - tax_cash - serv - capex[i]
        treso += flux
        row = {'label': c['label'], 'annee': c['annee'], 'volume': vol, 'ca': ca, 'couts_variables': cv,
               'charges_fixes': cf, 'charges': charges, 'ebe': ca - charges, 'dotations': dot_m[i], 'interets': it,
               'service_dette': serv, 'capital_rembourse': prin, 'deblocage': deb, 'impots_payes': tax_cash,
               'capex': capex[i], 'flux': flux, 'tresorerie': treso}
        mois.append(row)
        # fin d'exercice : calcul de l'impôt
        if i == N - 1 or cal[i + 1]['annee'] != c['annee']:
            R = [m for m in mois if m['annee'] == c['annee']]
            S = lambda k: sum(m[k] for m in R)
            rai = S('ebe') - S('dotations') - S('interets')
            if f['regime'] == 'forfait':
                imp = f['montant']; is_calc = None
            else:
                is_calc = f['taux_is'] * max(rai, 0)
                imp = max(is_calc, f['minimum_ca'] * S('ca'))
            acomptes = sum(m['impots_payes'] for m in R) - sum(v for k, v in solde_is_a_payer.items() if cal[k]['annee'] == c['annee'])
            solde = imp - acomptes
            if solde > 0.5:
                idx = i + cfg.get('mois_paiement_solde_is', 3)
                if idx < N:
                    solde_is_a_payer[idx] = solde_is_a_payer.get(idx, 0.0) + solde
            ys[c['annee']] = {'volume': S('volume'), 'ca': S('ca'), 'charges': S('charges'), 'ebe': S('ebe'),
                              'dotations': S('dotations'), 'resultat_exploitation': S('ebe') - S('dotations'),
                              'interets': S('interets'), 'resultat_avant_impot': rai, 'impot': imp, 'is_calcule': is_calc,
                              'resultat_net': rai - imp, 'caf': rai - imp + S('dotations'), 'service_dette': S('service_dette'),
                              'capital_rembourse': S('capital_rembourse'), 'capex': S('capex'), 'deblocages': S('deblocage'),
                              'tresorerie_fin': R[-1]['tresorerie'], 'tresorerie_min': min(m['tresorerie'] for m in R),
                              'couts_variables': {k: sum(m['couts_variables'].get(k, 0) for m in R) for k in R[0]['couts_variables']},
                              'charges_fixes': {k: sum(m['charges_fixes'].get(k, 0) for m in R) for k in R[0]['charges_fixes']},
                              'solde_is_restant': solde}
            y = ys[c['annee']]
            y['couverture_dette'] = (y['caf'] / y['service_dette']) if (y['service_dette'] > 0 and y['caf'] > 0) else None

    # Bilans
    bilans = {'ouverture': {'immobilisations': inv_ouv, 'tresorerie': fin['tresorerie_depart'],
                            'capital': fin['capital'], 'comptes_courants': fin.get('comptes_courants', 0.0)}}
    o = bilans['ouverture']
    o['actif'] = o['immobilisations'] + o['tresorerie']
    o['passif'] = o['capital'] + o['comptes_courants']
    o['ecart'] = o['actif'] - o['passif']
    cum_rn = 0.0; amort = 0.0; brut = inv_ouv
    for a in annees:
        y = ys[a]
        amort += y['dotations']; brut += y['capex']
        ran = cum_rn; cum_rn += y['resultat_net']
        last = max(i for i, c in enumerate(cal) if c['annee'] == a)
        dette = sum(r[last]['crd_fin'] for r, _ in emp_rows)
        # IS restant dû à la clôture (payé l'exercice suivant)
        dfisc = max(y['solde_is_restant'], 0.0)
        b = {'immobilisations_brutes': brut, 'amortissements_cumules': amort, 'immobilisations_nettes': brut - amort,
             'tresorerie': y['tresorerie_fin'], 'capital': fin['capital'], 'report_a_nouveau': ran, 'resultat': y['resultat_net'],
             'comptes_courants': fin.get('comptes_courants', 0.0), 'dettes_financieres': dette, 'dette_fiscale': dfisc}
        b['actif'] = b['immobilisations_nettes'] + b['tresorerie']
        b['passif'] = b['capital'] + ran + y['resultat_net'] + b['comptes_courants'] + dette + dfisc
        b['ecart'] = round(b['actif'] - b['passif'], 2)
        b['capitaux_propres'] = b['capital'] + ran + y['resultat_net']
        bilans[a] = b

    # Seuil de rentabilité (sur le dernier exercice complet)
    a = annees[-1]; y = ys[a]
    cv_tot = sum(y['couts_variables'].values()); cf_mens = sum(fixes.values())
    taux_mcv = (y['ca'] - cv_tot) / y['ca'] if y['ca'] else 0
    seuil = {'taux_marge_cv': taux_mcv, 'charges_fixes_mensuelles': cf_mens,
             'ca_mensuel_seuil': cf_mens / taux_mcv if taux_mcv > 0 else None}
    if len(prods) == 1:
        p = prods[0]; mu = p['prix'] - sum(p.get('couts_variables', {}).values())
        seuil['marge_unitaire'] = mu
        seuil['volume_mensuel_seuil'] = cf_mens / mu if mu > 0 else None
        dot_mens = sum(i['montant'] / i['duree_ans'] / 12 for i in invs if i.get('duree_ans', 0) > 0)
        fm = fisc[a]
        mu2 = mu - (fm.get('minimum_ca', 0) * p['prix'] if fm['regime'] == 'reel' else 0)
        seuil['volume_mensuel_seuil_avec_amort_et_impot_min'] = (cf_mens + dot_mens) / mu2 if mu2 > 0 else None

    return {'mois': mois, 'exercices': ys, 'bilans': bilans, 'seuil': seuil,
            'emprunts': [{'nom': e['nom'], 'annuite': a_, 'tableau': r} for e, (r, a_) in zip(fin.get('emprunts', []), emp_rows)],
            'tresorerie_min': min(m['tresorerie'] for m in mois),
            'mois_tresorerie_min': min(mois, key=lambda m: m['tresorerie'])['label']}


def applique(cfg, s):
    c = copy.deepcopy(cfg)
    t = s['type']
    for p in c['produits']:
        if t == 'prix_abs': p['prix'] = s['valeur']
        elif t == 'prix_pct': p['prix'] *= 1 + s['valeur']
        elif t == 'prix_ttc': p['prix'] /= 1 + s['valeur']
        elif t == 'volume_pct': p['volumes'] = [v * (1 + s['valeur']) for v in p['volumes']]
        elif t == 'cout_var_abs' and s['poste'] in p.get('couts_variables', {}): p['couts_variables'][s['poste']] = s['valeur']
        elif t == 'cout_var_pct' and s['poste'] in p.get('couts_variables', {}): p['couts_variables'][s['poste']] *= 1 + s['valeur']
    if t == 'charge_fixe_pct':
        c['charges_fixes_mensuelles'] = {k: v * (1 + s['valeur']) for k, v in c['charges_fixes_mensuelles'].items()}
    return c


def main():
    cfg = json.load(open(sys.argv[1], encoding='utf-8'))
    res = calcule(cfg)
    sens = [{'nom': 'Scénario de base', **{a: {k: res['exercices'][a][k] for k in ('ebe', 'resultat_net', 'caf', 'couverture_dette')} for a in res['exercices']}, 'tresorerie_min': res['tresorerie_min']}]
    for s in cfg.get('sensibilites', []):
        r = calcule(applique(cfg, s))
        sens.append({'nom': s['nom'], **{a: {k: r['exercices'][a][k] for k in ('ebe', 'resultat_net', 'caf', 'couverture_dette')} for a in r['exercices']}, 'tresorerie_min': r['tresorerie_min']})
    res['sensibilites'] = sens
    json.dump(res, open(sys.argv[2], 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    f = lambda x: f"{x:,.0f}".replace(',', ' ') if isinstance(x, (int, float)) else str(x)
    print(f"== {cfg.get('societe', '')} ==")
    for a, y in res['exercices'].items():
        cd = f"{y['couverture_dette']:.2f}" if y['couverture_dette'] is not None else 'n.s.'
        print(f"{a}: CA {f(y['ca'])} | EBE {f(y['ebe'])} | RN {f(y['resultat_net'])} | CAF {f(y['caf'])} | CAF/dette {cd} | Tréso fin {f(y['tresorerie_fin'])}")
    for k, b in res['bilans'].items():
        if abs(b['ecart']) > 1:
            print(f"ATTENTION bilan {k} déséquilibré : écart {f(b['ecart'])}")
    print(f"Trésorerie minimale : {f(res['tresorerie_min'])} ({res['mois_tresorerie_min']})")
    print('Seuil :', {k: (round(v, 3) if isinstance(v, float) else v) for k, v in res['seuil'].items()})
    for s in sens:
        print(' -', s['nom'], '| tréso min', f(s['tresorerie_min']), '|', ' ; '.join(f"{a} RN {f(s[a]['resultat_net'])}" for a in res['exercices']))


if __name__ == '__main__':
    main()
