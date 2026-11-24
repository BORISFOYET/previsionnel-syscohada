"""Transforme les données saisies sur le site en configuration du modèle financier.

Choix de modélisation (identiques dans l'Excel et dans la synthèse PDF) :
- horizon de 60 mois à partir du mois de démarrage choisi ;
- tous les investissements sont achetés au mois indiqué (pas d'actif au bilan d'ouverture) :
  le bilan d'ouverture ne contient que la trésorerie apportée (capital + comptes courants) ;
- prévisionnel établi hors taxes ; la TVA est affichée pour information ;
- impôt : régime réel avec minimum de perception du pays, ou forfait la première année.
"""
import copy
import json
import re
from pathlib import Path

PAYS = json.loads((Path(__file__).resolve().parent.parent / 'data' / 'pays.json').read_text(encoding='utf-8'))['pays']
NB_MOIS = 60
SALAIRES = 'Salaires bruts du personnel'
CHARGES_SOC = 'Charges sociales patronales'


class ErreurSaisie(ValueError):
    pass


def _num(v, defaut=0.0):
    if v in (None, ''):
        return defaut
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(' ', '').replace(' ', '').replace(',', '.').replace('%', '')
    try:
        return float(s)
    except ValueError:
        raise ErreurSaisie(f'Valeur numérique invalide : « {v} »')


def _pct(v, defaut=0.0):
    """Accepte 5 ou 0,05 pour 5 % : toute valeur > 1 est lue comme un pourcentage."""
    x = _num(v, defaut)
    return x / 100 if abs(x) > 1 else x


def normaliser_societe(nom):
    s = re.sub(r'[^A-Z0-9]+', ' ', (nom or '').upper()).strip()
    return re.sub(r'\b(SARL|SA|SAS|SASU|SUARL|SNC|GIE|ETS|ETABLISSEMENTS?)\b', '', s).strip() or s


def volumes(depart, croissance_mensuelle, croissance_annuelle):
    """Année 1 : croissance mensuelle composée. Années 2 à 5 : le volume mensuel de chaque mois
    est celui du même mois de l'année précédente × (1 + croissance annuelle)."""
    v = []
    for i in range(NB_MOIS):
        if i < 12:
            v.append(round(depart * (1 + croissance_mensuelle) ** i, 2))
        else:
            v.append(round(v[i - 12] * (1 + croissance_annuelle), 2))
    return v


def construire_config(d):
    """d : dictionnaire JSON envoyé par le formulaire. Retourne (cfg_excel, cfg_calcul)."""
    societe = (d.get('societe') or '').strip()
    if not societe:
        raise ErreurSaisie('Indiquez la raison sociale.')
    code = d.get('pays')
    if code not in PAYS:
        raise ErreurSaisie('Choisissez un pays.')
    pays = PAYS[code]
    debut = d.get('debut') or ''
    if not re.fullmatch(r'20\d\d-(0[1-9]|1[0-2])', debut):
        raise ErreurSaisie('Indiquez le mois de démarrage (AAAA-MM).')
    a0, m0 = map(int, debut.split('-'))
    annees = sorted({a0 + (m0 - 1 + i) // 12 for i in range(NB_MOIS)})

    produits = []
    for p in d.get('produits', []):
        nom = (p.get('nom') or '').strip()
        if not nom:
            continue
        cv = {}
        if _num(p.get('cout_achat')):
            cv['Achats et matières consommés'] = _num(p.get('cout_achat'))
        if _num(p.get('autres_couts')):
            cv['Autres coûts variables'] = _num(p.get('autres_couts'))
        produits.append({'nom': nom, 'unite': p.get('unite') or 'unité', 'prix': _num(p.get('prix')),
                         'volumes': volumes(_num(p.get('volume_depart')), _pct(p.get('croissance_mensuelle')),
                                            _pct(p.get('croissance_annuelle'))),
                         'couts_variables': cv})
    if not produits:
        raise ErreurSaisie('Ajoutez au moins un produit ou service.')
    # tous les produits doivent connaître tous les postes de coûts variables
    postes = sorted({k for p in produits for k in p['couts_variables']})
    for p in produits:
        for k in postes:
            p['couts_variables'].setdefault(k, 0.0)

    masse = sum(_num(x.get('effectif'), 1) * _num(x.get('salaire')) for x in d.get('personnel', []))
    fixes_excel, fixes_calc = {}, {}
    if masse:
        fixes_excel[SALAIRES] = masse
        fixes_excel[CHARGES_SOC] = {'taux_sur': SALAIRES}
        fixes_calc[SALAIRES] = masse
        fixes_calc[CHARGES_SOC] = masse * pays['charges_sociales_employeur']
    for c in d.get('charges_fixes', []):
        nom = (c.get('nom') or '').strip()
        if nom and _num(c.get('montant')):
            k = nom
            while k in fixes_excel:
                k += ' '
            fixes_excel[k] = fixes_calc[k] = _num(c.get('montant'))

    invs = []
    for iv in d.get('investissements', []):
        nom = (iv.get('nom') or '').strip()
        if nom and _num(iv.get('montant')):
            mois = int(_num(iv.get('mois'), 1)) - 1
            if not 0 <= mois < NB_MOIS:
                raise ErreurSaisie(f'Mois d\'acquisition invalide pour « {nom} » (1 à {NB_MOIS}).')
            invs.append({'nom': nom, 'montant': _num(iv.get('montant')), 'duree_ans': int(_num(iv.get('duree'))),
                         'ouverture': False, 'mise_en_service': mois})

    capital = _num(d.get('capital'))
    cca = _num(d.get('comptes_courants'))
    emprunts = []
    for e in d.get('emprunts', []):
        if not _num(e.get('montant')):
            continue
        duree = int(_num(e.get('duree'), 12)); differe = int(_num(e.get('differe'), 0))
        deb = int(_num(e.get('mois'), 1)) - 1
        if duree <= 0 or not 0 <= deb < NB_MOIS:
            raise ErreurSaisie('Emprunt : durée ou mois de déblocage invalide.')
        emprunts.append({'nom': (e.get('nom') or 'Emprunt bancaire').strip(), 'montant': _num(e.get('montant')),
                         'taux_annuel': _pct(e.get('taux')), 'duree_mois': duree, 'mois_deblocage': deb,
                         'mois_premiere_echeance': deb + differe + 1, 'interets_intercalaires': 'payes'})

    regime1 = d.get('regime_annee1', 'reel')
    fisc = {}
    for k, a in enumerate(annees):
        if k == 0 and regime1 == 'forfait':
            fisc[str(a)] = {'regime': 'forfait', 'montant': _num(d.get('montant_forfait'))}
        else:
            fisc[str(a)] = {'regime': 'reel', 'taux_is': pays['is'], 'minimum_ca': pays['minimum_ca']}

    base = {
        'societe': societe, 'devise': pays['devise_libelle'], 'debut': debut, 'nb_mois': NB_MOIS,
        'pays': {**pays, 'code': code},
        'produits': produits, 'investissements': invs,
        'financement': {'capital': capital, 'tresorerie_depart': capital + cca, 'comptes_courants': cca, 'emprunts': emprunts},
        'fiscalite': fisc, 'mois_paiement_solde_is': int(pays.get('mois_solde_is', 4)),
        'sensibilites': [
            {'nom': 'Volumes -15 %', 'type': 'volume_pct', 'valeur': -0.15},
            {'nom': 'Prix de vente -10 %', 'type': 'prix_pct', 'valeur': -0.10},
            {'nom': 'Charges fixes +10 %', 'type': 'charge_fixe_pct', 'valeur': 0.10},
        ],
    }
    cfg_excel = copy.deepcopy(base); cfg_excel['charges_fixes_mensuelles'] = fixes_excel
    cfg_calc = copy.deepcopy(base); cfg_calc['charges_fixes_mensuelles'] = fixes_calc
    return cfg_excel, cfg_calc
