"""Transforme les données saisies sur le site en configuration du modèle financier.

Choix de modélisation (identiques dans l'Excel et dans la synthèse PDF) :
- horizon de 5 exercices : du mois de démarrage choisi au 31 décembre de la 4e année suivante
  (le premier exercice peut être incomplet, comme dans le modèle du cabinet) ;
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
NB_MOIS = 60  # maximum (démarrage en janvier)
NB_EXERCICES = 5
SEP = ' — '
ACHATS = '60' + SEP + 'Achats consommés (matières et marchandises)'
AUTRES_VAR = '605' + SEP + 'Autres coûts variables'
SAL_EMP = '6611' + SEP + 'Salaires bruts des employés'
CS_EMP = '6641' + SEP + 'Charges sociales patronales (employés)'
SAL_DIR = '6612' + SEP + 'Rémunération brute du dirigeant'
CS_DIR = '6642' + SEP + 'Charges sociales patronales (dirigeant)'

# Comptes de charges proposés (plan SYSCOHADA révisé), dans l'ordre du compte de résultat du modèle
COMPTES = [
    ('604', 'Achats stockés de matières et fournitures consommables'), ('605', 'Autres achats'), ('608', "Achats d'emballages"),
    ('611', 'Transports sur achats'), ('612', 'Transports sur ventes'), ('613', 'Transports pour le compte de tiers'),
    ('614', 'Transports du personnel'), ('616', 'Transports de plis'), ('618', 'Autres frais de transport (voyages et déplacements)'),
    ('621', 'Sous-traitance générale'), ('622', 'Locations et charges locatives'), ('623', 'Redevances de crédit-bail et contrats assimilés'),
    ('624', 'Entretien, réparations et maintenance'), ('625', "Primes d'assurance"), ('626', 'Études, recherches et documentation'),
    ('627', 'Publicité, publications, relations publiques'), ('628', 'Frais de télécommunications'), ('631', 'Frais bancaires'),
    ('632', "Rémunérations d'intermédiaires et de conseils"), ('633', 'Frais de formation du personnel'),
    ('634', 'Redevances pour brevets, licences, logiciels et droits similaires'), ('635', 'Cotisations'),
    ('637', "Rémunérations de personnel extérieur à l'entreprise"), ('638', 'Autres charges externes'),
    ('65', 'Autres charges'), ('64', 'Impôts et taxes (patente, taxes locales…)'),
]
LIB_COMPTE = dict(COMPTES)

# Frais de démarrage : charges de la première période (non immobilisées en SYSCOHADA révisé)
FRAIS_DEMARRAGE = {
    'etudes': ('626', 'Études et recherches préliminaires'), 'consultants': ('632', 'Consultants et prestataires'),
    'permis': ('64', "Frais d'obtention de permis et autorisations"), 'formation': ('633', 'Formation et développement'),
    'communication': ('627', 'Enseigne et éléments de communication'), 'notaire': ('632', "Honoraires, frais de notaire ou d'avocat"),
    'experts': ('632', 'Honoraires autres experts'), 'compteurs': ('605', "Frais d'ouverture de compteurs"),
    'creation': ('64', "Frais d'enregistrement et de création de société"),
}

# Catégories d'immobilisations : (libellé, groupe du bilan, durée d'amortissement par défaut en années)
CATEGORIES = {
    'brevets': ('Brevets et propriété intellectuelle', 'Immobilisations incorporelles', 5),
    'logiciels': ('Logiciels, site internet et ERP', 'Immobilisations incorporelles', 3),
    'terrains': ('Terrains', 'Immobilisations corporelles', 0),
    'constructions': ('Constructions', 'Immobilisations corporelles', 20),
    'amenagements': ('Aménagements et installations', 'Immobilisations corporelles', 10),
    'materiel_industriel': ('Matériel et outillage industriel, agricole et commercial', 'Immobilisations corporelles', 5),
    'agencements': ('Agencement et aménagement du matériel', 'Immobilisations corporelles', 10),
    'transport': ('Matériel de transport', 'Immobilisations corporelles', 4),
    'bureau': ('Matériel de bureau et matériel informatique', 'Immobilisations corporelles', 3),
    'mobilier': ('Mobilier', 'Immobilisations corporelles', 10),
    'cautions': ('Dépôts et cautionnements versés', 'Immobilisations financières', 0),
}


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


def volumes(depart, croissance_mensuelle, croissance_annuelle, nb=NB_MOIS):
    """Année 1 : croissance mensuelle composée. Années 2 à 5 : le volume mensuel de chaque mois
    est celui du même mois de l'année précédente × (1 + croissance annuelle)."""
    v = []
    for i in range(nb):
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
    duree = int(_num(d.get('duree'), NB_EXERCICES))
    if duree not in (3, 5):
        raise ErreurSaisie('Choisissez une durée de 3 ou 5 ans.')
    nb = 12 * duree - (m0 - 1)
    annees = sorted({a0 + (m0 - 1 + i) // 12 for i in range(nb)})

    produits = []
    for p in d.get('produits', []):
        nom = (p.get('nom') or '').strip()
        if not nom:
            continue
        cv = {}
        if _num(p.get('cout_achat')):
            cv[ACHATS] = _num(p.get('cout_achat'))
        if _num(p.get('autres_couts')):
            cv[AUTRES_VAR] = _num(p.get('autres_couts'))
        produits.append({'nom': nom, 'unite': p.get('unite') or 'unité', 'prix': _num(p.get('prix')),
                         'volumes': volumes(_num(p.get('volume_depart')), _pct(p.get('croissance_mensuelle')),
                                            _pct(p.get('croissance_annuelle')), nb),
                         'couts_variables': cv})
    if not produits:
        raise ErreurSaisie('Ajoutez au moins un produit ou service.')
    # tous les produits doivent connaître tous les postes de coûts variables
    postes = [k for k in (ACHATS, AUTRES_VAR) if any(k in p['couts_variables'] for p in produits)]
    for p in produits:
        for k in postes:
            p['couts_variables'].setdefault(k, 0.0)

    fixes_excel, fixes_calc = {}, {}
    for dirigeant, sal, cs in ((False, SAL_EMP, CS_EMP), (True, SAL_DIR, CS_DIR)):
        masse = sum(_num(x.get('effectif'), 1) * _num(x.get('salaire')) for x in d.get('personnel', [])
                    if bool(x.get('dirigeant')) == dirigeant)
        if masse:
            fixes_excel[sal] = fixes_calc[sal] = masse
            fixes_excel[cs] = {'taux_sur': sal}
            fixes_calc[cs] = masse * pays['charges_sociales_employeur']
    for c in d.get('charges_fixes', []):
        compte = str(c.get('compte') or '638')
        if compte not in LIB_COMPTE:
            raise ErreurSaisie(f'Compte de charge inconnu : {compte}')
        nom = (c.get('nom') or '').strip() or LIB_COMPTE[compte]
        if _num(c.get('montant')):
            k = compte + SEP + nom
            while k in fixes_excel:
                k += ' '
            fixes_excel[k] = fixes_calc[k] = _num(c.get('montant'))
    ponctuelles = []
    for f in d.get('frais_demarrage', []):
        if not _num(f.get('montant')):
            continue
        t = f.get('type') or 'autre'
        if t != 'autre' and t not in FRAIS_DEMARRAGE:
            raise ErreurSaisie(f'Type de frais de démarrage inconnu : {t}')
        code, lib = FRAIS_DEMARRAGE.get(t, ('638', 'Autres frais de démarrage'))
        nom = code + SEP + ((f.get('nom') or '').strip() or lib) + ' (frais de démarrage)'
        while nom in [x['nom'] for x in ponctuelles] or nom in fixes_excel:
            nom += ' '
        mois = int(_num(f.get('mois'), 1)) - 1
        if not 0 <= mois < nb:
            raise ErreurSaisie(f'Mois invalide pour le frais « {nom} » (1 à {nb}).')
        ponctuelles.append({'nom': nom, 'montant': _num(f.get('montant')), 'mois': mois, 'rubrique': lib})
    if not fixes_excel and not ponctuelles and not postes:
        raise ErreurSaisie('Indiquez au moins une charge (coût variable, personnel ou charge fixe).')

    invs = []
    for iv in d.get('investissements', []):
        nom = (iv.get('nom') or '').strip()
        if nom and _num(iv.get('montant')):
            mois = int(_num(iv.get('mois'), 1)) - 1
            if not 0 <= mois < nb:
                raise ErreurSaisie(f'Mois d\'acquisition invalide pour « {nom} » (1 à {nb}).')
            cat = iv.get('categorie') if iv.get('categorie') in CATEGORIES else 'materiel_industriel'
            duree = int(_num(iv.get('duree'), CATEGORIES[cat][2]))
            if cat in ('terrains', 'cautions'):
                duree = 0
            invs.append({'nom': nom, 'montant': _num(iv.get('montant')), 'duree_ans': duree,
                         'ouverture': False, 'mise_en_service': mois, 'categorie': CATEGORIES[cat][0]})

    capital = _num(d.get('capital'))
    cca = _num(d.get('comptes_courants'))
    emprunts = []
    for e in d.get('emprunts', []):
        if not _num(e.get('montant')):
            continue
        duree = int(_num(e.get('duree'), 12)); differe = int(_num(e.get('differe'), 0))
        deb = int(_num(e.get('mois'), 1)) - 1
        if duree <= 0 or not 0 <= deb < nb:
            raise ErreurSaisie(f'Emprunt : durée ou mois de déblocage invalide (mois 1 à {nb}).')
        emprunts.append({'nom': (e.get('nom') or 'Emprunt bancaire').strip(), 'montant': _num(e.get('montant')),
                         'taux_annuel': _pct(e.get('taux')), 'duree_mois': duree, 'mois_deblocage': deb,
                         'mois_premiere_echeance': deb + differe + 1, 'interets_intercalaires': 'payes'})

    mode = 'reel' if d.get('regime_fiscal') == 'reel' else 'auto'
    synth = pays.get('synthetique') or {}
    fisc = {}
    for k, a in enumerate(annees):
        if True:
            fisc[str(a)] = {'regime': 'reel', 'mode': mode, 'synth_seuil': synth.get('seuil_ca', 0), 'synth_bareme': synth.get('bareme', []), 'taux_is': pays['is'], 'minimum_ca': pays['minimum_ca'],
                            'seuil_ca': pays.get('is_seuil_ca', 0), 'taux_majore': pays.get('is_majore', pays['is']),
                            'minimum_plancher': pays.get('minimum_plancher', 0), 'minimum_plafond': pays.get('minimum_plafond', 0)}

    base = {
        'societe': societe, 'porteur': (d.get('porteur') or '').strip(), 'devise': pays['devise_libelle'], 'debut': debut, 'nb_mois': nb,
        'contact': {k: (d.get(k) or '').strip() for k in ('forme_juridique', 'telephone', 'email', 'ville')},
        'pays': {**pays, 'code': code},
        'produits': produits, 'investissements': invs, 'charges_ponctuelles': ponctuelles,
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
