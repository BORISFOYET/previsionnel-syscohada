"""Règles de gestion : documents, numérotation, encaissements, caisse, stocks, livres, TVA, tableau de bord.

Chaque fonction reçoit une connexion `cx` (déjà dans une transaction) et l'identifiant de la société `sid` :
aucune donnée n'est lue ou écrite sans ce filtre.
"""
import json
import math
from datetime import date, timedelta
from pathlib import Path

from sqlalchemy import and_, func, insert, select, update
from sqlalchemy.exc import IntegrityError

from .db import articles, clients, clotures, documents, journal, lignes, stock_mvts, tresorerie, utilisateurs

PAYS = json.loads((Path(__file__).resolve().parent / 'data' / 'pays.json').read_text(encoding='utf-8'))['pays']

MODES = {'especes': 'Espèces', 'mobile_money': 'Mobile Money', 'banque': 'Banque'}

# Rubriques du livre de recettes et de dépenses (système minimal de trésorerie, SYSCOHADA révisé)
CATEGORIES_RECETTES = {
    'ventes': 'Ventes et prestations encaissées',
    'apport': "Apport de l'exploitant",
    'emprunt': 'Emprunt reçu',
    'autres_recettes': 'Autres recettes',
    'ecart_caisse': 'Écart de caisse (excédent)',
}
CATEGORIES_DEPENSES = {
    'achats_marchandises': 'Achats de marchandises',
    'achats_matieres': 'Achats de matières et fournitures',
    'transport': 'Transport',
    'loyer': 'Loyer',
    'eau_electricite': 'Eau et électricité',
    'telecom': 'Téléphone et internet',
    'salaires': 'Salaires et charges sociales',
    'impots_taxes': 'Impôts et taxes',
    'tva_reversee': "TVA reversée à l'État",
    'entretien': 'Entretien et réparations',
    'frais_bancaires': 'Frais bancaires et Mobile Money',
    'remboursement_emprunt': "Remboursement d'emprunt",
    'investissement': 'Achat de matériel (investissement)',
    'prelevement': "Prélèvement de l'exploitant",
    'remboursement_client': 'Remboursement client',
    'autres_depenses': 'Autres dépenses',
    'ecart_caisse': 'Écart de caisse (manquant)',
}
PREFIXES = {'devis': 'D', 'facture': 'F', 'avoir': 'AV'}


class ErreurMetier(Exception):
    pass


def tracer(cx, sid, uid, action, detail=''):
    cx.execute(insert(journal).values(societe_id=sid, utilisateur_id=uid, action=action, detail=str(detail)[:500]))


# ---------------------------------------------------------------- calculs
def arrondi(x):
    """Arrondi commercial (0,5 vers le haut), identique à celui de l'écran."""
    return int(math.floor(float(x) + 0.5 + 1e-9))


def calcul_ligne(q, prix, remise, tva):
    ht = arrondi(float(q) * int(prix) * (1 - float(remise or 0) / 100))
    return ht, arrondi(ht * float(tva))


def totaux(lgs):
    ht = tv = 0
    par_taux = {}
    for l in lgs:
        h, t = calcul_ligne(l['quantite'], l['prix'], l.get('remise', 0), l['tva'])
        ht += h
        tv += t
        p = par_taux.setdefault(round(float(l['tva']), 4), [0, 0])
        p[0] += h
        p[1] += t
    return {'ht': ht, 'tva': tv, 'ttc': ht + tv, 'par_taux': sorted(par_taux.items())}


def _nombre(v, defaut=0.0):
    try:
        return float(str(v).replace(' ', '').replace(' ', '').replace(',', '.'))
    except (TypeError, ValueError):
        return defaut


def entier(v, defaut=0):
    return int(round(_nombre(v, defaut)))


# ---------------------------------------------------------------- numérotation
def _prochain_numero(cx, sid, type_, annee):
    prefixe = f'{PREFIXES[type_]}-{annee}-'
    rows = cx.execute(select(documents.c.numero).where(documents.c.societe_id == sid, documents.c.type == type_,
                                                       documents.c.numero.like(prefixe + '%'))).scalars().all()
    n = max((int(r.rsplit('-', 1)[1]) for r in rows if r), default=0) + 1
    return f'{prefixe}{n:04d}'


def _attribuer_numero(cx, sid, doc_id, type_, d):
    for _ in range(5):
        num = _prochain_numero(cx, sid, type_, d.year)
        try:
            with cx.begin_nested():
                cx.execute(update(documents).where(documents.c.id == doc_id).values(numero=num))
            return num
        except IntegrityError:
            continue
    raise ErreurMetier('Numérotation impossible, réessayez.')


def _prochaine_piece(cx, sid):
    return (cx.execute(select(func.max(tresorerie.c.numero)).where(tresorerie.c.societe_id == sid)).scalar() or 0) + 1


# ---------------------------------------------------------------- lectures
def document(cx, sid, doc_id):
    d = cx.execute(select(documents).where(documents.c.id == doc_id, documents.c.societe_id == sid)).mappings().first()
    if not d:
        raise ErreurMetier('Document introuvable.')
    lgs = cx.execute(select(lignes).where(lignes.c.document_id == doc_id).order_by(lignes.c.ordre)).mappings().all()
    return dict(d), [dict(l) for l in lgs]


def article(cx, sid, art_id):
    a = cx.execute(select(articles).where(articles.c.id == art_id, articles.c.societe_id == sid)).mappings().first()
    if not a:
        raise ErreurMetier('Article introuvable.')
    return dict(a)


def _client_nom(cx, sid, client_id):
    if not client_id:
        return None, ''
    c = cx.execute(select(clients.c.id, clients.c.nom).where(clients.c.id == client_id,
                                                             clients.c.societe_id == sid)).first()
    if not c:
        raise ErreurMetier('Client introuvable.')
    return c.id, c.nom


# ---------------------------------------------------------------- lignes saisies
def preparer_lignes(cx, sid, brutes, prix_libre=True):
    """brutes : liste de dict {article_id, designation, quantite, prix, remise, tva}.
    prix_libre=False (employé) : prix, remise et TVA imposés par le catalogue."""
    out = []
    for i, b in enumerate(brutes):
        q = _nombre(b.get('quantite'))
        if q <= 0:
            continue
        art = None
        if b.get('article_id'):
            art = article(cx, sid, int(b['article_id']))
        des = (b.get('designation') or (art['nom'] if art else '')).strip()
        if not des:
            raise ErreurMetier(f'Ligne {i + 1} : indiquez une désignation.')
        if prix_libre or not art:
            if not prix_libre and not art:
                raise ErreurMetier('Choisissez les articles dans le catalogue.')
            prix = entier(b.get('prix'), art['prix_vente'] if art else 0)
            remise = min(max(_nombre(b.get('remise')), 0), 100)
            tva = _nombre(b.get('tva'), art['tva'] if art else 0)
        else:
            prix, remise, tva, des = art['prix_vente'], 0, art['tva'], art['nom']
        if prix < 0:
            raise ErreurMetier(f'Ligne {i + 1} : le prix ne peut pas être négatif.')
        out.append({'article_id': art['id'] if art else None, 'designation': des[:300], 'quantite': q,
                    'prix': prix, 'remise': remise, 'tva': tva, 'ordre': i})
    if not out:
        raise ErreurMetier('Ajoutez au moins une ligne avec une quantité.')
    return out


def _ecrire_lignes(cx, doc_id, lgs):
    cx.execute(lignes.delete().where(lignes.c.document_id == doc_id))
    cx.execute(insert(lignes), [{**l, 'document_id': doc_id} for l in lgs])


# ---------------------------------------------------------------- devis et factures
def enregistrer_document(cx, sid, uid, type_, client_id, d, lgs, note='', echeance=None, doc_id=None):
    """Crée ou modifie un devis (numéroté dès sa création) ou une facture brouillon (numérotée à la validation)."""
    if type_ not in ('devis', 'facture'):
        raise ErreurMetier('Type de document inconnu.')
    cid, cnom = _client_nom(cx, sid, client_id)
    if type_ == 'facture' and not cid:
        raise ErreurMetier('Choisissez un client (ou utilisez la vente au comptoir).')
    t = totaux(lgs)
    vals = dict(client_id=cid, client_nom=cnom, date=d, echeance=echeance, note=note or '',
                total_ht=t['ht'], total_tva=t['tva'], total_ttc=t['ttc'])
    if doc_id:
        doc, _ = document(cx, sid, doc_id)
        if doc['statut'] != 'brouillon':
            raise ErreurMetier('Ce document est validé : il ne peut plus être modifié.')
        cx.execute(update(documents).where(documents.c.id == doc_id).values(**vals))
    else:
        doc_id = cx.execute(insert(documents).values(societe_id=sid, type=type_, statut='brouillon', cree_par=uid,
                                                     **vals)).inserted_primary_key[0]
        if type_ == 'devis':
            _attribuer_numero(cx, sid, doc_id, 'devis', d)
    _ecrire_lignes(cx, doc_id, lgs)
    tracer(cx, sid, uid, f'{type_}_enregistre', f'#{doc_id} {t["ttc"]}')
    return doc_id


def statut_devis(cx, sid, uid, doc_id, statut):
    doc, _ = document(cx, sid, doc_id)
    if doc['type'] != 'devis' or doc['statut'] == 'converti' or statut not in ('brouillon', 'accepte', 'refuse'):
        raise ErreurMetier('Changement impossible.')
    cx.execute(update(documents).where(documents.c.id == doc_id).values(statut=statut))
    tracer(cx, sid, uid, 'devis_statut', f'{doc["numero"]} {statut}')


def convertir_devis(cx, sid, uid, doc_id):
    doc, lgs = document(cx, sid, doc_id)
    if doc['type'] != 'devis' or doc['statut'] in ('converti', 'refuse'):
        raise ErreurMetier('Ce devis ne peut pas être transformé en facture.')
    if not doc['client_id']:
        raise ErreurMetier('Ajoutez un client au devis avant de le transformer en facture.')
    lgs = [{k: l[k] for k in ('article_id', 'designation', 'quantite', 'prix', 'remise', 'tva', 'ordre')} for l in lgs]
    fid = enregistrer_document(cx, sid, uid, 'facture', doc['client_id'], date.today(), lgs,
                               note=f'Selon devis {doc["numero"]}')
    cx.execute(update(documents).where(documents.c.id == fid).values(origine_id=doc_id))
    cx.execute(update(documents).where(documents.c.id == doc_id).values(statut='converti'))
    tracer(cx, sid, uid, 'devis_converti', f'{doc["numero"]} -> facture #{fid}')
    return fid


def _sortie_stock(cx, sid, uid, lgs, d, doc_id, libelle, signe=-1, type_='sortie'):
    """Mouvements de stock des lignes d'articles suivis en stock. Retourne le coût total (CUMP)."""
    cout = 0.0
    for l in lgs:
        if not l.get('article_id'):
            continue
        a = article(cx, sid, l['article_id'])
        if a['type'] != 'produit':
            continue
        q = signe * float(l['quantite'])
        if q < 0 and a['stock'] + q < -1e-9:
            raise ErreurMetier(f'Stock insuffisant pour « {a["nom"]} » : il en reste {a["stock"]:g}. '
                               'Enregistrez d’abord l’entrée en stock.')
        cx.execute(insert(stock_mvts).values(societe_id=sid, article_id=a['id'], date=d, type=type_, quantite=q,
                                             cout_unitaire=a['cump'], libelle=libelle, document_id=doc_id,
                                             utilisateur_id=uid))
        cx.execute(update(articles).where(articles.c.id == a['id']).values(stock=articles.c.stock + q))
        cout += abs(q) * a['cump']
    return round(cout)


def valider_facture(cx, sid, uid, doc_id):
    doc, lgs = document(cx, sid, doc_id)
    if doc['type'] != 'facture' or doc['statut'] != 'brouillon':
        raise ErreurMetier('Seule une facture en brouillon peut être validée.')
    num = _attribuer_numero(cx, sid, doc_id, 'facture', doc['date'])
    cx.execute(update(documents).where(documents.c.id == doc_id).values(statut='valide'))
    _sortie_stock(cx, sid, uid, lgs, doc['date'], doc_id, f'Facture {num}')
    tracer(cx, sid, uid, 'facture_validee', f'{num} {doc["total_ttc"]}')
    return num


def supprimer_brouillon(cx, sid, uid, doc_id):
    doc, _ = document(cx, sid, doc_id)
    if doc['statut'] != 'brouillon' or doc['type'] != 'facture':
        raise ErreurMetier('Seule une facture en brouillon peut être supprimée.')
    cx.execute(documents.delete().where(documents.c.id == doc_id))
    tracer(cx, sid, uid, 'brouillon_supprime', f'#{doc_id}')


def annuler_par_avoir(cx, sid, uid, doc_id, motif=''):
    """Avoir total : la facture validée n'est jamais modifiée ; un avoir numéroté l'annule et remet le stock."""
    doc, lgs = document(cx, sid, doc_id)
    if doc['type'] != 'facture' or doc['statut'] != 'valide':
        raise ErreurMetier('Seule une facture validée peut être annulée par un avoir.')
    aid = cx.execute(insert(documents).values(
        societe_id=sid, type='avoir', statut='valide', date=date.today(), client_id=doc['client_id'],
        client_nom=doc['client_nom'], total_ht=doc['total_ht'], total_tva=doc['total_tva'],
        total_ttc=doc['total_ttc'], origine_id=doc_id, cree_par=uid,
        note=f'Avoir sur facture {doc["numero"]}' + (f' : {motif}' if motif else ''))).inserted_primary_key[0]
    _ecrire_lignes(cx, aid, [{k: l[k] for k in ('article_id', 'designation', 'quantite', 'prix', 'remise', 'tva', 'ordre')}
                             for l in lgs])
    num = _attribuer_numero(cx, sid, aid, 'avoir', date.today())
    cx.execute(update(documents).where(documents.c.id == doc_id).values(statut='annule'))
    _sortie_stock(cx, sid, uid, lgs, date.today(), aid, f'Avoir {num}', signe=1, type_='retour')
    tracer(cx, sid, uid, 'avoir', f'{num} sur {doc["numero"]}')
    return aid


# ---------------------------------------------------------------- trésorerie
def mouvement(cx, sid, uid, sens, mode, categorie, libelle, montant, d=None, tva=0, document_id=None, tiers=''):
    if sens not in ('recette', 'depense') or mode not in MODES:
        raise ErreurMetier('Mouvement invalide.')
    cats = CATEGORIES_RECETTES if sens == 'recette' else CATEGORIES_DEPENSES
    if categorie not in cats:
        raise ErreurMetier('Choisissez une rubrique.')
    montant = entier(montant)
    if montant <= 0:
        raise ErreurMetier('Le montant doit être positif.')
    d = d or date.today()
    if d > date.today():
        raise ErreurMetier('La date ne peut pas être dans le futur.')
    if mode == 'especes' and _jour_cloture(cx, sid, d):
        raise ErreurMetier(f'La caisse du {d:%d/%m/%Y} est clôturée : saisissez le mouvement à la date du jour.')
    tva = max(0, min(entier(tva), montant))
    for _ in range(5):
        try:
            with cx.begin_nested():
                mid = cx.execute(insert(tresorerie).values(
                    societe_id=sid, numero=_prochaine_piece(cx, sid), date=d, sens=sens, mode=mode, categorie=categorie,
                    libelle=(libelle or cats[categorie])[:300], montant=montant, tva=tva, document_id=document_id,
                    tiers=(tiers or '')[:200], utilisateur_id=uid)).inserted_primary_key[0]
            break
        except IntegrityError:
            continue
    else:
        raise ErreurMetier('Enregistrement impossible, réessayez.')
    tracer(cx, sid, uid, f'{sens}_{mode}', f'{montant} {libelle}')
    return mid


def annuler_mouvement(cx, sid, uid, mid):
    m = cx.execute(select(tresorerie).where(tresorerie.c.id == mid, tresorerie.c.societe_id == sid)).mappings().first()
    if not m or m['annule']:
        raise ErreurMetier('Mouvement introuvable.')
    if m['cloture_id']:
        raise ErreurMetier('Ce mouvement est dans une caisse clôturée : passez une écriture inverse.')
    if m['document_id']:
        cx.execute(update(documents).where(documents.c.id == m['document_id'])
                   .values(paye=documents.c.paye - m['montant']))
    cx.execute(update(tresorerie).where(tresorerie.c.id == mid).values(annule=True))
    tracer(cx, sid, uid, 'mouvement_annule', f'pièce {m["numero"]} {m["montant"]}')


def encaisser(cx, sid, uid, doc_id, montant, mode, d=None):
    doc, _ = document(cx, sid, doc_id)
    if doc['type'] != 'facture' or doc['statut'] != 'valide':
        raise ErreurMetier('On ne peut encaisser qu’une facture validée.')
    reste = doc['total_ttc'] - doc['paye']
    montant = entier(montant)
    if montant <= 0 or montant > reste:
        raise ErreurMetier(f'Montant à encaisser entre 1 et {reste}.')
    tva = round(doc['total_tva'] * montant / doc['total_ttc']) if doc['total_ttc'] else 0
    mid = mouvement(cx, sid, uid, 'recette', mode, 'ventes', f'Règlement facture {doc["numero"]}', montant, d,
                    tva=tva, document_id=doc_id, tiers=doc['client_nom'])
    cx.execute(update(documents).where(documents.c.id == doc_id).values(paye=documents.c.paye + montant))
    return mid


def client_comptoir(cx, sid):
    c = cx.execute(select(clients.c.id).where(clients.c.societe_id == sid, clients.c.nom == 'Client comptoir')).scalar()
    return c or cx.execute(insert(clients).values(societe_id=sid, nom='Client comptoir')).inserted_primary_key[0]


def vente_comptoir(cx, sid, uid, lgs, mode, client_id=None):
    """Vente immédiate : facture validée + encaissement total + sortie de stock, en une seule action."""
    cid = client_id or client_comptoir(cx, sid)
    fid = enregistrer_document(cx, sid, uid, 'facture', cid, date.today(), lgs)
    cx.execute(update(documents).where(documents.c.id == fid).values(comptoir=True))
    valider_facture(cx, sid, uid, fid)
    doc, _ = document(cx, sid, fid)
    if doc['total_ttc'] > 0:
        encaisser(cx, sid, uid, fid, doc['total_ttc'], mode)
    return fid


# ---------------------------------------------------------------- caisse
def _jour_cloture(cx, sid, d):
    return cx.execute(select(clotures.c.id).where(clotures.c.societe_id == sid, clotures.c.date >= d)).first() is not None


def soldes(cx, sid, jusqua=None):
    q = select(tresorerie.c.mode, tresorerie.c.sens, func.sum(tresorerie.c.montant)).where(
        tresorerie.c.societe_id == sid, tresorerie.c.annule.is_(False))
    if jusqua:
        q = q.where(tresorerie.c.date <= jusqua)
    s = {m: 0 for m in MODES}
    for mode, sens, tot in cx.execute(q.group_by(tresorerie.c.mode, tresorerie.c.sens)):
        s[mode] += (tot or 0) * (1 if sens == 'recette' else -1)
    return s


def cloturer_caisse(cx, sid, uid, compte, note='', d=None):
    """Clôture du jour : compare l'espèce comptée au solde théorique, enregistre l'écart et fige les mouvements."""
    d = d or date.today()
    if _jour_cloture(cx, sid, d):
        raise ErreurMetier('La caisse de ce jour est déjà clôturée.')
    theorique = soldes(cx, sid, d)['especes']
    compte = entier(compte)
    if compte < 0:
        raise ErreurMetier('Le montant compté ne peut pas être négatif.')
    ecart = compte - theorique
    if ecart:
        mouvement(cx, sid, uid, 'recette' if ecart > 0 else 'depense', 'especes', 'ecart_caisse',
                  f'Écart de caisse du {d:%d/%m/%Y}' + (f' ({note})' if note else ''), abs(ecart), d)
    cid = cx.execute(insert(clotures).values(societe_id=sid, date=d, utilisateur_id=uid, solde_theorique=theorique,
                                             solde_compte=compte, ecart=ecart, note=note or '')).inserted_primary_key[0]
    cx.execute(update(tresorerie).where(tresorerie.c.societe_id == sid, tresorerie.c.mode == 'especes',
                                        tresorerie.c.date <= d, tresorerie.c.cloture_id.is_(None)).values(cloture_id=cid))
    tracer(cx, sid, uid, 'cloture_caisse', f'théorique {theorique} compté {compte} écart {ecart}')
    return cid, ecart


# ---------------------------------------------------------------- stocks
def entree_stock(cx, sid, uid, art_id, quantite, cout, libelle='', paiement=None, d=None):
    """Entrée (achat) : met à jour le coût moyen pondéré ; enregistre la dépense si payée maintenant."""
    a = article(cx, sid, art_id)
    if a['type'] != 'produit':
        raise ErreurMetier('Les services ne se stockent pas.')
    q, cout = _nombre(quantite), _nombre(cout)
    if q <= 0 or cout < 0:
        raise ErreurMetier('Quantité positive et coût unitaire requis.')
    d = d or date.today()
    ancien = max(a['stock'], 0)
    cump = (ancien * a['cump'] + q * cout) / (ancien + q) if ancien + q else cout
    cx.execute(insert(stock_mvts).values(societe_id=sid, article_id=art_id, date=d, type='entree', quantite=q,
                                         cout_unitaire=cout, libelle=libelle or 'Achat', utilisateur_id=uid))
    cx.execute(update(articles).where(articles.c.id == art_id).values(stock=articles.c.stock + q, cump=cump,
                                                                       prix_achat=round(cout)))
    if paiement in MODES:
        ht = round(q * cout)
        mouvement(cx, sid, uid, 'depense', paiement, 'achats_marchandises', f'Achat {a["nom"]} x {q:g}',
                  ht + round(ht * a['tva']), d, tva=round(ht * a['tva']), tiers=libelle)
    tracer(cx, sid, uid, 'entree_stock', f'{a["nom"]} +{q:g} à {cout:g}')


def inventaire(cx, sid, uid, comptes):
    """comptes : {article_id: quantité comptée}. Enregistre les écarts et les retourne."""
    ecarts = []
    for art_id, qc in comptes.items():
        a = article(cx, sid, int(art_id))
        qc = _nombre(qc, None)
        if qc is None or a['type'] != 'produit':
            continue
        diff = qc - a['stock']
        if abs(diff) < 1e-9:
            continue
        cx.execute(insert(stock_mvts).values(societe_id=sid, article_id=a['id'], date=date.today(), type='inventaire',
                                             quantite=diff, cout_unitaire=a['cump'], libelle='Inventaire',
                                             utilisateur_id=uid))
        cx.execute(update(articles).where(articles.c.id == a['id']).values(stock=qc))
        ecarts.append({'article': a['nom'], 'theorique': a['stock'], 'compte': qc, 'ecart': diff,
                       'valeur': round(diff * a['cump'])})
    tracer(cx, sid, uid, 'inventaire', f'{len(ecarts)} écart(s)')
    return ecarts


# ---------------------------------------------------------------- livres, TVA, tableau de bord
def periode_mois(mois):
    """'2026-11' -> (date(2026,11,1), date(2026,11,30))."""
    a, m = (int(x) for x in mois.split('-'))
    debut = date(a, m, 1)
    fin = (date(a + (m == 12), m % 12 + 1, 1) - timedelta(days=1))
    return debut, fin


def livre(cx, sid, sens, debut, fin):
    q = (select(tresorerie, utilisateurs.c.nom.label('par'))
         .select_from(tresorerie.outerjoin(utilisateurs, utilisateurs.c.id == tresorerie.c.utilisateur_id))
         .where(tresorerie.c.societe_id == sid, tresorerie.c.date.between(debut, fin), tresorerie.c.annule.is_(False))
         .order_by(tresorerie.c.date, tresorerie.c.numero))
    if sens:
        q = q.where(tresorerie.c.sens == sens)
    return [dict(r) for r in cx.execute(q).mappings()]


def tva_periode(cx, sid, debut, fin):
    def somme(type_):
        r = cx.execute(select(func.coalesce(func.sum(documents.c.total_ht), 0), func.coalesce(func.sum(documents.c.total_tva), 0))
                       .where(documents.c.societe_id == sid, documents.c.type == type_,
                              documents.c.statut.in_(('valide', 'annule')), documents.c.numero.isnot(None),
                              documents.c.date.between(debut, fin))).first()
        return int(r[0]), int(r[1])
    fht, ftva = somme('facture')
    aht, atva = somme('avoir')
    deductible = cx.execute(select(func.coalesce(func.sum(tresorerie.c.tva), 0)).where(
        tresorerie.c.societe_id == sid, tresorerie.c.sens == 'depense', tresorerie.c.annule.is_(False),
        tresorerie.c.date.between(debut, fin))).scalar()
    collectee = ftva - atva
    return {'ca_ht': fht - aht, 'collectee': collectee, 'deductible': int(deductible),
            'a_reverser': collectee - int(deductible)}


def cout_ventes(cx, sid, debut, fin):
    """Coût des marchandises vendues (CUMP au moment de la vente), retours d'avoirs déduits."""
    r = cx.execute(select(func.coalesce(func.sum(-stock_mvts.c.quantite * stock_mvts.c.cout_unitaire), 0)).where(
        stock_mvts.c.societe_id == sid, stock_mvts.c.type.in_(('sortie', 'retour')),
        stock_mvts.c.date.between(debut, fin))).scalar()
    return round(r or 0)


def marges_articles(cx, sid, debut, fin):
    vente = func.sum(lignes.c.quantite * lignes.c.prix * (1 - func.coalesce(lignes.c.remise, 0) / 100.0))
    rows = cx.execute(select(lignes.c.article_id, lignes.c.designation, func.sum(lignes.c.quantite), vente)
                      .select_from(lignes.join(documents, documents.c.id == lignes.c.document_id))
                      .where(documents.c.societe_id == sid, documents.c.type == 'facture', documents.c.statut == 'valide',
                             documents.c.date.between(debut, fin))
                      .group_by(lignes.c.article_id, lignes.c.designation)).all()
    couts = dict(cx.execute(select(stock_mvts.c.article_id, func.sum(-stock_mvts.c.quantite * stock_mvts.c.cout_unitaire))
                            .where(stock_mvts.c.societe_id == sid, stock_mvts.c.type == 'sortie',
                                   stock_mvts.c.date.between(debut, fin))
                            .group_by(stock_mvts.c.article_id)).all())
    out = []
    for art_id, des, q, ca in rows:
        ca = round(ca or 0)
        c = round(couts.get(art_id) or 0) if art_id else 0
        out.append({'designation': des, 'quantite': q, 'ca': ca, 'cout': c, 'marge': ca - c,
                    'taux': (ca - c) / ca if ca else 0})
    return sorted(out, key=lambda x: -x['marge'])


def tableau_de_bord(cx, sid, mois):
    debut, fin = periode_mois(mois)
    tva = tva_periode(cx, sid, debut, fin)
    tot = lambda sens: int(cx.execute(select(func.coalesce(func.sum(tresorerie.c.montant), 0)).where(
        tresorerie.c.societe_id == sid, tresorerie.c.sens == sens, tresorerie.c.annule.is_(False),
        tresorerie.c.categorie != 'ecart_caisse', tresorerie.c.date.between(debut, fin))).scalar())
    impayes = cx.execute(select(documents.c.id, documents.c.numero, documents.c.client_nom, documents.c.date,
                                documents.c.echeance, (documents.c.total_ttc - documents.c.paye).label('reste'))
                         .where(documents.c.societe_id == sid, documents.c.type == 'facture', documents.c.statut == 'valide',
                                documents.c.total_ttc > documents.c.paye).order_by(documents.c.date)).mappings().all()
    alertes = cx.execute(select(articles.c.id, articles.c.nom, articles.c.stock, articles.c.seuil_alerte, articles.c.unite)
                         .where(articles.c.societe_id == sid, articles.c.actif.is_(True), articles.c.type == 'produit',
                                articles.c.stock <= articles.c.seuil_alerte)).mappings().all()
    ecarts = cx.execute(select(clotures.c.date, clotures.c.ecart, utilisateurs.c.nom)
                        .select_from(clotures.outerjoin(utilisateurs, utilisateurs.c.id == clotures.c.utilisateur_id))
                        .where(clotures.c.societe_id == sid, clotures.c.date.between(debut, fin), clotures.c.ecart != 0)
                        .order_by(clotures.c.date.desc())).mappings().all()
    par_vendeur = cx.execute(select(utilisateurs.c.nom, func.count(documents.c.id), func.sum(documents.c.total_ttc))
                             .select_from(documents.join(utilisateurs, utilisateurs.c.id == documents.c.cree_par))
                             .where(documents.c.societe_id == sid, documents.c.type == 'facture',
                                    documents.c.statut == 'valide', documents.c.date.between(debut, fin))
                             .group_by(utilisateurs.c.nom)).all()
    valeur_stock = cx.execute(select(func.coalesce(func.sum(articles.c.stock * articles.c.cump), 0)).where(
        articles.c.societe_id == sid, articles.c.type == 'produit', articles.c.stock > 0)).scalar()
    cv = cout_ventes(cx, sid, debut, fin)
    return {'mois': mois, 'debut': debut, 'fin': fin, 'ca_ht': tva['ca_ht'], 'tva': tva,
            'recettes': tot('recette'), 'depenses': tot('depense'), 'soldes': soldes(cx, sid),
            'impayes': [dict(i) for i in impayes], 'total_impayes': sum(i['reste'] for i in impayes),
            'alertes': [dict(a) for a in alertes], 'ecarts': [dict(e) for e in ecarts],
            'par_vendeur': [{'nom': n, 'nb': c, 'total': int(t or 0)} for n, c, t in par_vendeur],
            'cout_ventes': cv, 'marge': tva['ca_ht'] - cv, 'valeur_stock': round(valeur_stock or 0)}
