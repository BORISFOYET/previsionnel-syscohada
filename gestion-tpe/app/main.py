"""Kit de gestion TPE Account's : devis, factures, caisse, stocks, livres de recettes et de dépenses SYSCOHADA."""
import base64
import json
import os
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import func, insert, select, update
from sqlalchemy.exc import IntegrityError

from . import auth, db, exports, licences, metier
from .metier import CATEGORIES_DEPENSES, CATEGORIES_RECETTES, MODES, PAYS, ErreurMetier, entier
from .pdf import document_pdf, montant

ICI = Path(__file__).resolve().parent
BOUTIQUE = os.environ.get('URL_BOUTIQUE', 'https://chariow.com')
NOM = os.environ.get('NOM_PRODUIT', "Kit Gestion TPE")
FORMES = ['', 'EI', 'SARL', 'SARLU', 'SA', 'SAS', 'SASU', 'SNC', 'GIE', 'Coopérative']
LIBRE = {'/', '/connexion', '/deconnexion', '/activer', '/renouveler', '/manifest.webmanifest', '/sw.js'}

app = FastAPI(title=NOM, docs_url=None, redoc_url=None, openapi_url=None)
app.mount('/static', StaticFiles(directory=ICI / 'static'), name='static')
tpl = Jinja2Templates(directory=ICI / 'templates')
tpl.env.filters['m'] = lambda v: montant(v or 0)
tpl.env.filters['q'] = lambda v: f'{(v or 0):g}'.replace('.', ',')
tpl.env.filters['d'] = lambda v: v.strftime('%d/%m/%Y') if v else ''
tpl.env.filters['pct'] = lambda v: f'{(v or 0) * 100:g} %'.replace('.', ',')
tpl.env.globals.update(MODES=MODES, NOM=NOM, BOUTIQUE=BOUTIQUE, arrondi=metier.arrondi)


class Redirection(Exception):
    def __init__(self, url):
        self.url = url


@app.exception_handler(Redirection)
async def _redir(request, exc):
    return RedirectResponse(exc.url, 303)


def aller(url, ok='', erreur=''):
    sep = '&' if '?' in url else '?'
    if ok:
        url += f'{sep}ok={quote(ok)}'
    elif erreur:
        url += f'{sep}erreur={quote(erreur)}'
    return RedirectResponse(url, 303)


def contexte(request, gerant=False):
    """Retourne (utilisateur, société) de la session, ou redirige vers la connexion."""
    s = auth.lire(request.cookies.get(auth.COOKIE))
    if not s:
        raise Redirection('/')
    with db.moteur().connect() as cx:
        u = cx.execute(select(db.utilisateurs).where(db.utilisateurs.c.id == s['u'],
                                                     db.utilisateurs.c.societe_id == s['s'])).mappings().first()
        soc = cx.execute(select(db.societes).where(db.societes.c.id == s['s'])).mappings().first()
    if not u or not u['actif'] or not soc:
        raise Redirection('/deconnexion')
    if gerant and u['role'] != 'gerant':
        raise Redirection('/vente?erreur=' + quote('Réservé au gérant.'))
    soc = dict(soc)
    soc['expire'] = soc['expire_le'] < date.today()
    soc['jours_restants'] = (soc['expire_le'] - date.today()).days
    soc['devise_lib'] = PAYS.get(soc['pays'], {}).get('devise_libelle', soc['devise'])
    if soc['expire'] and request.method == 'POST' and request.url.path not in LIBRE:
        raise Redirection('/renouveler?erreur=' + quote('Abonnement expiré : votre espace est en lecture seule.'))
    return dict(u), soc


def _badges(sid):
    with db.moteur().connect() as cx:
        impayes = cx.execute(select(func.count()).where(
            db.documents.c.societe_id == sid, db.documents.c.type == 'facture', db.documents.c.statut == 'valide',
            db.documents.c.total_ttc > db.documents.c.paye)).scalar()
        alertes = cx.execute(select(func.count()).where(
            db.articles.c.societe_id == sid, db.articles.c.actif.is_(True), db.articles.c.type == 'produit',
            db.articles.c.stock <= db.articles.c.seuil_alerte)).scalar()
    return {'impayes': impayes, 'alertes': alertes}


def page(request, nom, u=None, soc=None, **kw):
    if soc:
        kw['nav_badges'] = _badges(soc['id'])
    kw.update(u=u, soc=soc, ok=request.query_params.get('ok', ''), erreur=kw.get('erreur') or request.query_params.get('erreur', ''),
              chemin=request.url.path, aujourdhui=date.today())
    return tpl.TemplateResponse(request, nom, kw, status_code=kw.pop('code_http', 200))


def tx():
    return db.moteur().begin()


def _date(v, defaut=None):
    try:
        return date.fromisoformat(v) if v else defaut
    except ValueError:
        return defaut


def _mois(v):
    try:
        datetime.strptime(v, '%Y-%m')
        return v
    except (TypeError, ValueError):
        return date.today().strftime('%Y-%m')


# ------------------------------------------------------------------ PWA
@app.get('/manifest.webmanifest')
def manifest():
    return Response(json.dumps({'name': NOM + " · Account's", 'short_name': 'Gestion TPE', 'start_url': '/tableau',
                                'display': 'standalone', 'background_color': '#F7F5FC', 'theme_color': '#331C77',
                                'icons': [{'src': '/static/icone-192.png', 'sizes': '192x192', 'type': 'image/png'},
                                          {'src': '/static/icone-512.png', 'sizes': '512x512', 'type': 'image/png'}]}),
                    media_type='application/manifest+json')


@app.get('/sw.js')
def sw():
    return Response((ICI / 'static' / 'sw.js').read_text(), media_type='application/javascript')


# ------------------------------------------------------------------ accès
@app.get('/', response_class=HTMLResponse)
def accueil(request: Request):
    if auth.lire(request.cookies.get(auth.COOKIE)):
        return RedirectResponse('/tableau', 303)
    return page(request, 'accueil.html', demo=licences.demo())


def _ouvrir_session(u):
    r = RedirectResponse('/tableau' if u['role'] == 'gerant' else '/vente', 303)
    r.set_cookie(auth.COOKIE, auth.jeton(u['id'], u['societe_id']), httponly=True, samesite='lax',
                 secure=os.environ.get('HTTPS', '1') == '1', max_age=auth.DUREE)
    return r


@app.post('/connexion')
async def connexion(request: Request):
    f = await request.form()
    ident = (f.get('identifiant') or '').strip().lower()
    with tx() as cx:
        u = cx.execute(select(db.utilisateurs).where(db.utilisateurs.c.identifiant == ident)).mappings().first()
        maintenant = datetime.now()
        if u and u['bloque_jusqua'] and u['bloque_jusqua'] > maintenant:
            return page(request, 'accueil.html', demo=licences.demo(), code_http=429,
                        erreur='Trop d’essais. Réessayez dans 15 minutes.')
        if not u or not u['actif'] or not auth.verifier(f.get('secret') or '', u['secret']):
            if u:
                echecs = (u['echecs'] or 0) + 1
                cx.execute(update(db.utilisateurs).where(db.utilisateurs.c.id == u['id']).values(
                    echecs=0 if echecs >= 5 else echecs,
                    bloque_jusqua=maintenant + timedelta(minutes=15) if echecs >= 5 else None))
            return page(request, 'accueil.html', demo=licences.demo(), code_http=401,
                        erreur='Identifiant ou mot de passe incorrect.')
        cx.execute(update(db.utilisateurs).where(db.utilisateurs.c.id == u['id']).values(echecs=0, bloque_jusqua=None))
        metier.tracer(cx, u['societe_id'], u['id'], 'connexion')
    return _ouvrir_session(u)


@app.get('/deconnexion')
def deconnexion():
    r = RedirectResponse('/', 303)
    r.delete_cookie(auth.COOKIE)
    return r


def _pays_liste():
    return sorted(({'code': k, **v} for k, v in PAYS.items()), key=lambda p: p['nom'])


@app.get('/activer', response_class=HTMLResponse)
def activer_form(request: Request):
    return page(request, 'activer.html', pays=_pays_liste(), formes=FORMES, demo=licences.demo(), f={})


@app.post('/activer')
async def activer(request: Request):
    f = dict(await request.form())
    def refus(msg):
        return page(request, 'activer.html', pays=_pays_liste(), formes=FORMES, demo=licences.demo(), f=f,
                    erreur=msg, code_http=400)
    raison = (f.get('raison_sociale') or '').strip()
    pays = f.get('pays')
    ident = (f.get('identifiant') or '').strip().lower()
    mdp = f.get('mot_de_passe') or ''
    if not raison or pays not in PAYS or not (f.get('nom') or '').strip() or len(ident) < 3:
        return refus('Remplissez tous les champs.')
    if len(mdp) < 8:
        return refus('Le mot de passe doit faire au moins 8 caractères.')
    if mdp != f.get('confirmation'):
        return refus('Les deux mots de passe ne sont pas identiques.')
    try:
        info = licences.verifier_neuve(f.get('cle'))
    except licences.ErreurLicence as e:
        return refus(str(e))
    with db.moteur().connect() as cx:
        if cx.execute(select(db.societes.c.id).where(db.societes.c.licence == info['cle'])).first():
            return refus('Cette clé a déjà servi à créer un espace. Connectez-vous avec votre identifiant.')
        if cx.execute(select(db.utilisateurs.c.id).where(db.utilisateurs.c.identifiant == ident)).first():
            return refus('Cet identifiant est déjà pris : choisissez-en un autre.')
    try:
        licences.activer(info['cle'], 'gestion:' + raison.upper()[:150])
    except licences.ErreurLicence as e:
        return refus(str(e))
    p = PAYS[pays]
    try:
        with tx() as cx:
            sid = cx.execute(insert(db.societes).values(
                licence=info['cle'], raison_sociale=raison[:200], pays=pays, devise=p['devise'],
                forme=f.get('forme', '') if f.get('forme') in FORMES else '',
                telephone=(f.get('telephone') or '')[:60], expire_le=licences.echeance(info['expire']))
            ).inserted_primary_key[0]
            uid = cx.execute(insert(db.utilisateurs).values(
                societe_id=sid, nom=f['nom'].strip()[:120], identifiant=ident[:160], secret=auth.hacher(mdp),
                role='gerant')).inserted_primary_key[0]
            metier.tracer(cx, sid, uid, 'espace_cree', f'{raison} ({pays})')
    except IntegrityError:
        return refus('Cette clé ou cet identifiant est déjà utilisé.')
    r = _ouvrir_session({'id': uid, 'societe_id': sid, 'role': 'gerant'})
    r.headers['location'] = '/societe?ok=' + quote('Espace créé. Complétez les informations de votre société.')
    return r


@app.get('/renouveler', response_class=HTMLResponse)
def renouveler_form(request: Request):
    u, soc = contexte(request, gerant=True)
    return page(request, 'renouveler.html', u, soc)


@app.post('/renouveler')
async def renouveler(request: Request):
    u, soc = contexte(request, gerant=True)
    f = await request.form()
    try:
        info = licences.verifier_neuve(f.get('cle'))
        with db.moteur().connect() as cx:
            if cx.execute(select(db.societes.c.id).where(db.societes.c.licence == info['cle'])).first() or \
                    cx.execute(select(db.renouvellements.c.id).where(db.renouvellements.c.licence == info['cle'])).first():
                raise licences.ErreurLicence('Cette clé a déjà été utilisée.')
        licences.activer(info['cle'], f'gestion-renouvellement:{soc["id"]}')
    except licences.ErreurLicence as e:
        return aller('/renouveler', erreur=str(e))
    depart = max(soc['expire_le'], date.today())
    nouvelle = licences.echeance(None, depart)
    with tx() as cx:
        cx.execute(insert(db.renouvellements).values(societe_id=soc['id'], licence=info['cle'],
                                                     ancienne_echeance=soc['expire_le'], nouvelle_echeance=nouvelle))
        cx.execute(update(db.societes).where(db.societes.c.id == soc['id']).values(expire_le=nouvelle))
        metier.tracer(cx, soc['id'], u['id'], 'renouvellement', f'jusqu’au {nouvelle:%d/%m/%Y}')
    return aller('/tableau', ok=f'Abonnement prolongé jusqu’au {nouvelle:%d/%m/%Y}.')


# ------------------------------------------------------------------ société et équipe
@app.get('/societe', response_class=HTMLResponse)
def societe_form(request: Request):
    u, soc = contexte(request, gerant=True)
    return page(request, 'societe.html', u, soc, pays=_pays_liste(), formes=FORMES)


@app.post('/societe')
async def societe(request: Request):
    u, soc = contexte(request, gerant=True)
    f = await request.form()
    vals = {k: (f.get(k) or '').strip()[:300] for k in ('raison_sociale', 'adresse', 'telephone', 'email', 'rccm',
                                                       'niu', 'regime')}
    if not vals['raison_sociale']:
        return aller('/societe', erreur='La raison sociale est obligatoire.')
    vals['forme'] = f.get('forme') if f.get('forme') in FORMES else ''
    vals['mentions'] = (f.get('mentions') or '')[:1500]
    vals['assujetti_tva'] = f.get('assujetti_tva') == '1'
    for k in ('couleur1', 'couleur2'):
        c = f.get(k) or ''
        if len(c) == 7 and c.startswith('#'):
            vals[k] = c
    if f.get('pays') in PAYS:
        vals['pays'], vals['devise'] = f['pays'], PAYS[f['pays']]['devise']
    logo = f.get('logo')
    if logo is not None and getattr(logo, 'filename', ''):
        contenu = await logo.read()
        if len(contenu) > 300_000:
            return aller('/societe', erreur='Logo trop lourd (300 Ko maximum).')
        typ = 'image/png' if contenu[:4] == b'\x89PNG' else 'image/jpeg' if contenu[:2] == b'\xff\xd8' else None
        if not typ:
            return aller('/societe', erreur='Logo : format PNG ou JPEG uniquement.')
        vals['logo'] = f'data:{typ};base64,' + base64.b64encode(contenu).decode()
    if f.get('retirer_logo') == '1':
        vals['logo'] = ''
    with tx() as cx:
        cx.execute(update(db.societes).where(db.societes.c.id == soc['id']).values(**vals))
        metier.tracer(cx, soc['id'], u['id'], 'societe_modifiee')
    return aller('/societe', ok='Informations enregistrées.')


@app.get('/equipe', response_class=HTMLResponse)
def equipe(request: Request):
    u, soc = contexte(request, gerant=True)
    with db.moteur().connect() as cx:
        membres = cx.execute(select(db.utilisateurs).where(db.utilisateurs.c.societe_id == soc['id'])
                             .order_by(db.utilisateurs.c.role.desc(), db.utilisateurs.c.nom)).mappings().all()
    return page(request, 'equipe.html', u, soc, membres=membres, max_employes=int(os.environ.get('MAX_EMPLOYES', 2)))


@app.post('/equipe')
async def equipe_ajout(request: Request):
    u, soc = contexte(request, gerant=True)
    f = await request.form()
    nom, ident, pin = (f.get('nom') or '').strip(), (f.get('identifiant') or '').strip().lower(), f.get('pin') or ''
    if not nom or len(ident) < 3:
        return aller('/equipe', erreur='Nom et identifiant (3 caractères minimum) obligatoires.')
    if not (pin.isdigit() and 4 <= len(pin) <= 8):
        return aller('/equipe', erreur='Le code PIN doit contenir 4 à 8 chiffres.')
    with tx() as cx:
        nb = cx.execute(select(func.count()).where(db.utilisateurs.c.societe_id == soc['id'],
                                                   db.utilisateurs.c.role == 'employe',
                                                   db.utilisateurs.c.actif.is_(True))).scalar()
        if nb >= int(os.environ.get('MAX_EMPLOYES', 2)):
            return aller('/equipe', erreur='Nombre maximum d’employés atteint pour votre licence.')
        if cx.execute(select(db.utilisateurs.c.id).where(db.utilisateurs.c.identifiant == ident)).first():
            return aller('/equipe', erreur='Cet identifiant est déjà pris.')
        cx.execute(insert(db.utilisateurs).values(societe_id=soc['id'], nom=nom[:120], identifiant=ident[:160],
                                                  secret=auth.hacher(pin), role='employe'))
        metier.tracer(cx, soc['id'], u['id'], 'employe_ajoute', nom)
    return aller('/equipe', ok=f'{nom} peut se connecter avec l’identifiant « {ident} » et son code PIN.')


@app.post('/equipe/{mid}/actif')
async def equipe_actif(request: Request, mid: int):
    u, soc = contexte(request, gerant=True)
    with tx() as cx:
        m = cx.execute(select(db.utilisateurs).where(db.utilisateurs.c.id == mid, db.utilisateurs.c.societe_id == soc['id'],
                                                     db.utilisateurs.c.role == 'employe')).mappings().first()
        if not m:
            return aller('/equipe', erreur='Employé introuvable.')
        cx.execute(update(db.utilisateurs).where(db.utilisateurs.c.id == mid).values(actif=not m['actif']))
        metier.tracer(cx, soc['id'], u['id'], 'employe_' + ('desactive' if m['actif'] else 'reactive'), m['nom'])
    return aller('/equipe', ok='Modification enregistrée.')


@app.post('/motdepasse')
async def mot_de_passe(request: Request):
    u, soc = contexte(request)
    f = await request.form()
    with tx() as cx:
        stocke = cx.execute(select(db.utilisateurs.c.secret).where(db.utilisateurs.c.id == u['id'])).scalar()
        if not auth.verifier(f.get('actuel') or '', stocke):
            return aller('/equipe' if u['role'] == 'gerant' else '/vente', erreur='Mot de passe actuel incorrect.')
        nouveau = f.get('nouveau') or ''
        if (u['role'] == 'gerant' and len(nouveau) < 8) or (u['role'] == 'employe' and not (nouveau.isdigit() and 4 <= len(nouveau) <= 8)):
            return aller('/equipe' if u['role'] == 'gerant' else '/vente',
                         erreur='Nouveau secret trop court (8 caractères pour le gérant, PIN de 4 à 8 chiffres).')
        cx.execute(update(db.utilisateurs).where(db.utilisateurs.c.id == u['id']).values(secret=auth.hacher(nouveau)))
        metier.tracer(cx, soc['id'], u['id'], 'secret_modifie')
    return aller('/equipe' if u['role'] == 'gerant' else '/vente', ok='Mot de passe modifié.')


# ------------------------------------------------------------------ clients et articles
@app.get('/clients', response_class=HTMLResponse)
def clients_liste(request: Request):
    u, soc = contexte(request)
    with db.moteur().connect() as cx:
        rows = cx.execute(select(db.clients).where(db.clients.c.societe_id == soc['id'], db.clients.c.actif.is_(True))
                          .order_by(db.clients.c.nom)).mappings().all()
        dus = dict(cx.execute(select(db.documents.c.client_id, func.sum(db.documents.c.total_ttc - db.documents.c.paye))
                              .where(db.documents.c.societe_id == soc['id'], db.documents.c.type == 'facture',
                                     db.documents.c.statut == 'valide').group_by(db.documents.c.client_id)).all())
    return page(request, 'clients.html', u, soc, clients=rows, dus=dus)


@app.post('/clients')
async def clients_enreg(request: Request):
    u, soc = contexte(request)
    f = await request.form()
    vals = {k: (f.get(k) or '').strip()[:300] for k in ('nom', 'telephone', 'email', 'adresse', 'niu')}
    if not vals['nom']:
        return aller('/clients', erreur='Le nom du client est obligatoire.')
    cid = entier(f.get('id'))
    with tx() as cx:
        if cid:
            if u['role'] != 'gerant':
                return aller('/clients', erreur='Réservé au gérant.')
            cx.execute(update(db.clients).where(db.clients.c.id == cid, db.clients.c.societe_id == soc['id']).values(**vals))
        else:
            cid = cx.execute(insert(db.clients).values(societe_id=soc['id'], **vals)).inserted_primary_key[0]
        metier.tracer(cx, soc['id'], u['id'], 'client_enregistre', vals['nom'])
    retour = f.get('retour') or '/clients'
    return aller(retour if retour.startswith('/') else '/clients', ok=f'Client « {vals["nom"]} » enregistré.')


@app.get('/articles', response_class=HTMLResponse)
def articles_liste(request: Request):
    u, soc = contexte(request)
    with db.moteur().connect() as cx:
        rows = cx.execute(select(db.articles).where(db.articles.c.societe_id == soc['id'], db.articles.c.actif.is_(True))
                          .order_by(db.articles.c.nom)).mappings().all()
    return page(request, 'articles.html', u, soc, articles=rows, tva_pays=PAYS[soc['pays']]['tva'])


@app.post('/articles')
async def articles_enreg(request: Request):
    u, soc = contexte(request, gerant=True)
    f = await request.form()
    nom = (f.get('nom') or '').strip()
    if not nom:
        return aller('/articles', erreur='Le nom de l’article est obligatoire.')
    tva = metier._nombre(f.get('tva'), PAYS[soc['pays']]['tva'] * 100) / 100
    if not soc['assujetti_tva']:
        tva = 0
    vals = dict(nom=nom[:200], reference=(f.get('reference') or '')[:60], unite=(f.get('unite') or 'unité')[:30],
                type='service' if f.get('type') == 'service' else 'produit', prix_vente=max(entier(f.get('prix_vente')), 0),
                tva=min(max(tva, 0), 1), seuil_alerte=max(metier._nombre(f.get('seuil_alerte')), 0))
    aid = entier(f.get('id'))
    with tx() as cx:
        if aid:
            cx.execute(update(db.articles).where(db.articles.c.id == aid, db.articles.c.societe_id == soc['id']).values(**vals))
            metier.tracer(cx, soc['id'], u['id'], 'article_modifie', f'{nom} prix {vals["prix_vente"]}')
        else:
            aid = cx.execute(insert(db.articles).values(societe_id=soc['id'], **vals)).inserted_primary_key[0]
            q = metier._nombre(f.get('stock_initial'))
            if q > 0 and vals['type'] == 'produit':
                metier.entree_stock(cx, soc['id'], u['id'], aid, q, metier._nombre(f.get('prix_achat')), 'Stock d’ouverture')
            metier.tracer(cx, soc['id'], u['id'], 'article_cree', nom)
    return aller('/articles', ok=f'Article « {nom} » enregistré.')


@app.post('/articles/{aid}/archiver')
async def articles_archiver(request: Request, aid: int):
    u, soc = contexte(request, gerant=True)
    with tx() as cx:
        cx.execute(update(db.articles).where(db.articles.c.id == aid, db.articles.c.societe_id == soc['id']).values(actif=False))
        metier.tracer(cx, soc['id'], u['id'], 'article_archive', f'#{aid}')
    return aller('/articles', ok='Article retiré du catalogue.')


def _catalogue(cx, sid):
    return [dict(a) for a in cx.execute(select(db.articles).where(db.articles.c.societe_id == sid,
                                                                    db.articles.c.actif.is_(True))
                                         .order_by(db.articles.c.nom)).mappings()]


def _clients(cx, sid):
    return cx.execute(select(db.clients.c.id, db.clients.c.nom).where(db.clients.c.societe_id == sid,
                                                                    db.clients.c.actif.is_(True))
                      .order_by(db.clients.c.nom)).mappings().all()


def _lignes_formulaire(f):
    g = f.getlist
    n = max(len(g('q')), len(g('des')))
    def at(k, i):
        v = g(k)
        return v[i] if i < len(v) else ''
    return [{'article_id': entier(at('art', i)) or None, 'designation': at('des', i), 'quantite': at('q', i),
             'prix': at('pu', i), 'remise': at('rem', i), 'tva': metier._nombre(at('tva', i)) / 100} for i in range(n)]


# ------------------------------------------------------------------ vente au comptoir
@app.get('/vente', response_class=HTMLResponse)
def vente_form(request: Request):
    u, soc = contexte(request)
    with db.moteur().connect() as cx:
        cat = _catalogue(cx, soc['id'])
        q = (select(db.documents).where(db.documents.c.societe_id == soc['id'], db.documents.c.type == 'facture',
                                        db.documents.c.comptoir.is_(True), db.documents.c.date == date.today())
             .order_by(db.documents.c.id.desc()))
        if u['role'] != 'gerant':
            q = q.where(db.documents.c.cree_par == u['id'])
        ventes = cx.execute(q).mappings().all()
    return page(request, 'vente.html', u, soc, catalogue=cat, catalogue_json=json.dumps(
        [{'id': a['id'], 'nom': a['nom'], 'prix': a['prix_vente'], 'tva': a['tva'], 'stock': a['stock'],
          'type': a['type'], 'unite': a['unite']} for a in cat], ensure_ascii=False), ventes=ventes)


@app.post('/vente')
async def vente(request: Request):
    u, soc = contexte(request)
    f = await request.form()
    try:
        with tx() as cx:
            brutes = [{'article_id': entier(a), 'quantite': q} for a, q in zip(f.getlist('art'), f.getlist('q'))]
            lgs = metier.preparer_lignes(cx, soc['id'], brutes, prix_libre=False)
            fid = metier.vente_comptoir(cx, soc['id'], u['id'], lgs, f.get('mode') if f.get('mode') in MODES else 'especes')
            doc, _ = metier.document(cx, soc['id'], fid)
    except ErreurMetier as e:
        return aller('/vente', erreur=str(e))
    return aller('/vente', ok=f'Vente {doc["numero"]} enregistrée : {montant(doc["total_ttc"], soc["devise_lib"])}.')


# ------------------------------------------------------------------ devis, factures, avoirs
TYPES = {'devis': 'Devis', 'facture': 'Factures', 'avoir': 'Avoirs'}


@app.get('/documents', response_class=HTMLResponse)
def documents_liste(request: Request, type: str = 'facture', statut: str = ''):
    u, soc = contexte(request)
    type = type if type in TYPES else 'facture'
    with db.moteur().connect() as cx:
        q = (select(db.documents).where(db.documents.c.societe_id == soc['id'], db.documents.c.type == type)
             .order_by(db.documents.c.date.desc(), db.documents.c.id.desc()).limit(300))
        if statut == 'impaye':
            q = q.where(db.documents.c.statut == 'valide', db.documents.c.total_ttc > db.documents.c.paye)
        elif statut:
            q = q.where(db.documents.c.statut == statut)
        docs = cx.execute(q).mappings().all()
    return page(request, 'documents.html', u, soc, docs=docs, type=type, titre=TYPES[type], statut=statut)


@app.get('/documents/nouveau', response_class=HTMLResponse)
def document_nouveau(request: Request, type: str = 'facture', client: int = 0):
    u, soc = contexte(request, gerant=True)
    return _formulaire(request, u, soc, {'type': type if type in ('devis', 'facture') else 'facture',
                                         'date': date.today(), 'client_id': client,
                                         'echeance': date.today() + timedelta(days=30)}, [])


def _formulaire(request, u, soc, doc, lgs, erreur=''):
    with db.moteur().connect() as cx:
        cat = _catalogue(cx, soc['id'])
        cl = _clients(cx, soc['id'])
    return page(request, 'document_form.html', u, soc, doc=doc, lignes=lgs, clients=cl, erreur=erreur,
                catalogue_json=json.dumps([{'id': a['id'], 'nom': a['nom'], 'prix': a['prix_vente'], 'tva': a['tva']}
                                           for a in cat], ensure_ascii=False),
                lignes_json=json.dumps([{'art': l.get('article_id') or '', 'des': l['designation'], 'q': l['quantite'],
                                         'pu': l['prix'], 'rem': l.get('remise') or 0, 'tva': round(l['tva'] * 100, 4)}
                                        for l in lgs], ensure_ascii=False),
                tva_defaut=(PAYS[soc['pays']]['tva'] * 100) if soc['assujetti_tva'] else 0)


@app.post('/documents')
async def document_enreg(request: Request):
    u, soc = contexte(request, gerant=True)
    f = await request.form()
    type_ = f.get('type') if f.get('type') in ('devis', 'facture') else 'facture'
    doc_id = entier(f.get('id')) or None
    d, ech = _date(f.get('date'), date.today()), _date(f.get('echeance'))
    brutes = _lignes_formulaire(f)
    try:
        with tx() as cx:
            lgs = metier.preparer_lignes(cx, soc['id'], brutes)
            doc_id = metier.enregistrer_document(cx, soc['id'], u['id'], type_, entier(f.get('client_id')) or None, d,
                                                 lgs, note=(f.get('note') or '')[:1000], echeance=ech, doc_id=doc_id)
            if f.get('valider') == '1' and type_ == 'facture':
                metier.valider_facture(cx, soc['id'], u['id'], doc_id)
    except ErreurMetier as e:
        lgs_aff = [{**b, 'quantite': metier._nombre(b['quantite']), 'prix': entier(b['prix']),
                    'remise': metier._nombre(b['remise'])} for b in brutes if b['designation'] or b['article_id']]
        return _formulaire(request, u, soc, {'id': doc_id, 'type': type_, 'date': d, 'echeance': ech,
                                             'client_id': entier(f.get('client_id')), 'note': f.get('note')},
                           lgs_aff, erreur=str(e))
    return aller(f'/documents/{doc_id}', ok='Document enregistré.')


@app.get('/documents/{doc_id}', response_class=HTMLResponse)
def document_voir(request: Request, doc_id: int):
    u, soc = contexte(request)
    try:
        with db.moteur().connect() as cx:
            doc, lgs = metier.document(cx, soc['id'], doc_id)
            reglements = cx.execute(select(db.tresorerie).where(db.tresorerie.c.document_id == doc_id,
                                                                db.tresorerie.c.societe_id == soc['id'])
                                    .order_by(db.tresorerie.c.date)).mappings().all()
            lies = cx.execute(select(db.documents.c.id, db.documents.c.type, db.documents.c.numero).where(
                db.documents.c.societe_id == soc['id'],
                (db.documents.c.origine_id == doc_id) | (db.documents.c.id == (doc['origine_id'] or -1)))).mappings().all()
    except ErreurMetier as e:
        return aller('/documents', erreur=str(e))
    t = metier.totaux(lgs)
    return page(request, 'document.html', u, soc, doc=doc, lignes=lgs, t=t, reglements=reglements, lies=lies,
                calcul=metier.calcul_ligne)


@app.get('/documents/{doc_id}/modifier', response_class=HTMLResponse)
def document_modifier(request: Request, doc_id: int):
    u, soc = contexte(request, gerant=True)
    try:
        with db.moteur().connect() as cx:
            doc, lgs = metier.document(cx, soc['id'], doc_id)
    except ErreurMetier as e:
        return aller('/documents', erreur=str(e))
    if doc['statut'] != 'brouillon':
        return aller(f'/documents/{doc_id}', erreur='Document validé : il ne peut plus être modifié.')
    return _formulaire(request, u, soc, doc, lgs)


@app.get('/documents/{doc_id}/pdf')
def document_pdf_route(request: Request, doc_id: int):
    u, soc = contexte(request)
    with db.moteur().connect() as cx:
        try:
            doc, lgs = metier.document(cx, soc['id'], doc_id)
        except ErreurMetier as e:
            return aller('/documents', erreur=str(e))
        if doc['client_id']:
            c = cx.execute(select(db.clients).where(db.clients.c.id == doc['client_id'])).mappings().first()
            doc.update(client_adresse=c['adresse'], client_telephone=c['telephone'], client_niu=c['niu'])
    contenu = document_pdf(soc, doc, lgs)
    nom = f'{doc["type"].capitalize()}_{doc["numero"] or "brouillon"}.pdf'
    return Response(contenu, media_type='application/pdf', headers={'Content-Disposition': f'inline; filename="{nom}"'})


@app.post('/documents/{doc_id}/{action}')
async def document_action(request: Request, doc_id: int, action: str):
    gerant_seul = action not in ('encaisser',)
    u, soc = contexte(request, gerant=gerant_seul)
    f = await request.form()
    cible, ok = f'/documents/{doc_id}', ''
    try:
        with tx() as cx:
            if action == 'valider':
                ok = f'Facture {metier.valider_facture(cx, soc["id"], u["id"], doc_id)} validée.'
            elif action == 'convertir':
                fid = metier.convertir_devis(cx, soc['id'], u['id'], doc_id)
                cible, ok = f'/documents/{fid}', 'Facture créée en brouillon à partir du devis : vérifiez-la puis validez-la.'
            elif action == 'statut':
                metier.statut_devis(cx, soc['id'], u['id'], doc_id, f.get('statut'))
                ok = 'Statut du devis mis à jour.'
            elif action == 'avoir':
                aid = metier.annuler_par_avoir(cx, soc['id'], u['id'], doc_id, (f.get('motif') or '')[:200])
                cible, ok = f'/documents/{aid}', 'Avoir créé : la facture est annulée et le stock remis.'
            elif action == 'supprimer':
                metier.supprimer_brouillon(cx, soc['id'], u['id'], doc_id)
                cible, ok = '/documents?type=facture', 'Brouillon supprimé.'
            elif action == 'encaisser':
                metier.encaisser(cx, soc['id'], u['id'], doc_id, f.get('montant'), f.get('mode'), _date(f.get('date'), date.today()))
                ok = 'Règlement enregistré.'
            else:
                raise ErreurMetier('Action inconnue.')
    except ErreurMetier as e:
        return aller(cible, erreur=str(e))
    return aller(cible, ok=ok)


# ------------------------------------------------------------------ caisse et trésorerie
@app.get('/caisse', response_class=HTMLResponse)
def caisse(request: Request, jour: str = ''):
    u, soc = contexte(request)
    j = _date(jour, date.today())
    with db.moteur().connect() as cx:
        mvts = metier.livre(cx, soc['id'], None, j, j)
        s = metier.soldes(cx, soc['id'])
        theorique = metier.soldes(cx, soc['id'], j)['especes']
        cl = cx.execute(select(db.clotures, db.utilisateurs.c.nom).select_from(
            db.clotures.outerjoin(db.utilisateurs, db.utilisateurs.c.id == db.clotures.c.utilisateur_id))
            .where(db.clotures.c.societe_id == soc['id']).order_by(db.clotures.c.date.desc()).limit(10)).mappings().all()
        cloture_jour = metier._jour_cloture(cx, soc['id'], j)
    return page(request, 'caisse.html', u, soc, jour=j, mvts=mvts, soldes=s, theorique=theorique, clotures=cl,
                cloture_jour=cloture_jour, cat_r=CATEGORIES_RECETTES, cat_d=CATEGORIES_DEPENSES)


@app.post('/caisse/mouvement')
async def caisse_mouvement(request: Request):
    u, soc = contexte(request)
    f = await request.form()
    sens = f.get('sens')
    if u['role'] != 'gerant' and sens == 'recette' and f.get('categorie') not in ('autres_recettes',):
        return aller('/caisse', erreur='Réservé au gérant.')
    try:
        with tx() as cx:
            metier.mouvement(cx, soc['id'], u['id'], sens, f.get('mode'), f.get('categorie'), (f.get('libelle') or '').strip(),
                             f.get('montant'), _date(f.get('date'), date.today()) if u['role'] == 'gerant' else date.today(),
                             tva=f.get('tva') or 0, tiers=(f.get('tiers') or '').strip())
    except ErreurMetier as e:
        return aller('/caisse', erreur=str(e))
    return aller('/caisse', ok=('Recette' if sens == 'recette' else 'Dépense') + ' enregistrée.')


@app.post('/caisse/cloture')
async def caisse_cloture(request: Request):
    u, soc = contexte(request)
    f = await request.form()
    try:
        with tx() as cx:
            _, ecart = metier.cloturer_caisse(cx, soc['id'], u['id'], f.get('compte'), (f.get('note') or '')[:300])
    except ErreurMetier as e:
        return aller('/caisse', erreur=str(e))
    if ecart:
        return aller('/caisse', erreur=f'Caisse clôturée avec un écart de {montant(ecart, soc["devise_lib"])} '
                                       f'({"excédent" if ecart > 0 else "manquant"}). Le gérant est informé.')
    return aller('/caisse', ok='Caisse clôturée, aucun écart.')


@app.post('/tresorerie/{mid}/annuler')
async def tresorerie_annuler(request: Request, mid: int):
    u, soc = contexte(request, gerant=True)
    try:
        with tx() as cx:
            metier.annuler_mouvement(cx, soc['id'], u['id'], mid)
    except ErreurMetier as e:
        return aller('/livres', erreur=str(e))
    return aller('/livres', ok='Mouvement annulé (visible dans l’historique).')


@app.get('/livres', response_class=HTMLResponse)
def livres(request: Request, mois: str = '', sens: str = 'recette'):
    u, soc = contexte(request, gerant=True)
    mois = _mois(mois)
    debut, fin = metier.periode_mois(mois)
    sens = sens if sens in ('recette', 'depense') else 'recette'
    with db.moteur().connect() as cx:
        mvts = metier.livre(cx, soc['id'], sens, debut, fin)
        tva = metier.tva_periode(cx, soc['id'], debut, fin)
    cats = CATEGORIES_RECETTES if sens == 'recette' else CATEGORIES_DEPENSES
    return page(request, 'livres.html', u, soc, mois=mois, sens=sens, mvts=mvts, cats=cats, tva=tva,
                total=sum(m['montant'] for m in mvts))


@app.get('/livres/export')
def livres_export(request: Request, periode: str = ''):
    u, soc = contexte(request, gerant=True)
    try:
        debut, fin = exports.mois_ou_annee(periode or date.today().strftime('%Y-%m'))
    except ValueError:
        return aller('/livres', erreur='Période invalide.')
    with db.moteur().connect() as cx:
        contenu = exports.livres(cx, soc, debut, fin)
    nom = f'Livres_{soc["raison_sociale"][:30].replace(" ", "_")}_{periode}.xlsx'
    return Response(contenu, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f'attachment; filename="{quote(nom)}"'})


@app.get('/export/tout')
def export_tout(request: Request):
    u, soc = contexte(request, gerant=True)
    with db.moteur().connect() as cx:
        contenu = exports.toutes_donnees(cx, soc)
    with tx() as cx:
        metier.tracer(cx, soc['id'], u['id'], 'export_complet')
    return Response(contenu, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f'attachment; filename="Donnees_{date.today():%Y%m%d}.xlsx"'})


# ------------------------------------------------------------------ stocks
@app.get('/stocks', response_class=HTMLResponse)
def stocks(request: Request):
    u, soc = contexte(request)
    with db.moteur().connect() as cx:
        cat = [a for a in _catalogue(cx, soc['id']) if a['type'] == 'produit']
        mvts = cx.execute(select(db.stock_mvts, db.articles.c.nom).select_from(
            db.stock_mvts.join(db.articles, db.articles.c.id == db.stock_mvts.c.article_id))
            .where(db.stock_mvts.c.societe_id == soc['id']).order_by(db.stock_mvts.c.id.desc()).limit(40)).mappings().all()
    return page(request, 'stocks.html', u, soc, articles=cat, mvts=mvts,
                valeur=round(sum(max(a['stock'], 0) * a['cump'] for a in cat)))


@app.post('/stocks/entree')
async def stocks_entree(request: Request):
    u, soc = contexte(request, gerant=True)
    f = await request.form()
    try:
        with tx() as cx:
            metier.entree_stock(cx, soc['id'], u['id'], entier(f.get('article_id')), f.get('quantite'), f.get('cout'),
                                (f.get('fournisseur') or '').strip(), f.get('paiement') or None)
    except ErreurMetier as e:
        return aller('/stocks', erreur=str(e))
    return aller('/stocks', ok='Entrée en stock enregistrée.')


@app.get('/stocks/inventaire', response_class=HTMLResponse)
def inventaire_form(request: Request):
    u, soc = contexte(request, gerant=True)
    with db.moteur().connect() as cx:
        cat = [a for a in _catalogue(cx, soc['id']) if a['type'] == 'produit']
    return page(request, 'inventaire.html', u, soc, articles=cat, ecarts=None)


@app.post('/stocks/inventaire', response_class=HTMLResponse)
async def inventaire(request: Request):
    u, soc = contexte(request, gerant=True)
    f = await request.form()
    comptes = {k[2:]: v for k, v in f.items() if k.startswith('a_') and str(v).strip() != ''}
    with tx() as cx:
        ecarts = metier.inventaire(cx, soc['id'], u['id'], comptes)
    with db.moteur().connect() as cx:
        cat = [a for a in _catalogue(cx, soc['id']) if a['type'] == 'produit']
    return page(request, 'inventaire.html', u, soc, articles=cat, ecarts=ecarts)


# ------------------------------------------------------------------ tableau de bord et historique
@app.get('/tableau', response_class=HTMLResponse)
def tableau(request: Request, mois: str = ''):
    u, soc = contexte(request)
    if u['role'] != 'gerant':
        return RedirectResponse('/vente', 303)
    mois = _mois(mois)
    with db.moteur().connect() as cx:
        tb = metier.complements_tableau(cx, soc['id'], metier.tableau_de_bord(cx, soc['id'], mois))
        marges = metier.marges_articles(cx, soc['id'], tb['debut'], tb['fin'])[:10]
        vide = not cx.execute(select(db.articles.c.id).where(db.articles.c.societe_id == soc['id'])).first()
    return page(request, 'tableau.html', u, soc, tb=tb, marges=marges, vide=vide)


@app.get('/historique', response_class=HTMLResponse)
def historique(request: Request):
    u, soc = contexte(request, gerant=True)
    with db.moteur().connect() as cx:
        rows = cx.execute(select(db.journal, db.utilisateurs.c.nom).select_from(
            db.journal.outerjoin(db.utilisateurs, db.utilisateurs.c.id == db.journal.c.utilisateur_id))
            .where(db.journal.c.societe_id == soc['id']).order_by(db.journal.c.id.desc()).limit(300)).mappings().all()
    return page(request, 'historique.html', u, soc, rows=rows)


# ------------------------------------------------------------------ recherche
@app.get('/recherche', response_class=HTMLResponse)
def recherche(request: Request, q: str = ''):
    u, soc = contexte(request)
    q = q.strip()[:80]
    res = {'documents': [], 'clients': [], 'articles': []}
    if len(q) >= 2:
        motif = f'%{q.lower()}%'
        with db.moteur().connect() as cx:
            res['documents'] = cx.execute(select(db.documents).where(
                db.documents.c.societe_id == soc['id'],
                func.lower(func.coalesce(db.documents.c.numero, '')).like(motif) | func.lower(db.documents.c.client_nom).like(motif))
                .order_by(db.documents.c.date.desc()).limit(20)).mappings().all()
            res['clients'] = cx.execute(select(db.clients).where(
                db.clients.c.societe_id == soc['id'], db.clients.c.actif.is_(True),
                func.lower(db.clients.c.nom).like(motif) | db.clients.c.telephone.like(motif)).limit(20)).mappings().all()
            res['articles'] = cx.execute(select(db.articles).where(
                db.articles.c.societe_id == soc['id'], db.articles.c.actif.is_(True),
                func.lower(db.articles.c.nom).like(motif) | func.lower(db.articles.c.reference).like(motif)).limit(20)).mappings().all()
    return page(request, 'recherche.html', u, soc, q=q, res=res)
