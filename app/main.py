"""Site Account's : génération en ligne d'un prévisionnel financier SYSCOHADA."""
import base64
import json
import os
import tempfile
from pathlib import Path

from fastapi import FastAPI, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import licences, seminaire
from .moteur.formulaire import CATEGORIES, COMPTES, FRAIS_DEMARRAGE, PAYS, ErreurSaisie, normaliser_societe
from .moteur.service import generer_livrables

ICI = Path(__file__).resolve().parent
MARQUE = os.environ.get('MARQUE', "Réalisé avec Account's")
BOUTIQUE = os.environ.get('URL_BOUTIQUE', 'https://chariow.com')
POLICES = ['Calibri', 'Arial', 'Century Gothic', 'Verdana', 'Trebuchet MS', 'Georgia', 'Cambria', 'Times New Roman']
COOKIE = 'acc_session'

app = FastAPI(title="Account's — Prévisionnel SYSCOHADA", docs_url=None, redoc_url=None)
app.mount('/static', StaticFiles(directory=ICI / 'static'), name='static')
tpl = Jinja2Templates(directory=ICI / 'templates')
app.include_router(seminaire.router)


def cle_session(request):
    return licences.lire_jeton(request.cookies.get(COOKIE))


@app.get('/', response_class=HTMLResponse)
def accueil(request: Request, erreur: str = ''):
    if cle_session(request):
        return RedirectResponse('/outil', 303)
    return tpl.TemplateResponse(request, 'accueil.html', {'erreur': erreur, 'boutique': BOUTIQUE, 'demo': licences.DEMO})


@app.post('/connexion')
def connexion(request: Request, cle: str = Form('')):
    try:
        licences.verifier(cle)
    except licences.ErreurLicence as e:
        return tpl.TemplateResponse(request, 'accueil.html', {'erreur': str(e), 'boutique': BOUTIQUE, 'demo': licences.DEMO}, status_code=401)
    r = RedirectResponse('/outil', 303)
    r.set_cookie(COOKIE, licences.jeton(cle.strip()), httponly=True, secure=os.environ.get('HTTPS', '1') == '1', samesite='lax', max_age=8 * 3600)
    return r


@app.get('/deconnexion')
def deconnexion():
    r = RedirectResponse('/', 303); r.delete_cookie(COOKIE); return r


@app.get('/outil', response_class=HTMLResponse)
def outil(request: Request):
    cle = cle_session(request)
    if not cle:
        return RedirectResponse('/', 303)
    try:
        info = licences.verifier(cle)
    except licences.ErreurLicence as e:
        return RedirectResponse(f'/?erreur={e}', 303)
    pays = sorted(({'code': k, **v} for k, v in PAYS.items()), key=lambda p: p['nom'])
    return tpl.TemplateResponse(request, 'outil.html', {'pays': pays, 'pays_json': json.dumps(PAYS, ensure_ascii=False),
                                                        'polices': POLICES, 'societe_liee': info['societe_liee'],
                                                        'listes_json': json.dumps({
                                                            'comptes': [[c, l] for c, l in COMPTES],
                                                            'categories': [[k, v[0], v[2]] for k, v in CATEGORIES.items()],
                                                            'frais': [[k, v[1]] for k, v in FRAIS_DEMARRAGE.items()] + [['autre', 'Autres frais de démarrage']],
                                                        }, ensure_ascii=False)})


@app.post('/generer')
async def generer(request: Request, donnees: str = Form(...), logo: UploadFile | None = None,
                  primaire: str = Form('#331C77'), secondaire: str = Form('#FF2FED'), police: str = Form('Calibri')):
    cle = cle_session(request)
    if not cle:
        return JSONResponse({'erreur': 'Session expirée : reconnectez-vous avec votre clé.'}, 401)
    try:
        d = json.loads(donnees)
        info = licences.verifier(cle)
        soc = normaliser_societe(d.get('societe'))
        if not soc:
            raise ErreurSaisie('Indiquez la raison sociale.')
        if info['societe_liee'] and info['societe_liee'] != soc[:180]:
            return JSONResponse({'erreur': "Cette licence est liée à une autre société. Une licence permet de produire le prévisionnel d'une seule société."}, 403)
        theme = {'primaire': primaire if primaire.startswith('#') and len(primaire) == 7 else '#331C77',
                 'secondaire': secondaire if secondaire.startswith('#') and len(secondaire) == 7 else '#FF2FED',
                 'police': police if police in POLICES else 'Calibri', 'marque': MARQUE}
        with tempfile.TemporaryDirectory() as tmp:
            if logo is not None and logo.filename:
                contenu = await logo.read()
                if len(contenu) > 3 * 1024 * 1024:
                    raise ErreurSaisie('Logo trop lourd (3 Mo maximum).')
                ext = '.png' if contenu[:4] == b'\x89PNG' else '.jpg' if contenu[:2] == b'\xff\xd8' else None
                if not ext:
                    raise ErreurSaisie('Logo : formats acceptés PNG ou JPG.')
                chemin = Path(tmp) / f'logo{ext}'; chemin.write_bytes(contenu); theme['logo'] = str(chemin)
            nom, octets, resume = generer_livrables(d, theme)
        if not info['societe_liee']:
            licences.lier(cle, soc)
    except (ErreurSaisie, licences.ErreurLicence) as e:
        return JSONResponse({'erreur': str(e)}, 400)
    except json.JSONDecodeError:
        return JSONResponse({'erreur': 'Données du formulaire illisibles.'}, 400)
    return Response(octets, media_type='application/zip', headers={
        'Content-Disposition': f'attachment; filename="{nom}"',
        'X-Resume': base64.b64encode(json.dumps(resume).encode()).decode()})


@app.get('/sante')
def sante():
    return {'ok': True, 'pays': len(PAYS), 'demo': licences.DEMO}
