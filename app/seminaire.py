"""Séminaires business plan : page de présentation, formulaire de candidature qualifiant, redirection vers Chariow."""
import asyncio
import os
from datetime import datetime, timezone

import httpx
from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from pathlib import Path

router = APIRouter()
tpl = Jinja2Templates(directory=Path(__file__).resolve().parent / 'templates')

# Réglages modifiables sans toucher au code, via les variables d'environnement Vercel.
URL_PAIEMENT = os.environ.get('SEMINAIRE_URL_CHARIOW', '')          # page du produit « Séminaires » sur Chariow
CODE_PROMO = os.environ.get('SEMINAIRE_CODE_PROMO', 'LANCEMENT25')  # code Chariow -25 %, limité à 5 utilisations
WEBHOOK = os.environ.get('SEMINAIRE_WEBHOOK', '')                   # Google Apps Script qui ajoute une ligne à la feuille
WHATSAPP = os.environ.get('SEMINAIRE_WHATSAPP', '')                 # numéro de contact au format 2376XXXXXXXX
PIXEL_META = os.environ.get('PIXEL_META', '1125923300431099')

SEMINAIRE = {
    'debut': 'samedi 5 décembre 2026',
    'places': 20,
    'prix': 69900,
    'prix_lancement': 52425,
    'nb_lancement': 5,
    'lieu': os.environ.get('SEMINAIRE_LIEU', ''),    # ex. « Locaux d'ACCOUNT'S, Douala » ou « En ligne (Zoom) »
    'video': os.environ.get('SEMINAIRE_VIDEO', ''),  # identifiant YouTube de la vidéo de présentation, facultatif
}

CHOIX = {
    'stade': ['Une idée', 'Un projet en préparation', 'Une entreprise de moins de 2 ans', 'Une entreprise de plus de 2 ans'],
    'objectif': ['Obtenir un crédit bancaire', 'Répondre à une subvention ou un appel à projets',
                 'Convaincre un associé ou un investisseur', 'Y voir clair avant de me lancer'],
    'delai': ['Dans les 3 mois', 'Dans 3 à 6 mois', 'Plus tard, je ne sais pas'],
    'presence': ['Oui, aux 4', 'Pas sûr', 'Non'],
    'engagement': ['Oui', 'Cela dépend du prix', 'Non, je cherche du gratuit'],
}


def qualifier(d):
    """Retenu si présent aux 4 séances (ou incertain), prêt à investir et projet décrit sérieusement."""
    motifs = []
    if d.get('presence') == 'Non':
        motifs.append('indisponible')
    if d.get('engagement') == 'Non, je cherche du gratuit':
        motifs.append('cherche du gratuit')
    if len((d.get('projet') or '').strip()) < 15:
        motifs.append('projet trop vague')
    return not motifs, motifs


def nettoyer(brut):
    d = {k: str(brut.get(k, '')).strip()[:600] for k in
         ['nom', 'whatsapp', 'ville', 'projet', 'stade', 'objectif', 'delai', 'presence', 'engagement',
          'utm_source', 'utm_campaign', 'utm_content']}
    manquants = [k for k in ['nom', 'whatsapp', 'ville', 'projet'] if not d[k]]
    manquants += [k for k, opts in CHOIX.items() if d[k] not in opts]
    return d, manquants


async def enregistrer(ligne):
    if not WEBHOOK:
        print('[seminaire] candidature (aucune feuille configurée) :', ligne, flush=True)
        return
    try:
        async with httpx.AsyncClient(timeout=8, follow_redirects=True) as c:
            await c.post(WEBHOOK, json=ligne)
    except httpx.HTTPError as e:
        print('[seminaire] échec envoi feuille :', e, ligne, flush=True)


@router.get('/seminaire', response_class=HTMLResponse)
def page(request: Request):
    return tpl.TemplateResponse(request, 'seminaire.html', {
        's': SEMINAIRE, 'choix': CHOIX, 'pixel': PIXEL_META, 'whatsapp': WHATSAPP})


@router.get('/seminaire/candidature', response_class=HTMLResponse)
def quiz(request: Request):
    return tpl.TemplateResponse(request, 'seminaire_quiz.html', {
        's': SEMINAIRE, 'choix': CHOIX, 'pixel': PIXEL_META, 'whatsapp': WHATSAPP})


@router.post('/seminaire/candidature')
async def candidature(request: Request):
    try:
        brut = await request.json()
    except ValueError:
        return JSONResponse({'erreur': 'Formulaire illisible.'}, 400)
    d, manquants = nettoyer(brut)
    if manquants:
        return JSONResponse({'erreur': 'Merci de répondre à toutes les questions.', 'champs': manquants}, 400)
    retenu, motifs = qualifier(d)
    ligne = {'date': datetime.now(timezone.utc).isoformat(timespec='seconds'), **d,
             'statut': 'Qualifié' if retenu else 'Non retenu', 'motifs': ', '.join(motifs)}
    await asyncio.shield(enregistrer(ligne))
    if not retenu:
        return {'retenu': False}
    return {'retenu': True, 'paiement': URL_PAIEMENT, 'code': CODE_PROMO}
