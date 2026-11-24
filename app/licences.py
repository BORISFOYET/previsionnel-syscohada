"""Vérification des licences Chariow et liaison d'une licence à une seule société.

Principe : le produit Chariow est de type « Licence » avec 1 activation maximum.
À la première génération, la licence est activée avec comme identifiant d'appareil
la raison sociale normalisée de la société. Ensuite, seules les générations pour
cette même société sont acceptées. Aucune base de données n'est nécessaire :
Chariow garde la trace de l'activation (visible dans le tableau de bord).

Mode démonstration : si CHARIOW_API_KEY n'est pas définie et que MODE_DEMO=1,
toute clé commençant par « DEMO- » est acceptée (pour les tests uniquement).
"""
import base64
import hashlib
import hmac
import json
import os
import time

import httpx
from urllib.parse import quote

API = 'https://api.chariow.com/v1'
CLE_API = os.environ.get('CHARIOW_API_KEY', '')
SECRET = os.environ.get('SECRET_KEY', 'a-changer-en-production').encode()
DEMO = os.environ.get('MODE_DEMO') == '1' and not CLE_API
PREFIXE = 'societe:'


class ErreurLicence(Exception):
    pass


def _get(chemin):
    r = httpx.get(f'{API}{chemin}', headers={'Authorization': f'Bearer {CLE_API}'}, timeout=15)
    if r.status_code == 404:
        raise ErreurLicence('Clé de licence inconnue. Vérifiez la clé reçue après votre achat.')
    if r.status_code >= 400:
        raise ErreurLicence('Vérification de la licence impossible pour le moment. Réessayez dans quelques minutes.')
    return r.json()


def verifier(cle):
    """Retourne {'cle', 'societe_liee' (normalisée ou None), 'expire'} ou lève ErreurLicence."""
    cle = (cle or '').strip()
    if not cle:
        raise ErreurLicence('Saisissez votre clé de licence.')
    if DEMO:
        if not cle.upper().startswith('DEMO-'):
            raise ErreurLicence('Mode démonstration : utilisez une clé commençant par DEMO-.')
        return {'cle': cle, 'societe_liee': None, 'expire': None}
    if not CLE_API:
        raise ErreurLicence('Service non configuré (clé API Chariow manquante).')
    data = _get(f'/licenses/{quote(cle, safe="")}')['data']
    if data.get('revoked_at') or data.get('status') == 'revoked':
        raise ErreurLicence('Cette licence a été révoquée.')
    if data.get('is_expired'):
        raise ErreurLicence('Cette licence a expiré. Renouvelez-la sur la boutique.')
    societe = None
    if (data.get('activations') or {}).get('count', 0) > 0:
        acts = _get(f'/licenses/{quote(cle, safe="")}/activations?per_page=20').get('data', [])
        for a in acts:
            dev = ((a.get('activated_by') or {}).get('device') or '')
            if dev.startswith(PREFIXE):
                societe = dev[len(PREFIXE):]
                break
    return {'cle': cle, 'societe_liee': societe, 'expire': data.get('expires_at')}


def lier(cle, societe_normalisee):
    """Active la licence pour cette société (première génération)."""
    if DEMO:
        return
    r = httpx.post(f'{API}/licenses/{quote(cle, safe="")}/activate', headers={'Authorization': f'Bearer {CLE_API}'},
                   json={'device_identifier': PREFIXE + societe_normalisee[:180]}, timeout=15)
    if r.status_code >= 400:
        msg = (r.json() or {}).get('message', '') if r.headers.get('content-type', '').startswith('application/json') else ''
        if 'does not require activation' in msg:
            raise ErreurLicence("Configuration Chariow : activez l'option « activation requise » sur le produit.")
        raise ErreurLicence('Cette licence est déjà utilisée pour une autre société.')


# ---- jeton de session signé (cookie) ----
def jeton(cle, duree=8 * 3600):
    corps = base64.urlsafe_b64encode(json.dumps({'k': cle, 'e': int(time.time()) + duree}).encode()).decode()
    sig = hmac.new(SECRET, corps.encode(), hashlib.sha256).hexdigest()
    return f'{corps}.{sig}'


def lire_jeton(j):
    try:
        corps, sig = (j or '').rsplit('.', 1)
        if not hmac.compare_digest(sig, hmac.new(SECRET, corps.encode(), hashlib.sha256).hexdigest()):
            return None
        d = json.loads(base64.urlsafe_b64decode(corps))
        return d['k'] if d['e'] > time.time() else None
    except Exception:
        return None
