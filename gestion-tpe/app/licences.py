"""Licences Chariow du kit de gestion.

Le produit Chariow est de type « Licence », 1 activation, activation requise, validité 12 mois.
- Création de l'espace : la clé est activée avec l'identifiant `gestion:<RAISON SOCIALE>` et
  enregistrée dans la table societes (une clé = une société, contrainte d'unicité).
- Renouvellement : une nouvelle clé du produit est activée avec `gestion-renouvellement:<id société>`
  et prolonge l'échéance d'un an.
Mode démonstration (MODE_DEMO=1 et pas de CHARIOW_API_KEY) : toute clé « DEMO-… » est acceptée.
"""
import os
from datetime import date, datetime

import httpx
from urllib.parse import quote

API = 'https://api.chariow.com/v1'


class ErreurLicence(Exception):
    pass


def _cle_api():
    return os.environ.get('CHARIOW_API_KEY', '')


def demo():
    return os.environ.get('MODE_DEMO') == '1' and not _cle_api()


def _get(chemin):
    r = httpx.get(f'{API}{chemin}', headers={'Authorization': f'Bearer {_cle_api()}'}, timeout=15)
    if r.status_code == 404:
        raise ErreurLicence('Clé de licence inconnue. Vérifiez la clé reçue après votre achat.')
    if r.status_code >= 400:
        raise ErreurLicence('Vérification de la licence impossible pour le moment. Réessayez dans quelques minutes.')
    return r.json()


def _date(v):
    if not v:
        return None
    try:
        return datetime.fromisoformat(str(v).replace('Z', '+00:00')).date()
    except ValueError:
        return None


def verifier_neuve(cle):
    """Vérifie qu'une clé est valide et pas encore activée. Retourne {'cle', 'expire'}."""
    cle = (cle or '').strip()
    if not cle:
        raise ErreurLicence('Saisissez votre clé de licence.')
    if demo():
        if not cle.upper().startswith('DEMO-'):
            raise ErreurLicence('Mode démonstration : utilisez une clé commençant par DEMO-.')
        return {'cle': cle, 'expire': None}
    if not _cle_api():
        raise ErreurLicence('Service non configuré (clé API Chariow manquante).')
    data = _get(f'/licenses/{quote(cle, safe="")}')['data']
    if data.get('revoked_at') or data.get('status') == 'revoked':
        raise ErreurLicence('Cette licence a été révoquée.')
    if data.get('is_expired'):
        raise ErreurLicence('Cette licence a expiré.')
    if (data.get('activations') or {}).get('count', 0) > 0:
        raise ErreurLicence('Cette clé a déjà été utilisée. Connectez-vous avec votre identifiant.')
    return {'cle': cle, 'expire': _date(data.get('expires_at'))}


def activer(cle, identifiant_appareil):
    if demo():
        return
    r = httpx.post(f'{API}/licenses/{quote(cle, safe="")}/activate', headers={'Authorization': f'Bearer {_cle_api()}'},
                   json={'device_identifier': identifiant_appareil[:190]}, timeout=15)
    if r.status_code >= 400:
        msg = (r.json() or {}).get('message', '') if r.headers.get('content-type', '').startswith('application/json') else ''
        if 'does not require activation' in msg:
            raise ErreurLicence("Configuration Chariow : activez l'option « activation requise » sur le produit.")
        raise ErreurLicence('Cette clé a déjà été utilisée.')


def echeance(expire, depuis=None):
    """Échéance retenue : celle de Chariow si elle existe, sinon un an après `depuis`."""
    if expire:
        return expire
    d = depuis or date.today()
    try:
        return d.replace(year=d.year + 1)
    except ValueError:  # 29 février
        return d.replace(year=d.year + 1, day=28)
