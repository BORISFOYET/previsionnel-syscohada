"""Mots de passe, PIN et cookie de session signé."""
import base64
import hashlib
import hmac
import json
import os
import secrets
import time

COOKIE = 'tpe_session'
DUREE = 12 * 3600


def _secret():
    return os.environ.get('SECRET_KEY', 'a-changer-en-production').encode()


def hacher(mot):
    sel = secrets.token_bytes(16)
    h = hashlib.scrypt(mot.encode(), salt=sel, n=2 ** 14, r=8, p=1)
    return 'scrypt$' + base64.b64encode(sel).decode() + '$' + base64.b64encode(h).decode()


def verifier(mot, stocke):
    try:
        _, sel, h = stocke.split('$')
        calc = hashlib.scrypt((mot or '').encode(), salt=base64.b64decode(sel), n=2 ** 14, r=8, p=1)
        return hmac.compare_digest(calc, base64.b64decode(h))
    except Exception:
        return False


def jeton(utilisateur_id, societe_id):
    corps = base64.urlsafe_b64encode(json.dumps({'u': utilisateur_id, 's': societe_id,
                                                 'e': int(time.time()) + DUREE}).encode()).decode()
    return corps + '.' + hmac.new(_secret(), corps.encode(), hashlib.sha256).hexdigest()


def lire(j):
    try:
        corps, sig = (j or '').rsplit('.', 1)
        if not hmac.compare_digest(sig, hmac.new(_secret(), corps.encode(), hashlib.sha256).hexdigest()):
            return None
        d = json.loads(base64.urlsafe_b64decode(corps))
        return d if d['e'] > time.time() else None
    except Exception:
        return None


def jeton_csrf(session_jeton):
    return hmac.new(_secret(), ('csrf' + (session_jeton or '')).encode(), hashlib.sha256).hexdigest()[:32]
