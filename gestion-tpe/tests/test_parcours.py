"""Parcours complet du kit, en mode démo, sur une base SQLite neuve."""
import io
import os
import re
import sys
from datetime import date, timedelta
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
os.environ.update(MODE_DEMO='1', HTTPS='0', SECRET_KEY='test')
os.environ.pop('CHARIOW_API_KEY', None)

from fastapi.testclient import TestClient  # noqa: E402
from sqlalchemy import select, update  # noqa: E402

from app import db  # noqa: E402
from app.main import app  # noqa: E402


@pytest.fixture()
def base(tmp_path):
    import sqlalchemy as sa
    url = os.environ.get('TEST_PG')
    if url:
        e = sa.create_engine(url)
        db.meta.drop_all(e); e.dispose()
        db.reinitialiser(url)
    else:
        db.reinitialiser(f'sqlite:///{tmp_path}/t.db')
    yield


def espace(cle='DEMO-1', ident='gerant@test.cm', pays='CM', raison='BOUTIQUE TEST'):
    c = TestClient(app)
    r = c.post('/activer', data={'cle': cle, 'raison_sociale': raison, 'pays': pays, 'forme': 'SARL', 'nom': 'Awa',
                                 'identifiant': ident, 'mot_de_passe': 'motdepasse1', 'confirmation': 'motdepasse1'},
               follow_redirects=False)
    assert r.status_code == 303, r.text
    return c


def article(c, nom, prix, stock=0, cout=0, type_='produit', seuil=0):
    r = c.post('/articles', data={'nom': nom, 'type': type_, 'prix_vente': prix, 'tva': '19.25', 'stock_initial': stock,
                                  'prix_achat': cout, 'seuil_alerte': seuil, 'unite': 'unité'})
    assert r.status_code == 200
    with db.moteur().connect() as cx:
        return cx.execute(select(db.articles.c.id).where(db.articles.c.nom == nom)).scalar()


def test_parcours_complet(base):
    c = espace()
    a1 = article(c, 'Sac de riz 25 kg', 15000, stock=10, cout=11000, seuil=3)
    a2 = article(c, 'Livraison', 2000, type_='service')

    # vente au comptoir en espèces : 2 sacs + 1 livraison
    r = c.post('/vente', data={'art': [a1, a2], 'q': ['2', '1'], 'mode': 'especes'})
    assert 'Vente F-' in r.text
    with db.moteur().connect() as cx:
        f = cx.execute(select(db.documents).where(db.documents.c.type == 'facture')).mappings().first()
        stock = cx.execute(select(db.articles.c.stock).where(db.articles.c.id == a1)).scalar()
    ht = 2 * 15000 + 2000
    assert f['total_ht'] == ht and f['total_ttc'] == ht + round(30000 * .1925) + round(2000 * .1925)
    assert f['paye'] == f['total_ttc'] and f['numero'] == f'F-{date.today().year}-0001'
    assert stock == 8

    # client + devis -> facture -> validation -> règlement partiel
    c.post('/clients', data={'nom': 'Hôtel du Lac', 'telephone': '699000000'})
    with db.moteur().connect() as cx:
        cid = cx.execute(select(db.clients.c.id).where(db.clients.c.nom == 'Hôtel du Lac')).scalar()
    r = c.post('/documents', data={'type': 'devis', 'client_id': cid, 'date': date.today().isoformat(),
                                   'art': [a1], 'des': ['Sac de riz 25 kg'], 'q': ['4'], 'pu': ['14000'],
                                   'rem': ['5'], 'tva': ['19.25']}, follow_redirects=False)
    devis_id = int(r.headers['location'].split('/')[2].split('?')[0])
    r = c.post(f'/documents/{devis_id}/convertir', follow_redirects=False)
    fid = int(r.headers['location'].split('/')[2].split('?')[0])
    assert 'validée' in c.post(f'/documents/{fid}/valider').text
    with db.moteur().connect() as cx:
        fac = cx.execute(select(db.documents).where(db.documents.c.id == fid)).mappings().first()
    assert fac['numero'].endswith('0002') and fac['total_ht'] == round(4 * 14000 * .95)
    c.post(f'/documents/{fid}/encaisser', data={'montant': 20000, 'mode': 'mobile_money'})
    # facture validée non modifiable
    r = c.post('/documents', data={'type': 'facture', 'id': fid, 'client_id': cid, 'date': date.today().isoformat(),
                                   'des': ['X'], 'q': ['1'], 'pu': ['1'], 'tva': ['0']})
    assert 'ne peut plus être modifié' in r.text

    # PDF
    r = c.get(f'/documents/{fid}/pdf')
    assert r.status_code == 200 and r.content[:4] == b'%PDF'

    # dépense + clôture de caisse avec un manquant de 500
    c.post('/caisse/mouvement', data={'sens': 'depense', 'mode': 'especes', 'categorie': 'transport',
                                      'libelle': 'Carburant', 'montant': 3000})
    with db.moteur().connect() as cx:
        from app import metier
        attendu = metier.soldes(cx, 1)['especes']
    assert attendu == f['total_ttc'] - 3000
    r = c.post('/caisse/cloture', data={'compte': attendu - 500})
    assert 'manquant' in r.text
    # caisse du jour figée
    r = c.post('/caisse/mouvement', data={'sens': 'depense', 'mode': 'especes', 'categorie': 'transport', 'montant': 100})
    assert 'clôturée' in r.text

    # avoir : facture annulée, stock remis
    c.post(f'/documents/{fid}/avoir', data={'motif': 'Erreur de quantité'})
    with db.moteur().connect() as cx:
        st = cx.execute(select(db.documents.c.statut).where(db.documents.c.id == fid)).scalar()
        stock = cx.execute(select(db.articles.c.stock).where(db.articles.c.id == a1)).scalar()
    assert st == 'annule' and stock == 8

    # inventaire : 7 comptés au lieu de 8
    r = c.post('/stocks/inventaire', data={f'a_{a1}': '7'})
    assert 'Écart total' in r.text

    # tableau de bord et livres
    r = c.get('/tableau')
    assert r.status_code == 200 and 'Écarts de caisse' in r.text
    r = c.get(f'/livres/export?periode={date.today():%Y-%m}')
    wb = openpyxl.load_workbook(io.BytesIO(r.content))
    assert {'Livre des recettes', 'Livre des dépenses', 'Récapitulatif', 'TVA'} <= set(wb.sheetnames)
    rec = wb['Livre des recettes']
    assert rec.max_row >= 6
    r = c.get('/export/tout')
    assert 'Historique' in openpyxl.load_workbook(io.BytesIO(r.content)).sheetnames
    for p in ['/vente', '/caisse', '/stocks', '/livres', '/clients', '/articles', '/societe', '/equipe',
              '/historique', '/documents?type=devis', '/documents?type=facture', '/documents?type=avoir',
              f'/documents/{devis_id}', '/stocks/inventaire', '/renouveler']:
        assert c.get(p).status_code == 200, p


def test_employe_limite(base):
    c = espace()
    a1 = article(c, 'Savon', 500, stock=20, cout=300)
    c.post('/equipe', data={'nom': 'Moussa', 'identifiant': 'moussa', 'pin': '4321'})
    e = TestClient(app)
    r = e.post('/connexion', data={'identifiant': 'moussa', 'secret': '4321'}, follow_redirects=False)
    assert r.headers['location'] == '/vente'
    # l'employé ne peut pas imposer son prix : le prix du catalogue est appliqué
    e.post('/vente', data={'art': [a1], 'q': ['2'], 'mode': 'especes', 'pu': ['1']})
    with db.moteur().connect() as cx:
        ttc = cx.execute(select(db.documents.c.total_ttc)).scalar()
    assert ttc == 1193
    for p in ['/articles', '/stocks/entree', '/documents', '/equipe']:
        r = e.post(p, data={'nom': 'x'}, follow_redirects=False)
        assert r.status_code == 303 and 'R%C3%A9serv%C3%A9' in r.headers['location'], p
    assert e.get('/livres', follow_redirects=False).status_code == 303


def test_isolement_des_societes(base):
    c1 = espace('DEMO-A', 'a@a.cm')
    article(c1, 'Article A', 1000, stock=5, cout=500)
    c2 = espace('DEMO-B', 'b@b.ci', pays='CI', raison='AUTRE')
    assert 'Article A' not in c2.get('/articles').text
    r = c1.post('/vente', data={'art': [1], 'q': ['1'], 'mode': 'especes'})
    with db.moteur().connect() as cx:
        doc_id = cx.execute(select(db.documents.c.id)).scalar()
    assert 'introuvable' in c2.get(f'/documents/{doc_id}/pdf', follow_redirects=False).headers['location']
    r = c2.get(f'/documents/{doc_id}', follow_redirects=False)
    assert r.status_code == 303 and 'introuvable' in r.headers['location']
    # vente avec un article d'une autre société refusée
    r = c2.post('/vente', data={'art': [1], 'q': ['1'], 'mode': 'especes'})
    assert 'introuvable' in r.text


def test_cle_unique_et_connexion(base):
    espace('DEMO-X', 'x@x.cm')
    c = TestClient(app)
    r = c.post('/activer', data={'cle': 'DEMO-X', 'raison_sociale': 'Y', 'pays': 'SN', 'nom': 'Y',
                                 'identifiant': 'y@y.sn', 'mot_de_passe': 'motdepasse1', 'confirmation': 'motdepasse1'})
    assert r.status_code == 400 and 'déjà servi' in r.text
    for _ in range(5):
        c.post('/connexion', data={'identifiant': 'x@x.cm', 'secret': 'faux'})
    r = c.post('/connexion', data={'identifiant': 'x@x.cm', 'secret': 'motdepasse1'})
    assert r.status_code == 429


def test_expiration_lecture_seule(base):
    c = espace()
    with db.moteur().begin() as cx:
        cx.execute(update(db.societes).values(expire_le=date.today() - timedelta(days=1)))
    assert 'lecture seule' in c.get('/tableau').text
    r = c.post('/clients', data={'nom': 'Nouveau'}, follow_redirects=False)
    assert r.headers['location'].startswith('/renouveler')
    assert c.get('/export/tout').status_code == 200
    r = c.post('/renouveler', data={'cle': 'DEMO-RENOUV'})
    with db.moteur().connect() as cx:
        assert cx.execute(select(db.societes.c.expire_le)).scalar() > date.today()
    r = c.post('/renouveler', data={'cle': 'DEMO-RENOUV'})
    assert 'déjà été utilisée' in r.text


def test_pdf_tous_pays(base):
    from app.metier import PAYS
    for i, code in enumerate(PAYS):
        c = espace(f'DEMO-P{i}', f'p{i}@t.org', pays=code, raison=f'SOC {code}')
        a = article(c, f'Produit {code}', 1234, stock=5, cout=1000)
        c.post('/vente', data={'art': [a], 'q': ['1'], 'mode': 'banque'})
        with db.moteur().connect() as cx:
            fid = cx.execute(select(db.documents.c.id).order_by(db.documents.c.id.desc())).scalar()
        assert c.get(f'/documents/{fid}/pdf').content[:4] == b'%PDF'


def test_stock_insuffisant_bloque(base):
    c = espace()
    a = article(c, 'Ciment', 5500, stock=3, cout=4800)
    r = c.post('/vente', data={'art': [a], 'q': ['5'], 'mode': 'especes'})
    assert 'Stock insuffisant' in r.text
    with db.moteur().connect() as cx:
        assert cx.execute(select(db.documents.c.id)).first() is None
        assert cx.execute(select(db.tresorerie.c.id)).first() is None


def test_arrondi_demi_superieur():
    from app.metier import calcul_ligne
    assert calcul_ligne(1, 25000, 0, 0.1925) == (25000, 4813)
