"""Base de données du kit de gestion.

En production : PostgreSQL (Supabase, Neon…) via la variable DATABASE_URL.
En local et dans les tests : SQLite (fichier gestion.db par défaut).
Toutes les tables métier portent societe_id : chaque requête filtre sur la société connectée.
Les montants sont des entiers (FCFA, GNF, KMF et CDF se gèrent sans centimes).
"""
import os

from sqlalchemy import (Boolean, Column, Date, DateTime, Float, ForeignKey, Index, Integer, MetaData, String, Table,
                        Text, UniqueConstraint, create_engine, func)

meta = MetaData()


def _url():
    url = os.environ.get('DATABASE_URL', 'sqlite:///gestion.db')
    if url.startswith('postgres://'):
        url = 'postgresql://' + url[len('postgres://'):]
    if url.startswith('postgresql://'):
        url = 'postgresql+psycopg://' + url[len('postgresql://'):]
    return url


def _id():
    return Column('id', Integer, primary_key=True)


def _soc():
    return Column('societe_id', Integer, ForeignKey('societes.id', ondelete='CASCADE'), nullable=False, index=True)


societes = Table('societes', meta, _id(),
                 Column('licence', String(120), nullable=False, unique=True),
                 Column('raison_sociale', String(200), nullable=False),
                 Column('pays', String(2), nullable=False),
                 Column('devise', String(10), nullable=False),
                 Column('forme', String(40), default=''),
                 Column('adresse', String(300), default=''),
                 Column('telephone', String(60), default=''),
                 Column('email', String(160), default=''),
                 Column('rccm', String(80), default=''),
                 Column('niu', String(80), default=''),
                 Column('regime', String(80), default=''),
                 Column('assujetti_tva', Boolean, default=True),
                 Column('mentions', Text, default=''),
                 Column('logo', Text, default=''),           # data URI (PNG/JPEG), 300 Ko max
                 Column('couleur1', String(7), default='#331C77'),
                 Column('couleur2', String(7), default='#FF2FED'),
                 Column('cree_le', DateTime, server_default=func.now()),
                 Column('expire_le', Date, nullable=False))

utilisateurs = Table('utilisateurs', meta, _id(), _soc(),
                     Column('nom', String(120), nullable=False),
                     Column('identifiant', String(160), nullable=False, unique=True),
                     Column('secret', String(300), nullable=False),  # mot de passe (gérant) ou PIN (employé), haché
                     Column('role', String(10), nullable=False),     # gerant | employe
                     Column('actif', Boolean, default=True),
                     Column('echecs', Integer, default=0),
                     Column('bloque_jusqua', DateTime))

clients = Table('clients', meta, _id(), _soc(),
                Column('nom', String(200), nullable=False),
                Column('telephone', String(60), default=''),
                Column('email', String(160), default=''),
                Column('adresse', String(300), default=''),
                Column('niu', String(80), default=''),
                Column('actif', Boolean, default=True))

articles = Table('articles', meta, _id(), _soc(),
                 Column('reference', String(60), default=''),
                 Column('nom', String(200), nullable=False),
                 Column('type', String(10), nullable=False, default='produit'),  # produit | service
                 Column('unite', String(30), default='unité'),
                 Column('prix_vente', Integer, nullable=False, default=0),      # HT
                 Column('prix_achat', Integer, nullable=False, default=0),      # HT, prix d'achat de référence
                 Column('tva', Float, nullable=False, default=0.18),
                 Column('stock', Float, nullable=False, default=0),
                 Column('cump', Float, nullable=False, default=0),             # coût unitaire moyen pondéré
                 Column('seuil_alerte', Float, nullable=False, default=0),
                 Column('actif', Boolean, default=True))

documents = Table('documents', meta, _id(), _soc(),
                  Column('type', String(10), nullable=False),     # devis | facture | avoir
                  Column('numero', String(40)),                   # attribué à la validation (factures) ou création (devis)
                  Column('date', Date, nullable=False),
                  Column('echeance', Date),
                  Column('client_id', Integer, ForeignKey('clients.id')),
                  Column('client_nom', String(200), default=''),
                  Column('statut', String(12), nullable=False),   # brouillon | valide | accepte | refuse | converti | annule
                  Column('total_ht', Integer, default=0),
                  Column('total_tva', Integer, default=0),
                  Column('total_ttc', Integer, default=0),
                  Column('paye', Integer, default=0),
                  Column('origine_id', Integer),                  # devis d'origine, ou facture corrigée par un avoir
                  Column('note', Text, default=''),
                  Column('comptoir', Boolean, default=False),
                  Column('cree_par', Integer),
                  Column('cree_le', DateTime, server_default=func.now()),
                  UniqueConstraint('societe_id', 'type', 'numero', name='uq_numero'))

lignes = Table('lignes', meta, _id(),
               Column('document_id', Integer, ForeignKey('documents.id', ondelete='CASCADE'), nullable=False, index=True),
               Column('article_id', Integer, ForeignKey('articles.id')),
               Column('designation', String(300), nullable=False),
               Column('quantite', Float, nullable=False),
               Column('prix', Integer, nullable=False),        # HT unitaire
               Column('remise', Float, default=0),             # en %
               Column('tva', Float, nullable=False),
               Column('ordre', Integer, default=0))

# Livre de recettes et de dépenses (= mouvements de trésorerie, système minimal de trésorerie)
tresorerie = Table('tresorerie', meta, _id(), _soc(),
                   Column('numero', Integer, nullable=False),     # n° de pièce chronologique par société
                   Column('date', Date, nullable=False),
                   Column('sens', String(8), nullable=False),     # recette | depense
                   Column('mode', String(14), nullable=False),    # especes | mobile_money | banque
                   Column('categorie', String(30), nullable=False),
                   Column('libelle', String(300), nullable=False),
                   Column('montant', Integer, nullable=False),    # TTC, toujours positif
                   Column('tva', Integer, default=0),             # TVA comprise (dépenses : TVA déductible)
                   Column('document_id', Integer, ForeignKey('documents.id')),
                   Column('tiers', String(200), default=''),
                   Column('utilisateur_id', Integer),
                   Column('cloture_id', Integer),
                   Column('annule', Boolean, default=False),
                   Column('cree_le', DateTime, server_default=func.now()),
                   UniqueConstraint('societe_id', 'numero', name='uq_piece'))

clotures = Table('clotures', meta, _id(), _soc(),
                 Column('date', Date, nullable=False),
                 Column('utilisateur_id', Integer),
                 Column('solde_theorique', Integer, nullable=False),
                 Column('solde_compte', Integer, nullable=False),
                 Column('ecart', Integer, nullable=False),
                 Column('note', Text, default=''),
                 Column('cree_le', DateTime, server_default=func.now()))

stock_mvts = Table('stock_mvts', meta, _id(), _soc(),
                   Column('article_id', Integer, ForeignKey('articles.id'), nullable=False, index=True),
                   Column('date', Date, nullable=False),
                   Column('type', String(12), nullable=False),    # entree | sortie | inventaire | retour
                   Column('quantite', Float, nullable=False),     # signée : + entrée, - sortie
                   Column('cout_unitaire', Float, default=0),
                   Column('libelle', String(300), default=''),
                   Column('document_id', Integer),
                   Column('utilisateur_id', Integer),
                   Column('cree_le', DateTime, server_default=func.now()))

journal = Table('journal', meta, _id(), _soc(),
                Column('quand', DateTime, server_default=func.now()),
                Column('utilisateur_id', Integer),
                Column('action', String(60), nullable=False),
                Column('detail', String(500), default=''))

renouvellements = Table('renouvellements', meta, _id(), _soc(),
                        Column('licence', String(120), nullable=False, unique=True),
                        Column('ancienne_echeance', Date), Column('nouvelle_echeance', Date),
                        Column('cree_le', DateTime, server_default=func.now()))

Index('ix_doc_soc_type', documents.c.societe_id, documents.c.type, documents.c.date)
Index('ix_treso_soc_date', tresorerie.c.societe_id, tresorerie.c.date)

_moteur = None


def moteur():
    global _moteur
    if _moteur is None:
        url = _url()
        kw = {'pool_pre_ping': True}
        if url.startswith('sqlite'):
            kw = {'connect_args': {'check_same_thread': False}}
        else:
            # fonctions serverless + pooler Supabase (mode transaction) : pas de pool local, pas de requêtes préparées
            from sqlalchemy.pool import NullPool
            kw.update(poolclass=NullPool, connect_args={'prepare_threshold': None})
        _moteur = create_engine(url, **kw)
        meta.create_all(_moteur)
    return _moteur


def reinitialiser(url=None):
    """Tests : repart sur une base neuve."""
    global _moteur
    if url:
        os.environ['DATABASE_URL'] = url
    _moteur = None
    return moteur()
