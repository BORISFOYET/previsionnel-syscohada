# Account's — Kit Gestion TPE (SYSCOHADA)

Application web de gestion pour les TPE des 17 pays de l'espace OHADA, ouverte avec une clé de licence Chariow.
Produit indépendant du prévisionnel (même dépôt pour l'instant, dossier `gestion-tpe/`, site séparé).

## Fonctions de la V1

| Module | Contenu |
|---|---|
| Accès | Activation de la clé Chariow (verrouillée sur la société), compte gérant (mot de passe), jusqu'à 2 employés (code PIN), blocage 15 min après 5 essais ratés |
| Ma société | Raison sociale, forme, pays (devise et TVA), RCCM, NIU/IFU, régime, logo, 2 couleurs, mentions de bas de facture |
| Clients, articles | Fichier clients avec reste dû, catalogue produits/services, TVA par article, stock d'ouverture |
| Vente au comptoir | Panier tactile : crée la facture validée, sort le stock, encaisse (espèces, Mobile Money, banque). Prix imposés par le catalogue pour les employés |
| Devis, factures, avoirs | Devis numérotés D-AAAA-0001, transformation en facture, factures F-AAAA-0001 numérotées à la validation sans trou, non modifiables ensuite, annulation par avoir AV-AAAA-0001 (stock remis), règlements partiels, PDF à la charte du client, partage WhatsApp |
| Caisse | Dépenses et recettes par rubrique, soldes par mode, clôture du jour avec écart (excédent ou manquant passé en écriture), mouvements figés après clôture |
| Stocks | Entrées d'achat (coût moyen pondéré, dépense enregistrée si payée), sorties automatiques, vente bloquée si stock insuffisant, alerte de seuil, inventaire avec écarts valorisés |
| Livres | Livre des recettes et livre des dépenses (système minimal de trésorerie du SYSCOHADA révisé), récapitulatif par rubrique, TVA collectée / déductible / à reverser (estimation), export Excel du mois ou de l'année |
| Tableau de bord | CA HT, marge, encaissé, dépensé, impayés, TVA, soldes, valeur du stock, écarts de caisse, stock bas, ventes par vendeur, marge par article |
| Contrôle | Historique non modifiable de toutes les actions, export complet des données en Excel |
| Abonnement | Échéance reprise de Chariow (ou 1 an), rappel 30 jours avant, lecture seule après expiration (consultation et export restent possibles), renouvellement par nouvelle clé |

## Licence Chariow

Créer un produit **Licence** : activation requise, 1 activation maximum, validité 12 mois. Créer un second produit
Licence identique pour le **renouvellement**. À la création de l'espace, la clé est activée avec l'identifiant
`gestion:<RAISON SOCIALE>` ; un renouvellement avec `gestion-renouvellement:<n° d'espace>`.

## Mise en ligne (Vercel + Supabase)

1. **Base de données** : créer un projet Supabase (formule gratuite), puis copier l'URL de connexion
   « Transaction pooler » (port 6543). Les tables sont créées automatiquement au premier lancement.
   Sauvegardes : quotidiennes sur la formule Pro de Supabase ; en formule gratuite, faire un export régulier.
2. **Vercel** : nouveau projet à partir du même dépôt GitHub, avec **Root Directory = `gestion-tpe`**.
3. Variables d'environnement :

| Variable | Valeur |
|---|---|
| `DATABASE_URL` | URL Supabase (postgresql://…) |
| `SECRET_KEY` | longue chaîne aléatoire (ne jamais la changer ensuite, sinon tout le monde est déconnecté) |
| `CHARIOW_API_KEY` | clé API Chariow (sans elle et avec `MODE_DEMO=1`, les clés `DEMO-…` sont acceptées) |
| `URL_BOUTIQUE` | lien de la page Chariow du kit |
| `MAX_EMPLOYES` | 2 par défaut |
| `MODE_DEMO` | `1` uniquement pour tester |

## Lancer en local et tester

```bash
pip install -r requirements.txt pytest
MODE_DEMO=1 HTTPS=0 uvicorn app.main:app --reload     # base SQLite gestion.db
python -m pytest -q tests                              # SQLite
TEST_PG=postgresql+psycopg://… python -m pytest -q tests  # PostgreSQL
```

## Limites connues

- Le kit ne remplace pas les systèmes de facturation certifiée là où ils sont obligatoires (e-MECeF au Bénin,
  FNE en Côte d'Ivoire…). La TVA affichée est une estimation à faire valider par le comptable.
- Avoir total uniquement (pas d'avoir partiel) ; pas de mode hors connexion pour les données.
