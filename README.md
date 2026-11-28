# Account's — Prévisionnel financier SYSCOHADA en ligne

Site qui produit, à partir des informations saisies par le client, un prévisionnel financier sur 5 exercices
(le premier part du mois de démarrage) pour une société de l'un des 17 pays de l'espace OHADA :

- un **Excel entièrement en formules**, au format du modèle de plan financier du cabinet : onglet « Plan
  financier à imprimer » (page de garde, investissements et financements initiaux, salaires, détail des
  amortissements, compte de résultat par compte SYSCOHADA, SIG, CAF, seuil de rentabilité, BFR, plan de
  financement, trésorerie de la première année), tableaux des immobilisations, bilans, et dossier de travail
  (hypothèses en cellules bleues, détail mensuel, crédit, amortissements) ;
- le **plan financier en PDF** : les mêmes pages, imprimées à l'identique (formules calculées par pycel,
  pages dessinées avec reportlab ; polices libres Carlito, Caladea et Liberation dans `app/data/polices`).

Les deux documents prennent le logo, les deux couleurs et la police choisis par le client, avec la mention
« Réalisé avec Account's » en pied de page.

## Accès par licence Chariow

1. Sur Chariow, créer un produit de type **Licence** (prix 25 000 FCFA), avec :
   - **activation requise** (ne pas cocher « sans activation ») ;
   - **1 activation maximum** ;
   - durée de validité : 12 mois (conseillé).
2. Dans la page du produit, indiquer au client l'adresse du site.
3. Le client saisit sa clé sur le site. À sa première génération, le site active la licence au nom de sa
   société (l'identifiant d'activation est `societe:<RAISON SOCIALE>`, visible dans le tableau de bord
   Chariow). Ensuite, la clé ne fonctionne que pour cette société, autant de fois que nécessaire.
4. En cas d'erreur de saisie de la raison sociale, révoquer la licence dans Chariow et en émettre une
   nouvelle.

## Mise en ligne (Render)

1. Créer un compte sur render.com et connecter ce dépôt GitHub (« New » → « Blueprint » : `render.yaml`
   est lu automatiquement).
2. Renseigner les variables d'environnement :
   - `CHARIOW_API_KEY` : clé secrète `sk_live_…` (tableau de bord Chariow → Développeurs), droits de lecture
     et d'activation des licences ;
   - `URL_BOUTIQUE` : lien de la page produit Chariow ;
   - `SECRET_KEY` : généré automatiquement.
3. Optionnel : brancher un nom de domaine (ex. `previsionnel.accounts…`).

## Mise en ligne (Vercel, au choix)

1. Créer un compte sur vercel.com, « Add New » → « Project », importer ce dépôt GitHub. Aucune commande de build
   à saisir : Vercel détecte FastAPI dans `app/main.py` (une seule fonction Python, 60 s maximum par génération).
2. Dans « Settings » → « Environment Variables », ajouter `CHARIOW_API_KEY`, `URL_BOUTIQUE`, `MARQUE` et
   `SECRET_KEY` (une longue chaîne aléatoire, à générer soi-même, par exemple avec `openssl rand -hex 32`).
3. Redéployer pour prendre en compte les variables.

Le plan gratuit Hobby fonctionne techniquement, mais les conditions de Vercel le réservent à un usage personnel
et non commercial. Pour un site qui livre un produit payant, prévoir le plan Pro (20 $/mois) dès que les ventes
sont régulières, ou Render (environ 7 $/mois).

## Développement local

```bash
pip install -r requirements.txt
MODE_DEMO=1 HTTPS=0 uvicorn app.main:app --reload   # clés de test : DEMO-xxx
python3 tests/test_generation.py                      # génère un cas pour chacun des 17 pays
```

## Paramètres fiscaux

`app/data/pays.json` : impôt sur les sociétés, minimum de perception (% du CA), TVA (pour information),
charges sociales patronales, avec la source de chaque taux. À mettre à jour à chaque loi de finances.
Limites actuelles : taux de droit commun uniquement ; montants planchers/plafonds du minimum de
perception non modélisés ; prévisionnel hors taxes, sans délais clients/fournisseurs (BFR).

## Séminaires business plan (`/seminaire`)

Page de présentation des 4 séminaires (`/seminaire`), puis candidature en mode quiz, une question par écran
(`/seminaire/candidature`), avec 6 questions et les coordonnées en dernier. Le prix n'apparaît
qu'après le formulaire, et seulement aux candidats retenus (présents aux 4 séances, prêts à investir, projet
décrit). Ceux-ci sont envoyés vers le produit Chariow ; les autres passent en liste d'attente.

Réglages (variables d'environnement Vercel, puis redéployer) :

- `SEMINAIRE_URL_CHARIOW` : lien de la page du produit « Séminaires business plan » sur Chariow ;
- `SEMINAIRE_CODE_PROMO` : code de réduction Chariow de 25 %, limité à 5 utilisations (par défaut `LANCEMENT25`) ;
- `SEMINAIRE_WEBHOOK` : adresse de la feuille Google qui enregistre les candidatures (voir `outils/feuille_candidatures.gs`) ;
- `SEMINAIRE_LIEU` : lieu affiché sur la page (ex. « Locaux d'ACCOUNT'S, Douala » ou « En ligne ») ;
- `SEMINAIRE_VIDEO` : identifiant YouTube de la vidéo de présentation (ex. `dQw4w9WgXcQ`), facultatif ;
- `SEMINAIRE_WHATSAPP` : numéro de contact, chiffres seulement (ex. `2376XXXXXXXX`), facultatif.

Côté Chariow : produit de type Service à 69 900 FCFA, quantité 20, et dans les réglages du produit le lien du
groupe WhatsApp en URL de redirection après achat et dans les instructions après achat.
Le pixel Meta suit `PageView`, `ViewContent`, `Lead` (candidat retenu) et `InitiateCheckout` (clic sur « Je réserve »).
