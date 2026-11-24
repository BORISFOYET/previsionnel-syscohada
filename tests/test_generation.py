"""Génère un prévisionnel pour chaque pays et vérifie l'équilibre des bilans (python et Excel)."""
import json, sys, zipfile, io
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app.moteur.formulaire import PAYS
from app.moteur.service import generer_livrables

DONNEES = {
    'societe': 'AGRO PLUS SARL', 'pays': 'CM', 'debut': '2027-01',
    'produits': [{'nom': 'Jus de fruits', 'unite': 'bouteille', 'prix': '1500', 'volume_depart': '3000', 'croissance_mensuelle': '4', 'croissance_annuelle': '12', 'cout_achat': '600', 'autres_couts': '150'},
                 {'nom': 'Livraison', 'unite': 'course', 'prix': '2000', 'volume_depart': '100', 'croissance_mensuelle': '2', 'croissance_annuelle': '5', 'cout_achat': '0', 'autres_couts': '800'}],
    'personnel': [{'poste': 'Gérant', 'effectif': '1', 'salaire': '300000'}, {'poste': 'Ouvriers', 'effectif': '4', 'salaire': '80000'}],
    'charges_fixes': [{'nom': 'Loyer', 'montant': '250000'}, {'nom': 'Électricité et eau', 'montant': '120000'}],
    'investissements': [{'nom': 'Presse et embouteilleuse', 'montant': '18000000', 'duree': '7', 'mois': '1'}, {'nom': 'Véhicule', 'montant': '9000000', 'duree': '4', 'mois': '7'}, {'nom': 'Terrain', 'montant': '5000000', 'duree': '0', 'mois': '1'}],
    'capital': '10000000', 'comptes_courants': '5000000',
    'emprunts': [{'nom': 'Crédit équipement', 'montant': '20000000', 'taux': '12', 'duree': '48', 'differe': '3', 'mois': '1'}],
    'regime_annee1': 'reel', 'montant_forfait': '0'}

if __name__ == '__main__':
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    for code in PAYS:
        nom, octets, res = generer_livrables({**DONNEES, 'pays': code}, {'primaire': '#1F4E79', 'secondaire': '#F2A900', 'police': 'Arial', 'marque': "Réalisé avec Account's"})
        z = zipfile.ZipFile(io.BytesIO(octets)); assert len(z.namelist()) == 2
        if out and code == 'CM':
            z.extractall(out)
        print(code, 'ok', res['tresorerie_min'], {a: v['resultat_net'] for a, v in res['exercices'].items()})
