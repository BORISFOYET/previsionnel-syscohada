"""Génère un prévisionnel pour chaque pays et vérifie l'équilibre des bilans (python et Excel)."""
import json, sys, zipfile, io
import openpyxl
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app.moteur.formulaire import PAYS
from app.moteur.service import generer_livrables

DONNEES = {
 "societe": "AGRO PLUS SARL",
 "pays": "CM",
 "debut": "2027-03",
 "produits": [
  {
   "nom": "Jus de fruits",
   "unite": "bouteille",
   "prix": "1500",
   "volume_depart": "3000",
   "croissance_mensuelle": "4",
   "croissance_annuelle": "12",
   "cout_achat": "600",
   "autres_couts": "150"
  },
  {
   "nom": "Livraison",
   "unite": "course",
   "prix": "2000",
   "volume_depart": "100",
   "croissance_mensuelle": "2",
   "croissance_annuelle": "5",
   "cout_achat": "0",
   "autres_couts": "800"
  }
 ],
 "personnel": [
  {
   "poste": "Gérant",
   "effectif": "1",
   "salaire": "300000",
   "dirigeant": True
  },
  {
   "poste": "Ouvriers",
   "effectif": "4",
   "salaire": "80000"
  }
 ],
 "charges_fixes": [
  {
   "compte": "622",
   "nom": "Loyer",
   "montant": "250000"
  },
  {
   "compte": "605",
   "nom": "Électricité et eau",
   "montant": "120000"
  },
  {
   "compte": "64",
   "nom": "Patente",
   "montant": "20000"
  },
  {
   "compte": "625",
   "nom": "Assurance",
   "montant": "30000"
  }
 ],
 "investissements": [
  {
   "categorie": "materiel_industriel",
   "nom": "Presse et embouteilleuse",
   "montant": "18000000",
   "duree": "7",
   "mois": "1"
  },
  {
   "categorie": "transport",
   "nom": "Véhicule",
   "montant": "9000000",
   "duree": "4",
   "mois": "7"
  },
  {
   "categorie": "terrains",
   "nom": "Terrain",
   "montant": "5000000",
   "mois": "1"
  },
  {
   "categorie": "logiciels",
   "nom": "Site internet",
   "montant": "800000",
   "duree": "3",
   "mois": "1"
  },
  {
   "categorie": "cautions",
   "nom": "Caution loyer",
   "montant": "750000",
   "mois": "1"
  },
  {
   "categorie": "bureau",
   "nom": "Ordinateurs",
   "montant": "1200000",
   "duree": "3",
   "mois": "20"
  }
 ],
 "capital": "10000000",
 "comptes_courants": "5000000",
 "emprunts": [
  {
   "nom": "Crédit équipement",
   "montant": "20000000",
   "taux": "12",
   "duree": "48",
   "differe": "3",
   "mois": "1"
  }
 ],
 "regime_fiscal": "auto",
 "porteur": "Jean Mbarga",
 "frais_demarrage": [
  {
   "type": "creation",
   "montant": "600000",
   "mois": "1"
  },
  {
   "type": "etudes",
   "montant": "1500000",
   "mois": "1"
  }
 ]
}

if __name__ == '__main__':
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    for code in PAYS:
        nom, octets, res = generer_livrables({**DONNEES, 'pays': code}, {'primaire': '#8B4513', 'secondaire': '#DDD4B9', 'police': 'Calibri', 'marque': "Réalisé avec Account's"})
        z = zipfile.ZipFile(io.BytesIO(octets)); assert len(z.namelist()) == 2 and any(n.startswith('Plan_financier') for n in z.namelist())
        xl = openpyxl.load_workbook(io.BytesIO(z.read([n for n in z.namelist() if n.endswith('.xlsx')][0])))
        devise = PAYS[code]['devise']
        for ws in xl.worksheets:
            assert ws.max_row > 1, f'onglet vide : {ws.title}'
            if devise != 'XAF':
                assert not any('XAF' in (c.number_format or '') for r in ws.iter_rows() for c in r), f'XAF dans {ws.title}'
        if out and code == 'CM':
            z.extractall(out)
        print(code, 'ok', res['tresorerie_min'], {a: v['resultat_net'] for a, v in res['exercices'].items()})
