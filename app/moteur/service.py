"""Génère les livrables (Excel en formules + synthèse PDF) à partir des données du formulaire."""
import io
import re
import tempfile
import zipfile
from pathlib import Path

from . import calcul
from .excel import generer as generer_excel
from .cabinet import charger_gabarit, remplir as remplir_cabinet
from .formulaire import construire_config
from .pdf import generer_pdf


def resultats(cfg_calc):
    res = calcul.calcule(cfg_calc)
    sens = [{'nom': 'Scénario de base', **{a: {k: res['exercices'][a][k] for k in ('ebe', 'resultat_net', 'caf', 'couverture_dette')}
                                            for a in res['exercices']}, 'tresorerie_min': res['tresorerie_min']}]
    for s in cfg_calc.get('sensibilites', []):
        r = calcul.calcule(calcul.applique(cfg_calc, s))
        sens.append({'nom': s['nom'], **{a: {k: r['exercices'][a][k] for k in ('ebe', 'resultat_net', 'caf', 'couverture_dette')}
                                         for a in r['exercices']}, 'tresorerie_min': r['tresorerie_min']})
    res['sensibilites'] = sens
    ecarts = [k for k, b in res['bilans'].items() if abs(b['ecart']) > 1]
    if ecarts:
        raise RuntimeError(f'Bilan déséquilibré : {ecarts}')
    return res


def generer_livrables(donnees, theme):
    """Retourne (nom_zip, octets_zip, resume)."""
    cfg_excel, cfg_calc = construire_config(donnees)
    res = resultats(cfg_calc)
    slug = re.sub(r'[^A-Za-z0-9]+', '_', cfg_calc['societe']).strip('_')[:40] or 'societe'
    with tempfile.TemporaryDirectory() as tmp:
        x = Path(tmp) / f'Previsionnel_{slug}.xlsx'
        p = Path(tmp) / f'Synthese_{slug}.pdf'
        generer_excel(cfg_excel, str(x), theme, p=10, apres=remplir_cabinet, classeur=charger_gabarit())
        generer_pdf(cfg_calc, res, str(p), theme)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
            z.write(x, x.name); z.write(p, p.name)
    resume = {a: {k: round(v) for k, v in y.items() if k in ('ca', 'ebe', 'resultat_net', 'caf', 'tresorerie_fin')}
              for a, y in res['exercices'].items()}
    return f'Previsionnel_{slug}.zip', buf.getvalue(), {'exercices': resume, 'tresorerie_min': round(res['tresorerie_min']),
                                                       'mois_tresorerie_min': res['mois_tresorerie_min']}
