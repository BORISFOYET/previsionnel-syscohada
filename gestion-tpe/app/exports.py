"""Exports Excel : livres de recettes et de dépenses (SMT), TVA, stock, et export complet des données."""
import io
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from sqlalchemy import select

from . import db
from .metier import CATEGORIES_DEPENSES, CATEGORIES_RECETTES, MODES, livre, periode_mois, tva_periode

FMT = '#,##0;[Red]-#,##0'


def _entete(ws, titres, couleur):
    ws.append(titres)
    for c in ws[ws.max_row]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=couleur.lstrip('#'))
        c.alignment = Alignment(vertical='center', wrap_text=True)
    ws.freeze_panes = ws.cell(row=ws.max_row + 1, column=1)


def _largeurs(ws, largeurs):
    for i, l in enumerate(largeurs, 1):
        ws.column_dimensions[get_column_letter(i)].width = l


def _titre(ws, soc, texte):
    ws.append([f'{soc["raison_sociale"]} — {texte}'])
    ws['A1'].font = Font(bold=True, size=13, color=soc['couleur1'].lstrip('#'))
    ws.append([f'Devise : {soc["devise"]} · Montants TTC · Édité le {date.today():%d/%m/%Y}'])
    ws.append([])


def livres(cx, soc, debut, fin):
    sid, c1 = soc['id'], soc['couleur1']
    wb = Workbook()
    for idx, (sens, nom, cats) in enumerate([('recette', 'Livre des recettes', CATEGORIES_RECETTES),
                                             ('depense', 'Livre des dépenses', CATEGORIES_DEPENSES)]):
        ws = wb.active if idx == 0 else wb.create_sheet()
        ws.title = nom
        _titre(ws, soc, f'{nom} du {debut:%d/%m/%Y} au {fin:%d/%m/%Y}')
        _entete(ws, ['Date', 'N° pièce', 'Libellé', 'Tiers', 'Rubrique', 'Mode'] + [MODES[m] for m in MODES]
                + ['Total', 'dont TVA', 'Saisi par'], c1)
        premier = ws.max_row + 1
        for m in livre(cx, sid, sens, debut, fin):
            ws.append([m['date'], m['numero'], m['libelle'], m['tiers'], cats.get(m['categorie'], m['categorie']),
                       MODES[m['mode']]] + [m['montant'] if m['mode'] == k else None for k in MODES]
                      + [m['montant'], m['tva'] or None, m['par'] or ''])
            ws.cell(row=ws.max_row, column=1).number_format = 'DD/MM/YYYY'
        dernier = ws.max_row
        ws.append(['TOTAL'] + [None] * 5 + [f'=SUM({get_column_letter(c)}{premier}:{get_column_letter(c)}{dernier})'
                                            for c in range(7, 12)])
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
        for row in ws.iter_rows(min_row=premier, min_col=7, max_col=11):
            for c in row:
                c.number_format = FMT
        _largeurs(ws, [11, 9, 38, 22, 30, 14, 13, 13, 13, 13, 11, 16])

    # récapitulatif par rubrique (présentation du système minimal de trésorerie)
    ws = wb.create_sheet('Récapitulatif')
    _titre(ws, soc, f'Récapitulatif des encaissements et décaissements du {debut:%d/%m/%Y} au {fin:%d/%m/%Y}')
    _entete(ws, ['Rubrique', 'Montant'], c1)
    for sens, cats, lib in [('recette', CATEGORIES_RECETTES, 'Recettes'), ('depense', CATEGORIES_DEPENSES, 'Dépenses')]:
        tot = {}
        for m in livre(cx, sid, sens, debut, fin):
            tot[m['categorie']] = tot.get(m['categorie'], 0) + m['montant']
        ws.append([lib.upper()])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
        for k, l in cats.items():
            if tot.get(k):
                ws.append([l, tot[k]])
        ws.append([f'Total {lib.lower()}', sum(tot.values())])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
        ws.append([])
    rec = sum(m['montant'] for m in livre(cx, sid, 'recette', debut, fin))
    dep = sum(m['montant'] for m in livre(cx, sid, 'depense', debut, fin))
    ws.append(['Excédent (ou déficit) de trésorerie de la période', rec - dep])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    for row in ws.iter_rows(min_col=2, max_col=2):
        for c in row:
            c.number_format = FMT
    _largeurs(ws, [52, 16])

    # TVA
    ws = wb.create_sheet('TVA')
    _titre(ws, soc, f'TVA du {debut:%d/%m/%Y} au {fin:%d/%m/%Y}')
    t = tva_periode(cx, sid, debut, fin)
    _entete(ws, ['Élément', 'Montant'], c1)
    for lib, v in [("Chiffre d'affaires HT facturé (avoirs déduits)", t['ca_ht']), ('TVA collectée', t['collectee']),
                   ('TVA déductible sur dépenses saisies', t['deductible']), ('TVA à reverser (estimation)', t['a_reverser'])]:
        ws.append([lib, v])
        ws.cell(row=ws.max_row, column=2).number_format = FMT
    ws.append([])
    ws.append(['Estimation à faire valider par votre comptable avant la déclaration.'])
    _largeurs(ws, [52, 16])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def toutes_donnees(cx, soc):
    """Export complet : une feuille par table, filtrée sur la société."""
    wb = Workbook()
    wb.remove(wb.active)
    tables = [('Clients', db.clients), ('Articles', db.articles), ('Documents', db.documents),
              ('Lignes', db.lignes), ('Recettes et dépenses', db.tresorerie), ('Clôtures de caisse', db.clotures),
              ('Mouvements de stock', db.stock_mvts), ('Historique', db.journal)]
    for nom, t in tables:
        ws = wb.create_sheet(nom)
        cols = [c.name for c in t.columns if c.name not in ('societe_id',)]
        _entete(ws, cols, soc['couleur1'])
        if t is db.lignes:
            q = select(t).join(db.documents, db.documents.c.id == t.c.document_id).where(db.documents.c.societe_id == soc['id'])
        else:
            q = select(t).where(t.c.societe_id == soc['id'])
        for r in cx.execute(q.order_by(t.c.id)).mappings():
            ws.append([r[c] for c in cols])
        _largeurs(ws, [14] * len(cols))
    ws = wb.create_sheet('Société', 0)
    for k, v in soc.items():
        if k != 'logo':
            ws.append([k, v])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def mois_ou_annee(valeur):
    """'2026-11' -> mois ; '2026' -> année entière."""
    if len(valeur) == 4:
        a = int(valeur)
        return date(a, 1, 1), date(a, 12, 31)
    return periode_mois(valeur)
