"""Construit app/data/gabarit_cabinet.xlsx à partir du modèle de plan financier du cabinet.

Chaque tableau imprimable du modèle est copié dans son propre onglet (styles, fusions, largeurs,
hauteurs), sans les données ni les formules de l'ancien projet : seuls les libellés sont gardés.
Onglets produits : TDM, INV, CR, PF, TRE, AMORT, BILAN. Le site les remplit à chaque génération.

Usage : python3 outils/construire_gabarit.py chemin/modele.xlsx app/data/gabarit_cabinet.xlsx
"""
import sys
from copy import copy
from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string as CI, get_column_letter as GL
from openpyxl.worksheet.pagebreak import Break, ColBreak

BLOCS = {  # nom : (onglet du modèle, 1re colonne, dernière colonne, 1re ligne, dernière ligne)
    'TDM': ('0_Table des matières ', 'A', 'M', 1, 20),
    'INV': ('02_Plan financier à imprimer', 'K', 'R', 1, 62),
    'CR': ('02_Plan financier à imprimer', 'AE', 'AN', 1, 62),
    'PF': ('02_Plan financier à imprimer', 'BP', 'BY', 1, 62),
    'TRE': ('02_Plan financier à imprimer', 'CA', 'CS', 1, 62),
    'AMORT': ('#68 Tabl amort 2025', 'A', 'L', 1, 85),
    'BILAN': ('6_Bilan Prévisionnel 010125', 'A', 'K', 1, 53),
}
DECALAGE = {'TDM': 0, 'AMORT': 0, 'BILAN': 0}  # les autres blocs commencent en colonne B


def copier(src, dst, a, b, r0, r1, dec):
    c0, c1 = CI(a), CI(b)
    off = c0 - 1 - dec  # colonne cible = colonne source - off
    for r in range(r0, r1 + 1):
        if src.row_dimensions[r].height:
            dst.row_dimensions[r].height = src.row_dimensions[r].height
        for c in range(c0, c1 + 1):
            s = src.cell(r, c)
            if type(s).__name__ == 'MergedCell':
                continue
            t = dst.cell(r, c - off)
            v = s.value
            if isinstance(v, str) and not v.startswith('='):
                t.value = v
            if s.has_style:
                t.font = copy(s.font); t.fill = copy(s.fill); t.border = copy(s.border)
                t.alignment = copy(s.alignment); t.number_format = s.number_format; t.protection = copy(s.protection)
    # bordures des cellules fusionnées
    for m in src.merged_cells.ranges:
        if m.min_col >= c0 and m.max_col <= c1 and m.min_row >= r0 and m.max_row <= r1:
            dst.merge_cells(start_row=m.min_row, start_column=m.min_col - off, end_row=m.max_row, end_column=m.max_col - off)
            for rr in range(m.min_row, m.max_row + 1):
                for cc in range(m.min_col, m.max_col + 1):
                    s = src.cell(rr, cc); t = dst.cell(rr, cc - off)
                    if s.has_style:
                        t.border = copy(s.border)
    for c in range(c0, c1 + 1):
        k = GL(c)
        if k in src.column_dimensions:
            d = src.column_dimensions[k]
            dst.column_dimensions[GL(c - off)].width = d.width
            dst.column_dimensions[GL(c - off)].hidden = d.hidden
    if dec:
        dst.column_dimensions['A'].width = 2.6
    dst.sheet_view.showGridLines = False
    dst.page_setup.orientation = src.page_setup.orientation or 'portrait'
    dst.page_setup.paperSize = src.page_setup.paperSize
    dst.sheet_properties.pageSetUpPr.fitToPage = True
    dst.page_setup.fitToWidth = 1; dst.page_setup.fitToHeight = 0
    dst.print_options.horizontalCentered = True
    return off


def main(modele, sortie):
    src = load_workbook(modele)
    wb = Workbook(); wb.remove(wb.active)
    for nom, (onglet, a, b, r0, r1) in BLOCS.items():
        ws = wb.create_sheet(nom)
        dec = DECALAGE.get(nom, 1)
        copier(src[onglet], ws, a, b, r0, r1, dec)
    # le budget de trésorerie s'imprime sur deux pages (mois 1 à 5, puis 6 à 12)
    tre = wb['TRE']; tre.page_setup.fitToWidth = 2; tre.page_setup.orientation = 'landscape'
    tre.col_breaks.append(Break(id=CI('CI') - (CI('CA') - 2)))
    wb['CR'].page_setup.orientation = 'portrait'
    wb.save(sortie)
    print('ok', wb.sheetnames)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
