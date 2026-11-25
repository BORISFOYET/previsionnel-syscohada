"""PDF des devis, factures et avoirs, aux couleurs et au logo de la société cliente."""
import base64
import io
from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from .metier import PAYS, calcul_ligne, totaux

ICI = Path(__file__).resolve().parent
pdfmetrics.registerFont(TTFont('Carlito', str(ICI / 'data' / 'Carlito-Regular.ttf')))
pdfmetrics.registerFont(TTFont('Carlito-Bold', str(ICI / 'data' / 'Carlito-Bold.ttf')))
TITRES = {'devis': 'DEVIS', 'facture': 'FACTURE', 'avoir': 'AVOIR'}
MARQUE = "Réalisé avec Account's"


def montant(v, devise=''):
    s = f'{int(round(v)):,}'.replace(',', ' ')
    return f'{s} {devise}'.strip()


def qte(v):
    return f'{v:g}'.replace('.', ',')


def _couleur(c, defaut):
    try:
        return HexColor(c or defaut)
    except Exception:
        return HexColor(defaut)


def _wrap(texte, police, taille, largeur):
    mots, lignes, cour = (texte or '').split(), [], ''
    for m in mots:
        essai = (cour + ' ' + m).strip()
        if pdfmetrics.stringWidth(essai, police, taille) <= largeur:
            cour = essai
        else:
            if cour:
                lignes.append(cour)
            cour = m
    if cour:
        lignes.append(cour)
    return lignes or ['']


def document_pdf(soc, doc, lgs):
    devise = PAYS.get(soc['pays'], {}).get('devise_libelle', soc['devise'])
    c1, c2 = _couleur(soc['couleur1'], '#331C77'), _couleur(soc['couleur2'], '#FF2FED')
    buf = io.BytesIO()
    pdf = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    m = 40
    titre = TITRES[doc['type']]
    pdf.setTitle(f'{titre} {doc["numero"] or "brouillon"} - {soc["raison_sociale"]}')

    def pied():
        pdf.setFillColor(c2)
        pdf.rect(0, 0, W, 4, stroke=0, fill=1)
        pdf.setFillColor(HexColor('#6B6585'))
        pdf.setFont('Carlito', 7.5)
        infos = ' · '.join(x for x in [soc['raison_sociale'] + (f' {soc["forme"]}' if soc['forme'] else ''),
                                       soc['adresse'], f'RCCM {soc["rccm"]}' if soc['rccm'] else '',
                                       f'NIU/IFU {soc["niu"]}' if soc['niu'] else '', soc['regime']] if x)
        pdf.drawCentredString(W / 2, 22, infos[:180])
        pdf.setFont('Carlito', 6.5)
        pdf.drawRightString(W - m, 10, MARQUE)

    # en-tête
    pdf.setFillColor(c1)
    pdf.rect(0, H - 6, W, 6, stroke=0, fill=1)
    y = H - 40
    if soc['logo']:
        try:
            img = ImageReader(io.BytesIO(base64.b64decode(soc['logo'].split(',', 1)[1])))
            iw, ih = img.getSize()
            h = 55
            w = min(160, iw * h / ih)
            pdf.drawImage(img, m, y - h + 10, width=w, height=w * ih / iw, mask='auto')
        except Exception:
            pass
    pdf.setFillColor(c1)
    pdf.setFont('Carlito-Bold', 24)
    pdf.drawRightString(W - m, y - 8, titre)
    pdf.setFont('Carlito', 10)
    pdf.setFillColor(HexColor('#1F1640'))
    pdf.drawRightString(W - m, y - 24, f'N° {doc["numero"] or "BROUILLON"}')
    pdf.drawRightString(W - m, y - 37, f'Date : {doc["date"]:%d/%m/%Y}')
    if doc.get('echeance'):
        pdf.drawRightString(W - m, y - 50, ('Valable jusqu’au : ' if doc['type'] == 'devis' else 'Échéance : ')
                            + f'{doc["echeance"]:%d/%m/%Y}')

    # émetteur / client
    y = H - 130
    pdf.setFont('Carlito-Bold', 11)
    pdf.setFillColor(c1)
    pdf.drawString(m, y, soc['raison_sociale'] + (f' {soc["forme"]}' if soc['forme'] else ''))
    pdf.setFont('Carlito', 9)
    pdf.setFillColor(HexColor('#1F1640'))
    yy = y - 13
    for t in [soc['adresse'], soc['telephone'], soc['email'],
              f'RCCM : {soc["rccm"]}' if soc['rccm'] else '', f'NIU / IFU : {soc["niu"]}' if soc['niu'] else '']:
        if t:
            pdf.drawString(m, yy, t[:70])
            yy -= 12
    pdf.setFillColor(HexColor('#F3F0FA'))
    pdf.roundRect(W / 2 + 10, y - 60, W / 2 - m - 10, 75, 6, stroke=0, fill=1)
    pdf.setFillColor(HexColor('#6B6585'))
    pdf.setFont('Carlito', 8)
    pdf.drawString(W / 2 + 22, y, 'CLIENT')
    pdf.setFillColor(HexColor('#1F1640'))
    pdf.setFont('Carlito-Bold', 11)
    pdf.drawString(W / 2 + 22, y - 14, (doc.get('client_nom') or '')[:45])
    pdf.setFont('Carlito', 9)
    yc = y - 27
    for t in [doc.get('client_adresse'), doc.get('client_telephone'),
              f'NIU / IFU : {doc["client_niu"]}' if doc.get('client_niu') else '']:
        if t:
            pdf.drawString(W / 2 + 22, yc, t[:55])
            yc -= 12

    # tableau des lignes
    y = min(yy, y - 70) - 20
    cols = [(m, 'Désignation'), (310, 'Qté'), (360, 'P.U. HT'), (430, 'Remise'), (470, 'TVA'), (W - m, 'Total HT')]

    def entete_tableau(y):
        pdf.setFillColor(c1)
        pdf.rect(m - 4, y - 6, W - 2 * m + 8, 20, stroke=0, fill=1)
        pdf.setFillColor(white)
        pdf.setFont('Carlito-Bold', 9)
        pdf.drawString(cols[0][0] + 2, y, cols[0][1])
        for x, lib in cols[1:]:
            pdf.drawRightString(x if x == W - m else x + 30, y, lib)
        return y - 22

    y = entete_tableau(y)
    pdf.setFont('Carlito', 9)
    for i, l in enumerate(lgs):
        ht, _ = calcul_ligne(l['quantite'], l['prix'], l['remise'], l['tva'])
        des = _wrap(l['designation'], 'Carlito', 9, 255)
        haut = 12 * len(des) + 4
        if y - haut < 170:
            pied()
            pdf.showPage()
            y = entete_tableau(H - 60)
            pdf.setFont('Carlito', 9)
        if i % 2:
            pdf.setFillColor(HexColor('#FAF9FD'))
            pdf.rect(m - 4, y - haut + 10, W - 2 * m + 8, haut, stroke=0, fill=1)
        pdf.setFillColor(HexColor('#1F1640'))
        for k, t in enumerate(des):
            pdf.drawString(m + 2, y - 12 * k, t)
        pdf.drawRightString(340, y, qte(l['quantite']))
        pdf.drawRightString(390, y, montant(l['prix']))
        pdf.drawRightString(460, y, f'{l["remise"]:g} %' if l['remise'] else '')
        pdf.drawRightString(500, y, f'{l["tva"] * 100:g} %'.replace('.', ','))
        pdf.drawRightString(W - m, y, montant(ht))
        y -= haut

    # totaux
    t = totaux(lgs)
    y -= 8
    if y < 190:
        pied()
        pdf.showPage()
        y = H - 80
    pdf.setStrokeColor(HexColor('#E3DEF2'))
    pdf.line(330, y + 6, W - m, y + 6)
    pdf.setFont('Carlito', 10)
    pdf.setFillColor(HexColor('#1F1640'))
    pdf.drawString(340, y - 8, 'Total HT')
    pdf.drawRightString(W - m, y - 8, montant(t['ht'], devise))
    y -= 22
    for taux, (base, tv) in t['par_taux']:
        pdf.drawString(340, y, f'TVA {taux * 100:g} % sur {montant(base)}'.replace('.', ','))
        pdf.drawRightString(W - m, y, montant(tv, devise))
        y -= 14
    pdf.setFillColor(c2)
    pdf.roundRect(330, y - 14, W - m - 330, 24, 5, stroke=0, fill=1)
    pdf.setFillColor(white)
    pdf.setFont('Carlito-Bold', 12)
    pdf.drawString(340, y - 6, 'Total TTC')
    pdf.drawRightString(W - m - 8, y - 6, montant(t['ttc'], devise))
    y -= 32
    if doc['type'] == 'facture' and doc.get('paye'):
        pdf.setFillColor(HexColor('#1F1640'))
        pdf.setFont('Carlito', 10)
        pdf.drawString(340, y, 'Déjà payé')
        pdf.drawRightString(W - m, y, montant(doc['paye'], devise))
        y -= 14
        pdf.setFont('Carlito-Bold', 10)
        pdf.drawString(340, y, 'Reste à payer')
        pdf.drawRightString(W - m, y, montant(doc['total_ttc'] - doc['paye'], devise))
        y -= 14

    # notes et mentions
    pdf.setFillColor(HexColor('#1F1640'))
    pdf.setFont('Carlito', 8.5)
    texte = []
    if doc.get('note'):
        texte += _wrap(doc['note'], 'Carlito', 8.5, W - 2 * m)
    if not soc['assujetti_tva']:
        texte.append('Entreprise non assujettie à la TVA.')
    if soc['mentions']:
        for para in soc['mentions'].splitlines():
            texte += _wrap(para, 'Carlito', 8.5, W - 2 * m)
    yt = max(y - 16, 60 + 11 * len(texte))
    for ligne in texte[:14]:
        pdf.drawString(m, yt, ligne)
        yt -= 11
    pied()
    pdf.save()
    return buf.getvalue()
