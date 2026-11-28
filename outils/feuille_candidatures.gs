/**
 * Feuille Google qui reçoit les candidatures aux séminaires business plan.
 * 1. Créer une feuille Google, puis Extensions > Apps Script, coller ce code et enregistrer.
 * 2. Déployer > Nouveau déploiement > type « Application Web », exécuter en tant que « Moi »,
 *    accès « Tout le monde ». Copier l'URL obtenue (https://script.google.com/macros/s/…/exec).
 * 3. Sur Vercel, variable d'environnement SEMINAIRE_WEBHOOK = cette URL, puis redéployer.
 */
const COLONNES = ['date', 'statut', 'motifs', 'nom', 'whatsapp', 'ville', 'projet', 'stade', 'objectif',
  'delai', 'presence', 'engagement', 'utm_source', 'utm_campaign', 'utm_content'];

function doPost(e) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (feuille.getLastRow() === 0) feuille.appendRow(COLONNES);
  const d = JSON.parse(e.postData.contents);
  feuille.appendRow(COLONNES.map(c => d[c] || ''));
  return ContentService.createTextOutput('ok');
}
