// Reçoit les diagnostics, crée ou complète une ligne du Google Sheet et envoie l'audit par email.
const EMAIL_DESTINATAIRE = 'A_REMPLACER@exemple.fr';

const STATUT_DIAGNOSTIC = 'Diagnostic seul';
const STATUT_CONTACT = 'Contact reçu';

// Les nouvelles colonnes s'ajoutent à la fin pour ne pas décaler les lignes existantes
const COLONNES = [
  ['date', 'Date'], ['prenom', 'Prénom'], ['nom', 'Nom'], ['email', 'Email'], ['telephone', 'Téléphone'],
  ['secteur', 'Secteur'], ['demandes', 'Demandes visées / mois'], ['ticket', 'Ticket moyen (€)'], ['closing', 'Closing (%)'],
  ['budgetMin', 'Budget Google min (€/mois)'], ['budgetCentral', 'Budget Google central (€/mois)'], ['budgetMax', 'Budget Google max (€/mois)'],
  ['rdv', 'RDV / mois'], ['signesMin', 'Signés min'], ['signesCentral', 'Signés central'], ['signesMax', 'Signés max'],
  ['caMin', 'CA min (€/mois)'], ['caCentral', 'CA central (€/mois)'], ['caMax', 'CA max (€/mois)'], ['moisCouverts', 'Mois de budget payés par 1 client'],
  ['utmSource', 'Source'], ['utmCampaign', 'Campagne'], ['utmContent', 'Post'],
  ['id', 'ID'], ['statut', 'Statut'], ['dateContact', 'Contact reçu le']
];

function doPost(e) {
  try {
    const lead = JSON.parse(e.postData.contents);
    if (lead.website) return reponse({ ok: true }); // pot de miel anti-robots
    if (!valide(lead)) return reponse({ ok: false, erreur: 'donnees_invalides' });
    ['date', 'statut', 'dateContact'].forEach(k => delete lead[k]);
    const verrou = LockService.getScriptLock();
    verrou.waitLock(20000);
    let premierContact;
    try { premierContact = enregistrer(lead); } finally { verrou.releaseLock(); }
    if (premierContact) envoyerEmail(lead);
    return reponse({ ok: true });
  } catch (err) {
    console.error(err);
    return reponse({ ok: false, erreur: 'erreur_serveur' });
  }
}

function aContact(lead) { return Boolean(lead.prenom || lead.nom || lead.email); }
function idValide(id) { return /^[a-z0-9][a-z0-9-]{7,63}$/i.test(String(id || '')); }

function valide(lead) {
  if (lead.id && !idValide(lead.id)) return false;
  if (!aContact(lead)) return Boolean(lead.id) && texte(lead.secteur);
  return texte(lead.prenom) && texte(lead.nom) && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(lead.email || ''));
}
function texte(v) { return typeof v === 'string' && v.trim().length > 0 && v.length <= 200; }

// Empêche une saisie comme =FORMULE ou +33… d'être interprétée par le Sheet
function cellule(v) {
  if (typeof v === 'number') return v;
  const s = String(v === undefined || v === null ? '' : v).slice(0, 500);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function colonne(cle) { return COLONNES.findIndex(c => c[0] === cle) + 1; }

function preparerEntetes(feuille) {
  if (feuille.getLastRow() > 0 && feuille.getLastColumn() >= COLONNES.length) return;
  feuille.getRange(1, 1, 1, COLONNES.length).setValues([COLONNES.map(c => c[1])]);
}

function trouverLigne(feuille, id) {
  const derniere = feuille.getLastRow();
  if (!id || derniere < 2) return 0;
  const ids = feuille.getRange(2, colonne('id'), derniere - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}

// Renvoie true la première fois que les coordonnées arrivent sur une ligne
function enregistrer(lead) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  preparerEntetes(feuille);
  const contact = aContact(lead);
  const maintenant = new Date().toISOString();
  const ligne = trouverLigne(feuille, lead.id);
  if (!ligne) {
    lead.date = maintenant;
    lead.statut = contact ? STATUT_CONTACT : STATUT_DIAGNOSTIC;
    if (contact) lead.dateContact = maintenant;
    feuille.appendRow(COLONNES.map(c => cellule(lead[c[0]])));
    return contact;
  }
  const dejaContact = feuille.getRange(ligne, colonne('statut')).getValue() === STATUT_CONTACT;
  if (contact && !dejaContact) { lead.statut = STATUT_CONTACT; lead.dateContact = maintenant; }
  // Seules les valeurs reçues sont écrites : le reste de la ligne ne bouge pas
  COLONNES.forEach((c, i) => {
    const v = lead[c[0]];
    if (v !== undefined && v !== null && v !== '') feuille.getRange(ligne, i + 1).setValue(cellule(v));
  });
  return contact && !dejaContact;
}

function envoyerEmail(lead) {
  const corps = COLONNES.filter(c => lead[c[0]] !== undefined && lead[c[0]] !== '').map(c => c[1] + ' : ' + lead[c[0]]).join('\n');
  MailApp.sendEmail({
    to: EMAIL_DESTINATAIRE,
    replyTo: lead.email,
    subject: 'Nouveau diagnostic : ' + lead.prenom + ' ' + lead.nom + ' (' + lead.secteur + ')',
    body: corps
  });
}

function reponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
