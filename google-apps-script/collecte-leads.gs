// Reçoit les diagnostics, ajoute une ligne au Google Sheet et envoie l'audit par email.
const EMAIL_DESTINATAIRE = 'A_REMPLACER@exemple.fr';

const COLONNES = [
  ['date', 'Date'], ['prenom', 'Prénom'], ['nom', 'Nom'], ['email', 'Email'], ['telephone', 'Téléphone'],
  ['secteur', 'Secteur'], ['demandes', 'Demandes visées / mois'], ['ticket', 'Ticket moyen (€)'], ['closing', 'Closing (%)'],
  ['budgetMin', 'Budget Google min (€/mois)'], ['budgetCentral', 'Budget Google central (€/mois)'], ['budgetMax', 'Budget Google max (€/mois)'],
  ['rdv', 'RDV / mois'], ['signesMin', 'Signés min'], ['signesCentral', 'Signés central'], ['signesMax', 'Signés max'],
  ['caMin', 'CA min (€/mois)'], ['caCentral', 'CA central (€/mois)'], ['caMax', 'CA max (€/mois)'], ['moisCouverts', 'Mois de budget payés par 1 client'],
  ['utmSource', 'Source'], ['utmCampaign', 'Campagne'], ['utmContent', 'Post']
];

function doPost(e) {
  try {
    const lead = JSON.parse(e.postData.contents);
    if (lead.website) return reponse({ ok: true }); // pot de miel anti-robots
    if (!valide(lead)) return reponse({ ok: false, erreur: 'donnees_invalides' });
    lead.date = new Date().toISOString();
    enregistrer(lead);
    envoyerEmail(lead);
    return reponse({ ok: true });
  } catch (err) {
    console.error(err);
    return reponse({ ok: false, erreur: 'erreur_serveur' });
  }
}

function valide(lead) {
  return texte(lead.prenom) && texte(lead.nom) && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(lead.email || ''));
}
function texte(v) { return typeof v === 'string' && v.trim().length > 0 && v.length <= 200; }

// Empêche une saisie comme =FORMULE ou +33… d'être interprétée par le Sheet
function cellule(v) {
  if (typeof v === 'number') return v;
  const s = String(v === undefined || v === null ? '' : v).slice(0, 500);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function enregistrer(lead) {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (feuille.getLastRow() === 0) feuille.appendRow(COLONNES.map(c => c[1]));
  feuille.appendRow(COLONNES.map(c => cellule(lead[c[0]])));
}

function envoyerEmail(lead) {
  const corps = COLONNES.map(c => c[1] + ' : ' + (lead[c[0]] === undefined ? '' : lead[c[0]])).join('\n');
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
