// ═══════════════════════════════════════════════════════════════════════════════
// CleanSwift — Google Apps Script v2
// Déployer : Web App > Exécuter en tant que MOI > Accès : Tout le monde
// Permissions requises : Sheets + Calendar + Gmail
// ═══════════════════════════════════════════════════════════════════════════════

const SHEET_ID   = 'REMPLACE_PAR_TON_SHEET_ID'; // ← ID de ton Google Sheet
const SHEET_NAME = 'Clients';
const LOG_SHEET  = 'Historique';
const CALENDAR_ID = 'primary'; // 'primary' = ton calendrier principal Google
// Pour un calendrier dédié CleanSwift : crée-le dans Google Calendar, copie son ID ici
// ex: 'abc123xyz@group.calendar.google.com'

const TIMEZONE = 'America/Edmonton';

const HEADERS = [
  'ID', 'Prénom', 'Nom', 'Téléphone', 'Email', 'Adresse', 'Quartier',
  'Type Service', 'Prix ($CAD)', 'Statut', 'Paiement',
  'Date Inscription', 'Dernier Service', 'Prochain Service', 'Heure Service',
  'Notes', 'Total Facturé ($)', 'Nb Services', 'ID Événement Calendrier', 'Dernière Mise à Jour'
];

// ── POINT D'ENTRÉE HTTP ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { action, client } = payload;
    let result;
    switch (action) {
      case 'CREATE':    result = createClient(client);       break;
      case 'UPDATE':    result = updateClient(client);       break;
      case 'DELETE':    result = deleteClient(client.id);    break;
      case 'SYNC_ALL':  result = syncAll(payload.clients);   break;
      case 'CALENDAR_CREATE': result = creerEvenement(client); break;
      case 'CALENDAR_DELETE': result = supprimerEvenement(client.calendarEventId); break;
      default: throw new Error('Action inconnue: ' + action);
    }
    return buildResponse({ success: true, ...result });
  } catch (err) {
    Logger.log('Erreur doPost: ' + err.message);
    return buildResponse({ success: false, error: err.message });
  }
}

function doGet(e) {
  if (e.parameter.action === 'GET_ALL') {
    return buildResponse({ success: true, clients: getAllClients() });
  }
  return buildResponse({ success: true, message: 'CleanSwift API v2 ✅ — Sheets + Calendar actifs' });
}

// ═══════════════════════════════════════════════════════════════════════════════
// GOOGLE CALENDAR
// ═══════════════════════════════════════════════════════════════════════════════

function creerEvenement(client) {
  if (!client.prochainService) throw new Error('Date de service manquante');

  const dateStr  = client.prochainService;            // '2026-06-20'
  const heureStr = client.heureService || '09:00';    // '09:00' par défaut

  // Construire les dates début / fin (2h par défaut)
  const [hh, mm] = heureStr.split(':').map(Number);
  const debut = new Date(dateStr + 'T00:00:00');
  debut.setHours(hh, mm, 0, 0);
  const fin = new Date(debut);
  fin.setHours(fin.getHours() + 2);

  const titre   = `🧹 CleanSwift — ${client.prenom} ${client.nom}`;
  const details = [
    `📍 Adresse : ${client.adresse || '—'}`,
    `📞 Téléphone : ${client.telephone || '—'}`,
    `💰 Prix : $${client.prixService || '—'} CAD`,
    `🔄 Fréquence : ${translateService(client.typeService)}`,
    client.notes ? `📝 Notes : ${client.notes}` : '',
    '',
    '— CleanSwift CRM'
  ].filter(Boolean).join('\n');

  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) throw new Error('Calendrier introuvable — vérifier CALENDAR_ID');

  // Supprimer l'ancien événement si mise à jour
  if (client.calendarEventId) {
    try {
      const old = cal.getEventById(client.calendarEventId);
      if (old) old.deleteEvent();
    } catch(e) { /* ignoré si déjà supprimé */ }
  }

  const event = cal.createEvent(titre, debut, fin, {
    description: details,
    location: client.adresse || '',
    guests: client.email || '',
    sendInvites: false,
  });

  // Rappel 1 jour avant (notification dans Google Calendar)
  event.addEmailReminder(24 * 60);   // 1440 min = 24h avant
  event.addPopupReminder(60);        // 60 min avant

  // Couleur selon statut
  try {
    event.setColor(CalendarApp.EventColor.CYAN); // Bleu propre pour CleanSwift
  } catch(e) {}

  // Sauvegarder l'ID de l'événement dans le Sheet
  updateCalendarEventId(client.id, event.getId());

  Logger.log(`Événement créé : ${event.getId()} pour ${client.prenom} ${client.nom}`);
  return {
    message: `Événement créé dans Google Calendar`,
    calendarEventId: event.getId(),
    eventTitle: titre,
    eventStart: debut.toISOString(),
    eventEnd: fin.toISOString()
  };
}

function supprimerEvenement(eventId) {
  if (!eventId) return { message: 'Aucun événement à supprimer' };
  try {
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);
    const event = cal.getEventById(eventId);
    if (event) { event.deleteEvent(); return { message: 'Événement supprimé du calendrier' }; }
    return { message: 'Événement déjà supprimé ou introuvable' };
  } catch(e) {
    return { message: 'Erreur suppression: ' + e.message };
  }
}

function updateCalendarEventId(clientId, eventId) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return;
    const rowIdx = findRowById(sheet, clientId);
    if (rowIdx > 0) {
      // Colonne 19 = 'ID Événement Calendrier'
      sheet.getRange(rowIdx, 19).setValue(eventId);
    }
  } catch(e) { Logger.log('updateCalendarEventId: ' + e.message); }
}

// ═══════════════════════════════════════════════════════════════════════════════
// CRUD SHEETS
// ═══════════════════════════════════════════════════════════════════════════════

function createClient(client) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_NAME);
  initHeaders(sheet);

  const existing = findRowById(sheet, client.id);
  if (existing > 0) return updateClient(client);

  const totalFacture = (client.historique || []).reduce((s, h) => s + (h.montant || 0), 0);
  const nbServices   = (client.historique || []).length;

  // Créer l'événement calendrier si date présente
  let calEventId = '';
  if (client.prochainService) {
    try {
      const calResult = creerEvenement(client);
      calEventId = calResult.calendarEventId || '';
    } catch(e) { Logger.log('Calendar non créé: ' + e.message); }
  }

  const row = [
    client.id, client.prenom || '', client.nom || '', client.telephone || '',
    client.email || '', client.adresse || '', client.quartier || '',
    translateService(client.typeService), client.prixService || 0,
    translateStatut(client.statut), translatePaiement(client.paiement),
    client.dateInscription || '', client.dernierService || '',
    client.prochainService || '', client.heureService || '09:00',
    client.notes || '', totalFacture, nbServices, calEventId,
    new Date().toLocaleString('fr-CA', { timeZone: TIMEZONE })
  ];

  sheet.appendRow(row);
  formatLastRow(sheet);
  logAction(ss, 'CRÉATION', client);
  if (client.historique && client.historique.length > 0) appendHistorique(ss, client);

  return {
    message: `Client ${client.prenom} ${client.nom} créé`,
    calendarEventId: calEventId,
    rowCount: sheet.getLastRow() - 1
  };
}

function updateClient(client) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_NAME);
  initHeaders(sheet);

  const rowIdx = findRowById(sheet, client.id);
  if (rowIdx < 0) return createClient(client);

  const totalFacture = (client.historique || []).reduce((s, h) => s + (h.montant || 0), 0);
  const nbServices   = (client.historique || []).length;

  // Recréer l'événement si la date a changé
  let calEventId = client.calendarEventId || sheet.getRange(rowIdx, 19).getValue() || '';
  if (client.prochainService) {
    try {
      client.calendarEventId = calEventId; // passer l'ancien ID pour le supprimer
      const calResult = creerEvenement(client);
      calEventId = calResult.calendarEventId || calEventId;
    } catch(e) { Logger.log('Calendar update: ' + e.message); }
  }

  const values = [
    client.id, client.prenom || '', client.nom || '', client.telephone || '',
    client.email || '', client.adresse || '', client.quartier || '',
    translateService(client.typeService), client.prixService || 0,
    translateStatut(client.statut), translatePaiement(client.paiement),
    client.dateInscription || '', client.dernierService || '',
    client.prochainService || '', client.heureService || '09:00',
    client.notes || '', totalFacture, nbServices, calEventId,
    new Date().toLocaleString('fr-CA', { timeZone: TIMEZONE })
  ];

  sheet.getRange(rowIdx, 1, 1, values.length).setValues([values]);
  formatRow(sheet, rowIdx, client.statut);
  logAction(ss, 'MISE À JOUR', client);

  return {
    message: `Client ${client.prenom} ${client.nom} mis à jour`,
    calendarEventId: calEventId,
    row: rowIdx
  };
}

function deleteClient(id) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_NAME);
  const rowIdx = findRowById(sheet, id);
  if (rowIdx > 0) {
    const nomClient  = sheet.getRange(rowIdx, 2).getValue() + ' ' + sheet.getRange(rowIdx, 3).getValue();
    const calEventId = sheet.getRange(rowIdx, 19).getValue();
    if (calEventId) supprimerEvenement(calEventId);
    sheet.deleteRow(rowIdx);
    logAction(ss, 'SUPPRESSION', { id, nom: nomClient });
    return { message: `Client ${nomClient} supprimé (+ événement calendrier)` };
  }
  return { message: 'Client non trouvé' };
}

function syncAll(clients) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_NAME);
  initHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  for (const client of clients) createClient(client);
  return { message: `${clients.length} clients synchronisés`, timestamp: new Date().toISOString() };
}

function getAllClients() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// DÉCLENCHEUR QUOTIDIEN — Rappels email
// Activer dans : Apps Script > Déclencheurs > Basé sur le temps > Quotidien > 7h–8h
// ═══════════════════════════════════════════════════════════════════════════════
function verifierRelancesQuotidiennes() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const data  = sheet.getDataRange().getValues();
  const today = new Date();
  const alertes = { retard: [], auj: [], demain: [], semaine: [] };

  for (let i = 1; i < data.length; i++) {
    const dateStr = data[i][13]; // Prochain Service
    if (!dateStr) continue;
    const prochainService = new Date(dateStr);
    const diff = Math.ceil((prochainService - today) / 86400000);
    const nom  = `${data[i][1]} ${data[i][2]}`;
    const tel  = data[i][3];
    const svc  = data[i][7];
    const entry = `• ${nom} (${tel}) — ${svc}`;
    if (diff < 0)  alertes.retard.push(entry + ` [${Math.abs(diff)}j de retard]`);
    else if (diff === 0) alertes.auj.push(entry);
    else if (diff === 1) alertes.demain.push(entry);
    else if (diff <= 7)  alertes.semaine.push(entry + ` [dans ${diff}j]`);
  }

  const total = Object.values(alertes).flat().length;
  if (total === 0) return;

  const body = [
    `Bonjour !\n\nVoici votre résumé CleanSwift du ${today.toLocaleDateString('fr-CA')} :\n`,
    alertes.retard.length  ? `🔴 EN RETARD (${alertes.retard.length})\n${alertes.retard.join('\n')}` : '',
    alertes.auj.length     ? `\n🚨 AUJOURD'HUI (${alertes.auj.length})\n${alertes.auj.join('\n')}` : '',
    alertes.demain.length  ? `\n⚠️ DEMAIN (${alertes.demain.length})\n${alertes.demain.join('\n')}` : '',
    alertes.semaine.length ? `\n📅 CETTE SEMAINE (${alertes.semaine.length})\n${alertes.semaine.join('\n')}` : '',
    '\n\nBonne journée !\n— CleanSwift CRM'
  ].filter(Boolean).join('\n');

  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: `🧹 CleanSwift — ${total} service(s) à planifier`,
    body
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// UTILITAIRES
// ═══════════════════════════════════════════════════════════════════════════════
function getOrCreateSheet(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

function initHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    const r = sheet.getRange(1, 1, 1, HEADERS.length);
    r.setFontWeight('bold').setBackground('#0057B8').setFontColor('#fff').setFontSize(11);
    sheet.setFrozenRows(1);
    const widths = [80,100,100,130,180,200,120,140,90,90,110,120,120,130,80,220,90,90,200,160];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
}

function formatLastRow(sheet) {
  const row = sheet.getLastRow();
  const statut = sheet.getRange(row, 10).getValue();
  formatRow(sheet, row, statut);
}

function formatRow(sheet, rowIdx, statut) {
  let bg = '#ffffff';
  if (statut === '🟢 Actif' || statut === 'actif')    bg = '#e8f5e9';
  if (statut === '🟡 Relance' || statut === 'relance') bg = '#fff8e1';
  if (statut === '⚪ Inactif' || statut === 'inactif') bg = '#f5f5f5';
  sheet.getRange(rowIdx, 1, 1, HEADERS.length).setBackground(bg);
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function appendHistorique(ss, client) {
  const sheet = getOrCreateSheet(ss, LOG_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID Client', 'Nom Client', 'Date Service', 'Montant ($)', 'Note', 'Enregistré le']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  }
  for (const h of client.historique) {
    sheet.appendRow([client.id, `${client.prenom} ${client.nom}`, h.date||'', h.montant||0, h.note||'',
      new Date().toLocaleString('fr-CA', { timeZone: TIMEZONE })]);
  }
}

function logAction(ss, action, client) {
  try {
    const sheet = getOrCreateSheet(ss, 'Journal');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Date/Heure', 'Action', 'Client ID', 'Nom', 'Détails']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#34a853').setFontColor('#fff');
    }
    sheet.appendRow([
      new Date().toLocaleString('fr-CA', { timeZone: TIMEZONE }),
      action, client.id||'', `${client.prenom||''} ${client.nom||''}`.trim(),
      JSON.stringify({ statut: client.statut, service: client.typeService, prix: client.prixService })
    ]);
  } catch(e) {}
}

function buildResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function translateService(s) {
  return { weekly:'📅 Hebdomadaire', biweekly:'📅 Bihebdomadaire', monthly:'📅 Mensuel', 'one-time':'🔹 Ponctuel' }[s] || s;
}
function translateStatut(s) {
  return { actif:'🟢 Actif', relance:'🟡 Relance', inactif:'⚪ Inactif' }[s] || s;
}
function translatePaiement(p) {
  return { 'payé':'✅ Payé', 'en attente':'🔵 En attente', 'en retard':'🔴 En retard' }[p] || p;
}
