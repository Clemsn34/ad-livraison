// ============================================================================
// RETOURS DE PIÈCES DÉTACHÉES (garantie ou échange standard)
// À commander → commandée chez le fournisseur → reçue → changée chez le client → ancienne pièce renvoyée → clôturé.
// Saisie par le technicien dans Dicter (mode « Pièces ») ou dans l'Agence (onglet « Retours pièces »). Visible par toute l'équipe AD.
// Chaque étape crée la tâche suivante et coche la précédente : Séverine commande, le technicien change la pièce, Séverine renvoie l'ancienne.
// ============================================================================
var SHEET_PIECES = 'Retours pièces détachées';
var PIECE_HEADERS = ['N°', 'Créé le', 'Créé par', 'Identifiant cabinet', 'Cabinet', 'Appareil', 'N° de série appareil', 'Fournisseur', 'Type de pièce', 'Référence',
  'Désignation', 'Nature de la panne', 'Mode', 'Statut', 'N° commande fournisseur', 'N° de retour / RMA', 'Suivi colis retour', 'Montant (€ HT)', 'Notes', 'Historique',
  'Commandée le', 'Reçue le', 'Changée le', 'Renvoyée le', 'Clôturé le', 'Technicien', 'Tâche en cours', 'Tâches', 'Casier', 'En casier le', 'Tâche d\'origine'];
var PIECE_STATUTS = ['À commander', 'Commandée', 'Reçue', 'En casier', 'Changée chez le client', 'Ancienne pièce renvoyée', 'Clôturé – avoir reçu', 'Clôturé – refusé par le fournisseur'];
var PIECE_FINIS = ['Clôturé – avoir reçu', 'Clôturé – refusé par le fournisseur'];
var PIECE_MODES = ['Garantie', 'Échange standard'];
var PIECE_DATE_COL = { 'Commandée': 20, 'Reçue': 21, 'En casier': 29, 'Changée chez le client': 22, 'Ancienne pièce renvoyée': 23, 'Clôturé – avoir reçu': 24, 'Clôturé – refusé par le fournisseur': 24 };
var PC = { n: 0, cree: 1, par: 2, cabId: 3, cab: 4, app: 5, serie: 6, frs: 7, type: 8, ref: 9, desig: 10, panne: 11, mode: 12, statut: 13, cmd: 14, rma: 15, suivi: 16,
  montant: 17, notes: 18, histo: 19, tech: 25, tache: 26, taches: 27, casier: 28, orig: 30 };
// Étape de la fiche → état « Pièces » affiché sur la tâche d'origine
var PIECE_ETAT_TACHE = { 'À commander': 'À commander', 'Commandée': 'Commandée', 'Reçue': 'Reçue – RDV à prendre', 'En casier': 'En casier' };

function pieceFeuille_() {
  var ss = getSpreadsheet_(), sh = ss.getSheetByName(SHEET_PIECES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PIECES);
    styleSheet_(sh, PIECE_HEADERS, [100, 95, 90, 100, 190, 170, 130, 130, 140, 130, 200, 260, 120, 190, 130, 130, 150, 90, 220, 320, 95, 95, 95, 95, 95, 100, 90, 140]);
    setChoices_(sh, PC.statut + 1, PIECE_STATUTS);
    setChoices_(sh, PC.mode + 1, PIECE_MODES);
  } else if (sh.getLastColumn() < PIECE_HEADERS.length) {
    // Colonnes ajoutées après coup (casier, tâche d'origine)
    var dep = sh.getLastColumn(); sh.getRange(1, dep + 1, 1, PIECE_HEADERS.length - dep).setValues([PIECE_HEADERS.slice(dep)]).setFontWeight('bold');
    setChoices_(sh, PC.statut + 1, PIECE_STATUTS);
  }
  return sh;
}
function pieceObj_(r, i) {
  var d = function (x) { var v = dateOrNull_(x); return v ? fmt_(v, 'dd/MM/yyyy') : ''; };
  var etapes = {}; Object.keys(PIECE_DATE_COL).forEach(function (k) { etapes[k] = d(r[PIECE_DATE_COL[k]]); });
  var st = String(r[PC.statut] || PIECE_STATUTS[0]), depuis = null;
  // Nombre de jours passés à l'étape en cours (alerte si ça traîne)
  var dEt = PIECE_DATE_COL[st] !== undefined ? dateOrNull_(r[PIECE_DATE_COL[st]]) : dateOrNull_(r[PC.cree]);
  if (dEt) { depuis = Math.floor((new Date().getTime() - dEt.getTime()) / 86400000); }
  return { ligne: i, n: String(r[PC.n]), cree: d(r[PC.cree]), par: String(r[PC.par] || ''), cabId: String(r[PC.cabId] || ''), cabinet: String(r[PC.cab] || ''),
    appareil: String(r[PC.app] || ''), serie: String(r[PC.serie] || '').replace(/^'/, ''), fournisseur: String(r[PC.frs] || ''), type: String(r[PC.type] || ''),
    reference: String(r[PC.ref] || '').replace(/^'/, ''), designation: String(r[PC.desig] || ''), panne: String(r[PC.panne] || ''), mode: String(r[PC.mode] || ''),
    statut: st, commande: String(r[PC.cmd] || '').replace(/^'/, ''), rma: String(r[PC.rma] || '').replace(/^'/, ''), suivi: String(r[PC.suivi] || '').replace(/^'/, ''),
    montant: r[PC.montant] === '' ? '' : String(r[PC.montant]), notes: String(r[PC.notes] || ''), historique: String(r[PC.histo] || ''), etapes: etapes,
    technicien: String(r[PC.tech] || ''), tache: String(r[PC.tache] || ''), casier: String(r[PC.casier] || ''), tacheOrigine: String(r[PC.orig] || ''), depuis: depuis };
}
function pieceListe_() {
  var sh = pieceFeuille_();
  var l = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, PIECE_HEADERS.length).getValues().map(function (r, i) { return pieceObj_(r, i + 2); }).filter(function (o) { return o.n.trim(); }) : [];
  var enCours = l.filter(function (o) { return PIECE_FINIS.indexOf(o.statut) < 0; }).reverse(), finis = l.filter(function (o) { return PIECE_FINIS.indexOf(o.statut) >= 0; }).slice(-40).reverse();
  return { ok: true, enCours: enCours, finis: finis, statuts: PIECE_STATUTS, modes: PIECE_MODES };
}
// Qui fait quoi : Séverine commande et renvoie (à défaut le gérant), le technicien change la pièce
function pieceBureau_(m) {
  var sev = equipe_().filter(function (x) { return sameName_(String(x.prenom)) === 'severine'; })[0];
  return sev ? sev.prenom : (membre_(COMPTE_SERVEUR, 'Clément') || m).prenom;
}
function pieceTache_(m, sh, L, r, titre, pour, etatPieces) {
  var nom = [r[PC.type], r[PC.ref]].filter(String).join(' réf. ') + (r[PC.app] ? ' (' + r[PC.app] + ')' : '');
  var detail = 'Fiche ' + r[PC.n] + ' (Dicter › Pièces / Agence › Retours pièces)\nPièce : ' + nom + (r[PC.desig] ? ' · ' + r[PC.desig] : '') + '\nFournisseur : ' + r[PC.frs] +
    ' · ' + r[PC.mode] + '\nPanne : ' + r[PC.panne] + (r[PC.serie] ? '\nN° de série appareil : ' + String(r[PC.serie]).replace(/^'/, '') : '') +
    (r[PC.cmd] ? '\nCommande fournisseur : ' + String(r[PC.cmd]).replace(/^'/, '') : '') + (r[PC.rma] ? '\nN° de retour : ' + String(r[PC.rma]).replace(/^'/, '') : '');
  var t = null; try { t = tacheAjouter_(m, { titre: titre + ' : ' + nom, cabinet: r[PC.cab], detail: detail, pour: [pour], pieces: etatPieces || '' }); } catch (e) {}
  if (t && t.ok) {
    sh.getRange(L, PC.tache + 1).setValue(t.n); r[PC.tache] = t.n;
    sh.getRange(L, PC.taches + 1).setValue(cut_([String(r[PC.taches] || ''), t.n].filter(String).join(', '), 300));
  }
  return t;
}
// Met à jour la colonne « Pièces » d'une tâche (sans passer par l'appli)
function tachePiecesServeur_(n, etat) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_TACHES); if (!sh || !n) { return; }
  var v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues(), col = TACHE_HEADERS.indexOf('Pièces') + 1;
  for (var i = 1; i < v.length; i++) { if (String(v[i][0]) === String(n)) { sh.getRange(i + 1, col).setValue(etat); return; } }
}
// Pour l'appli : n° de tâche → fiche pièce liée (la plus récente)
function pieceParTache_() {
  var sh = getSpreadsheet_().getSheetByName(SHEET_PIECES), map = {};
  if (!sh || sh.getLastRow() < 2) { return map; }
  sh.getRange(2, 1, sh.getLastRow() - 1, PIECE_HEADERS.length).getValues().forEach(function (r, i) {
    var o = pieceObj_(r, i + 2); if (!o.n.trim()) { return; }
    var f = { n: o.n, statut: o.statut, casier: o.casier, type: o.type, reference: o.reference, mode: o.mode, depuis: o.depuis };
    [o.tacheOrigine, o.tache].forEach(function (t) { if (t) { map[t] = f; } });
  });
  return map;
}
function pieceCreer_(m, o) {
  var s = function (k, n) { return cut_(String(o[k] == null ? '' : o[k]).trim(), n); };
  var type = s('type', 80), ref = s('reference', 80), cab = s('cabinet', 120), panne = s('panne', 1500), mode = PIECE_MODES.indexOf(o.mode) >= 0 ? o.mode : '';
  if (cab.length < 2) { return { ok: false, message: 'Indiquez le cabinet.' }; }
  if (!type && !ref) { return { ok: false, message: 'Indiquez le type de pièce ou sa référence.' }; }
  if (!mode) { return { ok: false, message: 'Choisissez : garantie ou échange standard.' }; }
  if (panne.length < 3) { return { ok: false, message: 'Décrivez la nature de la panne.' }; }
  var sh = pieceFeuille_(), n = prochainNumero_(sh, 'RP'), st = PIECE_STATUTS.indexOf(o.statut) >= 0 && PIECE_FINIS.indexOf(o.statut) < 0 ? o.statut : 'À commander';
  var row = []; for (var k = 0; k < PIECE_HEADERS.length; k++) { row.push(''); }
  row[PC.n] = n; row[PC.cree] = new Date(); row[PC.par] = m.prenom; row[PC.cabId] = s('cabId', 30); row[PC.cab] = cab; row[PC.app] = s('appareil', 120);
  row[PC.serie] = s('serie', 60) ? "'" + s('serie', 60) : ''; row[PC.frs] = s('fournisseur', 80); row[PC.type] = type; row[PC.ref] = ref ? "'" + ref : '';
  row[PC.desig] = s('designation', 200); row[PC.panne] = panne; row[PC.mode] = mode; row[PC.statut] = st; row[PC.cmd] = s('commande', 60) ? "'" + s('commande', 60) : '';
  row[PC.notes] = s('notes', 800); row[PC.tech] = s('technicien', 60) || m.prenom; row[PC.orig] = s('tacheOrigine', 30); row[PC.casier] = s('casier', 60);
  if (PIECE_DATE_COL[st] !== undefined) { row[PIECE_DATE_COL[st]] = new Date(); }
  row[PC.histo] = fmt_(new Date(), 'dd/MM/yyyy HH:mm') + ' · Créée : ' + st + ' (' + m.prenom + ')';
  sh.appendRow(row);
  var L = sh.getLastRow();
  // Tâche suivante selon l'étape de départ
  if (st === 'À commander') { pieceTache_(m, sh, L, row, 'Commander la pièce (' + mode.toLowerCase() + ')', pieceBureau_(m), 'À commander'); }
  else if (st === 'Commandée') { /* déjà commandée : on attend la réception */ }
  else if ((st === 'Reçue' || st === 'En casier') && !row[PC.orig]) { pieceTache_(m, sh, L, row, 'Pièce ' + (st === 'En casier' ? 'en casier' : 'reçue') + ' : la changer chez le client', row[PC.tech], PIECE_ETAT_TACHE[st]); }
  else if (st === 'Changée chez le client') { pieceTache_(m, sh, L, row, 'Renvoyer l\'ancienne pièce au fournisseur', pieceBureau_(m), ''); }
  if (row[PC.orig] && PIECE_ETAT_TACHE[st]) { try { tachePiecesServeur_(row[PC.orig], PIECE_ETAT_TACHE[st]); } catch (e) {} }
  try { log_('Agence', cab, row[PC.cabId], '', m.prenom, [type, ref].join(' '), 'PIÈCE ' + mode.toUpperCase(), n + ' · ' + cut_(panne, 80)); } catch (e) {}
  return { ok: true, n: n, fiche: pieceObj_(sh.getRange(L, 1, 1, PIECE_HEADERS.length).getValues()[0], L) };
}
function pieceMaj_(m, req) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_PIECES);
  if (!sh || sh.getLastRow() < 2) { return { ok: false, message: 'Fiche introuvable.' }; }
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, PIECE_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][PC.n]) !== String(req.n)) { continue; }
    var r = rows[i], L = i + 2, st = String(req.statut || ''), histo = [];
    if (st && PIECE_STATUTS.indexOf(st) < 0) { return { ok: false, message: 'Étape inconnue.' }; }
    var champs = { cabinet: PC.cab, appareil: PC.app, serie: PC.serie, fournisseur: PC.frs, type: PC.type, reference: PC.ref, designation: PC.desig, panne: PC.panne,
      mode: PC.mode, commande: PC.cmd, rma: PC.rma, suivi: PC.suivi, montant: PC.montant, notes: PC.notes, technicien: PC.tech, casier: PC.casier };
    var lib = { casier: 'Casier', commande: 'Commande fournisseur', rma: 'N° de retour', suivi: 'Suivi colis', montant: 'Montant', mode: 'Mode', reference: 'Référence' };
    Object.keys(champs).forEach(function (k) {
      if (req[k] === undefined) { return; }
      var v = cut_(String(req[k]).trim(), 1500);
      if (k === 'mode' && PIECE_MODES.indexOf(v) < 0) { return; }
      if (k === 'montant') { v = v === '' ? '' : (Number(String(v).replace(',', '.').replace(/[^\d.]/g, '')) || ''); }
      if (/^(serie|reference|commande|rma|suivi)$/.test(k) && v) { v = "'" + v; }
      if (String(r[champs[k]]).replace(/^'/, '') !== String(v).replace(/^'/, '')) {
        r[champs[k]] = v; sh.getRange(L, champs[k] + 1).setValue(v);
        if (lib[k] && v !== '') { histo.push(lib[k] + ' ' + String(v).replace(/^'/, '')); }
      }
    });
    if (st && st !== String(r[PC.statut])) {
      sh.getRange(L, PC.statut + 1).setValue(st); r[PC.statut] = st;
      if (PIECE_DATE_COL[st] !== undefined) { sh.getRange(L, PIECE_DATE_COL[st] + 1).setValue(dateOrNull_(req.le) || new Date()); }
      histo.unshift(st);
      // La tâche de l'étape précédente est faite, on crée celle de l'étape suivante.
      // Pièce mise en casier : la tâche « changer la pièce » reste ouverte. Tâche d'origine (appel du cabinet) : c'est elle, l'intervention.
      var orig = String(r[PC.orig] || '');
      if (String(r[PC.tache]) && st !== 'En casier') { try { tacheFaitServeur_(String(r[PC.tache]), m.prenom); } catch (e) {} sh.getRange(L, PC.tache + 1).setValue(''); r[PC.tache] = ''; }
      if (orig && PIECE_ETAT_TACHE[st]) { try { tachePiecesServeur_(orig, PIECE_ETAT_TACHE[st]); } catch (e) {} }
      if ((st === 'Reçue' || st === 'En casier') && !orig && !String(r[PC.tache])) { pieceTache_(m, sh, L, r, 'Pièce ' + (st === 'En casier' ? 'en casier' : 'reçue') + ' : la changer chez le client', String(r[PC.tech] || r[PC.par] || m.prenom), PIECE_ETAT_TACHE[st]); }
      else if (st === 'En casier' && String(r[PC.tache])) { try { tachePiecesServeur_(String(r[PC.tache]), 'En casier'); } catch (e) {} }
      if (st === 'Changée chez le client' && orig) { try { tacheFaitServeur_(orig, m.prenom); } catch (e) {} }
      if (st === 'Changée chez le client') { pieceTache_(m, sh, L, r, 'Renvoyer l\'ancienne pièce au fournisseur', pieceBureau_(m), ''); }
      if (st === 'Ancienne pièce renvoyée') { pieceTache_(m, sh, L, r, 'Suivre l\'avoir du fournisseur (pièce renvoyée)', pieceBureau_(m), ''); }
    }
    if (histo.length) { sh.getRange(L, PC.histo + 1).setValue(cut_(String(r[PC.histo] || '') + '\n' + fmt_(new Date(), 'dd/MM/yyyy HH:mm') + ' · ' + histo.join(' · ') + ' (' + m.prenom + ')', 4000)); }
    return { ok: true, fiche: pieceObj_(sh.getRange(L, 1, 1, PIECE_HEADERS.length).getValues()[0], L) };
  }
  return { ok: false, message: 'Fiche introuvable.' };
}
function pieceSuppr_(m, req) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_PIECES);
  if (!sh || sh.getLastRow() < 2) { return { ok: false, message: 'Fiche introuvable.' }; }
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, PIECE_HEADERS.length).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][PC.n]) !== String(req.n)) { continue; }
    if (String(v[i][PC.tache])) { try { tacheFaitServeur_(String(v[i][PC.tache]), m.prenom); } catch (e) {} }
    sh.deleteRow(i + 2);
    try { log_('Agence', v[i][PC.cab], v[i][PC.cabId], '', m.prenom, [v[i][PC.type], String(v[i][PC.ref]).replace(/^'/, '')].join(' '), 'PIÈCE SUPPRIMÉE', String(req.n)); } catch (e) {}
    return { ok: true };
  }
  return { ok: false, message: 'Fiche introuvable.' };
}
