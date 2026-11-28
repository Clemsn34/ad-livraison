// ============================================================================
// RÉPARATIONS D'INSTRUMENTS (turbines, contre-angles, détartreurs…) envoyés chez le fournisseur
// Récupéré chez le praticien → demande de devis faite → devis fait → devis accepté → retour demandé → retourné au client.
// Saisie dans Dicter (mode « Instruments ») ou demandée par le praticien depuis son appli.
// ============================================================================
var SHEET_INSTR = 'Réparations instruments';
var INSTR_HEADERS = ['N°', 'Créé le', 'Origine', 'Identifiant cabinet', 'Cabinet', 'Contact', 'Téléphone', 'E-mail', 'Marque', 'Type', 'Modèle', 'N° de série',
  'Panne', 'Date de la panne', 'Garantie', 'Info garantie', 'Garantie demandée', 'Statut', 'Fournisseur', 'N° de dossier fournisseur', 'Montant du devis',
  'Récupéré le', 'Demande de devis le', 'Devis fait le', 'Devis accepté le', 'Retour demandé le', 'Retourné le', 'Par', 'Notes', 'Historique', 'Tâche'];
var INSTR_STATUTS = ['À récupérer chez le praticien', 'Récupéré', 'Demande de devis faite', 'Devis fait', 'Devis accepté', 'Devis refusé', 'Retour demandé', 'Retourné au client'];
var INSTR_DATE_COL = { 'Récupéré': 21, 'Demande de devis faite': 22, 'Devis fait': 23, 'Devis accepté': 24, 'Retour demandé': 25, 'Retourné au client': 26 };
var IC = { n: 0, cree: 1, orig: 2, cabId: 3, cab: 4, contact: 5, tel: 6, email: 7, marque: 8, type: 9, modele: 10, serie: 11, panne: 12, datePanne: 13,
  gar: 14, garInfo: 15, garDem: 16, statut: 17, frs: 18, dossier: 19, montant: 20, par: 27, notes: 28, histo: 29, tache: 30 };
var INSTR_GARANTIES = ['Sous garantie', 'Hors garantie', 'À vérifier'];

function instrFeuille_() {
  var ss = getSpreadsheet_(), sh = ss.getSheetByName(SHEET_INSTR);
  if (!sh) {
    sh = ss.insertSheet(SHEET_INSTR);
    styleSheet_(sh, INSTR_HEADERS, [100, 95, 110, 100, 190, 140, 110, 170, 90, 110, 130, 120, 260, 95, 105, 220, 80, 170, 110, 120, 90, 95, 95, 95, 95, 95, 95, 90, 220, 300, 90]);
    setChoices_(sh, IC.statut + 1, INSTR_STATUTS);
    setChoices_(sh, IC.gar + 1, INSTR_GARANTIES);
  }
  return sh;
}
function instrObj_(r, i) {
  var d = function (x) { var v = dateOrNull_(x); return v ? fmt_(v, 'dd/MM/yyyy') : ''; };
  var etapes = {}; Object.keys(INSTR_DATE_COL).forEach(function (k) { etapes[k] = d(r[INSTR_DATE_COL[k]]); });
  return { ligne: i, n: String(r[IC.n]), cree: d(r[IC.cree]), origine: String(r[IC.orig] || ''), cabId: String(r[IC.cabId] || ''), cabinet: String(r[IC.cab] || ''),
    contact: String(r[IC.contact] || ''), tel: String(r[IC.tel] || ''), email: String(r[IC.email] || ''), marque: String(r[IC.marque] || ''), type: String(r[IC.type] || ''),
    modele: String(r[IC.modele] || ''), serie: String(r[IC.serie] || ''), panne: String(r[IC.panne] || ''), datePanne: d(r[IC.datePanne]),
    garantie: String(r[IC.gar] || ''), garantieInfo: String(r[IC.garInfo] || ''), garantieDemandee: r[IC.garDem] === true || String(r[IC.garDem]) === 'Oui',
    statut: String(r[IC.statut] || INSTR_STATUTS[1]), fournisseur: String(r[IC.frs] || ''), dossier: String(r[IC.dossier] || ''), montant: r[IC.montant] === '' ? '' : String(r[IC.montant]),
    etapes: etapes, par: String(r[IC.par] || ''), notes: String(r[IC.notes] || ''), historique: String(r[IC.histo] || ''), tache: String(r[IC.tache] || '') };
}
function instrListe_() {
  var sh = getSpreadsheet_().getSheetByName(SHEET_INSTR);
  var l = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, INSTR_HEADERS.length).getValues().map(function (r, i) { return instrObj_(r, i + 2); }).filter(function (o) { return o.n.trim(); }) : [];
  var enCours = l.filter(function (o) { return o.statut !== 'Retourné au client'; }).reverse(), finis = l.filter(function (o) { return o.statut === 'Retourné au client'; }).slice(-25).reverse();
  return { ok: true, enCours: enCours, finis: finis, statuts: INSTR_STATUTS };
}
// Fournisseur par défaut selon la marque (modifiable sur la fiche)
function instrFournisseur_(marque) {
  var m = sameName_(marque || '');
  if (/nsk/.test(m)) { return 'NSK France'; }
  if (/bien ?air/.test(m)) { return 'Bien-Air France'; }
  if (/acteon|satelec/.test(m)) { return 'Acteon'; }
  if (/w ?& ?h|w et h|wh\b/.test(m)) { return 'W&H France'; }
  if (/kavo/.test(m)) { return 'KaVo'; }
  if (/ems/.test(m)) { return 'EMS'; }
  return String(marque || '');
}
// Création d'une fiche (Dicter ou appli praticien). o.origine : 'Agence' ou 'Appli praticien'
function instrCreer_(m, o) {
  var s = function (k, n) { return cut_(String(o[k] == null ? '' : o[k]).trim(), n); };
  var marque = s('marque', 60), type = s('type', 60), modele = s('modele', 80);
  if (!marque && !type && !modele) { return { ok: false, message: 'Indiquez au moins la marque ou le type d\'instrument.' }; }
  var sh = instrFeuille_(), n = prochainNumero_(sh, 'RI'), appli = o.origine === 'Appli praticien';
  var st = appli ? INSTR_STATUTS[0] : (INSTR_STATUTS.indexOf(o.statut) >= 0 ? o.statut : 'Récupéré');
  var dp = dateOrNull_(o.datePanne) || new Date(), gar = INSTR_GARANTIES.indexOf(o.garantie) >= 0 ? o.garantie : 'À vérifier';
  var row = []; for (var k = 0; k < INSTR_HEADERS.length; k++) { row.push(''); }
  row[IC.n] = n; row[IC.cree] = new Date(); row[IC.orig] = appli ? 'Appli praticien' : 'Agence'; row[IC.cabId] = s('cabId', 30); row[IC.cab] = s('cabinet', 120);
  row[IC.contact] = s('contact', 120); row[IC.tel] = s('tel', 30) ? "'" + s('tel', 30) : ''; row[IC.email] = s('email', 120);
  row[IC.marque] = marque; row[IC.type] = type; row[IC.modele] = modele; row[IC.serie] = s('serie', 60) ? "'" + s('serie', 60) : '';
  row[IC.panne] = s('panne', 1500); row[IC.datePanne] = dp; row[IC.gar] = gar; row[IC.garInfo] = s('garantieInfo', 300); row[IC.garDem] = o.garantieDemandee ? 'Oui' : 'Non';
  row[IC.statut] = st; row[IC.frs] = s('fournisseur', 80) || instrFournisseur_(marque); row[IC.par] = m.prenom; row[IC.notes] = s('notes', 800);
  if (st === 'Récupéré') { row[INSTR_DATE_COL['Récupéré']] = new Date(); }
  row[IC.histo] = fmt_(new Date(), 'dd/MM/yyyy HH:mm') + ' · ' + st + ' (' + m.prenom + ')';
  sh.appendRow(row);
  try { log_('Agence', row[IC.cab], row[IC.cabId], '', m.prenom, [marque, type, modele].join(' '), 'INSTRUMENT', n + ' · ' + cut_(row[IC.panne], 80)); } catch (e) {}
  return { ok: true, n: n, fiche: instrObj_(row, sh.getLastRow()) };
}
function instrMaj_(m, req) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_INSTR);
  if (!sh || sh.getLastRow() < 2) { return { ok: false, message: 'Fiche introuvable.' }; }
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, INSTR_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][IC.n]) !== String(req.n)) { continue; }
    var r = rows[i], L = i + 2, st = String(req.statut || ''), histo = [];
    if (st && INSTR_STATUTS.indexOf(st) < 0) { return { ok: false, message: 'Étape inconnue.' }; }
    var champs = { cabinet: IC.cab, contact: IC.contact, tel: IC.tel, email: IC.email, marque: IC.marque, type: IC.type, modele: IC.modele, serie: IC.serie, panne: IC.panne,
      garantie: IC.gar, garantieInfo: IC.garInfo, fournisseur: IC.frs, dossier: IC.dossier, montant: IC.montant, notes: IC.notes };
    Object.keys(champs).forEach(function (k) {
      if (req[k] === undefined) { return; }
      var v = cut_(String(req[k]), 1500);
      if (k === 'garantie' && v && INSTR_GARANTIES.indexOf(v) < 0) { return; }
      if (k === 'montant') { v = v === '' ? '' : (Number(String(v).replace(',', '.').replace(/[^\d.]/g, '')) || ''); }
      if ((k === 'serie' || k === 'tel') && v) { v = "'" + v; }
      if (String(r[champs[k]]) !== String(v).replace(/^'/, '')) { r[champs[k]] = v; sh.getRange(L, champs[k] + 1).setValue(v); if (k === 'montant' && v !== '') { histo.push('Devis ' + v + ' €'); } }
    });
    if (req.datePanne !== undefined) { var dpn = dateOrNull_(req.datePanne); if (dpn) { sh.getRange(L, IC.datePanne + 1).setValue(dpn); } }
    if (req.garantieDemandee !== undefined) { sh.getRange(L, IC.garDem + 1).setValue(req.garantieDemandee ? 'Oui' : 'Non'); }
    if (st && st !== String(r[IC.statut])) {
      sh.getRange(L, IC.statut + 1).setValue(st); r[IC.statut] = st;
      if (INSTR_DATE_COL[st] !== undefined) { sh.getRange(L, INSTR_DATE_COL[st] + 1).setValue(dateOrNull_(req.le) || new Date()); }
      histo.push(st);
      // Instrument récupéré : la tâche « à récupérer » de la demande praticien est faite
      if (st === 'Récupéré' && String(r[IC.tache])) { try { tacheFaitServeur_(String(r[IC.tache]), m.prenom); } catch (e) {} }
    }
    if (histo.length) { sh.getRange(L, IC.histo + 1).setValue(cut_(String(r[IC.histo] || '') + '\n' + fmt_(new Date(), 'dd/MM/yyyy HH:mm') + ' · ' + histo.join(' · ') + ' (' + m.prenom + ')', 4000)); }
    return { ok: true, fiche: instrObj_(sh.getRange(L, 1, 1, INSTR_HEADERS.length).getValues()[0], L) };
  }
  return { ok: false, message: 'Fiche introuvable.' };
}
// Coche une tâche comme faite (sans passer par l'appli)
function tacheFaitServeur_(n, prenom) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_TACHES); if (!sh) { return; }
  var v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = 1; i < v.length; i++) { if (String(v[i][0]) === String(n)) { sh.getRange(i + 1, 9, 1, 3).setValues([['Fait', new Date(), prenom]]); return; } }
}
// Garantie : durée de l'onglet « Garanties » (instruments rotatifs, par marque si une ligne existe)
function instrDureeGarantie_(marque, type) {
  var mois = 24, sh = getSpreadsheet_().getSheetByName('Garanties');
  if (!sh || sh.getLastRow() < 2) { return mois; }
  var cat = /detartr|ultrason|lampe|polymer/.test(sameName_(type || '')) ? 'autre' : 'instruments rotatifs', mq = sameName_(marque || ''), gen = null, spe = null;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (x) {
    var c = sameName_(String(x[0] || '')), b = sameName_(String(x[1] || '')), d = Number(x[2]);
    if (!d || c.indexOf(cat) !== 0) { return; }
    if (!b) { gen = d; } else if (mq && (mq.indexOf(b) >= 0 || b.indexOf(mq) >= 0)) { spe = d; }
  });
  return spe || gen || mois;
}
// Cherche l'achat de l'instrument dans les factures Extrabat du cabinet (5 dernières années) et dans « Matériel praticiens installés »
function instrGarantie_(req) {
  var marque = String(req.marque || ''), type = String(req.type || ''), modele = String(req.modele || ''), serie = String(req.serie || '').trim();
  var duree = instrDureeGarantie_(marque, type), auj = new Date(), achat = null, source = '', libelle = '';
  var mots = sameName_(modele).split(' ').filter(function (w) { return w.length >= 2; });
  var colle = sameName_(modele).replace(/\s/g, ''), mq = sameName_(marque);
  var correspond = function (txt) {
    var t = sameName_(txt), tc = t.replace(/\s/g, '');
    if (serie && serie.length >= 4 && tc.indexOf(sameName_(serie).replace(/\s/g, '')) >= 0) { return 3; }
    if (colle.length >= 3 && tc.indexOf(colle) >= 0) { return 2; }
    if (mots.length && mots.every(function (w) { return t.indexOf(w) >= 0; })) { return 2; }
    return 0;
  };
  // 1. Matériel installé noté dans l'Agence
  try {
    var pm = getSpreadsheet_().getSheetByName('Matériel praticiens installés');
    if (pm && pm.getLastRow() > 1) {
      pm.getRange(2, 1, pm.getLastRow() - 1, 8).getValues().forEach(function (x) {
        if (achat) { return; }
        var memeCab = !req.cabinet || sameName_(String(x[1])).indexOf(sameName_(req.cabinet)) >= 0 || sameName_(req.cabinet).indexOf(sameName_(String(x[1]))) >= 0;
        if (memeCab && correspond(x[2] + ' ' + x[3] + ' ' + x[4]) && dateOrNull_(x[5])) { achat = dateOrNull_(x[5]); source = 'Matériel installé'; libelle = String(x[3] || x[2]); }
      });
    }
  } catch (e) {}
  // 2. Factures Extrabat du client
  var alerte = '';
  if (!achat && (modele || serie)) {
    try {
      var co = projetCoordonnees_({ cabId: req.cabId || '', nom: req.cabinet || req.contact || '' }), exId = co && co.exId;
      if (exId) {
        var debut = new Date(auj.getFullYear() - 5, auj.getMonth(), auj.getDate()), best = 0;
        for (var p = 1; p <= 5; p++) {
          var l = exbListe_(exbCall_('get', '/v1/pieces', { clients: exId, types: 'facture', include: 'lignes', date_debut: fmt_(debut, 'yyyy-MM-dd'), date_fin: fmt_(auj, 'yyyy-MM-dd'), nbitem: 50, page: p }));
          l.forEach(function (pc) {
            (pc.lignes || []).forEach(function (lg) {
              var sc = correspond((lg.code || '') + ' ' + (lg.description || lg.descriptionCondensee || ''));
              var dt = dateOrNull_(String(pc.date || '').slice(0, 10));
              if (sc && dt && (sc > best || (sc === best && (!achat || dt > achat)))) { best = sc; achat = dt; source = 'Facture ' + (pc.code || ''); libelle = String(lg.description || lg.code || '').split('\n')[0].slice(0, 80); }
            });
          });
          if (l.length < 50) { break; }
        }
      } else { alerte = 'Cabinet non retrouvé dans Extrabat.'; }
    } catch (e) { alerte = 'Extrabat : ' + e.message; }
  }
  if (!achat) { return { ok: true, garantie: 'À vérifier', info: (alerte ? alerte + ' ' : '') + 'Achat non retrouvé (garantie ' + duree + ' mois).', duree: duree }; }
  var fin = new Date(achat.getTime()); fin.setMonth(fin.getMonth() + duree);
  var sous = fin >= auj;
  return { ok: true, garantie: sous ? 'Sous garantie' : 'Hors garantie', achat: fmt_(achat, 'dd/MM/yyyy'), fin: fmt_(fin, 'dd/MM/yyyy'), duree: duree,
    info: (sous ? 'Sous garantie jusqu\'au ' : 'Garantie finie le ') + fmt_(fin, 'dd/MM/yyyy') + ' · acheté le ' + fmt_(achat, 'dd/MM/yyyy') + ' (' + source + (libelle ? ' : ' + libelle : '') + ')' };
}

// Appli praticiens : « J'ai un instrument à réparer » → fiche « À récupérer » + tâche pour Bruno (instrument à récupérer, devis à faire) + e-mail à l'agence
function reparationDemandePublic_(req) {
  var s = function (k, n) { return cut_(String(req[k] || '').trim(), n); };
  var prenom = s('prenom', 60), nom = s('nom', 80), tel = s('portable', 30), email = s('email', 120).toLowerCase(), cabinet = s('cabinet', 120), ville = s('ville', 80);
  var marque = s('marque', 60), type = s('type', 60), modele = s('modele', 80), serie = s('serie', 60), panne = s('panne', 1200), garDem = !!req.garantie;
  if (nom.length < 2) { return { ok: false, message: 'Indiquez votre nom.' }; }
  if (tel.replace(/\D/g, '').length < 9) { return { ok: false, message: 'Indiquez votre numéro de téléphone.' }; }
  if (email && !/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(email)) { return { ok: false, message: 'Vérifiez votre adresse e-mail.' }; }
  if (!type && !marque) { return { ok: false, message: 'Choisissez la marque ou le type d\'instrument.' }; }
  if (panne.length < 3) { return { ok: false, message: 'Décrivez la panne en quelques mots.' }; }
  var cache = CacheService.getScriptCache(), ck = 'rp_' + tel.replace(/\D/g, '').slice(-9), n = Number(cache.get(ck) || 0), all = Number(cache.get('rp_total') || 0);
  if (n >= 6 || all >= 60) { return { ok: false, message: 'Trop de demandes d\'un coup : réessayez plus tard ou appelez-nous au ' + AD_PHONE + '.' }; }
  cache.put(ck, String(n + 1), 3600); cache.put('rp_total', String(all + 1), 21600);
  var qui = (prenom ? prenom + ' ' : '') + nom, cab = cabinet || ('Dr ' + nom), instr = [type, marque, modele].filter(String).join(' ');
  var m = membre_(COMPTE_SERVEUR, 'Clément');
  // Cabinet connu de l'Agence (même nom) : on garde son identifiant
  var cabId = ''; try { var c = getSpreadsheet_().getSheetByName(SHEET_CLIENTS); if (c && c.getLastRow() > 1) { var nq = sameName_(cab); c.getRange(2, 1, c.getLastRow() - 1, 2).getValues().some(function (x) { if (sameName_(String(x[1])) === nq) { cabId = String(x[0]); return true; } return false; }); } } catch (e) {}
  var g = null; try { g = instrGarantie_({ marque: marque, type: type, modele: modele, serie: serie, cabinet: cab, cabId: cabId, contact: nom }); } catch (e) {}
  var f = instrCreer_(m, { origine: 'Appli praticien', cabId: cabId, cabinet: cab, contact: qui + (ville ? ' · ' + ville : ''), tel: tel, email: email, marque: marque, type: type, modele: modele, serie: serie,
    panne: panne, datePanne: new Date(), garantie: g ? g.garantie : 'À vérifier', garantieInfo: g ? g.info : '', garantieDemandee: garDem });
  if (!f.ok) { return f; }
  var detail = [qui, cabinet, ville, tel, email].filter(String).join(' · ') + '\nInstrument : ' + instr + (serie ? ' · N° de série ' + serie : '') + '\nPanne : ' + panne +
    '\nGarantie : ' + (g ? g.info : 'à vérifier') + (garDem ? ' · le praticien demande la prise sous garantie' : '') + '\nFiche ' + f.n + ' (Dicter › Instruments). Demande envoyée depuis l\'appli praticiens le ' + fmt_(new Date(), 'dd/MM/yyyy HH:mm') + '.';
  var eqp = equipe_(), bruno = eqp.filter(function (x) { return String(x.prenom).trim().toLowerCase() === 'bruno'; })[0];
  var pour = bruno ? bruno.prenom : (pourRole_('Responsable technique') || m.prenom), t = null;
  try { t = tacheAjouter_(m, { titre: 'Instrument à récupérer + devis à faire (appli) : ' + instr, cabinet: cab, detail: detail, pour: [pour] }); } catch (e) {}
  if (t && t.ok) { try { var sh = getSpreadsheet_().getSheetByName(SHEET_INSTR); sh.getRange(f.fiche.ligne, IC.tache + 1).setValue(t.n); } catch (e) {} }
  try {
    var html = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.55"><p><b>Instrument à réparer : demande depuis l\'appli praticiens</b></p><p>' + esc_(detail).replace(/\n/g, '<br>') + '</p>' + (t && t.ok ? '<p>Tâche ' + esc_(t.n) + ' créée pour ' + esc_(pour) + '.</p>' : '') + '</div>';
    MailApp.sendEmail('contact@alternativedentaire.fr', 'Instrument à réparer : ' + instr + ' – ' + cab, detail, { htmlBody: html, name: 'Appli Alternative Dentaire', replyTo: email || 'contact@alternativedentaire.fr' });
  } catch (e) {}
  try { log_('Appli praticiens', cab, cabId, '', qui, '', 'DEMANDE RÉPARATION', instr); } catch (e) {}
  return { ok: true, n: f.n, tache: t && t.n };
}
