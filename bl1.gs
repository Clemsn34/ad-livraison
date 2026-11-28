// ============================================================================
// BON DE LIVRAISON DES INSTALLATIONS : le matériel vendu, coché à réception (bureau) puis au chargement (technicien)
// Liste tirée de la commande Extrabat du praticien (à défaut son dernier devis), sinon du matériel de l'installation.
// Alerte dans l'Agence si la date d'installation approche et que tout n'est pas reçu.
// ============================================================================
var SHEET_BL = 'Bons de livraison';
var BL_HEADERS = ['Installation', 'Ligne', 'Désignation', 'Référence', 'Quantité', 'Source', 'Reçu le', 'Reçu par', 'Chargé le', 'Chargé par', 'Note'];
var BL_EXCLU = /main.d.(?:oe|œ)uvre|d[ée]placement|forfait|frais de port|port et emballage|livraison|formation|mise en service|param[ée]trage|acompte|remise|escompte|garantie|extension|contrat|abonnement|financement|installation des donn|transport/i;

function blFeuille_() {
  var ss = getSpreadsheet_(), sh = ss.getSheetByName(SHEET_BL);
  if (!sh) { sh = ss.insertSheet(SHEET_BL); styleSheet_(sh, BL_HEADERS, [100, 50, 340, 140, 70, 200, 95, 90, 95, 90, 200]); }
  return sh;
}
function blInstall_(n) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_INSTALL); if (!sh || sh.getLastRow() < 2) { return null; }
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, INSTALL_HEADERS.length).getValues();
  for (var i = 0; i < v.length; i++) { if (String(v[i][0]) === String(n)) { return { n: String(v[i][0]), cabId: String(v[i][2] || ''), cabinet: String(v[i][3] || ''), praticien: String(v[i][4] || ''), date: dateOrNull_(v[i][5]), materiel: String(v[i][6] || '') }; } }
  return null;
}
function blLignes_(n) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_BL); if (!sh || sh.getLastRow() < 2) { return []; }
  var d = function (x) { var v = dateOrNull_(x); return v ? fmt_(v, 'dd/MM/yyyy') : ''; };
  return sh.getRange(2, 1, sh.getLastRow() - 1, BL_HEADERS.length).getValues().map(function (r, i) { return { row: i + 2, r: r }; })
    .filter(function (x) { return String(x.r[0]) === String(n); })
    .map(function (x) { var r = x.r; return { ligne: Number(r[1]), designation: String(r[2]), reference: String(r[3]).replace(/^'/, ''), quantite: r[4] === '' ? '' : String(r[4]), source: String(r[5]),
      recuLe: d(r[6]), recuPar: String(r[7] || ''), chargeLe: d(r[8]), chargePar: String(r[9] || ''), note: String(r[10] || '') }; })
    .sort(function (a, b) { return a.ligne - b.ligne; });
}
// Lignes de la commande (sinon du devis) Extrabat du praticien
function blDepuisExtrabat_(ins) {
  var co = projetCoordonnees_({ cabId: ins.cabId, nom: ins.cabinet });
  if (!co || !co.exId) { return { ok: false, message: 'Cabinet non relié à Extrabat.' }; }
  var auj = new Date(), debut = new Date(auj.getFullYear() - 2, auj.getMonth(), auj.getDate());
  var types = ['commande', 'devis'];
  for (var t = 0; t < types.length; t++) {
    var l = [];
    try { l = exbListe_(exbCall_('get', '/v1/pieces', { clients: co.exId, types: types[t], include: 'lignes', date_debut: fmt_(debut, 'yyyy-MM-dd'), date_fin: fmt_(new Date(auj.getTime() + 365 * 864e5), 'yyyy-MM-dd'), nbitem: 50, page: 1 })); } catch (e) { continue; }
    if (!l.length) { continue; }
    // La plus récente, la plus grosse en cas d'égalité de date
    l.sort(function (a, b) { return String(b.date || '').localeCompare(String(a.date || '')) || (Number(b.totalTTC || b.montantTTC || 0) - Number(a.totalTTC || a.montantTTC || 0)); });
    var pc = l[0], lignes = [];
    (pc.lignes || []).forEach(function (lg) {
      var desc = String(lg.description || lg.descriptionCondensee || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
      var titre = desc.split(/\s{2,}|\. |\n/)[0].slice(0, 140) || String(lg.code || '');
      var qte = Number(lg.quantite || lg.qte || lg.quantity || 0);
      if (!titre || BL_EXCLU.test(titre)) { return; }
      if (!lg.code && !qte) { return; } // titres de section et commentaires
      lignes.push({ designation: titre, reference: String(lg.code || ''), quantite: qte || 1 });
    });
    if (lignes.length) { return { ok: true, lignes: lignes, source: (types[t] === 'commande' ? 'Commande ' : 'Devis ') + String(pc.code || pc.numero || '') + ' du ' + String(pc.date || '').slice(0, 10).split('-').reverse().join('/') }; }
  }
  return { ok: false, message: 'Aucune commande ni devis Extrabat trouvé pour ce cabinet.' };
}
function blConstruire_(ins, force) {
  var sh = blFeuille_();
  if (force && sh.getLastRow() > 1) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = v.length - 1; i >= 0; i--) { if (String(v[i][0]) === String(ins.n)) { sh.deleteRow(i + 2); } }
  }
  var x = null; try { x = blDepuisExtrabat_(ins); } catch (e) { x = { ok: false, message: e.message }; }
  var lignes = x && x.ok ? x.lignes : String(ins.materiel || '').split(/\s*[,;\n]\s*/).filter(String).map(function (m) { return { designation: m, reference: '', quantite: 1 }; });
  var source = x && x.ok ? x.source : 'Matériel de l\'installation' + (x && x.message ? ' (' + x.message + ')' : '');
  if (!lignes.length) { return { ok: false, message: (x && x.message) || 'Aucun matériel trouvé.' }; }
  sh.getRange(sh.getLastRow() + 1, 1, lignes.length, BL_HEADERS.length).setValues(lignes.map(function (l, i) {
    return [ins.n, i + 1, cut_(l.designation, 300), l.reference ? "'" + cut_(l.reference, 60) : '', l.quantite, source, '', '', '', '', ''];
  }));
  return { ok: true, source: source };
}
function blGet_(m, req) {
  var ins = blInstall_(req.install); if (!ins) { return { ok: false, message: 'Installation introuvable.' }; }
  var l = blLignes_(ins.n), info = '';
  if (!l.length || req.refaire) { var c = blConstruire_(ins, !!req.refaire); if (!c.ok) { return c; } info = c.source; l = blLignes_(ins.n); }
  return { ok: true, install: { n: ins.n, cabinet: ins.cabinet, praticien: ins.praticien, date: ins.date ? fmt_(ins.date, 'dd/MM/yyyy') : '' }, lignes: l, source: l.length ? l[0].source : info };
}
// Coche / décoche : recu (bureau, à réception) ou charge (technicien, dans le camion). ligne = 'tout' pour tout cocher.
function blCoche_(m, req) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_BL); if (!sh || sh.getLastRow() < 2) { return { ok: false, message: 'Bon introuvable.' }; }
  var col = req.champ === 'charge' ? 9 : 7, v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues(), fait = 0;
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) !== String(req.install) || (req.ligne !== 'tout' && Number(v[i][1]) !== Number(req.ligne))) { continue; }
    sh.getRange(i + 2, col, 1, 2).setValues([req.val ? [new Date(), m.prenom] : ['', '']]); fait++;
  }
  if (!fait) { return { ok: false, message: 'Ligne introuvable.' }; }
  try { log_('Agence', '', '', '', m.prenom, String(req.install), 'BON DE LIVRAISON', (req.champ === 'charge' ? 'Chargé' : 'Reçu') + ' ' + (req.val ? '✓' : '✗') + ' · ligne ' + req.ligne); } catch (e) {}
  return blGet_(m, { install: req.install });
}
function blAjouter_(m, req) {
  var ins = blInstall_(req.install); if (!ins) { return { ok: false, message: 'Installation introuvable.' }; }
  var des = cut_(String(req.designation || '').trim(), 300); if (des.length < 2) { return { ok: false, message: 'Indiquez le matériel.' }; }
  var sh = blFeuille_(), l = blLignes_(ins.n), n = l.reduce(function (s, x) { return Math.max(s, x.ligne); }, 0) + 1;
  sh.appendRow([ins.n, n, des, req.reference ? "'" + cut_(String(req.reference), 60) : '', Number(req.quantite) || 1, 'Ajouté par ' + m.prenom, '', '', '', '', '']);
  return blGet_(m, { install: ins.n });
}
// Pour l'appli : par installation, nombre de lignes / reçues / chargées
function blStats_() {
  var sh = getSpreadsheet_().getSheetByName(SHEET_BL), o = {};
  if (!sh || sh.getLastRow() < 2) { return o; }
  sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues().forEach(function (r) {
    var k = String(r[0]); if (!k) { return; }
    var s = o[k] || (o[k] = { total: 0, recus: 0, charges: 0 }); s.total++; if (r[6]) { s.recus++; } if (r[8]) { s.charges++; }
  });
  return o;
}
