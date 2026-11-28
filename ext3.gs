// TEMP ext3
// Remplit « Contrats matériel » (Actif = RAPPEL) à partir du parc « Matériel praticiens installés » :
// chaque appareil reconnu reçoit son article de révision du tarif entretien. Une date de révision commune par cabinet :
// prochain anniversaire de son achat le plus récent, au moins 330 jours plus tard (le contrat se propose un an avant).
function tmpExt3() {
  var ss = getSpreadsheet_(), auj = today_();
  var tarifs = {}; ss.getSheetByName(SHEET_TARIFS).getDataRange().getValues().slice(1).forEach(function (r) { tarifs[String(r[2])] = { cat: String(r[0]), art: String(r[1]), prix: Number(r[3]) || 0 }; });
  var n = function (s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' '); };
  var EXCL_INUTILE = /\b(option|cart|sellerie|expo|exposition|supplement|tablette|installation|pedale|bras|support|caisson|meuble|accessoire|ecran|kit regulation|bouchon|crachoir|vanne|recuperateur|imprimante|pied|teti[eè]re|eclairage|scialytique|cas1|pompe a vide|plateau)\b/;
  var regle = function (cat, t) {
    if (/fauteuil|unit/i.test(cat)) {
      if (/belmont/.test(t) && /eurus/.test(t)) { return / nu\b/.test(t) ? 'SE20250626-02' : 'SE003'; }
      if (/belmont/.test(t) && /(tb compass|clesta)/.test(t)) { return 'SE20241002-01'; }
      if (/belmont/.test(t) && /voyager/.test(t)) { return 'SE0174'; }
      if (/belmont/.test(t) && /progres/.test(t)) { return 'SE02'; }
      if (/(quetin|airel)/.test(t)) {
        if (/k2 evo\.?2/.test(t)) { return 'SE250909-2'; } if (/k2 evo/.test(t)) { return 'SE250809-6'; } if (/k2 neo\.?2/.test(t)) { return 'SE250909-1'; } if (/k2 one/.test(t)) { return 'MPP260302-02'; }
        if (/\bk2\b/.test(t)) { return 'SE012'; }
        if (/pe9 evo\.?2/.test(t)) { return 'SE250725-7'; } if (/pe9 evo/.test(t)) { return 'SE250725-6'; } if (/pe9 neo\.?2/.test(t)) { return 'SE250725-5'; } if (/pe9/.test(t)) { return 'SE250725-4'; }
        if (/pe8/.test(t)) { return 'SE260209-1'; }
      }
      if (/miglionico/.test(t)) { if (/(nice one|n1l)/.test(t)) { return 'MPF230322-10'; } if (/\b(ng|nt|nl)\b|flex/.test(t)) { return 'MPF20260505-01'; } }
      if (/\bxo\b/.test(t)) { if (/flow/.test(t)) { return 'SE20260120-05'; } if (/flex/.test(t)) { return 'SE251117-1'; } }
      if (/meunier carus/.test(t)) { if (/table dentaire/.test(t)) { return 'OR20250319-01'; } if (/unit dsa/.test(t)) { return 'SRP0060721'; } }
      if (/sirona/.test(t)) { if (/intego/.test(t)) { return 'MPP230725-1'; } if (/sinius/.test(t)) { return 'SE20250626-01'; } if (/\bc[1-4]\b/.test(t)) { return 'SE0115'; } }
      if (/planmeca|plan meca/.test(t) && /compact/.test(t)) { return 'SR20260623-01'; }
      return '';
    }
    if (/compresseur|aspiration/i.test(cat)) {
      if (/d[u]rr/.test(t)) { if (/duo/.test(t)) { return 'SC20250717-01'; } if (/tornado/.test(t)) { return 'SC20250717-02'; } if (/quattro/.test(t)) { return 'SC20250717-03'; } if (/primo/.test(t)) { return 'SC20251022-01'; } if (/(vsa|vs ?600|vs ?300|aspiration humide|\bvs\b)/.test(t)) { return 'SC20232211-02'; } return ''; }
      if (/cattani/.test(t)) { if (/micro.?smart/.test(t)) { return 'SC20232211-05'; } if (/turbo.?smart/.test(t)) { return 'SE20260121-01'; } if (/compresseur/.test(t)) { return /4 postes/.test(t) ? 'SC20232211-01' : 'SC20250701-01'; } return ''; }
      if (/ekom/.test(t)) { return 'OR20240729-01'; }
      if (/mgf/.test(t)) { if (/prime/.test(t)) { return 'SE20250729-01'; } if (/compresseur/.test(t)) { return 'SE20260120-01'; } return ''; }
      if (/meunier carus/.test(t) && /(aspiration|aspi.?jet|dsa aspi)/.test(t)) { return 'SRP0039179'; }
      return '';
    }
    if (/sterili/i.test(cat)) {
      if (/melag/.test(t)) {
        if (/melatherm ?10|mt ?10/.test(t)) { return 'SE11002220'; } if (/melatherm ?20|mt ?20/.test(t)) { return 'SE111'; } if (/melaseal pro/.test(t)) { return 'SE20260727-01'; }
        if (/\b(40|44) ?b|40\/44/.test(t)) { return 'SE11332'; } if (/\b(41|43) ?b/.test(t)) { return 'SE13322'; } if (/\b(23|31) ?b/.test(t)) { return 'SE11100'; } if (/\b(24|30) ?b/.test(t)) { return 'SE11121'; } if (/\b323\b/.test(t)) { return 'SE20260120-03'; }
        return '';
      }
      if (/nsk/.test(t)) { if (/icare/.test(t)) { return 'SC20250701-02'; } if (/mx ?230/.test(t)) { return 'SE250911-1'; } if (/iclave/.test(t)) { return /24/.test(t) ? 'SE251029' : 'SE13329'; } return ''; }
      if (/euronda/.test(t)) { if (/eurosafe/.test(t)) { return 'SE13358-01'; } if (/\be ?10\b/.test(t)) { return 'SE13348'; } return ''; }
      if (/gamasonic/.test(t)) { if (/snc ?17/.test(t)) { return 'SE20240103-02'; } if (/snc ?30/.test(t)) { return 'SC20232211-03'; } return ''; }
      if (/sirona/.test(t) && /dac/.test(t)) { return 'SE0120'; }
      return '';
    }
    if (/panoramique|cone/i.test(cat)) {
      if (/carestream|\bcs ?\d/.test(t)) { return /ceph|\bc\b/.test(t) ? 'SR20260325-01' : 'SE0012'; }
      if (/vatech/.test(t)) { return 'SE20250626-03'; }
      return '';
    }
    return '';
  };
  // Cabinets de l'Agence reliés à Extrabat : EXB_<cabinet> = n° client Extrabat
  var props = PropertiesService.getScriptProperties().getProperties(), versCab = {};
  Object.keys(props).forEach(function (k) { if (/^EXB_CAB/.test(k) && /^\d+$/.test(String(props[k]))) { versCab[String(props[k])] = k.slice(4); } });
  var cm = ss.getSheetByName(SHEET_MATERIEL), deja = {};
  cm.getDataRange().getValues().slice(1).forEach(function (r) { if (String(r[0])) { deja[String(r[0])] = 1; } });
  var parc = ss.getSheetByName(SHEET_PRATMAT).getDataRange().getValues().slice(1), cabs = {}, sansArticle = {}, vus = 0;
  parc.forEach(function (r) {
    var cle = String(r[0] || ''), prat = String(r[1] || ''), cat = String(r[2] || ''), mod = String(r[3] || ''), t = n(mod), d = dateOrNull_(r[5]);
    if (!/^EXB-\d+$/.test(cle) || /alternative dentaire/i.test(prat)) { return; }
    if (!/fauteuil|unit|compresseur|aspiration|sterili|panoramique|cone/i.test(cat)) { return; }
    vus++;
    var code = regle(cat, t);
    // Accessoire / option (le début du libellé le dit) : pas une révision à part
    var tete = t.split(' ').slice(0, 4).join(' ');
    if (code && code !== 'SE20260727-01' && /\b(option|cart|sellerie|supplement|tablette|installation|bras|support|caisson|meuble|meubles|accessoires?|bouchon|vanne|imprimante|pied|eclairage|scialytique|pompe|ensemble|ecran|bac|soudeuse|boitier)\b/.test(tete)) { return; }
    if (!code || !tarifs[code]) { sansArticle[cat + ' | ' + mod.split(' ').slice(0, 6).join(' ')] = (sansArticle[cat + ' | ' + mod.split(' ').slice(0, 6).join(' ')] || 0) + 1; return; }
    var exId = cle.slice(4), id = versCab[exId] || cle;
    if (deja[id]) { return; }
    var c = cabs[id] || (cabs[id] = { nom: prat, items: [], last: null });
    c.items.push({ code: code, serie: String(r[4] || '').replace(/^'/, ''), mod: mod, d: d });
    if (d && (!c.last || d > c.last)) { c.last = d; }
  });
  var lignes = [], nbCab = 0, ventil = {};
  Object.keys(cabs).forEach(function (id) {
    var c = cabs[id]; if (!c.items.length) { return; }
    var base = c.last || auj, A = new Date(auj.getFullYear(), base.getMonth(), base.getDate());
    while (A.getTime() <= auj.getTime()) { A = new Date(A.getFullYear() + 1, A.getMonth(), A.getDate()); }
    if ((A.getTime() - auj.getTime()) / 864e5 < 330) { A = new Date(A.getFullYear() + 1, A.getMonth(), A.getDate()); }
    var debut = new Date(A.getFullYear() - 1, A.getMonth(), A.getDate()), mois = fmt_(A, 'yyyy-MM'); ventil[mois] = (ventil[mois] || 0) + 1; nbCab++;
    c.items.forEach(function (x) {
      var tf = tarifs[x.code];
      lignes.push([id, c.nom, tf.cat, tf.art, x.code, x.serie ? "'" + x.serie : '', tf.prix, debut, '', '', A, 'RAPPEL', '',
        'Pas de contrat : révision à proposer. Appareil : ' + cut_(x.mod, 90) + (x.d ? ' (acheté le ' + fmt_(x.d, 'dd/MM/yyyy') + ')' : '') + '. Date de base pour la révision du ' + fmt_(A, 'dd/MM/yyyy') + '.', '']);
    });
  });
  if (lignes.length) { cm.getRange(cm.getLastRow() + 1, 1, lignes.length, MATERIEL_HEADERS.length).setValues(lignes); }
  Logger.log(vus + ' appareils examinés · ' + lignes.length + ' lignes créées pour ' + nbCab + ' cabinets · révisions par mois : ' + JSON.stringify(ventil));
  Logger.log('Sans article (' + Object.keys(sansArticle).length + ') :\n' + Object.keys(sansArticle).sort().map(function (k) { return sansArticle[k] + ' ' + k; }).join('\n'));
}
// FIN TEMP ext3
