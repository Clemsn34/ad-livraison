// Contrat d'entretien à proposer UN AN avant la révision (prélèvement mensuel qui finance la révision) :
// appareils « RAPPEL » (sans contrat) dont la révision tombe dans les 12 mois → par cabinet, une fois par échéance :
// devis Extrabat (révisions avec leur liste d'opérations + UN déplacement selon la zone) et tâche pour Clément.
var AGENCE_ADRESSE = '10 rue de la Prade, 34880 Lavérune, France';
// Zones de déplacement (distance par la route depuis l'agence, aller simple) → article Extrabat (prix HT)
var ZONES_DEPLACEMENT = [
  { max: 25, code: 'FM1', lib: 'Forfait déplacement en zone 1 (Montpellier et périphérie)', ht: 48.3333 },
  { max: 70, code: 'Forfait-Z2', lib: 'Forfait déplacement en zone 2 (de Nîmes à Béziers)', ht: 72.5 },
  { max: 100, code: 'Forfait-Z3', lib: 'Forfait déplacement en zone 3 (de Narbonne à Alès)', ht: 104.1667 },
  { max: 99999, code: 'Forfait-HZ', lib: 'Forfait déplacement supérieur à 200 km aller-retour', ht: 133.3333 }];
// Adresse du cabinet (Extrabat) → distance par la route → zone. Gardé en mémoire par cabinet.
function zoneDeplacement_(exId) {
  var props = PropertiesService.getScriptProperties(), k = 'ZONE_' + exId, deja = props.getProperty(k);
  if (deja) { try { return JSON.parse(deja); } catch (e) {} }
  var c = exbCall_('get', '/v3/client/' + encodeURIComponent(exId), { include: 'telephone,adresse' }), adr = (exbResume_(c).adresse || '').replace(/\n/g, ', ');
  if (!adr) { return null; }
  var km = null;
  try { var d = Maps.newDirectionFinder().setOrigin(AGENCE_ADRESSE).setDestination(adr + ', France').setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections(); km = Math.round(d.routes[0].legs[0].distance.value / 1000); } catch (e) { return null; }
  var z = ZONES_DEPLACEMENT.filter(function (x) { return km <= x.max; })[0], o = { km: km, code: z.code, lib: z.lib, ht: z.ht, adresse: adr };
  props.setProperty(k, JSON.stringify(o));
  return o;
}
// Description Extrabat des articles de révision (liste des opérations / pièces), colonne F de l'onglet Tarifs entretien
function revDescriptions_() {
  var sh = getSpreadsheet_().getSheetByName(SHEET_TARIFS), o = {};
  if (!sh || sh.getLastColumn() < 6) { return o; }
  sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(function (r) { if (r[2] && r[5]) { o[String(r[2])] = String(r[5]); } });
  return o;
}
function contratsAProposer_() {
  var sh = getSpreadsheet_().getSheetByName(SHEET_MATERIEL); if (!sh) { return 0; }
  var rows = sh.getDataRange().getValues(), today = today_(), par = {}, props = PropertiesService.getScriptProperties();
  for (var i = 1; i < rows.length; i++) {
    if (!String(rows[i][CM.id]).trim()) { continue; }
    var it = materielItem_(rows[i], i + 1, today);
    if (!it.rappel || !it._prochaine) { continue; }
    var j = Math.round((it._prochaine.getTime() - today.getTime()) / 86400000);
    if (j > 366 || j < 0) { continue; }
    var id = String(rows[i][CM.id]), p = par[id] || (par[id] = { cab: String(rows[i][CM.cab] || id), items: [], prochaine: it._prochaine });
    p.items.push(it); if (it._prochaine.getTime() < p.prochaine.getTime()) { p.prochaine = it._prochaine; }
  }
  var gerant = membre_(COMPTE_SERVEUR, 'Clément'), n = 0, descs = null, t0 = Date.now();
  var r2 = function (v) { return Math.round(v * 100) / 100; }, eur = function (v) { return String(r2(v)).replace('.', ',') + ' €'; };
  Object.keys(par).forEach(function (id) {
    if (Date.now() - t0 > 300000) { return; } // le reste passera la nuit suivante
    var p = par[id], k = 'CTRPROP_' + id + '_' + fmt_(p.prochaine, 'yyyyMM');
    if (props.getProperty(k)) { return; }
    if (!descs) { descs = revDescriptions_(); }
    var exId = exbLien_(id) || (/^EXB-\d+$/.test(id) ? id.slice(4) : '');
    var zone = null; if (exId) { try { zone = zoneDeplacement_(exId); } catch (e) {} }
    var totalRev = p.items.reduce(function (s, x) { return s + (Number(x.prix) || 0); }, 0), dep = zone ? r2(zone.ht * 1.2) : 0;
    var total = totalRev + dep, mens = Math.ceil(total / 12 * 100) / 100, codeDv = 'CTR-' + id + '-' + fmt_(p.prochaine, 'yyyyMM'), dvInfo = '';
    try {
      if (!exId) { dvInfo = 'Devis non créé : cabinet pas encore relié à Extrabat (Clients › fiche › Voir dans Extrabat).'; }
      else {
        var el = [];
        p.items.forEach(function (x) {
          var ttc = r2(Number(x.prix) || 0), ht = r2(ttc / 1.2), ops = descs[x.code] || '';
          el.push({ code: x.code || '', description: '<mep style="font-weight: bold; color: navy;">' + (x.modele || x.appareil) + '</mep>' +
              ([x.repere, x.serie ? 'N° de série ' + x.serie : ''].filter(String).length ? '<br>' + [x.repere, x.serie ? 'N° de série ' + x.serie : ''].filter(String).join(' · ') : '') + (ops ? '<br>' + ops : ''),
            quantite: 1, puht: ht, totalHt: ht, totalEscompte: 0, totalNet: ht, tauxTva: 20, totalTva: r2(ttc - ht), totalTtc: ttc, ordre: el.length + 1, tenueStock: false, prixMini: 0, nomenclature: false, poids: 0, prixAchat: 0, quantiteConditionnement: 1, tauxRemise: 0, numLigne: el.length + 1 });
        });
        if (zone) {
          var dht = r2(zone.ht);
          el.push({ code: zone.code, description: '<mep style="font-weight: bold; color: navy;">' + zone.lib + '</mep><br>Un seul déplacement pour toutes les révisions (' + zone.km + ' km de l\'agence)',
            quantite: 1, puht: dht, totalHt: dht, totalEscompte: 0, totalNet: dht, tauxTva: 20, totalTva: r2(dep - dht), totalTtc: dep, ordre: el.length + 1, tenueStock: false, prixMini: 0, nomenclature: false, poids: 0, prixAchat: 0, quantiteConditionnement: 1, tauxRemise: 0, numLigne: el.length + 1 });
        }
        var dv = { extrabatFrsId: exbShop_(), type: 1, code: codeDv, date: fmt_(today, 'yyyy-MM-dd'),
          titre: 'Contrat d\'entretien annuel – révision du ' + fmt_(p.prochaine, 'dd/MM/yyyy') + ' – ' + eur(mens) + ' TTC / mois',
          // Totaux à 0 : Extrabat les recalcule en additionnant les lignes (sinon ils sont comptés deux fois)
          totalHT: 0, totalTTC: 0, totalTVA: 0, escompte: 0, portHT: 0, portTTC: 0, portTVA: 0, blFrsNum: '', blFrsDate: null, adresseFacturation: '', adresseLivraison: '', client: Number(exId),
          commentaire: 'Contrat d\'entretien annuel : révision complète de tout le matériel ci-dessus, prévue le ' + fmt_(p.prochaine, 'dd/MM/yyyy') + ', déplacement compris. Règlement par prélèvement mensuel de ' + eur(mens) + ' TTC pendant 12 mois. Pour valider : nous retourner le devis signé, le RIB du cabinet, l\'autorisation de prélèvement (mandat SEPA) signée et les conditions générales du contrat d\'entretien signées.',
          lignes: el.map(exbLigneNette_) };
        var exist = null; try { exist = exbListe_(exbCall_('get', '/v1/pieces', { clients: exId, types: 'devis', date_debut: fmt_(new Date(today.getTime() - 400 * 864e5), 'yyyy-MM-dd'), date_fin: fmt_(new Date(today.getTime() + 30 * 864e5), 'yyyy-MM-dd'), nbitem: 50, page: 1 })).filter(function (c) { return String(c.code) === codeDv; })[0]; } catch (e3) {}
        var rd = exist || exbCall_('post', '/v1/client/' + encodeURIComponent(exId) + '/devis', {}, dv);
        dvInfo = 'Devis Extrabat ' + codeDv + ' ' + (exist ? 'déjà existant' : 'créé (brouillon)') + '.' + (zone ? '' : ' Adresse du cabinet introuvable : ajouter le déplacement à la main.');
        try { log_('Extrabat', p.cab, id, '', 'Claude', '', 'DEVIS CONTRAT', codeDv + ' · ' + eur(total) + ' TTC / an'); } catch (e2) {}
      }
    } catch (e) { dvInfo = 'Devis Extrabat non créé : ' + e.message; }
    var detail = dvInfo + '\nEnvoyer : devis, mandat SEPA, conditions générales. Récupérer : devis signé, RIB, mandat SEPA signé, CG signées. Puis saisir RIB et mandat à la banque, lancer les prélèvements, passer les appareils de RAPPEL à OUI (Contrats matériel).' +
      '\nRévision prévue le ' + fmt_(p.prochaine, 'dd/MM/yyyy') + ' :\n' +
      p.items.map(function (x) { return '- ' + String(x.modele || x.appareil).replace(/R[ée]vision annuelle\s*/i, '') + (x.repere ? ' (' + x.repere + ')' : '') + (x.prix ? ' : ' + eur(x.prix) : ''); }).join('\n') +
      (zone ? '\n- Déplacement zone ' + zone.code + ' (' + zone.km + ' km) : ' + eur(dep) : '') +
      '\nTotal : ' + eur(total) + ' TTC par an, soit ' + eur(mens) + ' TTC par mois.';
    var t = null; try { t = tacheAjouter_(gerant, { titre: 'Proposer le contrat d\'entretien (' + eur(mens) + ' TTC / mois)', cabinet: p.cab, detail: detail, pour: [gerant.prenom] }); } catch (e) {}
    props.setProperty(k, (t && t.n) || fmt_(today, 'yyyy-MM-dd')); n++;
  });
  return n;
}
