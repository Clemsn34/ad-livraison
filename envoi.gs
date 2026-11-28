
// =====================================================================
// ENVOI DU DEVIS AU PRATICIEN DEPUIS L'APPLI (28/11)
// PDF du devis aux couleurs AD (mêmes lignes et prix que le brouillon Extrabat), proposition de financement,
// documentation du fauteuil et de la panoramique (brochures « … – Brochure » du Drive), message rédigé dans l'appli.
// Copie à l'expéditeur avec le nom du praticien et la version (v1, v2… à chaque renvoi d'un devis modifié).
// =====================================================================
function projetEnvoyerDevis_(m, c) {
  var cat = catalogueProjet_(), r2 = function (v) { return Math.round(v * 100) / 100; }, tot = 0, sansPrix = 0, lignesH = '';
  var props = PropertiesService.getScriptProperties(), kv = 'DEVV_' + c.numero, version = Number(props.getProperty(kv) || 0) + 1;
  props.setProperty(kv, String(version));
  var vtxt = 'v' + version, titreDoc = 'Devis ' + c.numero + ' ' + vtxt;
  var cell = 'padding:6px 8px;border-bottom:1px solid #E3E8F0;font-size:11.5px;vertical-align:top';
  c.groupes.forEach(function (G) {
    lignesH += '<tr><td colspan="4" style="padding:12px 8px 4px;font-weight:800;color:#0F2347;font-size:12.5px;border-bottom:2px solid #E3A72F">' + esc_(G.g) + '</td></tr>';
    G.l.forEach(function (l) {
      var p = prixCatalogue_(cat, l) || {}, u = p.prix || 0, t = l.offert ? 0 : r2(u * l.qte);
      if (!u && !l.offert) { sansPrix++; }
      tot += t;
      var lib = l.equip + (l.marque ? ' – ' + l.marque : '') + (l.modele ? ' ' + l.modele.replace(/ · /, ' – ').replace(/ › /g, ' ') : '') + (l.note ? ' (' + l.note + ')' : '');
      lignesH += '<tr><td style="' + cell + '"><b>' + esc_(lib) + '</b></td><td style="' + cell + ';text-align:center">' + l.qte + '</td><td style="' + cell + ';text-align:right;white-space:nowrap">' + (u ? euros_(u) : '—') + '</td><td style="' + cell + ';text-align:right;white-space:nowrap">' +
        (l.offert ? '<b style="color:#1E8449">OFFERT</b>' : u ? euros_(t) : 'Sur devis') + '</td></tr>';
      l.options.forEach(function (o) {
        var po = prixOption_(cat, l, o), uo = po.prix || 0, to = l.offert || po.offert ? 0 : r2(uo * l.qte); tot += to;
        lignesH += '<tr><td style="' + cell + ';padding-left:22px;color:#374151">+ ' + esc_(libOpt_(o)) + '</td><td style="' + cell + ';text-align:center">' + l.qte + '</td><td style="' + cell + ';text-align:right;white-space:nowrap">' + (uo ? euros_(uo) : '') + '</td><td style="' + cell + ';text-align:right;white-space:nowrap">' +
          (l.offert || po.offert ? '<b style="color:#1E8449">OFFERT</b>' : uo ? euros_(to) : 'Inclus') + '</td></tr>';
      });
    });
  });
  tot = r2(tot);
  var logo = 'data:image/png;base64,' + LOGO_PNG_B64, praticien = (c.pr.prenom ? c.pr.prenom + ' ' : '') + c.nom;
  var html = '<html><head><meta charset="utf-8"><style>body{font-family:Arial,Helvetica,sans-serif;color:#1F2937;margin:28px}</style></head><body>' +
    '<table style="width:100%;border-collapse:collapse"><tr><td style="width:80px"><img src="' + logo + '" width="70" height="70"></td>' +
    '<td><div style="font-size:20px;font-weight:800;color:#0F2347">ALTERNATIVE DENTAIRE</div><div style="font-size:11px;color:#5A6880">10 rue de la Prade, 34880 Lavérune · ' + AD_PHONE + ' · contact@alternativedentaire.fr</div></td>' +
    '<td style="text-align:right"><div style="font-size:18px;font-weight:800;color:#0F2347">DEVIS</div><div style="font-size:12px">N° ' + esc_(c.numero) + ' – ' + vtxt + '</div><div style="font-size:12px">Le ' + fmt_(new Date(), 'dd/MM/yyyy') + '</div></td></tr></table>' +
    '<div style="height:3px;background:#E3A72F;margin:12px 0 14px"></div>' +
    '<table style="width:100%"><tr><td style="font-size:12px;color:#5A6880">Établi par ' + esc_(m.prenom + (m.nom ? ' ' + m.nom : '')) + '</td><td style="text-align:right;font-size:13px"><b>' + esc_(praticien) + '</b>' + (c.ville ? '<br>' + esc_(c.ville) : '') + (c.email ? '<br>' + esc_(c.email) : '') + '</td></tr></table>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:14px"><tr style="background:#0F2347;color:#fff"><th style="text-align:left;padding:7px 8px;font-size:11.5px">Désignation</th><th style="padding:7px 8px;font-size:11.5px">Qté</th><th style="text-align:right;padding:7px 8px;font-size:11.5px">PU TTC</th><th style="text-align:right;padding:7px 8px;font-size:11.5px">Total TTC</th></tr>' + lignesH + '</table>' +
    '<table style="width:100%;margin-top:14px"><tr><td></td><td style="width:260px;background:#F4F6FA;border:2px solid #0F2347;border-radius:6px;padding:10px 12px;text-align:right">' +
    '<div style="font-size:11px;color:#5A6880">Total HT : ' + euros_(r2(tot / 1.2)) + ' · TVA 20 % : ' + euros_(r2(tot - tot / 1.2)) + '</div><div style="font-size:18px;font-weight:800;color:#0F2347">' + euros_(tot) + ' TTC</div>' +
    (sansPrix ? '<div style="font-size:10.5px;color:#8A5A00">+ ' + sansPrix + ' ligne(s) chiffrée(s) sur devis</div>' : '') + '</td></tr></table>' +
    (c.note ? '<p style="font-size:12px;margin-top:14px"><b>Remarques :</b> ' + esc_(c.note).replace(/\n/g, '<br>') + '</p>' : '') +
    '<p style="font-size:11px;color:#5A6880;margin-top:16px">Devis valable 1 mois. Installation, mise en service et formation à l\'utilisation comprises sauf mention contraire. Proposition de financement jointe le cas échéant.</p>' +
    '<div style="margin-top:18px;border-top:1px solid #DCE3EE;padding-top:8px;font-size:11px;color:#0F2347"><b>Dépannage en moins de 3 h</b> pour toute panne bloquante · <b>Cabinet clé en main</b>, un seul interlocuteur du plan à la mise en service · <b>Plus de 24 ans d\'expérience</b> dans le secteur dentaire</div>' +
    '<div style="margin-top:18px;font-size:12px">Bon pour accord (date, signature et cachet) :</div></body></html>';
  var nomFich = titreDoc + ' - ' + c.nom.replace(/[\\\/:*?"<>|]/g, ' ').trim();
  var pdf = Utilities.newBlob(html, 'text/html', 'devis.html').getAs('application/pdf').setName(nomFich + '.pdf');
  var pj = [pdf];
  if (c.finPdf) { pj.push(c.finPdf); }
  // Documentation : brochure du fauteuil et de la panoramique / du cône beam (fichiers « … – Brochure » du Drive)
  var docs = [];
  if (c.envoi.docs) {
    var vus = {};
    c.groupes.forEach(function (G) { G.l.forEach(function (l) {
      if (!/fauteuil|panoramique|c[ôo]ne beam/i.test(l.equip) || !l.marque) { return; }
      var mods = String(l.modele || '').split(/ · /)[0].split(/ › | – |\s+\(|\s+/).map(function (x) { return x.replace(/[^\wÀ-ÿ.\-]/g, ''); })
        .filter(function (x) { return x.length >= 2 && !/^(avec|sans|pour|mural|contrat|panoramique|standard|de|et|la|le|ans?)$/i.test(x); });
      var essais = [], m1 = mods.slice(0, 3);
      if (/panoramique|c[ôo]ne/i.test(l.equip)) { var cs = /CS\s?(\d{4})/i.exec(l.modele || ''); if (cs) { essais.push('CS' + cs[1]); essais.push('CS ' + cs[1]); } }
      m1.forEach(function (x) { essais.push(x); });
      for (var i = 0; i < essais.length; i++) {
        var q = essais[i]; if (vus[q]) { break; }
        var it = DriveApp.searchFiles("title contains '" + q.replace(/'/g, "\\'") + "' and title contains 'Brochure' and mimeType = 'application/pdf' and trashed = false"), f = null;
        while (it.hasNext()) { var x = it.next(); if (x.getSize() < 9000000 && !/ – (EN|DE|IT|ES) /.test(x.getName())) { f = x; break; } }
        if (f) { vus[q] = 1; if (!docs.some(function (d) { return d.getId() === f.getId(); })) { docs.push(f); } break; }
      }
    }); });
    docs.slice(0, 4).forEach(function (f) { pj.push(f.getBlob()); });
  }
  var tu = !!c.envoi.tu, msg = cut_(String(c.envoi.message || '').trim(), 5000);
  if (msg.length < 20) { throw new Error('Message trop court.'); }
  var corpsHtml = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1F2937;line-height:1.6;max-width:640px">' +
    esc_(msg).split(/\n{2,}/).map(function (p) { return '<p>' + p.replace(/\n/g, '<br>') + '</p>'; }).join('') +
    '<p style="margin-top:18px">' + esc_(m.prenom + (m.nom ? ' ' + m.nom : '')) + '<br><b style="color:#0F2347">Alternative Dentaire</b><br>' + AD_PHONE + ' · contact@alternativedentaire.fr</p>' +
    '<p style="margin-top:14px;color:#6B7280;font-size:12px">Dépannage en moins de 3 h pour toute panne bloquante · Cabinet clé en main, un seul interlocuteur · Plus de 24 ans d\'expérience<br>' +
    'Showroom : 10 rue de la Prade, 34880 Lavérune (5 min de Montpellier) · du lundi au vendredi, 9 h – 17 h</p></div>';
  var sujet = cut_(String(c.envoi.sujet || '').trim(), 150) || ((tu ? 'Ton devis' : 'Votre devis') + ' Alternative Dentaire');
  var opts = { htmlBody: corpsHtml, name: 'Alternative Dentaire', replyTo: 'contact@alternativedentaire.fr', attachments: pj };
  try { if (GmailApp.getAliases().indexOf('contact@alternativedentaire.fr') >= 0) { opts.from = 'contact@alternativedentaire.fr'; } } catch (e) {}
  var envoyer = function (to, s, o) { if (o.from) { GmailApp.sendEmail(to, s, msg, o); } else { MailApp.sendEmail(to, s, msg, o); } };
  envoyer(c.email, sujet, opts);
  // Copie à l'expéditeur : nom du praticien et version dans l'objet
  var copie = String(m.email || '').trim() || COMPTE_SERVEUR;
  var opts2 = { htmlBody: '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;background:#FFF8E8;border:1px solid #E3A72F;border-radius:8px;padding:8px 12px;margin-bottom:12px">Copie du mail envoyé à <b>' + esc_(praticien) + '</b> (' + esc_(c.email) + ') · devis <b>' + esc_(c.numero) + ' ' + vtxt + '</b> · ' + euros_(tot) + ' TTC</div>' + corpsHtml,
    name: 'Alternative Dentaire', attachments: pj };
  if (opts.from) { opts2.from = opts.from; }
  try { envoyer(copie, '[Copie] Devis ' + c.numero + ' ' + vtxt + ' – ' + praticien, opts2); } catch (e) {}
  return { version: vtxt, total: tot, pj: pj.map(function (b) { return b.getName(); }), docs: docs.length };
}
