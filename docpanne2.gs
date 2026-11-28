  if (op === 'doc_panne') {
    // 1) Recherche plein texte Google Drive : le code (toutes ses écritures) ET la marque / le modèle DANS le texte des manuels
    // 2) À défaut, le code dans les dossiers dont le nom correspond au modèle (ancienne méthode)
    var code = String(req.code || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 8);
    if (!code) { return { ok: false, message: 'Code d\'erreur manquant.' }; }
    var STOP = /^(le|la|les|de|du|des|un|une|sur|mon|ma|mes|avec|et|a|au|aux|en|pour|il|y|est|qui|fait|affiche|marque|modele|appareil|machine|dentaire|cabinet|chez|the|of)$/;
    var qm = docNorm_(req.q), toks = qm.split(' ').filter(function (w) { return w.length >= 1 && !STOP.test(w); });
    // « vs 600 » -> « vs600 » (modèle), mots de 2 lettres et plus
    var mots = []; for (var ti = 0; ti < toks.length; ti++) { var tk = toks[ti]; if (/^[a-z]{1,3}$/.test(tk) && /^\d/.test(toks[ti + 1] || '')) { mots.push(tk + toks[ti + 1]); ti++; } else if (tk.length >= 2) { mots.push(tk); } }
    mots = mots.slice(0, 5);
    var marques = mots.filter(function (w) { return /^[a-z]{3,}$/.test(w); });
    var lettres = code.replace(/\d.*$/, ''), num = code.replace(/^[A-Z]*/, ''), n0 = String(Number(num.replace(/[A-Z]$/, '')) || num);
    var variantes = [];
    if (lettres) {
      variantes = [code, lettres + ' ' + num, lettres + '-' + num, lettres + '.' + num];
      if (n0 !== num) { variantes.push(lettres + n0, lettres + ' ' + n0); } else if (n0.length === 1) { variantes.push(lettres + '0' + n0, lettres + '00' + n0); } else if (n0.length === 2) { variantes.push(lettres + '0' + n0); }
    } else {
      ['Error', 'Erreur', 'Err', 'E', 'Code', 'Alarm', 'Alarme', 'Défaut', 'Fault', 'Fehler', 'Störung', 'Malfunction', 'Errore', 'Allarme', 'Panne', 'Message'].forEach(function (p) { variantes.push(p + ' ' + n0); });
      variantes.push('E' + n0, 'E0' + n0, 'Err' + n0, 'A' + n0);
    }
    var vq = '(' + variantes.map(function (v) { return 'fullText contains \'"' + v.replace(/['"\\]/g, '') + '"\''; }).join(' or ') + ')';
    var trouves = {}, nb = 0;
    var tente = function (q, max) {
      try {
        var it = DriveApp.searchFiles(q + ' and mimeType = \'application/pdf\' and trashed = false'), vus = 0;
        while (it.hasNext() && vus < max) {
          var f = it.next(); vus++;
          var pid = '', pp = f.getParents(); if (pp.hasNext()) { pid = pp.next().getId(); }
          if (!parId[pid]) { continue; } // hors de la documentation
          var cle = f.getName().toLowerCase() + '|' + f.getSize(); if (trouves[cle]) { continue; }
          trouves[cle] = { id: f.getId(), nom: f.getName(), chemin: parId[pid][2], taille: f.getSize() }; nb++;
        }
      } catch (e) {}
    };
    // Mot cherché dans le texte, avec ses variantes (« durr » = « dürr », « vs600 » = « vs 600 »)
    var mq = function (w) {
      var v = [w], m2 = /^([a-z]+)(\d+)$/.exec(w);
      if (m2) { v.push(m2[1] + ' ' + m2[2], m2[1] + '-' + m2[2]); }
      if (/ue|u/.test(w) && w === 'durr') { v.push('dürr', 'duerr'); }
      if (w === 'kavo' || w === 'melag' || w === 'cattani') { v = [w]; }
      return '(' + v.map(function (x) { return 'fullText contains \'"' + x.replace(/['"\\]/g, '') + '"\''; }).join(' or ') + ')';
    };
    if (mots.length) { tente(vq + ' and ' + mots.map(mq).join(' and '), 80); }
    if (nb < 4 && marques.length) { tente(vq + ' and ' + marques.map(mq).join(' and '), 80); }
    if (nb < 4 && mots.length > 1) { tente(vq + ' and (' + mots.map(mq).join(' or ') + ')', 80); }
    var bases = [];
    if (nb < 4 && qm.length >= 2) {
      var cand = [], comp2 = mots.join('');
      idx.forEach(function (r) {
        var nn = docNorm_(r[1]), ch = docNorm_(r[2]);
        if (!mots.some(function (w) { return ch.indexOf(w) >= 0; })) { return; }
        var sc = 0; mots.forEach(function (w) { if (nn.indexOf(w) >= 0) { sc += 10; } else if (ch.indexOf(w) >= 0) { sc += 4; } }); if (comp2 && nn.replace(/ /g, '').indexOf(comp2) >= 0) { sc += 15; }
        if (/^(manuel|plan|bulletin|information|garantie|notice|technique|client|photo)/.test(nn)) { sc -= 8; }
        cand.push({ r: r, sc: sc - r[4] * 0.5 });
      });
      cand.sort(function (a, b) { return b.sc - a.sc; });
      bases = cand.slice(0, 3).map(function (x) { return x.r; });
      var ids = {}, file = bases.map(function (r) { return r[0]; });
      while (file.length && Object.keys(ids).length < 300) { var x = file.shift(); if (ids[x]) { continue; } ids[x] = 1; idx.forEach(function (r) { if (r[3] === x) { file.push(r[0]); } }); }
      var liste = Object.keys(ids);
      for (var b = 0; b < liste.length && nb < 12; b += 15) {
        tente(vq + ' and (' + liste.slice(b, b + 15).map(function (fid) { return '\'' + fid + '\' in parents'; }).join(' or ') + ')', 30);
      }
    }
    var fl = Object.keys(trouves).map(function (k) { return trouves[k]; });
    fl.forEach(function (f) {
      var nm = docNorm_(f.nom), ch = docNorm_(f.chemin), sc = 0;
      var nmc = nm.replace(/ /g, ''), chc = ch.replace(/ /g, ''), vu = 0;
      mots.forEach(function (w) { if (nmc.indexOf(w) >= 0) { sc += 10; vu++; } else if (chc.indexOf(w) >= 0) { sc += 6; vu++; } });
      if (marques.length && marques.some(function (w) { return chc.indexOf(w) >= 0 || nmc.indexOf(w) >= 0; })) { sc += 12; } else if (marques.length) { sc -= 15; } // bonne marque dans le chemin
      if (/code|erreur|error|errore|fehler|storung|alarm|fault|defaut|trouble|diagnos/.test(nm)) { sc += 15; } else if (/erreur|error|panne|trouble|diagnos|service|techni|maintenance|repar|repair|depannage/.test(nm + ' ' + ch)) { sc += 6; }
      if (/catalogue|tarif|brochure|prix|offre|promo|presentation|plaquette/.test(nm)) { sc -= 20; }
      if (f.taille > 25 * 1048576) { sc -= 30; }
      f.sc = sc;
    });
    fl.sort(function (a, b) { return b.sc - a.sc; });
    var dossiers = bases.length ? bases : [];
    return { ok: true, fichiers: fl.filter(function (f) { return f.taille < 25 * 1048576; }).slice(0, 6).map(function (f) { return { id: f.id, nom: f.nom, chemin: f.chemin, taille: f.taille }; }),
      dossiers: dossiers.map(function (r) { return { id: r[0], nom: r[1], chemin: r[2] }; }), variantes: variantes.slice(0, 6) };
  }
