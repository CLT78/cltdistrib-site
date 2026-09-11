/**
 * CLTDistrib — Prospection assistée
 * ---------------------------------------------------------------------------
 * Ce code tourne dans votre compte Google, pas sur le site public. C'est ce
 * qui lui donne le droit d'envoyer depuis votre Gmail, de lire vos réponses,
 * d'interroger une IA et de consulter les sites des entreprises.
 *
 * Installation : voir la page « Application » du site.
 * ---------------------------------------------------------------------------
 */

/* ====================== À COMPLÉTER (une seule fois) ====================== */

var MOI = {
  nom: "Clément Pires Valente",
  fonction: "Fondateur",
  entreprise: "CLTDistrib",
  email: "clementpiresvalente@gmail.com",
  tel: "",                       // votre numéro, ex. "06 12 34 56 78"
  zone: "",                      // ex. "Mantes-la-Jolie et 30 km autour"
  site: "https://clt78.github.io/cltdistrib-site/",
  logo: "https://clt78.github.io/cltdistrib-site/images/signature-logo.png"
};

var REGLAGES = {
  joursAvantRelance: 7,          // délai avant bascule en « à relancer »
  relanceAuto: false,            // true = la relance part sans validation
  maxEnvoisParJour: 40,          // garde-fou (Gmail gratuit : ~100/jour)
  heureTournee: 8,               // heure du contrôle quotidien
  fuseau: "Europe/Paris"
};

/* Ce que vous vendez : sert à l'IA pour écrire des mails justes.
   Ne mettez ici que des choses vraies. */
var OFFRE = {
  resume: "CLTDistrib installe et exploite des distributeurs de boissons et de snacks " +
          "chez des entreprises et des établissements qui disposent d'un emplacement.",
  pourLeClient: [
    "Aucun investissement : la machine est fournie, installée et entretenue par CLTDistrib.",
    "Aucune gestion : réassort, maintenance et suivi des stocks sont assurés par CLTDistrib.",
    "Le client met à disposition un emplacement et une alimentation électrique.",
    "Étude de l'emplacement gratuite et sans engagement."
  ],
  interdits: "Ne jamais inventer de chiffres, de références clients, de témoignages, " +
             "de certifications, de nombre de machines installées ni de durée d'existence."
};

/* ============================ CONSTANTES ============================ */

var CLE_FEUILLE = "ID_FEUILLE";
var CLE_GEMINI  = "CLE_GEMINI";
var CLE_MODELE  = "MODELE_GEMINI";

var COLONNES = ["id", "entreprise", "ville", "activite", "type", "siren", "effectif",
  "site", "email", "emailSource", "tel", "contactNom", "contactFonction", "statut",
  "dateEnvoi", "dateRelance", "dateReponse", "objet", "corps", "note", "pourquoi", "historique"];

var STATUTS = {
  a_envoyer:  "À envoyer",
  envoye:     "Mail envoyé",
  a_relancer: "À relancer",
  relance:    "Relancé",
  reponse:    "Réponse reçue",
  rdv:        "Rendez-vous",
  gagne:      "Emplacement signé",
  refus:      "Sans suite"
};

var API_ENTREPRISES = "https://recherche-entreprises.api.gouv.fr/search";

/* ============================ POINT D'ENTRÉE ============================ */

function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("CLTDistrib — Prospection")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(nom) {
  return HtmlService.createHtmlOutputFromFile(nom).getContent();
}

/* ============================ STOCKAGE ============================ */

function proprietes() {
  return PropertiesService.getScriptProperties();
}

function feuille() {
  var props = proprietes();
  var id = props.getProperty(CLE_FEUILLE);
  var classeur = null;

  if (id) {
    try { classeur = SpreadsheetApp.openById(id); } catch (e) { classeur = null; }
  }
  if (!classeur) {
    classeur = SpreadsheetApp.create("CLTDistrib — Prospection");
    props.setProperty(CLE_FEUILLE, classeur.getId());
  }

  var f = classeur.getSheetByName("Prospects");
  if (!f) {
    f = classeur.getSheets()[0];
    f.setName("Prospects");
  }
  if (f.getLastRow() === 0) {
    f.getRange(1, 1, 1, COLONNES.length).setValues([COLONNES]).setFontWeight("bold");
    f.setFrozenRows(1);
  }
  return f;
}

function lireTout() {
  var f = feuille();
  if (f.getLastRow() < 2) return [];
  var valeurs = f.getRange(2, 1, f.getLastRow() - 1, COLONNES.length).getValues();
  return valeurs.map(function (ligne, i) {
    var o = { _ligne: i + 2 };
    COLONNES.forEach(function (c, j) {
      var v = ligne[j];
      o[c] = (v instanceof Date) ? Utilities.formatDate(v, REGLAGES.fuseau, "yyyy-MM-dd") : String(v == null ? "" : v);
    });
    return o;
  }).filter(function (o) { return o.id; });
}

function ecrireLigne(o) {
  var f = feuille();
  var ligne = COLONNES.map(function (c) { return o[c] == null ? "" : o[c]; });
  if (o._ligne) {
    f.getRange(o._ligne, 1, 1, COLONNES.length).setValues([ligne]);
  } else {
    f.appendRow(ligne);
    o._ligne = f.getLastRow();
  }
  return o;
}

function trouverParId(id) {
  var tous = lireTout();
  for (var i = 0; i < tous.length; i++) if (tous[i].id === id) return tous[i];
  return null;
}

function nouvelId() {
  return "p" + new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 6);
}

function ajouterHistorique(o, texte) {
  var h = [];
  try { h = o.historique ? JSON.parse(o.historique) : []; } catch (e) { h = []; }
  h.push({ date: aujourdhui(), texte: texte });
  o.historique = JSON.stringify(h.slice(-40));
}

function aujourdhui() {
  return Utilities.formatDate(new Date(), REGLAGES.fuseau, "yyyy-MM-dd");
}

function joursDepuis(iso) {
  if (!iso) return null;
  var d = new Date(iso + "T00:00:00");
  if (isNaN(d.getTime())) return null;
  return Math.floor((new Date(aujourdhui() + "T00:00:00").getTime() - d.getTime()) / 86400000);
}

/* ============================ API POUR L'INTERFACE ============================ */

function etat() {
  var tous = lireTout();
  return {
    moi: MOI,
    reglages: REGLAGES,
    statuts: STATUTS,
    cleIA: !!proprietes().getProperty(CLE_GEMINI),
    manquants: champsManquants(),
    prospects: tous.map(nettoyerPourInterface),
    urlFeuille: urlFeuille()
  };
}

function champsManquants() {
  var m = [];
  if (!MOI.tel) m.push("votre numéro de téléphone");
  if (!MOI.zone) m.push("votre zone d'intervention");
  return m;
}

function urlFeuille() {
  var id = proprietes().getProperty(CLE_FEUILLE);
  return id ? "https://docs.google.com/spreadsheets/d/" + id : "";
}

function nettoyerPourInterface(o) {
  var r = {};
  COLONNES.forEach(function (c) { r[c] = o[c]; });
  r.jours = joursDepuis(o.dateRelance || o.dateEnvoi);
  r.statutLibelle = STATUTS[o.statut] || o.statut;
  try { r.historique = o.historique ? JSON.parse(o.historique) : []; } catch (e) { r.historique = []; }
  return r;
}

function enregistrerCleIA(cle) {
  cle = String(cle || "").trim();
  if (!cle) throw new Error("Clé vide.");
  proprietes().setProperty(CLE_GEMINI, cle);
  proprietes().deleteProperty(CLE_MODELE);
  modeleIA();   // vérifie tout de suite que la clé fonctionne
  return true;
}

function majProspect(id, champs) {
  var o = trouverParId(id);
  if (!o) throw new Error("Entreprise introuvable.");
  Object.keys(champs || {}).forEach(function (k) {
    if (COLONNES.indexOf(k) >= 0 && k !== "id") o[k] = champs[k];
  });
  ecrireLigne(o);
  return nettoyerPourInterface(o);
}

function changerStatut(id, statut) {
  var o = trouverParId(id);
  if (!o) throw new Error("Entreprise introuvable.");
  o.statut = statut;
  if (statut === "envoye" && !o.dateEnvoi) o.dateEnvoi = aujourdhui();
  if (statut === "relance") o.dateRelance = aujourdhui();
  if (statut === "reponse") o.dateReponse = aujourdhui();
  ajouterHistorique(o, "Statut : " + (STATUTS[statut] || statut));
  ecrireLigne(o);
  return nettoyerPourInterface(o);
}

function supprimerProspect(id) {
  var o = trouverParId(id);
  if (!o) return true;
  feuille().deleteRow(o._ligne);
  return true;
}

/* ============================ IA (Gemini) ============================ */

function cleIA() {
  var cle = proprietes().getProperty(CLE_GEMINI);
  if (!cle) throw new Error("Aucune clé IA enregistrée. Ouvrez « Réglages » et collez votre clé Google AI Studio.");
  return cle;
}

/* Les noms de modèles changent au fil du temps : on demande la liste à
   Google et on garde le meilleur modèle rapide disponible. */
function modeleIA() {
  var props = proprietes();
  var garde = props.getProperty(CLE_MODELE);
  if (garde) return garde;

  var rep = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models?key=" + encodeURIComponent(cleIA()),
    { muteHttpExceptions: true });
  if (rep.getResponseCode() !== 200) {
    throw new Error("Clé IA refusée par Google (code " + rep.getResponseCode() + "). Vérifiez-la dans « Réglages ».");
  }
  var liste = (JSON.parse(rep.getContentText()).models || []).filter(function (m) {
    return (m.supportedGenerationMethods || []).indexOf("generateContent") >= 0;
  }).map(function (m) { return String(m.name).replace("models/", ""); })
    .filter(function (n) { return !/embedding|aqa|vision|tts|image|live|audio/.test(n); });

  function score(n) {
    var v = n.match(/gemini-(\d+)(?:\.(\d+))?/);
    var version = v ? parseInt(v[1], 10) * 10 + (v[2] ? parseInt(v[2], 10) : 0) : 0;
    var famille = /flash/.test(n) ? 3 : (/pro/.test(n) ? 2 : 1);
    var propre = /(latest|exp|preview|thinking)/.test(n) ? 0 : 1;
    return version * 100 + famille * 10 + propre;
  }
  liste.sort(function (a, b) { return score(b) - score(a); });

  if (!liste.length) throw new Error("Aucun modèle IA disponible avec cette clé.");
  props.setProperty(CLE_MODELE, liste[0]);
  return liste[0];
}

function appelerIA(prompt, avecRecherche) {
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" +
    modeleIA() + ":generateContent?key=" + encodeURIComponent(cleIA());

  var outils = avecRecherche ? [[{ google_search: {} }], [{ google_search_retrieval: {} }], null] : [null];
  var derniere = "";

  for (var i = 0; i < outils.length; i++) {
    var corps = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.35, maxOutputTokens: 8192 }
    };
    if (outils[i]) corps.tools = outils[i];

    var rep = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(corps),
      muteHttpExceptions: true
    });
    var code = rep.getResponseCode();
    var txt = rep.getContentText();

    if (code === 200) {
      var j = JSON.parse(txt);
      var cand = (j.candidates || [])[0] || {};
      var parts = ((cand.content || {}).parts || []);
      var texte = parts.map(function (p) { return p.text || ""; }).join("\n").trim();
      var sources = (((cand.groundingMetadata || {}).groundingChunks) || [])
        .map(function (c) { return c.web ? { titre: c.web.title, url: c.web.uri } : null; })
        .filter(Boolean);
      if (!texte) throw new Error("L'IA n'a rien renvoyé (réponse peut-être bloquée). Reformulez votre demande.");
      return { texte: texte, sources: sources, recherche: !!outils[i] };
    }
    derniere = code + " " + txt.slice(0, 300);
    /* 400 sur l'outil de recherche : on retente sans, plutôt que d'échouer. */
    if (code !== 400) break;
  }
  throw new Error("L'IA a refusé la requête (" + derniere + ").");
}

/* Les réponses arrivent parfois entourées de texte ou de balises de code. */
function extraireJson(texte) {
  var t = String(texte || "").replace(/```(?:json)?/gi, "").trim();
  var essais = [t];
  var d = t.indexOf("["), f = t.lastIndexOf("]");
  if (d >= 0 && f > d) essais.push(t.slice(d, f + 1));
  d = t.indexOf("{"); f = t.lastIndexOf("}");
  if (d >= 0 && f > d) essais.push(t.slice(d, f + 1));
  for (var i = 0; i < essais.length; i++) {
    try { return JSON.parse(essais[i]); } catch (e) { /* on essaie la forme suivante */ }
  }
  return null;
}

/* ============================ RECHERCHE DE PROSPECTS ============================ */

function chercherProspects(demande, dejaVus) {
  demande = String(demande || "").trim();
  if (!demande) throw new Error("Dites ce que vous cherchez.");

  var connus = (dejaVus || lireTout().map(function (o) { return o.entreprise; }))
    .filter(Boolean).slice(0, 200).join(", ");

  var prompt = [
    "Tu aides un exploitant de distributeurs automatiques à trouver des lieux d'implantation en France.",
    "",
    "CE QU'IL PROPOSE : " + OFFRE.resume,
    OFFRE.pourLeClient.map(function (l) { return "- " + l; }).join("\n"),
    "",
    "SA DEMANDE : " + demande,
    MOI.zone ? "SA ZONE HABITUELLE : " + MOI.zone : "",
    "",
    "Cherche sur le web des établissements RÉELS correspondant à cette demande.",
    "Privilégie ceux qui ont du personnel ou du public sur place et peu de commerces à proximité,",
    "et qui n'ont vraisemblablement pas encore de distributeur : petites et moyennes structures,",
    "zones d'activité isolées, lieux d'attente. Écarte les très grands sites qui en ont presque toujours,",
    "les administrations, et les entreprises sans local recevant du personnel.",
    "",
    connus ? "À EXCLURE (déjà dans sa liste) : " + connus : "",
    "",
    "Réponds UNIQUEMENT par un tableau JSON, sans texte autour, de 8 à 15 objets :",
    '[{"entreprise":"nom exact","ville":"commune","codePostal":"","activite":"activité en clair",',
    '"taille":"ordre de grandeur du personnel, ex. environ 20 à 30 personnes, ou vide si inconnu",',
    '"site":"https://… ou vide","pourquoi":"une phrase : pourquoi ce lieu est pertinent et pourquoi il n\'a probablement pas encore de distributeur",',
    '"certitude":"haute|moyenne|faible"}]',
    "",
    "RÈGLES ABSOLUES :",
    "- N'invente aucun établissement. Si tu n'es pas sûr qu'il existe, ne le mets pas.",
    "- Ne mets JAMAIS d'adresse email : elles seront vérifiées ailleurs.",
    "- N'invente pas de chiffre d'affaires, d'effectif précis ni de détail que tu ne sais pas.",
    "- \"certitude\" décrit ta confiance sur le fait que ce lieu n'a pas déjà un distributeur : c'est une estimation, pas un fait.",
    "- Tout en français."
  ].filter(function (l) { return l !== ""; }).join("\n");

  var rep = appelerIA(prompt, true);
  var liste = extraireJson(rep.texte);
  if (!liste || !liste.length) {
    throw new Error("L'IA n'a pas renvoyé de liste exploitable. Reformulez en précisant la ville et le type de lieu.");
  }
  if (!Array.isArray(liste)) liste = [liste];

  var connusMin = {};
  lireTout().forEach(function (o) { connusMin[minuscule(o.entreprise)] = true; });

  var propres = [];
  liste.forEach(function (c) {
    if (!c || !c.entreprise) return;
    var nom = String(c.entreprise).trim();
    if (connusMin[minuscule(nom)]) return;
    propres.push({
      entreprise: nom,
      ville: String(c.ville || "").trim(),
      codePostal: String(c.codePostal || "").trim(),
      activite: String(c.activite || "").trim(),
      taille: String(c.taille || "").trim(),
      site: normaliserUrl(c.site || ""),
      pourquoi: String(c.pourquoi || "").trim(),
      certitude: /haute|moyenne|faible/.test(c.certitude) ? c.certitude : "moyenne"
    });
  });

  propres = verifierAupresDeLEtat(propres);

  return {
    candidats: propres,
    sources: rep.sources,
    recherche: rep.recherche,
    avertissement: rep.recherche ? "" :
      "L'IA n'a pas pu consulter le web pour cette réponse : vérifiez chaque établissement avant d'écrire."
  };
}

function minuscule(s) {
  return String(s == null ? "" : s).toLowerCase().replace(/\s+/g, " ").trim();
}

function normaliserUrl(u) {
  u = String(u || "").trim();
  if (!u) return "";
  if (!/^https?:\/\//i.test(u)) u = "https://" + u;
  return u.replace(/\/+$/, "");
}

/* Une IA peut se tromper de nom : on confronte chaque candidat au registre
   public des entreprises, et on affiche le résultat de ce contrôle. */
function verifierAupresDeLEtat(candidats) {
  candidats.forEach(function (c) {
    try {
      var params = "?q=" + encodeURIComponent(c.entreprise) + "&etat_administratif=A&per_page=1";
      if (/^\d{5}$/.test(c.codePostal)) params += "&code_postal=" + c.codePostal;
      var rep = UrlFetchApp.fetch(API_ENTREPRISES + params, { muteHttpExceptions: true });
      if (rep.getResponseCode() !== 200) { c.verifie = null; return; }
      var r = (JSON.parse(rep.getContentText()).results || [])[0];
      if (!r) { c.verifie = false; return; }
      var siege = r.siege || {};
      c.verifie = true;
      c.siren = r.siren || "";
      c.nomOfficiel = r.nom_complet || "";
      c.effectif = libelleTranche(r.tranche_effectif_salarie);
      c.naf = r.activite_principale || "";
      if (!c.ville && siege.libelle_commune) c.ville = siege.libelle_commune;
      if (!c.codePostal && siege.code_postal) c.codePostal = siege.code_postal;
      c.adresse = siege.adresse || "";
    } catch (e) {
      c.verifie = null;
    }
  });
  return candidats;
}

function libelleTranche(code) {
  var t = {
    NN: "effectif non déclaré", "00": "pas de salarié", "01": "1 à 2 salariés", "02": "3 à 5 salariés",
    "03": "6 à 9 salariés", "11": "10 à 19 salariés", "12": "20 à 49 salariés", "21": "50 à 99 salariés",
    "22": "100 à 199 salariés", "31": "200 à 249 salariés", "32": "250 à 499 salariés",
    "41": "500 à 999 salariés", "42": "1 000 à 1 999 salariés", "51": "2 000 à 4 999 salariés",
    "52": "5 000 à 9 999 salariés", "53": "10 000 salariés et plus"
  };
  return t[String(code || "")] || "";
}

/* ============================ AJOUT À LA LISTE ============================ */

function ajouterCandidats(candidats) {
  var ajoutes = [];
  (candidats || []).forEach(function (c) {
    var o = { _ligne: 0 };
    COLONNES.forEach(function (k) { o[k] = ""; });
    o.id = nouvelId();
    o.entreprise = c.entreprise || "";
    o.ville = c.ville || "";
    o.activite = c.activite || "";
    o.type = typeDepuisActivite(c.activite + " " + (c.naf || ""));
    o.siren = c.siren || "";
    o.effectif = c.effectif || c.taille || "";
    o.site = c.site || "";
    o.statut = "a_envoyer";
    o.pourquoi = c.pourquoi || "";
    o.note = [c.adresse, c.naf ? "NAF " + c.naf : "", c.certitude ? "estimation : " + c.certitude : ""]
      .filter(Boolean).join(" · ");
    ajouterHistorique(o, "Proposée par l'assistant" + (c.verifie ? " et vérifiée au registre des entreprises" : ""));
    ecrireLigne(o);
    ajoutes.push(nettoyerPourInterface(o));
  });
  return ajoutes;
}

function typeDepuisActivite(txt) {
  var t = " " + minuscule(txt).replace(/[^a-zà-ÿ0-9]+/g, " ") + " ";
  var table = [
    ["Salle de sport", ["sport", "fitness", "musculation", "gym", "piscine", "padel", "tennis"]],
    ["Hôtel / résidence", ["hôtel", "hotel", "résidence", "camping", "auberge"]],
    ["Centre de formation", ["formation", "auto école", "auto-école", "cfa", "école", "lycée"]],
    ["Entrepôt / logistique", ["logistique", "entrepôt", "entrepot", "transport", "messagerie", "stockage"]],
    ["Usine / atelier", ["usine", "industrie", "atelier", "fabrication", "production", "métallerie",
      "chaudronnerie", "menuiserie", "imprimerie", "chantier", "bâtiment", "travaux", "mécanique"]],
    ["Établissement accueillant du public", ["laverie", "lavage", "pressing", "contrôle technique",
      "garage", "carrosserie", "concession", "clinique", "ehpad", "laboratoire", "magasin", "commerce"]],
    ["Gare / lieu de passage", ["gare", "station", "péage", "aire"]]
  ];
  for (var i = 0; i < table.length; i++) {
    for (var j = 0; j < table[i][1].length; j++) {
      if (t.indexOf(" " + table[i][1][j] + " ") >= 0) return table[i][0];
    }
  }
  return "Entreprise / bureaux";
}

/* ============================ RECHERCHE DE L'ADRESSE EMAIL ============================ */

var PREFIXES_UTILES = ["contact", "info", "infos", "commercial", "accueil", "direction",
  "secretariat", "secretariat", "administration", "bonjour", "hello", "agence", "siege"];

function recupererPage(url) {
  try {
    var rep = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: false,
      headers: { "User-Agent": "Mozilla/5.0 (compatible; CLTDistrib prospection)" }
    });
    if (rep.getResponseCode() !== 200) return "";
    var type = String(rep.getHeaders()["Content-Type"] || rep.getHeaders()["content-type"] || "");
    if (type && type.indexOf("text") < 0 && type.indexOf("html") < 0) return "";
    return rep.getContentText().slice(0, 400000);
  } catch (e) {
    return "";
  }
}

/* Lit les pages publiques du site de l'entreprise et en extrait l'adresse de
   contact. C'est exactement la recherche que vous faisiez à la main. */
function trouverEmail(id, siteFourni) {
  var o = id ? trouverParId(id) : null;
  var site = normaliserUrl(siteFourni || (o ? o.site : ""));
  if (!site) return { email: "", source: "", essais: [], message: "Renseignez d'abord le site internet." };

  var domaine = site.replace(/^https?:\/\//, "").replace(/^www\./, "").split("/")[0];
  var pages = [site, site + "/contact", site + "/contacts", site + "/nous-contacter",
    site + "/contact.html", site + "/contact.php", site + "/mentions-legales", site + "/fr/contact"];

  var trouves = {}, visitees = [];
  for (var i = 0; i < pages.length; i++) {
    if (Object.keys(trouves).length >= 8) break;
    var html = recupererPage(pages[i]);
    if (!html) continue;
    visitees.push(pages[i]);
    var brut = html.replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(parseInt(d, 10)); })
      .replace(/\s*(?:\[at\]|\(at\)|&#64;)\s*/gi, "@");
    (brut.match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g) || []).forEach(function (e) {
      e = e.toLowerCase();
      if (/\.(png|jpe?g|gif|svg|webp|css|js|ico)$/.test(e)) return;
      if (/(sentry|wixpress|squarespace|example\.|domain\.|votre|votresite|nom@|email@|adresse@|no-?reply|noreply|postmaster|abuse|webmaster@wordpress)/.test(e)) return;
      trouves[e] = (trouves[e] || 0) + 1;
    });
  }

  var liste = Object.keys(trouves);
  if (!liste.length) {
    return { email: "", source: "", essais: visitees,
      message: "Aucune adresse trouvée sur ce site. Elle est peut-être derrière un formulaire : le téléphone sera plus rapide." };
  }

  function note(e) {
    var dom = e.split("@")[1] || "";
    var prefixe = e.split("@")[0];
    var n = 0;
    if (dom === domaine || dom.indexOf(domaine) >= 0 || domaine.indexOf(dom) >= 0) n += 100;
    var rang = PREFIXES_UTILES.indexOf(prefixe);
    if (rang >= 0) n += 50 - rang;
    if (/^[a-z]+\.[a-z]+$/.test(prefixe)) n += 5;   // prenom.nom : souvent une vraie personne
    return n + trouves[e];
  }
  liste.sort(function (a, b) { return note(b) - note(a); });

  if (o) {
    o.email = liste[0];
    o.emailSource = "trouvée sur " + domaine;
    if (!o.site) o.site = site;
    ajouterHistorique(o, "Adresse trouvée sur le site : " + liste[0]);
    ecrireLigne(o);
  }
  return { email: liste[0], autres: liste.slice(1, 5), source: domaine, essais: visitees, message: "" };
}

/* ============================ RÉDACTION DU MAIL ============================ */

function genererMail(id, type) {
  var o = trouverParId(id);
  if (!o) throw new Error("Entreprise introuvable.");
  type = type === "relance" ? "relance" : "premier";

  var contexte = [
    "Établissement : " + o.entreprise,
    o.ville ? "Commune : " + o.ville : "",
    o.activite ? "Activité : " + o.activite : "",
    o.type ? "Type de lieu : " + o.type : "",
    o.effectif ? "Taille : " + o.effectif : "",
    o.contactNom ? "Nom du contact : " + o.contactNom : "",
    o.contactFonction ? "Fonction du contact : " + o.contactFonction : "",
    o.pourquoi ? "Pourquoi ce lieu est pertinent : " + o.pourquoi : ""
  ].filter(Boolean).join("\n");

  var prompt = [
    "Tu écris un email de prospection commerciale en français, pour " + MOI.entreprise + ".",
    "",
    "L'OFFRE : " + OFFRE.resume,
    OFFRE.pourLeClient.map(function (l) { return "- " + l; }).join("\n"),
    MOI.zone ? "Zone d'intervention : " + MOI.zone : "",
    "",
    "LE DESTINATAIRE :",
    contexte,
    "",
    type === "relance"
      ? "C'est une RELANCE : un premier message est resté sans réponse il y a environ " +
        (joursDepuis(o.dateEnvoi) || REGLAGES.joursAvantRelance) + " jours. Reste courtois, bref, " +
        "propose une sortie honorable (« si le sujet n'est pas d'actualité, dites-le moi »)."
      : "C'est un PREMIER contact : il ne vous connaît pas.",
    "",
    "RÈGLES :",
    "- Vouvoiement. Ton sobre et professionnel, jamais commercial agressif, aucun superlatif.",
    "- 120 à 170 mots pour le corps, phrases courtes, aucune liste à puces.",
    "- Commence par une observation juste et concrète sur SON type de lieu (ce que vivent ses équipes ou ses visiteurs), pas par une présentation de l'entreprise.",
    "- Explique en une phrase que l'installation, le réassort et l'entretien sont pris en charge, et que de son côté il n'y a qu'un emplacement et une prise électrique.",
    "- " + (o.contactFonction ? "Adapte un argument à sa fonction (" + o.contactFonction + ")." : "N'invente pas la fonction du destinataire."),
    "- Termine par une demande simple : une visite gratuite de l'emplacement, sans engagement.",
    "- " + OFFRE.interdits,
    "- Pas de signature, pas de « Cordialement » : ils sont ajoutés automatiquement.",
    "- Objet : court, factuel, sans majuscules inutiles, sans mot publicitaire, sans point d'exclamation.",
    "",
    'Réponds UNIQUEMENT par du JSON : {"objet":"…","corps":"…"} avec de vrais retours à la ligne \\n dans le corps.'
  ].filter(function (l) { return l !== ""; }).join("\n");

  var rep = appelerIA(prompt, false);
  var j = extraireJson(rep.texte);
  if (!j || !j.objet || !j.corps) throw new Error("L'IA n'a pas renvoyé de mail exploitable. Réessayez.");

  o.objet = String(j.objet).trim();
  o.corps = String(j.corps).replace(/\r/g, "").trim();
  ecrireLigne(o);
  return { objet: o.objet, corps: o.corps, type: type };
}

/* ============================ ENVOI ============================ */

function signatureTexte() {
  return [
    "Bien cordialement,",
    "",
    MOI.nom + (MOI.fonction ? " — " + MOI.fonction : ""),
    MOI.entreprise,
    MOI.tel ? "Tél. " + MOI.tel : "",
    MOI.email,
    MOI.site
  ].filter(Boolean).join("\n");
}

function signatureHtml() {
  var l = [];
  l.push('<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;font-family:Arial,Helvetica,sans-serif;margin-top:22px">');
  l.push('<tr><td style="padding:0 0 10px 0">');
  l.push('<img src="' + MOI.logo + '" width="190" alt="' + MOI.entreprise + '" style="display:block;border:0">');
  l.push('</td></tr><tr><td style="border-top:2px solid #0060FC;padding-top:10px;font-size:13px;line-height:1.55;color:#333">');
  l.push('<strong style="color:#050B1A;font-size:14px">' + MOI.nom + '</strong><br>');
  if (MOI.fonction) l.push('<span style="color:#5A6478">' + MOI.fonction + " — " + MOI.entreprise + "</span><br>");
  if (MOI.tel) l.push('<a href="tel:' + MOI.tel.replace(/[^+\d]/g, "") + '" style="color:#333;text-decoration:none">' + MOI.tel + "</a><br>");
  l.push('<a href="mailto:' + MOI.email + '" style="color:#0060FC;text-decoration:none">' + MOI.email + "</a><br>");
  l.push('<a href="' + MOI.site + '" style="color:#0060FC;text-decoration:none">' + MOI.site.replace(/^https?:\/\//, "") + "</a>");
  if (MOI.zone) l.push('<br><span style="color:#8A93A5;font-size:12px">' + MOI.zone + "</span>");
  l.push('<br><span style="color:#A0A8B8;font-size:11px;padding-top:6px;display:inline-block">' +
    "Vous ne souhaitez plus être contacté ? Répondez « STOP », je n'écrirai plus.</span>");
  l.push("</td></tr></table>");
  return l.join("");
}

function texteEnHtml(txt) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#222">' +
    String(txt || "").split(/\n{2,}/).map(function (p) {
      return "<p style=\"margin:0 0 14px 0\">" + echapper(p).replace(/\n/g, "<br>") + "</p>";
    }).join("") + "</div>";
}

function echapper(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function compteurDuJour() {
  var props = proprietes();
  var cle = "ENVOIS_" + aujourdhui();
  return { cle: cle, n: parseInt(props.getProperty(cle) || "0", 10) };
}

function envoyerMail(id, objet, corps) {
  var o = trouverParId(id);
  if (!o) throw new Error("Entreprise introuvable.");
  if (!o.email) throw new Error("Aucune adresse email pour " + o.entreprise + ".");

  var c = compteurDuJour();
  if (c.n >= REGLAGES.maxEnvoisParJour) {
    throw new Error("Limite du jour atteinte (" + REGLAGES.maxEnvoisParJour + " envois). Reprenez demain.");
  }
  var restant = MailApp.getRemainingDailyQuota();
  if (restant < 1) throw new Error("Gmail n'autorise plus d'envoi aujourd'hui. Reprenez demain.");

  objet = String(objet || o.objet || "").trim();
  corps = String(corps || o.corps || "").trim();
  if (!objet || !corps) throw new Error("Le mail est vide : générez-le d'abord.");

  GmailApp.sendEmail(o.email, objet, corps + "\n\n" + signatureTexte(), {
    name: MOI.nom,
    replyTo: MOI.email,
    htmlBody: texteEnHtml(corps) + signatureHtml()
  });
  proprietes().setProperty(c.cle, String(c.n + 1));

  var relance = (o.statut === "a_relancer" || o.statut === "envoye" || o.statut === "relance");
  o.objet = objet;
  o.corps = corps;
  if (relance) {
    o.statut = "relance";
    o.dateRelance = aujourdhui();
    ajouterHistorique(o, "Relance envoyée à " + o.email);
  } else {
    o.statut = "envoye";
    o.dateEnvoi = aujourdhui();
    ajouterHistorique(o, "Mail envoyé à " + o.email);
  }
  ecrireLigne(o);
  return nettoyerPourInterface(o);
}

/* ============================ SUIVI AUTOMATIQUE ============================ */

function tourneeQuotidienne() {
  var resultat = { reponses: [], bascules: [], relances: [] };
  var tous = lireTout();

  /* 1. les réponses reçues passent avant tout : on ne relance pas quelqu'un
        qui a déjà répondu. */
  tous.forEach(function (o) {
    if (["envoye", "a_relancer", "relance"].indexOf(o.statut) < 0 || !o.email) return;
    if (aRepondu(o)) {
      o.statut = "reponse";
      o.dateReponse = aujourdhui();
      ajouterHistorique(o, "Réponse repérée dans Gmail");
      ecrireLigne(o);
      resultat.reponses.push(o.entreprise);
    }
  });

  /* 2. bascule en « à relancer » au bout du délai */
  lireTout().forEach(function (o) {
    if (o.statut !== "envoye") return;
    var j = joursDepuis(o.dateEnvoi);
    if (j === null || j < REGLAGES.joursAvantRelance) return;
    o.statut = "a_relancer";
    ajouterHistorique(o, "Sans réponse depuis " + j + " jours : à relancer");
    ecrireLigne(o);
    resultat.bascules.push(o.entreprise);
  });

  /* 3. relance automatique, seulement si vous l'avez demandée */
  if (REGLAGES.relanceAuto) {
    lireTout().forEach(function (o) {
      if (o.statut !== "a_relancer" || !o.email) return;
      if (compteurDuJour().n >= REGLAGES.maxEnvoisParJour) return;
      try {
        genererMail(o.id, "relance");
        envoyerMail(o.id, null, null);
        resultat.relances.push(o.entreprise);
      } catch (e) {
        ajouterHistorique(o, "Relance automatique impossible : " + e.message);
        ecrireLigne(o);
      }
    });
  }
  proprietes().setProperty("DERNIERE_TOURNEE", new Date().toISOString());
  return resultat;
}

function aRepondu(o) {
  try {
    var depuis = o.dateRelance || o.dateEnvoi;
    var fils = GmailApp.search('from:' + o.email + ' newer_than:90d', 0, 5);
    if (!fils.length) {
      var domaine = String(o.email).split("@")[1];
      if (!domaine || /gmail|orange|wanadoo|free|hotmail|outlook|yahoo|laposte|sfr/.test(domaine)) return false;
      fils = GmailApp.search('from:@' + domaine + ' newer_than:90d', 0, 5);
      if (!fils.length) return false;
    }
    var limite = depuis ? new Date(depuis + "T00:00:00").getTime() : 0;
    for (var i = 0; i < fils.length; i++) {
      var msgs = fils[i].getMessages();
      for (var k = 0; k < msgs.length; k++) {
        var de = String(msgs[k].getFrom()).toLowerCase();
        if (de.indexOf(String(MOI.email).toLowerCase()) >= 0) continue;
        if (msgs[k].getDate().getTime() >= limite) return true;
      }
    }
  } catch (e) { /* Gmail indisponible : on ne conclut rien */ }
  return false;
}

function installerDeclencheur() {
  retirerDeclencheur();
  ScriptApp.newTrigger("tourneeQuotidienne").timeBased()
    .everyDays(1).atHour(REGLAGES.heureTournee).create();
  return "Contrôle quotidien installé vers " + REGLAGES.heureTournee + " h.";
}

function retirerDeclencheur() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "tourneeQuotidienne") ScriptApp.deleteTrigger(t);
  });
  return "Contrôle quotidien retiré.";
}

function declencheurInstalle() {
  return ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === "tourneeQuotidienne";
  });
}

function infosTechniques() {
  return {
    declencheur: declencheurInstalle(),
    derniereTournee: proprietes().getProperty("DERNIERE_TOURNEE") || "",
    envoisAujourdhui: compteurDuJour().n,
    quotaGmail: MailApp.getRemainingDailyQuota(),
    modele: proprietes().getProperty(CLE_MODELE) || "",
    feuille: urlFeuille()
  };
}

/* ============================ TRAITEMENT PAR LOT ============================ */

/* Trouve le site officiel quand il manque, pour pouvoir y chercher l'adresse. */
function trouverSite(id) {
  var o = trouverParId(id);
  if (!o) throw new Error("Entreprise introuvable.");
  if (o.site) return o.site;

  var prompt = [
    "Donne l'adresse du site internet officiel de cet établissement français :",
    o.entreprise + (o.ville ? ", à " + o.ville : "") + (o.activite ? " (" + o.activite + ")" : ""),
    "",
    "Réponds uniquement par l'URL, rien d'autre. Si tu ne la connais pas avec certitude,",
    "réponds exactement : AUCUN"
  ].join("\n");

  var rep = appelerIA(prompt, true);
  var t = String(rep.texte).trim();
  var m = t.match(/https?:\/\/[^\s"'<>)]+/);
  var url = m ? normaliserUrl(m[0]) : "";
  if (!url || /aucun/i.test(t)) return "";

  o.site = url;
  ajouterHistorique(o, "Site trouvé : " + url);
  ecrireLigne(o);
  return url;
}

/* Pour chaque entreprise : adresse email si elle manque, puis rédaction du
   mail. Traité par petits paquets pour rester dans le temps imparti. */
function preparerLot(ids) {
  var sortie = [];
  (ids || []).slice(0, 6).forEach(function (id) {
    var o = trouverParId(id);
    if (!o) return;
    var r = { id: id, entreprise: o.entreprise, email: o.email, pret: false, message: "" };
    try {
      if (!o.email) {
        if (!o.site) { try { trouverSite(id); } catch (e) { /* sans site, on continue */ } }
        o = trouverParId(id);
        if (o.site) {
          var t = trouverEmail(id, null);
          if (!t.email) r.message = t.message;
        }
        o = trouverParId(id);
      }
      r.email = o.email;
      if (!o.email) {
        r.message = r.message || "Aucune adresse trouvable" + (o.tel ? " : appelez le " + o.tel + "." : ".");
        sortie.push(r);
        return;
      }
      genererMail(id, o.statut === "a_relancer" || o.statut === "relance" ? "relance" : "premier");
      r.pret = true;
    } catch (e) {
      r.message = e.message;
    }
    sortie.push(r);
  });
  return sortie;
}

function envoyerLot(ids) {
  var sortie = [];
  (ids || []).slice(0, 12).forEach(function (id) {
    var o = trouverParId(id);
    if (!o) return;
    try {
      if (!o.objet || !o.corps) genererMail(id, o.statut === "a_relancer" ? "relance" : "premier");
      envoyerMail(id, null, null);
      sortie.push({ id: id, entreprise: o.entreprise, envoye: true, message: "" });
    } catch (e) {
      sortie.push({ id: id, entreprise: o.entreprise, envoye: false, message: e.message });
    }
  });
  return sortie;
}
