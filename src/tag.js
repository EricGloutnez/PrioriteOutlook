/*
 * tag.js — Logique de la balise de priorité dans l'objet du courriel.
 *
 * Format (pastille de couleur + balise au début de l'objet, suivies d'un espace) :
 *   🟠 [P1] Objet                    niveau seul (aussi URGENT, P2, P3)
 *   🟠 [M_12345_P1] Objet            type (M = Multi, R = Résidentiel), numéro, niveau
 *   🟣 [CONFIDENTIEL] Objet          confidentiel seul
 *   🟣 [CONFIDENTIEL_R_12345] Objet  confidentiel en avant de tout, puis le projet
 *
 * La pastille est facultative à la lecture : une balise tapée à la main
 * sans pastille (sur le téléphone, par exemple) est reconnue.
 *
 * Utilisé par le panneau (taskpane.js). Le fichier launchevent.js contient
 * une copie de la même expression régulière, parce que l'Outlook classique
 * exécute les événements dans un environnement JavaScript seul, sans page
 * HTML pour charger d'autres fichiers. Garder les deux synchronisées.
 */
(function (global) {
  "use strict";

  var NIVEAUX = ["URGENT", "P1", "P2", "P3", "CONFIDENTIEL"];
  var TYPES = { M: "Multi", R: "Résidentiel" };
  var PASTILLES = {
    URGENT: "\uD83D\uDD34",       // 🔴
    P1: "\uD83D\uDFE0",           // 🟠
    P2: "\uD83D\uDFE1",           // 🟡
    P3: "\uD83D\uDFE2",           // 🟢
    CONFIDENTIEL: "\uD83D\uDFE3"  // 🟣
  };
  // Une pastille (n'importe laquelle des cinq) suivie d'espaces, juste avant la balise.
  var REGEX_PASTILLE_FIN = /(?:\uD83D\uDD34|\uD83D\uDFE0|\uD83D\uDFE1|\uD83D\uDFE2|\uD83D\uDFE3)\s*$/;
  // Une pastille isolée au point d'insertion (à retirer avant d'ajouter la nouvelle).
  var REGEX_PASTILLE_DEBUT = /^(?:\uD83D\uDD34|\uD83D\uDFE0|\uD83D\uDFE1|\uD83D\uDFE2|\uD83D\uDFE3)\s*/;

  // Une seule balise par objet ; insensible à la casse pour tolérer
  // une saisie manuelle comme [p1] sur le téléphone.
  var REGEX_BALISE = /\[(?:(CONFIDENTIEL)(?:_([MR])_(\d+))?|(?:([MR])_(\d+)_)?(URGENT|P1|P2|P3))\]/i;

  // Préfixes de réponse et de transfert placés avant la balise.
  var REGEX_PREFIXES = /^(?:\s*(?:RE|TR|FW|FWD|RÉF|REF)\s*:\s*)+/i;

  /**
   * Lit la balise présente dans l'objet.
   * texte/index : la balise entre crochets seule ;
   * debut : position de départ en incluant la pastille qui la précède, s'il y en a une.
   * @returns {{niveau:string, type:string|null, numero:string|null,
   *            texte:string, index:number, debut:number} | null}
   */
  function lire(objet) {
    var m = REGEX_BALISE.exec(objet || "");
    if (!m) return null;
    var confidentiel = !!m[1];
    var type = confidentiel ? m[2] : m[4];
    var avant = objet.slice(0, m.index);
    var pastille = REGEX_PASTILLE_FIN.exec(avant);
    return {
      niveau: confidentiel ? "CONFIDENTIEL" : m[6].toUpperCase(),
      type: type ? type.toUpperCase() : null,
      numero: (confidentiel ? m[3] : m[5]) || null,
      texte: m[0],
      index: m.index,
      debut: pastille ? pastille.index : m.index
    };
  }

  /** Construit la balise, ex. [P1], [M_12345_P1] ou [CONFIDENTIEL_M_12345]. */
  function construire(niveau, type, numero) {
    if (NIVEAUX.indexOf(niveau) === -1) {
      throw new Error("Niveau inconnu : " + niveau);
    }
    if (numero) {
      if (!/^\d+$/.test(numero)) throw new Error("Le numéro doit contenir des chiffres seulement.");
      if (!TYPES[type]) throw new Error("Choisir Multi ou Résidentiel.");
      return niveau === "CONFIDENTIEL"
        ? "[CONFIDENTIEL_" + type + "_" + numero + "]"
        : "[" + type + "_" + numero + "_" + niveau + "]";
    }
    return "[" + niveau + "]";
  }

  /** Marque complète à écrire dans l'objet : pastille + espace + balise. */
  function marque(niveau, type, numero) {
    return PASTILLES[niveau] + " " + construire(niveau, type, numero);
  }

  /**
   * Applique la marque (pastille + balise) : remplace une marque existante à sa position
   * (même après RE:, TR:, FW:), sinon l'ajoute au début de l'objet,
   * après les préfixes de réponse ou de transfert.
   */
  function appliquer(objet, balise) {
    objet = objet || "";
    var existante = lire(objet);
    if (existante) {
      return objet.slice(0, existante.debut) + balise + objet.slice(existante.index + existante.texte.length);
    }
    var prefixes = (REGEX_PREFIXES.exec(objet) || [""])[0];
    var reste = objet.slice(prefixes.length).replace(/^\s+/, "").replace(REGEX_PASTILLE_DEBUT, "");
    var debut = prefixes ? prefixes.replace(/\s*$/, " ") : "";
    return debut + balise + (reste ? " " + reste : "");
  }

  /** Retire la pastille, la balise et l'espace qui suit, s'il y a lieu. */
  function retirer(objet) {
    objet = objet || "";
    var existante = lire(objet);
    if (!existante) return objet;
    var avant = objet.slice(0, existante.debut);
    var apres = objet.slice(existante.index + existante.texte.length);
    if (apres.charAt(0) === " ") apres = apres.slice(1);
    return avant + apres;
  }

  global.PrioriteTag = {
    NIVEAUX: NIVEAUX,
    TYPES: TYPES,
    PASTILLES: PASTILLES,
    REGEX_BALISE: REGEX_BALISE,
    lire: lire,
    construire: construire,
    marque: marque,
    appliquer: appliquer,
    retirer: retirer
  };
})(typeof window !== "undefined" ? window : this);
